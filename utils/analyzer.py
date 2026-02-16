import os
import re
import io
import tempfile
import logging
import platform
import shutil
from datetime import datetime
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches
import traceback
from PIL import Image, ImageEnhance, ImageOps

logging.basicConfig(level=logging.WARNING, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------
# Tesseract detection (cross-platform)
# ---------------------------
TESSERACT_AVAILABLE = False
OCR_LANGUAGES = "rus+eng"

def _try_set_tessdata_prefix():
    """
    На Streamlit Cloud иногда tesseract есть, но tessdata не находится.
    Поставим TESSDATA_PREFIX если найдём типичные пути.
    """
    candidates = [
        "/usr/share/tesseract-ocr/5/tessdata",
        "/usr/share/tesseract-ocr/4.00/tessdata",
        "/usr/share/tesseract-ocr/tessdata",
        "/usr/share/tessdata",
    ]
    for p in candidates:
        if os.path.isdir(p) and os.path.exists(os.path.join(p, "eng.traineddata")):
            os.environ["TESSDATA_PREFIX"] = p
            return p
    return None

try:
    import pytesseract

    if platform.system().lower() == "windows":
        win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(win_path):
            pytesseract.pytesseract.tesseract_cmd = win_path
            TESSERACT_AVAILABLE = True
        else:
            tpath = shutil.which("tesseract")
            if tpath:
                pytesseract.pytesseract.tesseract_cmd = tpath
                TESSERACT_AVAILABLE = True
    else:
        tpath = shutil.which("tesseract")
        if tpath:
            pytesseract.pytesseract.tesseract_cmd = tpath
            _try_set_tessdata_prefix()
            TESSERACT_AVAILABLE = True

    if TESSERACT_AVAILABLE:
        try:
            langs = pytesseract.get_languages(config="")
            if "rus" in langs and "eng" in langs:
                OCR_LANGUAGES = "rus+eng"
            elif "rus" in langs:
                OCR_LANGUAGES = "rus"
            else:
                OCR_LANGUAGES = "eng"
        except Exception:
            OCR_LANGUAGES = "rus+eng"

except Exception:
    TESSERACT_AVAILABLE = False
    OCR_LANGUAGES = "rus+eng"


class PresentationAnalyzer:
    def __init__(self, pptx_path: str, enable_ocr: bool = True):
        self.pptx_path = pptx_path
        self.enable_ocr = bool(enable_ocr)

        self.results = []
        self.used_fonts = set()
        self.ocr_languages = OCR_LANGUAGES
        self.analysis_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.selected_slides_range = "all"

        self.settings = {
            "min_text_length_for_ocr": 3,
            "max_text_chars": 1000,
            "min_image_area_percentage": 0.7,
            "ocr_alternate_min_confidence": 35,
            "max_ocr_text_length": 5000,
            "ocr_max_images_per_slide": 6,  # ограничение для скорости
        }

    # ---------------------------
    # Main
    # ---------------------------
    def analyze_selected_slides(self, slides_range="all"):
        try:
            self.selected_slides_range = slides_range
            prs = Presentation(self.pptx_path)
            total_slides = len(prs.slides)

            slides_to_analyze = self.parse_slides_range(slides_range, total_slides)
            if not slides_to_analyze:
                return [], {}

            stats = {
                "has_animations": False,
                "has_transitions": False,
                "fonts_count": 0,
                "background_issues": 0,
                "text_on_images": 0,
                "total_images": 0,
                "ocr_used": False,
                "ocr_text_found": 0,
                "total_ocr_characters": 0,
                "selected_slides_count": len(slides_to_analyze),
                "selected_slides_range": slides_range,
                "total_slides_in_presentation": total_slides,
                "tesseract_available": bool(TESSERACT_AVAILABLE),
                "ocr_enabled": bool(self.enable_ocr),
            }

            stats["has_transitions"] = self.check_presentation_transitions(prs)

            for slide_num in slides_to_analyze:
                slide = prs.slides[slide_num - 1]
                r = self.analyze_slide(slide, slide_num)
                self.results.append(r)

                if r["Анимации"] == "✗":
                    stats["has_animations"] = True
                if r["Фон"] == "✗":
                    stats["background_issues"] += 1
                if r["Текст_на_изобр"] == "Да":
                    stats["text_on_images"] += 1
                    stats["ocr_text_found"] += 1
                if r["Изображения"] > 0:
                    stats["total_images"] += r["Изображения"]
                if r.get("OCR_текст"):
                    stats["total_ocr_characters"] += len(r["OCR_текст"])
                    stats["ocr_used"] = True

            self.analyze_fonts()
            stats["fonts_count"] = len(self.used_fonts)

            return self.results, stats

        except Exception:
            logger.exception("Ошибка анализа")
            return [], {}

    # ---------------------------
    # Conformance (твоя логика)
    # ---------------------------
    def calculate_conformance_percentage(self, results, presentation_stats):
        try:
            total_slides = len(results)
            weights = {
                "background": 15,
                "fonts": 15,
                "text_overload": 10,
                "text_on_images": 15,
                "animations": 15,
                "transitions": 10,
                "slide_compliance": 20,
            }
            total_possible = sum(weights.values())
            achieved_score = 0

            bg_issues = presentation_stats.get("background_issues", 0)
            bg_score = ((total_slides - bg_issues) / total_slides * weights["background"]) if total_slides else weights["background"]
            achieved_score += bg_score

            fonts_count = presentation_stats.get("fonts_count", 0)
            if fonts_count <= 2:
                fonts_score = weights["fonts"]
            elif fonts_count <= 3:
                fonts_score = weights["fonts"] * 0.5
            else:
                fonts_score = 0
            achieved_score += fonts_score

            text_issues = sum(1 for r in results if r["Текст"] == "✗")
            text_score = ((total_slides - text_issues) / total_slides * weights["text_overload"]) if total_slides else weights["text_overload"]
            achieved_score += text_score

            text_on_images = presentation_stats.get("text_on_images", 0)
            images_score = ((total_slides - text_on_images) / total_slides * weights["text_on_images"]) if total_slides else weights["text_on_images"]
            achieved_score += images_score

            anim_issues = sum(1 for r in results if r["Анимации"] == "✗")
            anim_score = ((total_slides - anim_issues) / total_slides * weights["animations"]) if total_slides else weights["animations"]
            achieved_score += anim_score

            transition_issues = 1 if presentation_stats.get("has_transitions") else 0
            transition_score = weights["transitions"] if transition_issues == 0 else 0
            achieved_score += transition_score

            compliant_slides = 0
            for r in results:
                if (
                    r["Фон"] == "✓" and
                    r["Шрифты"] == "✓" and
                    r["Текст"] == "✓" and
                    r["Текст_на_изобр"] == "Нет" and
                    r["Анимации"] == "✓"
                ):
                    compliant_slides += 1

            slide_score = ((compliant_slides / total_slides) * weights["slide_compliance"]) if total_slides else weights["slide_compliance"]
            achieved_score += slide_score

            percentage = round((achieved_score / total_possible) * 100, 1)

            if percentage >= 90:
                readiness_level, readiness_color, readiness_emoji = "отлично", "#27ae60", "🎉"
            elif percentage >= 75:
                readiness_level, readiness_color, readiness_emoji = "хорошо", "#2ecc71", "👍"
            elif percentage >= 60:
                readiness_level, readiness_color, readiness_emoji = "удовлетворительно", "#f39c12", "⚠️"
            elif percentage >= 40:
                readiness_level, readiness_color, readiness_emoji = "требует доработки", "#e74c3c", "🔧"
            else:
                readiness_level, readiness_color, readiness_emoji = "критически низкая", "#c0392b", "🚨"

            can_send = percentage >= 57

            recommendations = []
            if percentage < 57:
                recommendations.append("Рекомендуется доработать презентацию перед отправкой дизайнерам")
            if bg_issues > 0:
                recommendations.append(f"Исправьте фон на {bg_issues} слайдах")
            if fonts_count > 2:
                recommendations.append(f"Уменьшите количество шрифтов с {fonts_count} до 2")
            if text_issues > 0:
                recommendations.append(f"Уменьшите текст на {text_issues} слайдах")
            if text_on_images > 0:
                recommendations.append(f"Уберите текст с изображений на {text_on_images} слайдах")
            if anim_issues > 0:
                recommendations.append(f"Удалите анимации с {anim_issues} слайдов")
            if transition_issues > 0:
                recommendations.append("Удалите переходы между слайдами")

            user_message = (
                f"🎉 Ваша презентация соответствует критериям на {percentage}%. Презентация готова для отправки дизайнерам!"
                if can_send else
                f"⚠️ Ваша презентация соответствует критериям на {percentage}%. Если Вы планируете отправлять дизайнерам, рекомендуется её доработать."
            )

            return {
                "percentage": percentage,
                "readiness_level": readiness_level,
                "readiness_color": readiness_color,
                "readiness_emoji": readiness_emoji,
                "can_send_to_designers": can_send,
                "criteria_details": {
                    "background": {"score": round(bg_score, 1), "max": weights["background"], "issues": bg_issues},
                    "fonts": {"score": round(fonts_score, 1), "max": weights["fonts"], "fonts_count": fonts_count},
                    "text_overload": {"score": round(text_score, 1), "max": weights["text_overload"], "issues": text_issues},
                    "text_on_images": {"score": round(images_score, 1), "max": weights["text_on_images"], "issues": text_on_images},
                    "animations": {"score": round(anim_score, 1), "max": weights["animations"], "issues": anim_issues},
                    "transitions": {"score": round(transition_score, 1), "max": weights["transitions"], "has_issues": transition_issues > 0},
                    "slide_compliance": {"score": round(slide_score, 1), "max": weights["slide_compliance"], "compliant": compliant_slides, "total": total_slides},
                },
                "recommendations": recommendations,
                "user_message": user_message,
                "total_possible_score": total_possible,
                "achieved_score": round(achieved_score, 1),
                "compliant_slides": compliant_slides,
                "total_slides": total_slides,
            }
        except Exception:
            return None

    # ---------------------------
    # Slide parsing
    # ---------------------------
    def parse_slides_range(self, slides_range, total_slides):
        slides_to_analyze = []
        try:
            if not slides_range or str(slides_range).lower() == "all":
                return list(range(1, total_slides + 1))

            slides_range = str(slides_range).strip()
            if slides_range.isdigit():
                n = int(slides_range)
                return [n] if 1 <= n <= total_slides else []

            slides_range = slides_range.replace(" ", "")

            if "," in slides_range:
                for part in slides_range.split(","):
                    if "-" in part:
                        a, b = part.split("-", 1)
                        if a.isdigit() and b.isdigit():
                            start, end = int(a), int(b)
                            slides_to_analyze.extend(range(start, min(end, total_slides) + 1))
                    elif part.isdigit():
                        n = int(part)
                        if 1 <= n <= total_slides:
                            slides_to_analyze.append(n)
            elif "-" in slides_range:
                a, b = slides_range.split("-", 1)
                if a.isdigit() and b.isdigit():
                    start, end = int(a), int(b)
                    slides_to_analyze = list(range(start, min(end, total_slides) + 1))

            slides_to_analyze = sorted(set(slides_to_analyze))
            if not slides_to_analyze:
                slides_to_analyze = list(range(1, total_slides + 1))
                self.selected_slides_range = "all"

        except Exception:
            slides_to_analyze = list(range(1, total_slides + 1))
            self.selected_slides_range = "all"

        return slides_to_analyze

    # ---------------------------
    # Slide analysis
    # ---------------------------
    def analyze_slide(self, slide, slide_num):
        r = {
            "Слайд": slide_num,
            "Статус": "OK",
            "Нарушения": [],
            "Шрифты": "✓",
            "Текст": "✓",
            "Анимации": "✓",
            "Переходы": "✓",
            "Фон": "✓",
            "Изображения": 0,
            "Текст_на_изобр": "Нет",
            "Текст_дет": "",
            "Элементы": len(slide.shapes),
            "OCR_текст": "",
            "OCR_уверенность": 0,
            "OCR_метод": "",
            "OCR_изображений_с_текстом": 0,
        }

        if not self.check_background_comprehensive(slide):
            r["Фон"] = "✗"
            r["Нарушения"].append("ФОН")

        overload, char_count = self.check_text_improved(slide)
        r["Текст_дет"] = f"{char_count} симв."
        if overload:
            r["Текст"] = "✗"
            r["Нарушения"].append(f"ТЕКСТ({char_count})")

        if self.check_animations_improved(slide):
            r["Анимации"] = "✗"
            r["Нарушения"].append("АНИМАЦИИ")

        has_text_on_images, image_count, ocr_data = self.check_images_enhanced(slide)
        r["Изображения"] = image_count

        if ocr_data:
            r["OCR_текст"] = (ocr_data.get("text") or "")[: self.settings["max_ocr_text_length"]]
            r["OCR_уверенность"] = ocr_data.get("confidence", 0)
            r["OCR_метод"] = ocr_data.get("method", "")
            r["OCR_изображений_с_текстом"] = ocr_data.get("images_with_text", 0)

        if has_text_on_images:
            r["Текст_на_изобр"] = "Да"
            r["Нарушения"].append("ТЕКСТ_НА_ИЗОБР")

        self.collect_fonts(slide)

        if r["Нарушения"]:
            r["Статус"] = ", ".join(r["Нарушения"])

        return r

    def check_presentation_transitions(self, prs):
        try:
            for slide in prs.slides:
                slide_xml = str(slide.element.xml).lower()
                if "p:transition" in slide_xml or "transition" in slide_xml:
                    return True
        except Exception:
            pass
        return False

    def check_background_comprehensive(self, slide):
        try:
            if slide.background:
                fill = slide.background.fill
                if fill.type == 1:
                    if hasattr(fill.fore_color, "rgb"):
                        color = fill.fore_color.rgb
                        if hasattr(color, "r"):
                            if not (color.r == 255 and color.g == 255 and color.b == 255):
                                return False
                        elif color != RGBColor(255, 255, 255):
                            return False
                elif fill.type != 0:
                    return False

            try:
                slide_width = slide.width if hasattr(slide, "width") else Inches(10)
                slide_height = slide.height if hasattr(slide, "height") else Inches(7.5)
                slide_area = slide_width * slide_height

                for shape in slide.shapes:
                    try:
                        shape_area = shape.width * shape.height
                        if shape_area > slide_area * self.settings["min_image_area_percentage"]:
                            if hasattr(shape, "fill"):
                                fill = shape.fill
                                if fill.type == 1 and hasattr(fill.fore_color, "rgb"):
                                    color = fill.fore_color.rgb
                                    if hasattr(color, "r"):
                                        if not (color.r == 255 and color.g == 255 and color.b == 255):
                                            return False
                                    elif color != RGBColor(255, 255, 255):
                                        return False
                    except Exception:
                        continue
            except Exception:
                pass

            try:
                slide_xml = str(slide.element.xml).lower()
                for hex_color in re.findall(r"#[0-9a-f]{6}", slide_xml):
                    if hex_color not in ("#ffffff", "#ffffff00"):
                        return False
            except Exception:
                pass

            return True
        except Exception:
            return False

    def check_text_improved(self, slide):
        try:
            total_chars = 0
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame and shape.text_frame.text:
                    text = shape.text_frame.text.strip()
                    if text and len(text) > 1:
                        total_chars += len(re.sub(r"\s+", " ", text))
            return total_chars > self.settings["max_text_chars"], total_chars
        except Exception:
            return False, 0

    def check_animations_improved(self, slide):
        try:
            xml = str(slide.element.xml).lower()
            patterns = [
                r"<p:anim\s", r"p:ctn", r"p:seq", r"p:par",
                r"dur=['\"]", r"accel=['\"]", r"decel=['\"]",
                r"<p:custanim\s", r"<p:set\s", r"animate\s",
                r"animation\s", r"animbullet\s", r"animeffect\s",
            ]
            for p in patterns:
                if re.search(p, xml):
                    return True
            return False
        except Exception:
            return False

    # ---------------------------
    # Images + OCR (ускорение)
    # ---------------------------
    def check_images_enhanced(self, slide):
        try:
            image_info = []
            text_shapes = []

            def process_shape(shape):
                # group shapes
                if hasattr(shape, "shapes"):
                    for sub in shape.shapes:
                        process_shape(sub)
                    return

                if hasattr(shape, "image"):
                    try:
                        image_info.append({
                            "shape": shape,
                            "id": id(shape),
                            "left": shape.left,
                            "top": shape.top,
                            "right": shape.left + shape.width,
                            "bottom": shape.top + shape.height,
                            "width": shape.width,
                            "height": shape.height,
                            "format": shape.image.ext,
                        })
                    except Exception:
                        return

                if hasattr(shape, "text_frame") and shape.text_frame:
                    t = (shape.text_frame.text or "").strip()
                    if t:
                        try:
                            text_shapes.append({
                                "left": shape.left,
                                "top": shape.top,
                                "right": shape.left + shape.width,
                                "bottom": shape.top + shape.height,
                                "text": t,
                                "char_count": len(t),
                            })
                        except Exception:
                            return

            for sh in slide.shapes:
                process_shape(sh)

            if not image_info:
                return False, 0, None

            # 1) быстрый сигнал: overlap
            overlap_found = False
            for txt in text_shapes:
                if txt["char_count"] < self.settings["min_text_length_for_ocr"]:
                    continue
                for img in image_info:
                    if self.shapes_overlap(txt, img):
                        overlap_found = True
                        break
                if overlap_found:
                    break

            # OCR выключен -> только overlap
            if not self.enable_ocr:
                return overlap_found, len(image_info), None

            # overlap нет -> OCR не делаем
            if not overlap_found:
                return False, len(image_info), None

            # OCR доступен?
            if not TESSERACT_AVAILABLE:
                # есть признаки текста, но OCR недоступен
                return True, len(image_info), None

            # лимитируем кол-во картинок для OCR
            images_for_ocr = image_info[: self.settings["ocr_max_images_per_slide"]]

            ocr_results = self.check_images_with_multiple_ocr_methods(images_for_ocr)

            combined_text = ""
            total_conf = 0
            images_with_text = 0
            best_method = ""
            best_conf = 0

            for _, (text, conf, method) in ocr_results.items():
                if self.is_meaningful_text(text) and conf > self.settings["ocr_alternate_min_confidence"]:
                    images_with_text += 1
                    total_conf += conf
                    if combined_text:
                        combined_text += "\n\n---\n"
                    combined_text += text
                    if conf > best_conf:
                        best_conf = conf
                        best_method = method

            if combined_text:
                avg_conf = total_conf / images_with_text if images_with_text else 0
                return True, len(image_info), {
                    "text": combined_text,
                    "confidence": avg_conf,
                    "method": best_method or "multiple",
                    "images_with_text": images_with_text,
                }

            # overlap был, но OCR не нашёл -> оставим как “есть риск текста”
            return True, len(image_info), None

        except Exception:
            return False, 0, None

    def shapes_overlap(self, a, b):
        try:
            overlap_x = not (a["right"] <= b["left"] or a["left"] >= b["right"])
            overlap_y = not (a["bottom"] <= b["top"] or a["top"] >= b["bottom"])
            return overlap_x and overlap_y
        except Exception:
            return False

    def check_images_with_multiple_ocr_methods(self, image_info):
        results = {}
        for img in image_info:
            try:
                shape = img["shape"]
                if shape.width < 50 or shape.height < 50:
                    continue

                best = self.try_multiple_ocr_methods(shape.image.blob)
                if best:
                    text, conf, method = best
                    results[img["id"]] = (text, conf, method)
            except Exception:
                continue
        return results

    def try_multiple_ocr_methods(self, image_data):
        best_text, best_conf, best_method = "", 0, ""

        methods = [
            {"name": "PSM6", "config": f"--oem 3 --psm 6 -l {self.ocr_languages}", "pre": "standard"},
            {"name": "PSM3", "config": f"--oem 3 --psm 3 -l {self.ocr_languages}", "pre": "standard"},
            {"name": "PSM11", "config": f"--oem 3 --psm 11 -l {self.ocr_languages}", "pre": "high_contrast"},
        ]

        for m in methods:
            try:
                img = self.preprocess_for_ocr(image_data, m["pre"])
                if img is None:
                    continue

                data = pytesseract.image_to_data(img, config=m["config"], output_type=pytesseract.Output.DICT)

                parts, confs = [], []
                for j in range(len(data["text"])):
                    t = (data["text"][j] or "").strip()
                    if t and len(t) > 1:
                        parts.append(t)
                        if data["conf"][j] != "-1":
                            confs.append(float(data["conf"][j]))

                if not parts:
                    continue

                text = self.clean_ocr_text(" ".join(parts))
                conf = sum(confs) / len(confs) if confs else 0

                if text and conf > best_conf and self.quick_text_quality_check(text, conf):
                    best_text, best_conf, best_method = text, conf, m["name"]

            except Exception:
                continue

        if best_text and best_conf > self.settings["ocr_alternate_min_confidence"]:
            return best_text, best_conf, best_method
        return None

    def preprocess_for_ocr(self, image_data, method="standard"):
        try:
            img = Image.open(io.BytesIO(image_data))

            if img.mode in ("RGBA", "LA", "P"):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "RGBA":
                    bg.paste(img, mask=img.split()[3])
                else:
                    bg.paste(img)
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            img = img.convert("L")

            if method == "standard":
                img = ImageEnhance.Sharpness(img).enhance(2.0)
                img = ImageEnhance.Contrast(img).enhance(1.5)
                img = ImageOps.autocontrast(img, cutoff=2)
            elif method == "high_contrast":
                img = ImageEnhance.Contrast(img).enhance(3.0)
                img = ImageOps.autocontrast(img, cutoff=5)
                img = img.point(lambda p: 255 if p > 200 else 0)

            return img
        except Exception:
            return None

    def clean_ocr_text(self, text):
        if not text:
            return ""
        text = text.strip()
        text = text.replace("ё", "е").replace("Ё", "Е")
        text = text.replace("—", "-").replace("–", "-")
        text = text.replace("«", '"').replace("»", '"').replace("„", '"').replace("“", '"').replace("”", '"')
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def quick_text_quality_check(self, text, confidence):
        if not text or len(text) < 10:
            return False
        russian_letters = sum(1 for c in text if "а" <= c.lower() <= "я" or c in "ёе")
        total_letters = sum(1 for c in text if c.isalpha())
        if total_letters == 0:
            return False
        ratio = russian_letters / total_letters
        if confidence < 50 and ratio < 0.6:
            return False
        if ratio < 0.35:
            return False
        return True

    def is_meaningful_text(self, text):
        if not text:
            return False
        text = self.clean_ocr_text(text)
        return len(text) >= 20

    # ---------------------------
    # Fonts
    # ---------------------------
    def collect_fonts(self, slide):
        try:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    for p in shape.text_frame.paragraphs:
                        for run in p.runs:
                            name = getattr(run.font, "name", None)
                            if name and name.strip():
                                self.used_fonts.add(name.strip())
        except Exception:
            pass

    def analyze_fonts(self):
        try:
            filtered = set()
            system_fonts = [
                "+mj-lt", "+mn-lt", "calibri", "tahoma", "arial",
                "times", "verdana", "cambria", "segoe ui", "consolas",
                "courier new", "georgia", "impact", "trebuchet ms",
            ]
            for f in self.used_fonts:
                fl = f.lower()
                if any(s in fl for s in system_fonts):
                    continue
                filtered.add(f)

            font_count = len(filtered)
            for r in self.results:
                if font_count > 2:
                    r["Шрифты"] = "✗"
                    if "ШРИФТЫ" not in r["Нарушения"]:
                        r["Нарушения"].append(f"ШРИФТЫ({font_count})")
                        r["Статус"] = ", ".join(r["Нарушения"])
        except Exception:
            pass

    # ---------------------------
    # Word report (оставляем твою реализацию как есть, если она уже у тебя ниже)
    # ---------------------------
    def generate_word_report(self, results, presentation_stats, output_path=None):
        """
        Оставь здесь свою текущую generate_word_report (из твоего файла),
        она у тебя рабочая.
        """
        from docx import Document
        from docx.shared import Pt as DocxPt, RGBColor as DocxRGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()
        title = doc.add_heading("Отчет анализа презентации", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph(f"Файл: {os.path.basename(self.pptx_path)}")
        doc.add_paragraph(f"Дата анализа: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph(f"Всего слайдов: {presentation_stats.get('total_slides_in_presentation', len(results))}")
        doc.add_paragraph(f"Проанализировано: {len(results)}")
        doc.add_paragraph(f"Диапазон: {self.selected_slides_range}")
        doc.add_paragraph(f"OCR включен: {'Да' if presentation_stats.get('ocr_enabled') else 'Нет'}")
        doc.add_paragraph(f"Tesseract доступен: {'Да' if presentation_stats.get('tesseract_available') else 'Нет'}")

        # таблицы/статистика — можешь оставить как у тебя, я сократила чтобы не раздувать ответ
        # если хочешь — вставлю 1:1 твой полный отчётный блок.

        if output_path is None:
            output_path = f"report_{self.analysis_timestamp}.docx"
        doc.save(output_path)
        return output_path
