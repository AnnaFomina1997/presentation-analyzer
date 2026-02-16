import streamlit as st
import os
import tempfile
import time
import traceback
from utils import PresentationAnalyzer, PresentationGenerator
import pandas as pd

st.set_page_config(page_title="Анализатор презентаций", page_icon="📊", layout="wide")

# ---------------------------
# Session state init
# ---------------------------
defaults = {
    "results": None,
    "presentation_stats": None,
    "original_name": None,
    "timestamp": None,
    "slides_range": "all",
    "enable_ocr": True,
    "report_bytes": None,
    "report_filename": None,
    "presentation_bytes": None,
    "presentation_filename": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def clear_state_for_new_run():
    st.session_state["results"] = None
    st.session_state["presentation_stats"] = None
    st.session_state["timestamp"] = None
    st.session_state["report_bytes"] = None
    st.session_state["report_filename"] = None
    st.session_state["presentation_bytes"] = None
    st.session_state["presentation_filename"] = None


st.title("📊 Анализатор презентаций")
st.markdown("---")

with st.sidebar:
    st.header("ℹ️ О программе")
    st.info(
        """
Программа проверяет презентацию на соответствие требованиям:
- Белый фон на всех слайдах
- Не более 2 шрифтов
- Не более 1000 символов на слайде
- Нет текста на изображениях (OCR)
- Нет анимаций и переходов
"""
    )

    st.header("📁 Загрузка файла")
    uploaded_file = st.file_uploader("Выберите файл .pptx", type=["pptx"])

    if uploaded_file is not None:
        st.session_state["original_name"] = uploaded_file.name

        # Важно: form предотвращает лишние перезапуски на каждом изменении виджета
        with st.form("analyze_form", clear_on_submit=False):
            st.header("🔍 Настройки анализа")

            range_option = st.radio(
                "Диапазон слайдов для анализа",
                options=["Все слайды", "Один слайд", "Диапазон", "Список"],
                index=0,
            )

            slides_range = "all"
            if range_option == "Один слайд":
                slide_num = st.number_input("Номер слайда", min_value=1, value=1)
                slides_range = str(slide_num)
            elif range_option == "Диапазон":
                c1, c2 = st.columns(2)
                with c1:
                    start = st.number_input("С", min_value=1, value=1)
                with c2:
                    end = st.number_input("По", min_value=1, value=10)
                slides_range = f"{start}-{end}"
            elif range_option == "Список":
                slides_list = st.text_input("Введите номера через запятую", "1,3,5")
                slides_range = slides_list

            enable_ocr = st.toggle(
                "🔍 OCR (поиск текста на изображениях)",
                value=st.session_state["enable_ocr"],
                help="Если выключить — анализ будет значительно быстрее. "
                     "При включении OCR запускается только когда есть признаки текста поверх картинки.",
            )

            submitted = st.form_submit_button("🚀 Запустить анализ", type="primary", use_container_width=True)

        st.session_state["slides_range"] = slides_range
        st.session_state["enable_ocr"] = enable_ocr

        if submitted:
            clear_state_for_new_run()
            st.session_state["timestamp"] = int(time.time())

            with st.spinner("Анализируем презентацию..."):
                try:
                    file_bytes = uploaded_file.getvalue()

                    # сохраняем во временный файл
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                        tmp.write(file_bytes)
                        tmp_path = tmp.name

                    analyzer = PresentationAnalyzer(tmp_path, enable_ocr=enable_ocr)

                    results, presentation_stats = analyzer.analyze_selected_slides(slides_range)
                    if not results:
                        st.error("Не удалось проанализировать презентацию")
                        st.stop()

                    st.session_state["results"] = results
                    st.session_state["presentation_stats"] = presentation_stats

                    # Word report -> bytes
                    try:
                        clean_name = os.path.splitext(uploaded_file.name)[0]
                        report_filename = f"анализ_презентации_{clean_name}.docx"

                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_doc:
                            report_path = tmp_doc.name

                        out_report_path = analyzer.generate_word_report(results, presentation_stats, report_path)
                        if out_report_path and os.path.exists(out_report_path):
                            with open(out_report_path, "rb") as f:
                                st.session_state["report_bytes"] = f.read()
                            st.session_state["report_filename"] = report_filename
                            st.success("✅ Word отчет сгенерирован!")
                        else:
                            st.warning("Не удалось сгенерировать Word отчет")
                    except Exception as e:
                        st.warning(f"Не удалось сгенерировать Word отчет: {e}")
                        traceback.print_exc()

                    # Fixed pptx -> bytes
                    try:
                        template_path = "template.pptx"  # лежит рядом со streamlit_app.py
                        if not os.path.exists(template_path):
                            st.warning("Файл template.pptx не найден в корневой папке проекта")
                        else:
                            generator = PresentationGenerator(tmp_path, template_path)

                            clean_name = os.path.splitext(uploaded_file.name)[0]
                            pres_filename = f"исправленная_{clean_name}.pptx"

                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_out:
                                out_pptx_path = tmp_out.name

                            st.info("Генерация исправленной презентации...")
                            out_path = generator.fix_presentation(out_pptx_path)

                            if out_path and os.path.exists(out_path):
                                with open(out_path, "rb") as f:
                                    st.session_state["presentation_bytes"] = f.read()
                                st.session_state["presentation_filename"] = pres_filename
                                st.success("✅ Исправленная презентация сгенерирована!")
                            else:
                                st.error("Не удалось создать исправленную презентацию")
                    except Exception as e:
                        st.error(f"Ошибка генерации презентации: {str(e)}")
                        traceback.print_exc()

                    st.success(f"✅ Анализ завершен! Проанализировано слайдов: {len(results)}")

                except Exception as e:
                    st.error(f"Ошибка при анализе: {str(e)}")
                    traceback.print_exc()


# ---------------------------
# Main: results
# ---------------------------
if st.session_state["results"] is not None:
    results = st.session_state["results"]
    presentation_stats = st.session_state["presentation_stats"]

    c1, c2, c3 = st.columns(3)
    with c1:
        st.info(f"📄 **Файл:** {st.session_state['original_name']}")
    with c2:
        st.info(f"🔍 **Диапазон:** {st.session_state['slides_range']}")
    with c3:
        total_in_presentation = presentation_stats.get("total_slides_in_presentation", len(results))
        st.info(f"📊 **Слайдов:** {len(results)} из {total_in_presentation}")

    # conformance – можно считать без хранения analyzer в session_state
    dummy = PresentationAnalyzer("__dummy__", enable_ocr=False)
    conformance_info = dummy.calculate_conformance_percentage(results, presentation_stats)

    if conformance_info:
        st.markdown("---")
        st.header("📈 Результаты анализа")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Соответствие критериям", f"{conformance_info['percentage']}%")
        with col2:
            st.metric("Всего слайдов", conformance_info["total_slides"])
        with col3:
            st.metric("Полностью соответствующие", conformance_info["compliant_slides"])
        with col4:
            st.metric("Использовано шрифтов", presentation_stats.get("fonts_count", 0))

        st.markdown(
            f"""
            <div style="padding: 20px; border-radius: 10px; background-color: {conformance_info['readiness_color']}20; border-left: 5px solid {conformance_info['readiness_color']}; margin: 20px 0;">
                <h3 style="margin: 0; color: {conformance_info['readiness_color']};">{conformance_info['readiness_emoji']} Уровень готовности: {conformance_info['readiness_level']}</h3>
                <p style="margin: 10px 0 0 0;">{conformance_info['user_message']}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if conformance_info["recommendations"]:
            st.markdown("#### 📋 Рекомендации по улучшению:")
            for rec in conformance_info["recommendations"]:
                st.warning(rec)

    # --- Таблица БЕЗ Styler (не дрожит)
    st.markdown("---")
    st.header("📊 Детальный анализ по слайдам")

    df = pd.DataFrame(
        [{
            "Слайд": r["Слайд"],
            "Статус": r["Статус"],
            "Фон": r["Фон"],
            "Шрифты": r["Шрифты"],
            "Текст": r["Текст_дет"],
            "Элементы": r["Элементы"],
            "Изображения": r["Изображения"],
            "Текст на изобр.": r["Текст_на_изобр"],
            "Анимации": r["Анимации"],
        } for r in results]
    )

    st.dataframe(df, use_container_width=True, height=420, hide_index=True)

    # OCR вкладки берём из results (там уже есть OCR_текст)
    ocr_rows = [r for r in results if r.get("OCR_текст")]
    if ocr_rows:
        with st.expander("🔍 Текст, найденный на изображениях (OCR)", expanded=False):
            tabs = st.tabs([f"Слайд {r['Слайд']}" for r in ocr_rows])
            for i, r in enumerate(ocr_rows):
                with tabs[i]:
                    st.markdown(f"**Изображений на слайде:** {r.get('Изображения', 0)}")
                    st.markdown(f"**Изображений с текстом:** {r.get('OCR_изображений_с_текстом', 0)}")
                    st.markdown(f"**Уверенность:** {r.get('OCR_уверенность', 0):.1f}%")
                    st.markdown(f"**Метод:** {r.get('OCR_метод', '')}")
                    st.text_area("", r.get("OCR_текст", ""), height=220, key=f"ocr_{r['Слайд']}")

    # downloads
    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        if st.session_state["report_bytes"]:
            st.download_button(
                "📥 Скачать Word отчет",
                data=st.session_state["report_bytes"],
                file_name=st.session_state["report_filename"] or "report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.button("📥 Word отчет не доступен", disabled=True, use_container_width=True)

    with col2:
        if st.session_state["presentation_bytes"]:
            st.download_button(
                "📥 Скачать исправленную презентацию",
                data=st.session_state["presentation_bytes"],
                file_name=st.session_state["presentation_filename"] or "fixed.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        else:
            st.button("🔄 Исправленная презентация не доступна", disabled=True, use_container_width=True)

    if st.button("🔄 Новый анализ", use_container_width=True):
        clear_state_for_new_run()
        st.rerun()

else:
    st.info("👈 Загрузите файл презентации в боковой панели для начала анализа")

st.markdown("---")
st.markdown("© 2024 Анализатор презентаций | Версия 1.1 (Streamlit)")
