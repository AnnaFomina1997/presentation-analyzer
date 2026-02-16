import streamlit as st
import os
import re
import io
import tempfile
import time
import traceback
from datetime import datetime
from utils import PresentationAnalyzer, PresentationGenerator
import pandas as pd

# Настройка страницы
st.set_page_config(
    page_title="Анализатор презентаций",
    page_icon="📊",
    layout="wide"
)

# Инициализация session state для хранения данных между перезагрузками
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'presentation_stats' not in st.session_state:
    st.session_state['presentation_stats'] = None
if 'analyzer' not in st.session_state:
    st.session_state['analyzer'] = None
if 'original_file' not in st.session_state:
    st.session_state['original_file'] = None
if 'original_name' not in st.session_state:
    st.session_state['original_name'] = None
if 'timestamp' not in st.session_state:
    st.session_state['timestamp'] = None
if 'slides_range' not in st.session_state:
    st.session_state['slides_range'] = 'all'
if 'report_path' not in st.session_state:
    st.session_state['report_path'] = None
if 'presentation_path' not in st.session_state:
    st.session_state['presentation_path'] = None

# Заголовок
st.title("📊 Анализатор презентаций")
st.markdown("---")

# Боковая панель с информацией
with st.sidebar:
    st.header("ℹ️ О программе")
    st.info("""
    Программа проверяет презентацию на соответствие требованиям:
    - Белый фон на всех слайдах
    - Не более 2 шрифтов
    - Не более 1000 символов на слайде
    - Нет текста на изображениях
    - Нет анимаций и переходов
    """)
    
    st.header("📁 Загрузка файла")
    uploaded_file = st.file_uploader("Выберите файл .pptx", type=['pptx'])
    
    if uploaded_file is not None:
        st.session_state['original_name'] = uploaded_file.name
        
        st.header("🔍 Настройки анализа")
        
        # Выбор диапазона слайдов
        range_option = st.radio(
            "Диапазон слайдов для анализа",
            options=["Все слайды", "Один слайд", "Диапазон", "Список"],
            index=0
        )
        
        slides_range = 'all'
        if range_option == "Один слайд":
            slide_num = st.number_input("Номер слайда", min_value=1, value=1)
            slides_range = str(slide_num)
        elif range_option == "Диапазон":
            col1, col2 = st.columns(2)
            with col1:
                start = st.number_input("С", min_value=1, value=1)
            with col2:
                end = st.number_input("По", min_value=1, value=10)
            slides_range = f"{start}-{end}"
        elif range_option == "Список":
            slides_list = st.text_input("Введите номера через запятую", "1,3,5")
            slides_range = slides_list
        
        st.session_state['slides_range'] = slides_range
        
        # Кнопка для запуска анализа
        if st.button("🚀 Запустить анализ", type="primary", use_container_width=True):
            with st.spinner("Анализируем презентацию..."):
                try:
                    # Сохраняем загруженный файл во временную папку
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    st.session_state['original_file'] = tmp_path
                    st.session_state['timestamp'] = int(time.time())
                    
                    # Создаем анализатор
                    analyzer = PresentationAnalyzer(tmp_path)
                    
                    # Анализируем слайды
                    results, presentation_stats = analyzer.analyze_selected_slides(slides_range)
                    
                    if results:
                        # Сохраняем результаты в session_state
                        st.session_state['results'] = results
                        st.session_state['presentation_stats'] = presentation_stats
                        st.session_state['analyzer'] = analyzer
                        
                        # Генерируем Word отчет
                        try:
                            report_filename = f"report_{st.session_state['timestamp']}_{os.path.splitext(uploaded_file.name)[0]}.docx"
                            report_path = os.path.join(tempfile.gettempdir(), report_filename)
                            generated_report_path = analyzer.generate_word_report(results, presentation_stats, report_path)
                            
                            if generated_report_path and os.path.exists(generated_report_path):
                                st.session_state['report_path'] = generated_report_path
                                st.success(f"✅ Word отчет сгенерирован!")
                        except Exception as e:
                            st.warning(f"Не удалось сгенерировать Word отчет: {e}")
                        
                        # Генерируем исправленную презентацию
                        try:
                            generator = PresentationGenerator(tmp_path, "template.pptx")
                            presentation_filename = f"fixed_{st.session_state['timestamp']}_{os.path.splitext(uploaded_file.name)[0]}.pptx"
                            presentation_path = os.path.join(tempfile.gettempdir(), presentation_filename)
                            
                            generated_presentation_path = generator.fix_presentation(presentation_path)
                            
                            if generated_presentation_path and os.path.exists(generated_presentation_path):
                                st.session_state['presentation_path'] = generated_presentation_path
                                st.success(f"✅ Исправленная презентация сгенерирована!")
                        except Exception as e:
                            st.warning(f"Не удалось сгенерировать исправленную презентацию: {e}")
                        
                        st.success(f"✅ Анализ завершен! Проанализировано слайдов: {len(results)}")
                        st.rerun()
                    else:
                        st.error("Не удалось проанализировать презентацию")
                        
                except Exception as e:
                    st.error(f"Ошибка при анализе: {str(e)}")
                    traceback.print_exc()

# Основная область - показываем результаты, если они есть
if st.session_state['results'] is not None:
    results = st.session_state['results']
    presentation_stats = st.session_state['presentation_stats']
    analyzer = st.session_state['analyzer']
    
    # Информация о файле
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info(f"📄 **Файл:** {st.session_state['original_name']}")
    with col2:
        st.info(f"🔍 **Диапазон:** {st.session_state['slides_range']}")
    with col3:
        total_in_presentation = presentation_stats.get('total_slides_in_presentation', len(results))
        st.info(f"📊 **Слайдов:** {len(results)} из {total_in_presentation}")
    
    # Рассчитываем процент соответствия
    conformance_info = analyzer.calculate_conformance_percentage(results, presentation_stats)
    
    if conformance_info:
        st.markdown("---")
        st.header("📈 Результаты анализа")
        
        # Большие метрики
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Соответствие критериям", f"{conformance_info['percentage']}%")
        with col2:
            st.metric("Всего слайдов", conformance_info['total_slides'])
        with col3:
            st.metric("Полностью соответствующие", conformance_info['compliant_slides'])
        with col4:
            st.metric("Использовано шрифтов", presentation_stats.get('fonts_count', 0))
        
        # Показываем уровень готовности
        st.markdown(f"""
        <div style="padding: 20px; border-radius: 10px; background-color: {conformance_info['readiness_color']}20; border-left: 5px solid {conformance_info['readiness_color']}; margin: 20px 0;">
            <h3 style="margin: 0; color: {conformance_info['readiness_color']};">{conformance_info['readiness_emoji']} Уровень готовности: {conformance_info['readiness_level']}</h3>
            <p style="margin: 10px 0 0 0;">{conformance_info['user_message']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Рекомендации
        if conformance_info['recommendations']:
            st.markdown("#### 📋 Рекомендации по улучшению:")
            for rec in conformance_info['recommendations']:
                st.warning(rec)
        
        # Детальная статистика
        with st.expander("📊 Детальная статистика по критериям"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Фон:**")
                bg_score = conformance_info['criteria_details']['background']['score']
                bg_max = conformance_info['criteria_details']['background']['max']
                bg_issues = conformance_info['criteria_details']['background']['issues']
                st.progress(bg_score/bg_max, text=f"{bg_score}/{bg_max} баллов")
                st.caption(f"Слайдов с нарушением: {bg_issues}")
                
                st.markdown("**Шрифты:**")
                fonts_score = conformance_info['criteria_details']['fonts']['score']
                fonts_max = conformance_info['criteria_details']['fonts']['max']
                fonts_count = conformance_info['criteria_details']['fonts']['fonts_count']
                st.progress(fonts_score/fonts_max, text=f"{fonts_score}/{fonts_max} баллов")
                st.caption(f"Использовано шрифтов: {fonts_count}")
                
                st.markdown("**Текстовая перегрузка:**")
                text_score = conformance_info['criteria_details']['text_overload']['score']
                text_max = conformance_info['criteria_details']['text_overload']['max']
                text_issues = conformance_info['criteria_details']['text_overload']['issues']
                st.progress(text_score/text_max, text=f"{text_score}/{text_max} баллов")
                st.caption(f"Слайдов с нарушением: {text_issues}")
            
            with col2:
                st.markdown("**Текст на изображениях:**")
                img_score = conformance_info['criteria_details']['text_on_images']['score']
                img_max = conformance_info['criteria_details']['text_on_images']['max']
                img_issues = conformance_info['criteria_details']['text_on_images']['issues']
                st.progress(img_score/img_max, text=f"{img_score}/{img_max} баллов")
                st.caption(f"Слайдов с нарушением: {img_issues}")
                
                st.markdown("**Анимации:**")
                anim_score = conformance_info['criteria_details']['animations']['score']
                anim_max = conformance_info['criteria_details']['animations']['max']
                anim_issues = conformance_info['criteria_details']['animations']['issues']
                st.progress(anim_score/anim_max, text=f"{anim_score}/{anim_max} баллов")
                st.caption(f"Слайдов с нарушением: {anim_issues}")
                
                st.markdown("**Переходы:**")
                trans_score = conformance_info['criteria_details']['transitions']['score']
                trans_max = conformance_info['criteria_details']['transitions']['max']
                has_trans = conformance_info['criteria_details']['transitions']['has_issues']
                st.progress(trans_score/trans_max, text=f"{trans_score}/{trans_max} баллов")
                st.caption(f"Есть переходы: {'Да' if has_trans else 'Нет'}")
        
        # Детальная таблица по слайдам
        st.markdown("---")
        st.header("📊 Детальный анализ по слайдам")
        
        # Создаем DataFrame для таблицы
        df_data = []
        for r in results:
            df_data.append({
                'Слайд': r['Слайд'],
                'Статус': r['Статус'],
                'Фон': r['Фон'],
                'Шрифты': r['Шрифты'],
                'Текст': r['Текст_дет'],
                'Элементы': r['Элементы'],
                'Изображения': r['Изображения'],
                'Текст на изобр.': r['Текст_на_изобр'],
                'Анимации': r['Анимации']
            })
        
        df = pd.DataFrame(df_data)
        
        # Функция для подсветки ячеек
        def highlight_cells(val):
            if val == '✗':
                return 'color: red; font-weight: bold'
            elif val == '✓':
                return 'color: green; font-weight: bold'
            elif val == 'Да':
                return 'color: red; font-weight: bold'
            elif val == 'Нет':
                return 'color: green; font-weight: bold'
            return ''
        
        styled_df = df.style.map(highlight_cells, subset=['Фон', 'Шрифты', 'Текст на изобр.', 'Анимации'])
        st.dataframe(styled_df, use_container_width=True, height=400)
        
        # OCR текст, если есть
        if analyzer.full_ocr_texts:
            with st.expander("🔍 Текст, найденный на изображениях (OCR)", expanded=False):
                tabs = st.tabs([f"Слайд {slide_num}" for slide_num in analyzer.full_ocr_texts.keys()])
                
                for i, (slide_num, ocr_data) in enumerate(analyzer.full_ocr_texts.items()):
                    with tabs[i]:
                        st.markdown(f"**Изображений на слайде:** {ocr_data.get('image_count', 0)}")
                        st.markdown(f"**Изображений с текстом:** {ocr_data.get('images_with_text', 0)}")
                        st.markdown(f"**Уверенность:** {ocr_data.get('confidence', 0):.1f}%")
                        st.markdown(f"**Метод:** {ocr_data.get('method', 'unknown')}")
                        st.markdown("**Текст:**")
                        st.text_area("", ocr_data.get('text', ''), height=200, key=f"ocr_{slide_num}")
        
        # Кнопки для скачивания
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.session_state['report_path'] and os.path.exists(st.session_state['report_path']):
                with open(st.session_state['report_path'], 'rb') as f:
                    report_data = f.read()
                
                st.download_button(
                    label="📥 Скачать Word отчет",
                    data=report_data,
                    file_name=f"анализ_презентации_{os.path.splitext(st.session_state['original_name'])[0]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.button("📥 Word отчет не доступен", disabled=True, use_container_width=True)
        
        with col2:
            if st.session_state['presentation_path'] and os.path.exists(st.session_state['presentation_path']):
                with open(st.session_state['presentation_path'], 'rb') as f:
                    pres_data = f.read()
                
                st.download_button(
                    label="🔄 Скачать исправленную презентацию",
                    data=pres_data,
                    file_name=f"исправленная_{st.session_state['original_name']}",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            else:
                st.button("🔄 Исправленная презентация не доступна", disabled=True, use_container_width=True)
        
        # Кнопка для нового анализа
        if st.button("🔄 Новый анализ", use_container_width=True):
            for key in ['results', 'presentation_stats', 'analyzer', 'original_file', 
                       'original_name', 'timestamp', 'slides_range', 'report_path', 'presentation_path']:
                if key in st.session_state:
                    st.session_state[key] = None
            st.rerun()

else:
    # Если файл не загружен, показываем приветствие
    st.info("👈 Загрузите файл презентации в боковой панели для начала анализа")
    
    # Пример того, как будет выглядеть результат
    st.markdown("### Пример результата:")
    example_df = pd.DataFrame({
        'Слайд': [1, 2, 3],
        'Статус': ['OK', 'ТЕКСТ(1200)', 'ФОН, ТЕКСТ_НА_ИЗОБР'],
        'Фон': ['✓', '✓', '✗'],
        'Шрифты': ['✓', '✓', '✓'],
        'Текст': ['540 симв.', '1200 симв.', '320 симв.'],
        'Изображения': [0, 2, 1],
        'Текст на изобр.': ['Нет', 'Нет', 'Да']
    })
    st.dataframe(example_df, use_container_width=True)

# Подвал
st.markdown("---")
st.markdown("© 2024 Анализатор презентаций | Версия 1.0 (Streamlit)")