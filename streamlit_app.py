import streamlit as st
import os
import tempfile
import time
from datetime import datetime
from utils import PresentationAnalyzer, PresentationGenerator
import pandas as pd

# Настройка страницы
st.set_page_config(
    page_title="Анализатор презентаций",
    page_icon="📊",
    layout="wide"
)

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
    
    st.header("🔍 Настройки анализа")
    slides_range = st.text_input(
        "Диапазон слайдов для анализа",
        value="all",
        help="Примеры: all, 1, 1-5, 1,3,5-7"
    )

# Основная область
if uploaded_file is not None:
    # Сохраняем загруженный файл во временную папку
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
    
    # Показываем информацию о файле
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info(f"📄 **Файл:** {uploaded_file.name}")
    with col2:
        file_size = len(uploaded_file.getvalue()) / (1024*1024)
        st.info(f"📦 **Размер:** {file_size:.2f} MB")
    with col3:
        st.info(f"🔍 **Диапазон:** {slides_range}")
    
    # Кнопка для запуска анализа
    if st.button("🚀 Запустить анализ", type="primary", use_container_width=True):
        with st.spinner("Анализируем презентацию..."):
            try:
                # Создаем анализатор
                analyzer = PresentationAnalyzer(tmp_path)
                
                # Анализируем слайды
                results, presentation_stats = analyzer.analyze_selected_slides(slides_range)
                
                if results:
                    # Сохраняем результаты в session_state
                    st.session_state['results'] = results
                    st.session_state['presentation_stats'] = presentation_stats
                    st.session_state['analyzer'] = analyzer
                    st.session_state['tmp_path'] = tmp_path
                    
                    st.success(f"✅ Анализ завершен! Проанализировано слайдов: {len(results)}")
                else:
                    st.error("Не удалось проанализировать презентацию")
                    
            except Exception as e:
                st.error(f"Ошибка при анализе: {str(e)}")
    
    # Если есть результаты, показываем их
    if 'results' in st.session_state:
        results = st.session_state['results']
        presentation_stats = st.session_state['presentation_stats']
        analyzer = st.session_state['analyzer']
        
        # Рассчитываем процент соответствия
        conformance_info = analyzer.calculate_conformance_percentage(results, presentation_stats)
        
        if conformance_info:
            # Показываем общий результат
            st.markdown("---")
            st.header("📈 Результаты анализа")
            
            # Большая метрика с процентом
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Соответствие критериям", f"{conformance_info['percentage']}%")
            with col2:
                st.metric("Всего слайдов", conformance_info['total_slides'])
            with col3:
                st.metric("Полностью соответствующие", conformance_info['compliant_slides'])
            
            # Показываем уровень готовности
            st.markdown(f"""
            <div style="padding: 20px; border-radius: 10px; background-color: {conformance_info['readiness_color']}20; border-left: 5px solid {conformance_info['readiness_color']};">
                <h3 style="margin: 0; color: {conformance_info['readiness_color']};">{conformance_info['readiness_emoji']} Уровень готовности: {conformance_info['readiness_level']}</h3>
                <p style="margin: 10px 0 0 0;">{conformance_info['user_message']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Рекомендации
            if conformance_info['recommendations']:
                st.markdown("#### 📋 Рекомендации по улучшению:")
                for rec in conformance_info['recommendations']:
                    st.warning(rec)
            
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
            st.dataframe(df, use_container_width=True, height=400)
            
            # OCR текст, если есть
            if analyzer.full_ocr_texts:
                with st.expander("🔍 Текст, найденный на изображениях (OCR)"):
                    for slide_num, ocr_data in analyzer.full_ocr_texts.items():
                        st.markdown(f"**Слайд {slide_num}**")
                        st.markdown(f"*Уверенность: {ocr_data['confidence']:.1f}%*")
                        st.text(ocr_data['text'][:1000] + "..." if len(ocr_data['text']) > 1000 else ocr_data['text'])
                        st.markdown("---")
            
            # Кнопки для скачивания
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Скачать Word отчет", use_container_width=True):
                    with st.spinner("Генерируем отчет..."):
                        # Создаем временный файл для отчета
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_report:
                            report_path = analyzer.generate_word_report(results, presentation_stats, tmp_report.name)
                            
                            if report_path and os.path.exists(report_path):
                                with open(report_path, 'rb') as f:
                                    st.download_button(
                                        label="✅ Нажмите для скачивания",
                                        data=f,
                                        file_name=f"отчет_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        use_container_width=True
                                    )
            
            with col2:
                if st.button("🔄 Создать исправленную презентацию", use_container_width=True):
                    with st.spinner("Генерируем презентацию по шаблону..."):
                        try:
                            # Создаем генератор
                            generator = PresentationGenerator(tmp_path, "template.pptx")
                            
                            # Создаем временный файл для результата
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_result:
                                result_path = generator.fix_presentation(tmp_result.name)
                                
                                if result_path and os.path.exists(result_path):
                                    with open(result_path, 'rb') as f:
                                        st.download_button(
                                            label="✅ Нажмите для скачивания",
                                            data=f,
                                            file_name=f"исправленная_{uploaded_file.name}",
                                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                            use_container_width=True
                                        )
                        except Exception as e:
                            st.error(f"Ошибка при создании презентации: {str(e)}")

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
st.markdown("© 2024 Анализатор презентаций | Версия 1.0")