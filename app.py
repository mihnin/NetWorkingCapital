import streamlit as st
import pandas as pd
import numpy as np
import io # Для работы с байтовыми потоками Excel

# --- Информация о шаблоне Excel ---
EXCEL_TEMPLATE_INFO_UPDATED = """
**Руководство по подготовке Excel файла для загрузки:**

Ваш Excel файл должен содержать один лист (приложение будет читать **первый лист** в файле) со следующей структурой:

1.  **Обязательные столбцы (названия должны быть точными, регистр важен):**
    * **Столбец A: `Статья`**
        * Содержит наименования статей Чистого Оборотного Капитала (ЧОК).
        * *Пример: "Денежные средства (ДС)", "Дебиторская задолженность (ДЗ)" и т.д.*
    * **Столбец B: `Тип`**
        * Указывает тип статьи. Допустимые значения (текст, заглавными буквами):
            * **`ОА`** (для Оборотных Активов)
            * **`КО`** (для Краткосрочных Обязательств)

2.  **Столбцы с данными по периодам (начиная со столбца C):**
    * **Названия столбцов:** Вы можете называть их так, как вам удобно для обозначения периода (например, "Янв 2024", "Месяц 1", "Квартал 1 2023"). Приложение не анализирует саму дату из названия, а ориентируется на ключевые слова.
    * **Ключевые слова для автоматического определения типа данных (важно!):**
        * Чтобы столбец был распознан как **фактические данные**, его название должно содержать слово **`Факт`** (без учета регистра, например, "Янв 2024 Факт", "факт за январь", "Q1 Факт").
        * Чтобы столбец был распознан как **прогнозные данные**, его название должно содержать слово **`Прогноз`** (без учета регистра, например, "Янв 2025 Прогноз", "прогноз на январь", "Q1 Прогноз").
    * **Содержимое:** Эти столбцы должны содержать только числовые значения. Пустые ячейки или нечисловые значения будут интерпретированы как отсутствующие данные (NaN) и могут повлиять на расчеты. Рекомендуется заменять пропуски нулями, если это уместно для вашей статьи.
    * **Пример названий:** "Янв 2023 Факт", "Фев 2023 Факт", ..., "Дек 2023 Факт", "Янв 2024 Прогноз".

3.  **Количество столбцов с данными по периодам:**
    * **Минимум:** Для полноценной работы всех функций анализа рекомендуется иметь хотя бы **один столбец с фактическими данными** и **один столбец с прогнозными данными**.
    * **Максимум:** Приложение теоретически не накладывает жестких ограничений. Однако, очень большое количество столбцов может замедлить обработку.

4.  **Прочее:**
    * Убедитесь, что на листе нет объединенных ячеек в области данных.
    * Данные должны начинаться с первой строки (строка 1 для заголовков, строка 2 для первой статьи ЧОК).

**Вы можете скачать шаблон с демонстрационными данными, нажав кнопку ниже в боковой панели, чтобы увидеть пример правильной структуры.**
"""

# --- 1. ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ ---
# ... (все функции get_demo_data, generate_template_excel_bytes, load_external_data, calculate_period_totals без изменений из предыдущей версии)
def get_demo_data():
    data = {
        'Статья': [
            'Денежные средства (ДС)', 'Дебиторская задолженность (ДЗ)',
            'Сырье и материалы (СиМ)', 'Незавершенное производство (НЗП)', 'Готовая продукция (ГП)',
            'Прочие оборотные активы',
            'Кредиторская задолженность (КЗ)', 'Краткосрочные кредиты и займы',
            'Налоги (к уплате)', 'Прочие краткосрочные обязательства'
        ],
        'Тип': ['ОА', 'ОА', 'ОА', 'ОА', 'ОА', 'ОА', 'КО', 'КО', 'КО', 'КО'],
        'Q1 2024 Факт': [500, 1200, 300, 200, 400, 50,  700, 400, 50, 150],
        'Q2 2024 Факт': [550, 1300, 320, 210, 420, 55,  750, 420, 60, 160],
        'Q3 2024 Факт': [520, 1250, 310, 205, 405, 52,  720, 410, 55, 155],
        'Q4 2024 Факт': [600, 1400, 350, 230, 450, 60,  800, 450, 70, 170],
        'Q1 2025 Прогноз': [620, 1450, 360, 240, 460, 65, 820, 460, 75, 175]
    }
    return pd.DataFrame(data)

def generate_template_excel_bytes():
    df_template = get_demo_data()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_template.to_excel(writer, sheet_name="Шаблон_ЧОК_Данные", index=False)
    return output.getvalue()

def load_external_data(uploaded_file):
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0)
            if 'Статья' not in df.columns or 'Тип' not in df.columns:
                st.error("Ошибка: В файле отсутствуют столбцы 'Статья' и/или 'Тип'.")
                return None
            if not df['Тип'].isin(['ОА', 'КО']).all():
                st.error("Ошибка: Столбец 'Тип' может содержать только 'ОА' или 'КО'.")
                return None
            data_cols = df.columns.drop(['Статья', 'Тип'])
            if not data_cols.tolist():
                 st.error("Ошибка: В файле отсутствуют столбцы с данными по периодам.")
                 return None
            for col in data_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            st.success("Файл успешно загружен!")
            return df
        except Exception as e:
            st.error(f"Ошибка при чтении файла: {e}")
            return None
    return None

def calculate_period_totals(df_articles, columns_to_calculate):
    period_totals_data = {}
    for col_name in columns_to_calculate:
        if col_name not in df_articles.columns: continue
        oa_total = df_articles.loc[df_articles['Тип'] == 'ОА', col_name].sum()
        co_total = df_articles.loc[df_articles['Тип'] == 'КО', col_name].sum()
        working_capital = oa_total - co_total
        period_totals_data[col_name] = {
            'Итого ОА': oa_total, 'Итого КО': co_total, 'ЧОК': working_capital
        }
    return period_totals_data

# --- 2. РАСЧЕТ СУЩЕСТВЕННОСТИ (НЕСКОЛЬКО МЕТОДОВ) ---
def calculate_materiality(df_articles, period_totals_data, data_columns_list, method="vs_CHOK"):
    materiality_data = {'Статья': df_articles['Статья'].tolist()}
    if method == "within_OA_CO":
        materiality_data['Тип'] = df_articles['Тип'].tolist()
    for col_name in data_columns_list:
        if col_name not in df_articles.columns: continue
        col_materiality = []
        base_chok = period_totals_data.get(col_name, {}).get('ЧОК', 0)
        base_total_components = df_articles[col_name].abs().sum() if method == "vs_TotalComponents" else 0
        base_total_oa = period_totals_data.get(col_name, {}).get('Итого ОА', 0) if method == "within_OA_CO" else 0
        base_total_co = period_totals_data.get(col_name, {}).get('Итого КО', 0) if method == "within_OA_CO" else 0
        for _, row in df_articles.iterrows():
            article_value = row[col_name]
            base_value_for_calc = 0
            if method == "vs_CHOK": base_value_for_calc = base_chok
            elif method == "vs_TotalComponents": base_value_for_calc = base_total_components
            elif method == "within_OA_CO":
                if row['Тип'] == 'ОА': base_value_for_calc = base_total_oa
                elif row['Тип'] == 'КО': base_value_for_calc = base_total_co
            if pd.isna(article_value) or base_value_for_calc == 0:
                col_materiality.append(np.nan)
            else:
                col_materiality.append((np.abs(article_value) / np.abs(base_value_for_calc)) * 100)
        materiality_data[f'Сущ-ть ({col_name.split(" ")[0]}) (%)'] = col_materiality
    df_result = pd.DataFrame(materiality_data)
    if method == "within_OA_CO" and 'Тип' in df_result.columns:
        df_result = df_result.drop(columns=['Тип'])
    return df_result

# --- 3. ОТКЛОНЕНИЯ ПРОГНОЗА ---
def calculate_forecast_deviations(df_articles, period_totals_data, forecast_col_name, base_col_name):
    # ... (код из предыдущего ответа, он корректен)
    deviations_data = {'Статья': df_articles['Статья'].tolist()}
    abs_deviations_list, rel_deviations_list = [], []
    if base_col_name not in df_articles.columns or forecast_col_name not in df_articles.columns:
        st.warning(f"Колонки для отклонений ('{base_col_name}'/'{forecast_col_name}') не найдены.")
        empty_df_articles = pd.DataFrame(deviations_data) 
        empty_df_summary = pd.DataFrame(columns=['Показатель', 'Прогноз', f'Факт ({base_col_name.split(" ")[0]})', 'Абс. откл.', 'Отн. откл. (%)'])
        return empty_df_articles, empty_df_summary
    for _, row in df_articles.iterrows():
        forecast_value, base_value = row[forecast_col_name], row[base_col_name]
        if pd.isna(forecast_value) or pd.isna(base_value):
            abs_dev, rel_dev = np.nan, np.nan
        else:
            abs_dev = forecast_value - base_value
            rel_dev = (abs_dev / base_value) * 100 if base_value != 0 else (np.nan if forecast_value != 0 else 0)
        abs_deviations_list.append(abs_dev)
        rel_deviations_list.append(rel_dev)
    deviations_data[f'Абс. откл. (Прогноз - {base_col_name.split(" ")[0]})'] = abs_deviations_list
    deviations_data[f'Отн. откл. (Прогноз - {base_col_name.split(" ")[0]}) (%)'] = rel_deviations_list
    df_deviations_result = pd.DataFrame(deviations_data)
    summary_dev_rows = []
    for indicator in ['Итого ОА', 'Итого КО', 'ЧОК']:
        prog_data = period_totals_data.get(forecast_col_name, {})
        base_data = period_totals_data.get(base_col_name, {})
        prog_val, base_val = prog_data.get(indicator, np.nan), base_data.get(indicator, np.nan)
        if pd.isna(prog_val) or pd.isna(base_val):
            abs_d, rel_d = np.nan, np.nan
        else:
            abs_d = prog_val - base_val
            rel_d = (abs_d / base_val) * 100 if base_val != 0 else (np.nan if prog_val !=0 else 0)
        summary_dev_rows.append({
            'Показатель': indicator, 'Прогноз': prog_val, f'Факт ({base_col_name.split(" ")[0]})': base_val,
            'Абс. откл.': abs_d, 'Отн. откл. (%)': rel_d
        })
    return df_deviations_result, pd.DataFrame(summary_dev_rows)

# --- 4. ДОПУСТИМЫЙ ДИАПАЗОН ОШИБКИ ПРОГНОЗА СТАТЬИ ---
def calculate_allowed_article_error_range(df_articles, forecast_col_name, forecast_chok_value, wc_deviation_perc_limit=5):
    # ... (код из предыдущего ответа, он корректен)
    if forecast_col_name not in df_articles.columns:
        st.warning(f"Прогнозная колонка '{forecast_col_name}' отсутствует.")
        return pd.DataFrame({'Статья': df_articles['Статья'].tolist()}), np.nan
    if pd.isna(forecast_chok_value):
        st.warning("Прогнозный ЧОК не определен, анализ чувствительности невозможен.")
        return pd.DataFrame({'Статья': df_articles['Статья'].tolist()}), np.nan
    max_abs_wc_deviation_allowed = np.abs(forecast_chok_value * (wc_deviation_perc_limit / 100.0))
    allowed_error_ranges_data = {'Статья': df_articles['Статья'].tolist()}
    error_range_percentages = []
    for _, row in df_articles.iterrows():
        forecast_article_value = row[forecast_col_name]
        if pd.isna(forecast_article_value) or forecast_article_value == 0:
            err_range_perc = np.inf
        else:
            err_range_perc = (max_abs_wc_deviation_allowed / np.abs(forecast_article_value)) * 100
        error_range_percentages.append(err_range_perc)
    allowed_error_ranges_data[f'Макс. ошибка статьи (+/- %) для откл. ЧОК до {wc_deviation_perc_limit}%'] = error_range_percentages
    return pd.DataFrame(allowed_error_ranges_data), max_abs_wc_deviation_allowed

# --- 5. ФУНКЦИЯ ВЫГРУЗКИ В EXCEL ---
def dfs_to_excel_bytes(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df_data in dfs_dict.items():
            if df_data is not None and isinstance(df_data, pd.DataFrame) and not df_data.empty:
                # Сокращаем имя листа, если оно слишком длинное
                safe_sheet_name = sheet_name[:31]
                index_val = True 
                if safe_sheet_name.startswith("Данные_статьи") or \
                   safe_sheet_name.startswith("Сущ_") or \
                   safe_sheet_name.startswith("Отклонения_статьи") or \
                   safe_sheet_name.startswith("Допустимые_ошибки"):
                    index_val = False
                df_data.to_excel(writer, sheet_name=safe_sheet_name, index=index_val)
    return output.getvalue()

# --- STREAMLIT APP ---
def main():
    st.set_page_config(layout="wide", page_title="Расширенный Анализ ЧОК")
    st.title("Расширенный Анализ Чистого Оборотного Капитала (ЧОК)")
    st.markdown("Этот инструмент предназначен для анализа структуры ЧОК, оценки точности прогнозов и выявления статей, оказывающих наибольшее влияние на итоговый ЧОК. Используйте боковую панель для загрузки данных и настройки параметров анализа.")

    if 'df_main_articles' not in st.session_state:
        st.session_state.df_main_articles = get_demo_data()
        st.session_state.data_source = "демо"

    st.sidebar.header("Управление данными")
    uploaded_file = st.sidebar.file_uploader("Загрузить Excel (см. шаблон ниже)", type=["xlsx", "xls"], key="file_uploader")
    col1, col2 = st.sidebar.columns(2)
    if col1.button("Использ. демо-данные", key="use_demo_data_btn", use_container_width=True): # Сократил название кнопки
        st.session_state.df_main_articles = get_demo_data()
        st.session_state.data_source = "демо"
        st.rerun()
    template_bytes = generate_template_excel_bytes()
    col2.download_button(label="Скачать шаблон", data=template_bytes, file_name="ЧОК_анализ_шаблон_данных.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_template_btn", use_container_width=True)

    if uploaded_file:
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get('last_uploaded_file_id') != current_file_id:
            df_custom = load_external_data(uploaded_file)
            if df_custom is not None:
                st.session_state.df_main_articles = df_custom
                st.session_state.data_source = "загруженный файл"
                st.session_state.last_uploaded_file_id = current_file_id
                st.rerun()
            else:
                st.session_state.last_uploaded_file_id = None
    df_main_articles = st.session_state.df_main_articles
    st.sidebar.caption(f"Источник данных: {st.session_state.data_source}")
    with st.sidebar.expander("Инструкция и формат шаблона Excel", expanded=False):
        st.markdown(EXCEL_TEMPLATE_INFO_UPDATED)

    available_cols = df_main_articles.columns.drop(['Статья', 'Тип']).tolist()
    fact_actual_columns_all = [col for col in available_cols if "факт" in col.lower()]
    forecast_actual_columns_available = [col for col in available_cols if "прогноз" in col.lower()]

    st.sidebar.header("Настройки анализа")
    selected_fact_columns = st.sidebar.multiselect("Фактические периоды для анализа:", options=fact_actual_columns_all, default=fact_actual_columns_all)
    base_for_deviation_analysis = st.sidebar.selectbox("Базовый факт. период для сравнения:", options=fact_actual_columns_all,
                                                       index=len(fact_actual_columns_all)-1 if fact_actual_columns_all else 0, disabled=not fact_actual_columns_all)
    forecast_actual_column_selected = st.sidebar.selectbox("Прогнозный период для анализа:", options=forecast_actual_columns_available,
                                                           index=0 if forecast_actual_columns_available else 0, disabled=not forecast_actual_columns_available)
    chok_deviation_limit_percentage = st.sidebar.slider("Допуст. лимит откл. ЧОК (%):", 1, 25, 5, 1, key="chok_limit_slider")
    materiality_method_key = st.sidebar.selectbox(
        "Метод расчета существенности:", options=["vs_CHOK", "vs_TotalComponents", "within_OA_CO"],
        format_func=lambda x: {"vs_CHOK": "vs_CHOK (рычаг влияния)", 
                               "vs_TotalComponents": "vs_TotalComponents (доля в общем объеме)",
                               "within_OA_CO": "within_OA_CO (структура ОА/КО)"}[x],
        key="materiality_method_selector"
    )
    
    active_value_columns = sorted(list(set(selected_fact_columns + ([forecast_actual_column_selected] if forecast_actual_column_selected else []))))
    active_value_columns = [col for col in active_value_columns if col in df_main_articles.columns]
    all_period_totals = calculate_period_totals(df_main_articles, active_value_columns)

    st.header("1. Данные по статьям и итоги ЧОК (тыс. руб.)")
    # ... (код отображения Раздела 1 с подробными пояснениями из предыдущего ответа) ...
    st.markdown("""
    Эта таблица отображает ваши исходные или загруженные данные по статьям Чистого Оборотного Капитала (ЧОК).
    - **Статья:** Наименование компонента.
    - **Тип:** "ОА" (Оборотные Активы) или "КО" (Краткосрочные Обязательства).
    - **Колонки с периодами:** Значения по статьям за соответствующие периоды.
    **Итоги по периодам:**
    - **Итого ОА:** Общая сумма всех оборотных активов.
    - **Итого КО:** Общая сумма всех краткосрочных обязательств.
    - **ЧОК (Чистый Оборотный Капитал):** Рассчитывается как `Итого ОА - Итого КО`. Это ключевой показатель, отражающий способность компании финансировать свою текущую операционную деятельность и выполнять краткосрочные обязательства. Положительный ЧОК обычно указывает на наличие достаточных краткосрочных активов для покрытия краткосрочных обязательств.
    """)
    df_main_for_export, df_totals_for_export = pd.DataFrame(), pd.DataFrame()
    if active_value_columns:
        common_formatters = {col: "{:.0f}" for col in active_value_columns}
        st.subheader("Детализация по статьям:")
        cols_to_show_main = ['Статья', 'Тип'] + active_value_columns
        df_main_for_export = df_main_articles[cols_to_show_main]
        st.dataframe(df_main_for_export.style.format(common_formatters))

        st.subheader("Итоги по периодам (ОА, КО, ЧОК):")
        totals_display_rows = [{'Показатель': name, **{col: all_period_totals.get(col, {}).get(name, np.nan) for col in active_value_columns}}
                               for name in ['Итого ОА', 'Итого КО', 'ЧОК']]
        df_totals_for_export = pd.DataFrame(totals_display_rows).set_index('Показатель')
        st.dataframe(df_totals_for_export.style.format(common_formatters))
    else:
        st.info("Выберите периоды для отображения в боковой панели.")

    st.header(f"2. Анализ существенности статей ЧОК")
    st.markdown("""
    **Для чего нужен анализ существенности, если есть анализ допустимых ошибок (Раздел 4)?**

    Анализ допустимых ошибок (Раздел 4) напрямую показывает, насколько точным должен быть прогноз каждой статьи, чтобы итоговый ЧОК оставался в заданных пределах. Это ваш **ключевой ориентир** для определения требуемой точности.

    Анализ существенности (этот Раздел 2) служит **важным дополнением**:
    1.  **Объясняет "почему":** Он помогает понять, *почему* для одних статей допустимый процент ошибки мал (они чувствительны), а для других — велик. Часто это связано с их "весом" или влиянием на ЧОК или его компоненты.
    2.  **Помогает расставить приоритеты:** Если статья имеет низкую существенность (например, по методу "vs_CHOK" она составляет всего 1-2% от ЧОК), то даже если ее относительная ошибка прогноза будет заметной, ее влияние на общий ЧОК в абсолютном выражении, скорее всего, будет незначительным. **В таких случаях вы можете принять решение не тратить чрезмерные ресурсы на "идеальный" прогноз именно этой статьи**, а сфокусироваться на тех, которые и существенны, и чувствительны (см. Раздел 4).
    3.  **Дает структурное понимание:** Разные методы существенности показывают статью с разных сторон (ее "рычаг" на ЧОК, долю в общем объеме, вес внутри группы ОА/КО), что дает более полную картину для принятия управленческих решений.

    Выберите метод ниже, чтобы детальнее изучить разные аспекты существенности ваших статей ЧОК.
    """)
    st.subheader(f"Метод: { {'vs_CHOK': 'vs_CHOK (рычаг влияния)', 'vs_TotalComponents': 'vs_TotalComponents (доля в общем объеме)', 'within_OA_CO': 'within_OA_CO (структура ОА/КО)'}[materiality_method_key]}")
    
    if materiality_method_key == "vs_CHOK":
        st.markdown("""*Формула: `(|Статья| / |ЧОК периода|) * 100%`. Показывает "рычаг" статьи на ЧОК. Сумма может быть >100%. Высокий % = статья сильно влияет на изменение ЧОК.*""")
    elif materiality_method_key == "vs_TotalComponents":
        st.markdown("""*Формула: `(|Статья| / Сумма_модулей_всех_статей_ОА_и_КО_периода|) * 100%`. Показывает долю статьи в общем объеме компонентов ЧОК. Сумма всех = 100%.*""")
    elif materiality_method_key == "within_OA_CO":
        st.markdown("""*Формула: `(|Статья ОА| / |Итого ОА|) * 100%` и `(|Статья КО| / |Итого КО|) * 100%`. Показывает структуру внутри ОА и КО. Суммы по ОА и КО = 100% соответственно.*""")

    df_materiality_calculated = pd.DataFrame()
    if selected_fact_columns:
        df_materiality_calculated = calculate_materiality(df_main_articles, all_period_totals, selected_fact_columns, materiality_method_key)
        if not df_materiality_calculated.empty:
            materiality_format = {col: "{:.2f}%" for col in df_materiality_calculated.columns if 'Сущ-ть' in col}
            st.dataframe(df_materiality_calculated.style.format(materiality_format))
    else:
        st.info("Выберите фактические периоды для расчета существенности.")

    st.header(f"3. Анализ точности прогноза")
    # ... (код Раздела 3 с подробными пояснениями из предыдущего ответа) ...
    st.markdown("""
    Эта секция помогает оценить точность вашего прогноза путем сравнения прогнозных значений с фактическими данными за выбранный базовый период.
    - **Абсолютное отклонение:** Разница между прогнозом и фактом (`Прогноз - Факт`). Показывает ошибку в денежном выражении.
    - **Относительное отклонение (%):** Абсолютное отклонение в процентах от факта (`(Абс. откл. / Факт) * 100%`). Показывает масштаб ошибки.
    Анализ этих отклонений помогает выявлять систематические ошибки и улучшать модели прогнозирования.
    """)
    df_article_deviations, df_summary_indicator_deviations = pd.DataFrame(), pd.DataFrame()
    if forecast_actual_column_selected and base_for_deviation_analysis and \
       forecast_actual_column_selected in df_main_articles.columns and \
       base_for_deviation_analysis in df_main_articles.columns:
        st.markdown(f"Сравнение прогноза **{forecast_actual_column_selected}** с фактом **{base_for_deviation_analysis}**")
        df_article_deviations, df_summary_indicator_deviations = calculate_forecast_deviations(
            df_main_articles, all_period_totals, forecast_actual_column_selected, base_for_deviation_analysis
        )
        st.subheader("Отклонения прогноза по статьям:")
        if not df_article_deviations.empty:
            dev_art_fmt = {c: "{:.0f}" for c in df_article_deviations.columns if 'Абс.' in c}
            dev_art_fmt.update({c: "{:.2f}%" for c in df_article_deviations.columns if 'Отн.' in c})
            st.dataframe(df_article_deviations.style.format(dev_art_fmt))
        st.subheader("Отклонения прогноза по итоговым показателям:")
        if not df_summary_indicator_deviations.empty:
            dev_sum_fmt = {'Прогноз': "{:.0f}", f'Факт ({base_for_deviation_analysis.split(" ")[0]})': "{:.0f}",
                           'Абс. откл.': "{:.0f}", 'Отн. откл. (%)': "{:.2f}%"}
            st.dataframe(df_summary_indicator_deviations.set_index('Показатель').style.format(dev_sum_fmt))
    else:
        st.info("Выберите корректный прогнозный и базовый фактический период для анализа отклонений.")

    st.header(f"4. Анализ чувствительности прогноза ЧОК к ошибкам в статьях")
    # ... (код Раздела 4 с подробными пояснениями из предыдущего ответа) ...
    df_allowed_article_errors = pd.DataFrame()
    if forecast_actual_column_selected and forecast_actual_column_selected in df_main_articles.columns:
        forecasted_chok_value = all_period_totals.get(forecast_actual_column_selected, {}).get('ЧОК', np.nan)
        df_allowed_article_errors, max_abs_chok_dev_value = calculate_allowed_article_error_range(
            df_main_articles, forecast_actual_column_selected, forecasted_chok_value, chok_deviation_limit_percentage
        )
        if not pd.isna(forecasted_chok_value):
             st.markdown(f"""
            Анализ для прогноза **{forecast_actual_column_selected}**. Прогнозный ЧОК = **{forecasted_chok_value:.0f} тыс. руб.**
            Макс. допуст. абс. отклонение для ЧОК при лимите в {chok_deviation_limit_percentage}%: **+/- {max_abs_chok_dev_value:.2f} тыс. руб.**
            """)
        st.markdown("""
        Этот анализ напрямую отвечает на вопрос: **"Какой процент ошибок я могу допустить при прогнозе той или иной статьи ЧОК, чтобы общий ЧОК не изменился больше заданного лимита?"**
        Он показывает, на сколько процентов может ошибиться прогноз по **одной конкретной статье** (при условии, что все остальные статьи спрогнозированы идеально точно),
        чтобы общее отклонение прогнозного ЧОК от его же первоначального прогнозного значения не превысило заданный вами лимит (сейчас **+/- {chok_deviation_limit_percentage}%**).
        
        **Интерпретация и как это использовать для улучшения прогнозов (совместно с Разделом 2 "Существенность"):**
        - **Малый процент в столбце "Макс. ошибка статьи (+/- %)":** Указывает на **высокую чувствительность** ЧОК к прогнозу этой статьи. Даже небольшая относительная ошибка в прогнозе этой статьи приведет к существенному отклонению итогового ЧОК. 
            - **Если такая статья еще и обладает высокой существенностью** (например, по методу "vs_CHOK" в Разделе 2), то она **требует максимальной точности прогнозирования**. Это ваш главный приоритет для улучшения.
        - **Большой процент (или "Любая..."):** Указывает на **низкую чувствительность**. Прогноз по этой статье может иметь большую относительную погрешность, прежде чем это значительно повлияет на общий ЧОК. 
            - **Если такая статья имеет низкую существенность**, вы можете **допустить для нее менее точный прогноз**, если ресурсы ограничены. "Любая (прогноз статьи=0)" означает, что прогнозное значение статьи равно нулю, и ее относительная ошибка (если она остается нулевой по абсолютной величине) не влияет на ЧОК.
        
        **Стратегии фокусировки усилий по прогнозированию:**
        1.  **Приоритет 1:** Статьи с **малым** допустимым процентом ошибки (высокочувствительные) **И** высокой существенностью (особенно "vs_CHOK"). Здесь точность критична.
        2.  **Приоритет 2:** Статьи с малым допустимым процентом ошибки, но средней/низкой существенностью. Они все еще важны из-за чувствительности.
        3.  **Меньший приоритет:** Статьи с большим допустимым процентом ошибки **И** низкой существенностью. Здесь можно допустить большую погрешность прогноза.
        4.  **Осторожно с "перекрытием" ошибок:** Пытаться компенсировать ошибку в важной статье за счет других – рискованно и сложно. Лучше стремиться к точности по ключевым статьям.
        """)
        if not df_allowed_article_errors.empty:
            allowed_err_fmt = {c: "{:.2f}%" for c in df_allowed_article_errors.columns if 'Макс. ошибка' in c}
            st.dataframe(df_allowed_article_errors.replace(np.inf, "Любая (прогноз статьи=0)").style.format(allowed_err_fmt))
    else:
        st.info("Выберите прогнозный период для анализа чувствительности.")


    st.header("5. Выгрузка данных в Excel")
    # ... (код Раздела 5 с подробными пояснениями из предыдущего ответа) ...
    st.markdown("Нажмите кнопку ниже, чтобы скачать все рассчитанные таблицы (на основе текущих настроек в боковой панели) в одном Excel файле. Каждая таблица будет размещена на отдельном листе.")
    dfs_for_export = {}
    if not df_main_for_export.empty: dfs_for_export["Данные_статьи"] = df_main_for_export
    if not df_totals_for_export.empty: dfs_for_export["Данные_итоги"] = df_totals_for_export
    if not df_materiality_calculated.empty: 
        sheet_name_materiality = f"Сущ_{materiality_method_key}"[:31] # ИСПРАВЛЕНО: сокращаем имя листа
        dfs_for_export[sheet_name_materiality] = df_materiality_calculated
    if not df_article_deviations.empty: dfs_for_export["Отклонения_статьи"] = df_article_deviations
    if not df_summary_indicator_deviations.empty: 
        dfs_for_export["Отклонения_итоги"] = df_summary_indicator_deviations.set_index('Показатель') if 'Показатель' in df_summary_indicator_deviations else df_summary_indicator_deviations
    if not df_allowed_article_errors.empty: dfs_for_export["Допустимые_ошибки"] = df_allowed_article_errors

    if dfs_for_export: 
        excel_bytes = dfs_to_excel_bytes(dfs_for_export)
        st.download_button(
           label="📥 Скачать все расчеты в Excel",
           data=excel_bytes,
           file_name=f"анализ_чок_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
           key="download_excel_button" 
        )
    else:
        st.info("Нет данных для формирования Excel файла. Выберите периоды и/или загрузите данные для проведения расчетов.")


if __name__ == '__main__':
    main()