import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO
import requests
from openpyxl import load_workbook

st.set_page_config(page_title="Отчет ТТК", layout="wide")
st.title("📞 Отчет ТТК — Мэтчинг звонков и визитов (60 минут)")

st.markdown("""
**Инструкция:**
1. Загрузите файл **выгрузки из Метрики** — [ссылка на Метрику](https://metrika.yandex.ru/stat/visits)  
   Важно: при формировании отчёта выберите **период с начала месяца до последнего воскресенья**, скачайте файл в формате **XLSX**
2. Загрузите файл **звонков из CRM** — обычно называется *«Детальный отчет»*, его присылает КС в рабочий чат
3. Нажмите на кнопку ниже, чтобы получить файл со списком совпадений
""")

metrika_file = st.file_uploader("📊 Загрузите выгрузку из Метрики", type="xlsx")
calls_file = st.file_uploader("📞 Загрузите звонки из CRM", type="xlsx")

st.header("🧾 План-Факт: Ввод данных")

report_date_range = st.date_input("📅 Период отчета", value=(pd.to_datetime("today").replace(day=1), pd.to_datetime("today")), format="DD.MM.YYYY")

col1, col2, col3 = st.columns(3)
with col1:
    search_cost = st.number_input("💰 Расход Поиск (с НДС)", min_value=0.0, step=100.0)
with col2:
    search_impressions = st.number_input("👁 Показы Поиск", min_value=0, step=100)
with col3:
    search_clicks = st.number_input("🖱 Клики Поиск", min_value=0, step=1)

col4, col5 = st.columns(2)
with col4:
    search_conversions = st.number_input("📩 Заявки Поиск (по Метрике)", min_value=0, step=1)

col6, col7, col8 = st.columns(3)
with col6:
    rsya_cost = st.number_input("💰 Расход РСЯ (с НДС)", min_value=0.0, step=100.0)
with col7:
    rsya_impressions = st.number_input("👁 Показы РСЯ", min_value=0, step=100)
with col8:
    rsya_clicks = st.number_input("🖱 Клики РСЯ", min_value=0, step=1)

col9 = st.columns(1)[0]
with col9:
    rsya_conversions = st.number_input("📩 Заявки РСЯ (по Метрике)", min_value=0, step=1)

def normalize_region(s):
    return str(s).strip().lower().replace('г.', '').replace('-', '').replace('ё', 'е').replace(' ', '')

def process_visits(df):
    for i, row in df.iterrows():
        if str(row.iloc[0]).strip().lower().startswith('дата и время визита'):
            df.columns = row
            df = df.iloc[i+1:]
            break
    df = df.dropna(how='all')
    df = df[~df.iloc[:, 0].astype(str).str.contains('итого', case=False, na=False)]
    df.columns = df.columns.str.strip()
    df['visit_time'] = pd.to_datetime(df['Дата и время визита'], errors='coerce')
    df['region'] = df['Город'].apply(normalize_region)
    df = df.dropna(subset=['visit_time', 'region'])
    df['visit_end'] = df['visit_time'] + timedelta(minutes=60)
    return df

def process_calls(df):
    df.columns = df.columns.str.strip()
    df['call_time'] = pd.to_datetime(df['Дата'].astype(str) + ' ' + df['Время'].astype(str), errors='coerce')
    df['region'] = df['Город'].apply(normalize_region)
    if '№ тел.' in df.columns:
        df = df.rename(columns={'№ тел.': 'Телефон'})
    df = df.dropna(subset=['call_time', 'region'])
    return df

def match_data(calls, visits):
    merged = pd.merge(
        calls[['call_time', 'region']],
        visits[['visit_time', 'visit_end', 'region']],
        on='region',
        how='inner'
    )
    merged = merged[
        (merged['call_time'] >= merged['visit_time']) &
        (merged['call_time'] <= merged['visit_end'])
    ].copy()

    merged['time_diff'] = (merged['call_time'] - merged['visit_time']).abs()
    merged = merged.sort_values('time_diff').drop_duplicates(subset=['visit_time', 'region'])
    merged = merged.drop(columns=['time_diff'])
    merged['Call DateTime'] = merged['call_time']
    return merged[['call_time', 'visit_time', 'region']].rename(columns={
        'call_time': 'Время звонка',
        'visit_time': 'Время визита'
    }).drop_duplicates()

result_df = pd.DataFrame()

if not result_df.empty:
    st.success(f"✅ Найдено совпадений: {len(result_df)}")

if not metrika_file or not calls_file:
    st.info("ℹ️ Отчет будет без совпадений — загрузите оба файла для мэтчинга.")
else:
    st.info("ℹ️ Отчет ожидает генерации — нажмите кнопку ниже.")

if st.button("🚀 Сгенерировать отчет"):
    with st.spinner("🔄 Обрабатываем данные..."):
        try:
            if metrika_file and calls_file:
                visits_raw = pd.read_excel(metrika_file, header=None)
                visits_df = process_visits(visits_raw)
                calls_df = pd.read_excel(calls_file)
                calls_df = process_calls(calls_df)
                result_df = match_data(calls_df, visits_df)
            else:
                result_df = pd.DataFrame()  # пустая таблица совпадений

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:

                # Добавляем шаблон "План-Факт" из GitHub как лист
                try:
                    template_url = "https://github.com/Exepicode/ttk/raw/main/ТТК-шаблон-отчета.xlsx"
                    response = requests.get(template_url)
                    if response.status_code == 200:
                        template_excel = BytesIO(response.content)
                        wb_template = load_workbook(template_excel, data_only=False)

                        if wb_template.sheetnames:
                            source_ws = wb_template.worksheets[0]
                            source_ws.title = "План-Факт"
                            target_ws = writer.book.create_sheet("План-Факт")

                            for row in source_ws.iter_rows():
                                if all(cell.value in (None, "") for cell in row):
                                    continue
                                for cell in row:
                                    new_cell = target_ws.cell(row=cell.row, column=cell.column)
                                    if cell.data_type == 'f' or (isinstance(cell.value, str) and cell.value.startswith("=")):
                                        new_cell.value = cell.value
                                    else:
                                        new_cell.value = cell.value
                                    try:
                                        if cell.has_style:
                                            new_cell.font = cell.font.copy()
                                            new_cell.border = cell.border.copy()
                                            new_cell.fill = cell.fill.copy()
                                            new_cell.number_format = cell.number_format
                                            new_cell.protection = cell.protection.copy()
                                            new_cell.alignment = cell.alignment.copy()
                                    except Exception as style_error:
                                        st.warning(f"⚠️ Стиль не скопирован для ячейки {cell.coordinate}: {style_error}")

                            for col_letter, dim in source_ws.column_dimensions.items():
                                target_ws.column_dimensions[col_letter].width = dim.width
                            for row_idx, dim in source_ws.row_dimensions.items():
                                target_ws.row_dimensions[row_idx].height = dim.height
                            target_ws.row_dimensions[7].height = 1
                            target_ws.column_dimensions['F'].width = 15
                        else:
                            st.warning("⚠️ В шаблоне отсутствуют листы")
                    else:
                        st.warning(f"⚠️ Не удалось скачать шаблон: статус {response.status_code}")
                except Exception as e:
                    st.warning(f"⚠️ Ошибка при вставке шаблона 'План-Факт': {e}")

                try:
                    plan_fact_ws = writer.book["План-Факт"]
                    plan_fact_ws["D4"] = f"{report_date_range[0].strftime('%d.%m.%Y')} – {report_date_range[1].strftime('%d.%m.%Y')}"
                    plan_fact_ws["F8"] = search_cost
                    plan_fact_ws["F9"] = rsya_cost
                    plan_fact_ws["H8"] = search_impressions
                    plan_fact_ws["H9"] = rsya_impressions
                    plan_fact_ws["J8"] = search_clicks
                    plan_fact_ws["J9"] = rsya_clicks
                    plan_fact_ws["P8"] = search_conversions
                    plan_fact_ws["P9"] = rsya_conversions
                    plan_fact_ws["Q8"] = len(result_df)
                except Exception as e:
                    st.warning(f"⚠️ Не удалось записать данные в 'План-Факт': {e}")

                result_df.to_excel(writer, sheet_name="Совпадения", index=False)
                # Устанавливаем ширину столбцов на листе "Совпадения"
                worksheet = writer.sheets["Совпадения"]
                for column_cells in worksheet.columns:
                    max_length = 20
                    col_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max_length
                if metrika_file:
                    visits_raw.to_excel(writer, sheet_name="Метрика", index=False, header=False)
                if calls_file:
                    pd.read_excel(calls_file).to_excel(writer, sheet_name="Звонки", index=False)

            st.download_button(
                label="📥 Скачать отчет (XLSX)",
                data=output.getvalue(),
                file_name="Отчет_ТТК.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Ошибка обработки: {e}")

if not result_df.empty:
    st.success(f"✅ Найдено совпадений: {len(result_df)}")
