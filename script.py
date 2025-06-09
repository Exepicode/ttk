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

if metrika_file and calls_file:
    with st.spinner("🔄 Обрабатываем данные..."):
        try:
            visits_raw = pd.read_excel(metrika_file, header=None)
            visits_df = process_visits(visits_raw)
            calls_df = pd.read_excel(calls_file)
            calls_df = process_calls(calls_df)
            result_df = match_data(calls_df, visits_df)

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
                                    if cell.data_type == 'f':  # если формула
                                        new_cell.value = f"={cell.value}"
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
                        else:
                            st.warning("⚠️ В шаблоне отсутствуют листы")
                    else:
                        st.warning(f"⚠️ Не удалось скачать шаблон: статус {response.status_code}")
                except Exception as e:
                    st.warning(f"⚠️ Ошибка при вставке шаблона 'План-Факт': {e}")

                result_df.to_excel(writer, sheet_name="Совпадения", index=False)
                # Устанавливаем ширину столбцов на листе "Совпадения"
                worksheet = writer.sheets["Совпадения"]
                for column_cells in worksheet.columns:
                    max_length = 20
                    col_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max_length
                visits_raw.to_excel(writer, sheet_name="Метрика", index=False, header=False)
                pd.read_excel(calls_file).to_excel(writer, sheet_name="Звонки", index=False)

            st.success(f"✅ Найдено совпадений: {len(result_df)}")
            st.download_button("📥 Скачать Отчет ТТК", data=output.getvalue(), file_name="Отчет_ТТК.xlsx")
        except Exception as e:
            st.error(f"❌ Ошибка обработки: {e}")
