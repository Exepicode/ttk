import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

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
                result_df.to_excel(writer, sheet_name="Совпадения", index=False)
                # Устанавливаем ширину столбцов на листе "Совпадения"
                worksheet = writer.sheets["Совпадения"]
                for column_cells in worksheet.columns:
                    max_length = 20
                    col_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max_length
                visits_raw.to_excel(writer, sheet_name="Метрика", index=False, header=False)
                pd.read_excel(calls_file).to_excel(writer, sheet_name="Звонки", index=False)

                # Кастомное создание листа "План-Факт"
                import openpyxl
                from openpyxl.styles import Font, Alignment

                # Список заголовков
                plan_fact_headers = [
                    "#", "Сайт", "Место размещения на сайте и таргетинги", "Расход план", "Расход факт",
                    "Показы план", "Показы факт", "Клики план", "Клики факт", "CTR % план", "CTR % факт",
                    "CPC план", "CPC факт", "Заявки план", "Заявки метрика", "Звонки", "Все заявки",
                    "CR % факт", "CPL план", "CPL факт"
                ]
                # Примерные данные (2 строки)
                plan_fact_data = [
                    [
                        1, "yandex.ru", "Главная, Москва", "100 000 ₽", "95 000 ₽",
                        "1 000 000", "950 000", "10 000", "9 500", "1,00%", "1,00%",
                        "10,00 ₽", "10,00 ₽", "500", "480", "50", "530", "5,6%", "200,00 ₽", "198,00 ₽"
                    ],
                    [
                        2, "vk.com", "Лента, Санкт-Петербург", "50 000 ₽", "48 000 ₽",
                        "500 000", "480 000", "5 000", "4 800", "1,00%", "1,00%",
                        "10,00 ₽", "10,00 ₽", "250", "240", "20", "260", "5,4%", "200,00 ₽", "192,00 ₽"
                    ]
                ]
                # Итоговая строка (суммы, выделить жирным)
                plan_fact_totals = [
                    "Итого", "", "",
                    "150 000 ₽", "143 000 ₽",
                    "1 500 000", "1 430 000",
                    "15 000", "14 300",
                    "1,00%", "1,00%",
                    "10,00 ₽", "10,00 ₽",
                    "750", "720", "70", "790",
                    "5,5%", "200,00 ₽", "190,00 ₽"
                ]

                # Сначала создаем DataFrame для данных и заголовков (без верхних 3 строк)
                plan_fact_df = pd.DataFrame(plan_fact_data, columns=plan_fact_headers)

                # Записываем в Excel начиная с строки 5 (то есть startrow=4)
                plan_fact_df.to_excel(writer, sheet_name="План-Факт", index=False, startrow=4)
                # Теперь открываем рабочий лист openpyxl
                ws = writer.sheets["План-Факт"]

                # Верхние ячейки
                ws["A1"] = "Клиент"
                ws["B1"] = "ТТК-Связь"
                ws["A2"] = "Продукт/Кампания"
                ws["B2"] = "Услуги домашнего интернета и телевидения."
                ws["A3"] = "Период кампании"
                ws["B3"] = "01.05-26.05"

                # Шапка (выравнивание, жирный)
                header_font = Font(bold=True)
                for col_idx in range(1, len(plan_fact_headers) + 1):
                    cell = ws.cell(row=5, column=col_idx)
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                # Итоговая строка (жирный, выравнивание)
                total_row_idx = 5 + len(plan_fact_data) + 1  # 5 (шапка) + 2 (данные) + 1 = 8
                for col_idx in range(1, len(plan_fact_totals) + 1):
                    cell = ws.cell(row=total_row_idx, column=col_idx)
                    cell.value = plan_fact_totals[col_idx - 1]
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Выравнивание всех данных
                for row in ws.iter_rows(min_row=6, max_row=total_row_idx, min_col=1, max_col=len(plan_fact_headers)):
                    for cell in row:
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                # Форматирование ширины столбцов для читаемости
                for col_idx, header in enumerate(plan_fact_headers, 1):
                    width = 20
                    if header in ("#", "Сайт"):
                        width = 12
                    elif header in ("Место размещения на сайте и таргетинги",):
                        width = 32
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

            st.success(f"✅ Найдено совпадений: {len(result_df)}")
            st.download_button("📥 Скачать Отчет ТТК", data=output.getvalue(), file_name="Отчет_ТТК.xlsx")
        except Exception as e:
            st.error(f"❌ Ошибка обработки: {e}")
