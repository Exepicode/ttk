import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Отчет ТТК", layout="wide")
st.title("📞 Мэтчинг звонков с визитами (60 минут)")

st.markdown("""
**Инструкция:**
1. Загрузите файл **Метрики** (XLSX, начиная с 8-й строки)
2. Загрузите файл **Звонков** (XLSX)
3. Нажмите кнопку — получите файл с совпадениями
""")

metrika_file = st.file_uploader("Загрузите файл Метрики", type=["xlsx"], key="metrika")
calls_file = st.file_uploader("Загрузите файл Звонков", type=["xlsx"], key="calls")

def normalize_region(region):
    return (
        str(region).strip().lower()
        .replace("г.", "")
        .replace("-", "")
        .replace("ё", "е")
        .replace(" ", "")
    )

def match_data(visits_df, calls_df):
    # Обработка звонков
    calls_df.columns = calls_df.columns.str.strip().str.replace('\ufeff', '')
    calls_df['call_datetime'] = pd.to_datetime(
        calls_df.iloc[:, 0].astype(str) + ' ' + calls_df.iloc[:, 1].astype(str), errors='coerce'
    )
    calls_df['region'] = calls_df.iloc[:, 2].astype(str).apply(normalize_region)
    calls_df['phone'] = calls_df.iloc[:, 3].astype(str)
    calls_df = calls_df.dropna(subset=['call_datetime', 'region'])

    # Обработка визитов
    for i, row in visits_df.iterrows():
        if str(row.iloc[0]).strip().lower().startswith('дата и время визита'):
            visits_df.columns = row
            visits_df = visits_df.iloc[i+1:]
            break

    visits_df = visits_df.dropna(how='all')
    visits_df = visits_df[~visits_df.iloc[:, 0].astype(str).str.contains('итого', case=False, na=False)]

    visits_df['visit_datetime'] = pd.to_datetime(visits_df.iloc[:, 0], errors='coerce')
    visits_df['region'] = visits_df.iloc[:, 1].astype(str).apply(normalize_region)
    visits_df = visits_df.dropna(subset=['visit_datetime', 'region'])

    # Сопоставление
    visits_df['visit_end'] = visits_df['visit_datetime'] + timedelta(minutes=60)

    merged = pd.merge(
        calls_df,
        visits_df[['visit_datetime', 'visit_end', 'region']],
        on='region',
        how='inner'
    )

    matches = merged[
        (merged['call_datetime'] >= merged['visit_datetime']) &
        (merged['call_datetime'] <= merged['visit_end'])
    ].copy()

    matches = matches.groupby('call_datetime').first().reset_index()

    result = matches[['call_datetime', 'visit_datetime', 'region', 'phone']]
    result.columns = ['Время звонка', 'Время визита', 'Регион', 'Телефон']

    return result

if metrika_file and calls_file:
    with st.spinner("Обрабатываем файлы..."):
        try:
            visits_raw = pd.read_excel(metrika_file, header=None)
            calls_df = pd.read_excel(calls_file, header=None)

            result_df = match_data(visits_raw, calls_df)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name="Совпадения")
                pd.read_excel(metrika_file, skiprows=7).to_excel(writer, index=False, sheet_name="Метрика")
                calls_df.to_excel(writer, index=False, sheet_name="Звонки")

            st.success(f"Найдено совпадений: {len(result_df)}")
            st.download_button("📥 Скачать результат", data=output.getvalue(), file_name="Отчет_ТТК.xlsx")
        except Exception as e:
            st.error(f"Ошибка обработки: {e}")
