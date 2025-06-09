import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Мэтчинг Звонков и Метрики", layout="wide")
st.title("\U0001F4F1 Мэтчинг Звонков с Метрикой (Окно 60 минут)")

st.markdown("""
**Инструкция:**
1. Загрузите файл **Метрика.xlsx** (начиная с 8-й строки)
2. Загрузите файл **Звонки.xlsx**
3. Нажмите кнопку ниже для сопоставления звонков с визитами.
4. Скачайте Excel с результатами.
""")

metrika_file = st.file_uploader("Загрузите файл Метрики (XLSX)", type=["xlsx"], key="metrika")
calls_file = st.file_uploader("Загрузите файл Звонков (XLSX)", type=["xlsx"], key="calls")

def match_data(visits_df, calls_df):
    # Преобразование звонков
    calls_df['call_datetime'] = pd.to_datetime(calls_df.iloc[:, 0], errors='coerce')
    calls_df['call_time'] = calls_df['call_datetime'].dt.time
    calls_df['call_date'] = calls_df['call_datetime'].dt.date
    calls_df['region'] = calls_df.iloc[:, 1].astype(str).str.lower().str.strip()

    # Преобразование визитов
    visits_df = visits_df.dropna(subset=[visits_df.columns[0]])
    visits_df['visit_datetime'] = pd.to_datetime(visits_df.iloc[:, 0], errors='coerce')
    visits_df['region'] = visits_df.iloc[:, 1].astype(str).str.lower().str.strip()

    # Мэтчинг
    matched = []
    for _, visit in visits_df.iterrows():
        for _, call in calls_df.iterrows():
            if (
                call['region'] in visit['region']
                and 0 <= (call['call_datetime'] - visit['visit_datetime']).total_seconds() <= 3600
            ):
                matched.append({
                    'Call Time': call['call_time'],
                    'Call Date': call['call_date'],
                    'Region': call['region'],
                    'Visit Time': visit['visit_datetime'],
                    'Visit Region': visit['region']
                })
    return pd.DataFrame(matched)

if metrika_file and calls_file:
    with st.spinner("Загружаем и сопоставляем данные..."):
        try:
            visits = pd.read_excel(metrika_file, skiprows=7)
            calls = pd.read_excel(calls_file)
            result_df = match_data(visits, calls)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Совпадения')
                visits.to_excel(writer, index=False, sheet_name='Метрика')
                calls.to_excel(writer, index=False, sheet_name='Звонки')

            st.success(f"Найдено совпадений: {len(result_df)}")
            st.download_button("Скачать результат (XLSX)", data=output.getvalue(), file_name="Результат_Мэтчинга.xlsx")
        except Exception as e:
            st.error(f"Ошибка обработки: {e}")
