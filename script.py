import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

st.set_page_config(page_title="Мэтчинг Звонков", layout="wide")
st.title("📞 Мэтчинг звонков и Метрики (60 минут)")

st.markdown("""
**Инструкция:**
1. Загрузите файл **Метрика.xlsx** (в таблице данные с 8 строки)
2. Загрузите файл **Звонки.xlsx**
3. Нажмите на кнопку ниже — и получите Excel с совпадениями.
""")

metrika_file = st.file_uploader("📊 Загрузите файл Метрики", type="xlsx")
calls_file = st.file_uploader("📞 Загрузите файл Звонков", type="xlsx")

def normalize_region(s):
    return str(s).strip().lower().replace('г.', '').replace('-', '').replace('ё', 'е').replace(' ', '')

def process_visits(df):
    # Найдём заголовок
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
    df = df.dropna(subset=['call_time', 'region'])
    return df

def match_data(calls, visits):
    merged = pd.merge(calls, visits[['visit_time', 'visit_end', 'region']], on='region', how='inner')
    merged = merged[
        (merged['call_time'] >= merged['visit_time']) &
        (merged['call_time'] <= merged['visit_end'])
    ].copy()
    merged['Call Time'] = merged['call_time'].dt.time
    merged['Call Date'] = merged['call_time'].dt.date
    final = merged.groupby('call_time').first().reset_index()
    return final[['Call Time', 'Call Date', 'region', 'visit_time']]

if metrika_file and calls_file:
    with st.spinner("🔄 Обрабатываем..."):
        try:
            visits_raw = pd.read_excel(metrika_file, header=None)
            visits_df = process_visits(visits_raw)
            calls_df = pd.read_excel(calls_file)
            calls_df = process_calls(calls_df)
            result_df = match_data(calls_df, visits_df)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name="Совпадения", index=False)
                visits_raw.to_excel(writer, sheet_name="Метрика", index=False, header=False)
                calls_df.to_excel(writer, sheet_name="Звонки", index=False)

            st.success(f"✅ Найдено совпадений: {len(result_df)}")
            st.download_button("📥 Скачать Excel", data=output.getvalue(), file_name="Результат_мэтчинга.xlsx")
        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
