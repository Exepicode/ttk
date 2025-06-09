import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

st.set_page_config(page_title="Отчет ТТК", layout="wide")
st.title("📞 Отчет ТТК — Мэтчинг звонков и визитов (60 минут)")

st.markdown("""
**Инструкция:**
1. Загрузите файл **выгрузки из Метрики** (где есть дата и время визита и город)
2. Загрузите файл **звонков из CRM** (где есть дата, время и город звонка)
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
    df = df.dropna(subset=['call_time', 'region'])
    return df

def match_data(calls, visits):
    merged = pd.merge(calls, visits[['visit_time', 'visit_end', 'region']], on='region', how='inner')
    merged = merged[
        (merged['call_time'] >= merged['visit_time']) &
        (merged['call_time'] <= merged['visit_end'])
    ].copy()
    merged['Call DateTime'] = merged['call_time']
    columns = ['Call DateTime', 'visit_time', 'region']
    if 'Телефон' in merged.columns:
        columns.append('Телефон')
    return merged[columns].drop_duplicates()

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
                visits_raw.to_excel(writer, sheet_name="Метрика", index=False, header=False)
                calls_df.to_excel(writer, sheet_name="Звонки", index=False)

            st.success(f"✅ Найдено совпадений: {len(result_df)}")
            st.download_button("📥 Скачать Отчет ТТК", data=output.getvalue(), file_name="Отчет_ТТК.xlsx")
        except Exception as e:
            st.error(f"❌ Ошибка обработки: {e}")
