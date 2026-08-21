import streamlit as st
import pandas as pd

st.set_page_config(page_title="Form Suket Lahir", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

EXCEL_FILE = "data/SUKET LAHIR.xlsx"

try:
    df = pd.read_excel(EXCEL_FILE, sheet_name="ISIAN DATA", header=None)

    st.subheader("Data Utama")
    no_surat = st.text_input("Nomor Surat", value=str(df.iloc[0, 0]) if pd.notna(df.iloc[0, 0]) else "")
    nama_kk = st.text_input("Nama Kepala Keluarga", value=str(df.iloc[4, 4]) if pd.notna(df.iloc[4, 4]) else "")
    no_kk = st.text_input("Nomor KK", value=str(df.iloc[5, 4]) if pd.notna(df.iloc[5, 4]) else "")
    nama_bayi = st.text_input("Nama Bayi", value=str(df.iloc[9, 4]) if pd.notna(df.iloc[9, 4]) else "")

    if st.button("Simpan"):
        st.success("Data berhasil dimuat!")
except Exception as e:
    st.error(f"File Excel belum ditemukan atau error: {e}")
