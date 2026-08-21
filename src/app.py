import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Form Suket Lahir", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

# Menunjuk langsung ke file Excel yang ada di dalam folder src
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "SUKET LAHIR.xlsx")

if not os.path.exists(EXCEL_FILE):
    st.error(f"⚠️ File tidak ditemukan di lokasi: {EXCEL_FILE}")
else:
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="ISIAN DATA", header=None)
        
        st.success("Berhasil membaca data dari file Excel!")
        st.subheader("Data Utama")
        
        no_surat = st.text_input("Nomor Surat", value=str(df.iloc[0, 0]) if pd.notna(df.iloc[0, 0]) else "")
        nama_kk = st.text_input("Nama Kepala Keluarga", value=str(df.iloc[4, 4]) if pd.notna(df.iloc[4, 4]) else "")
        no_kk = st.text_input("Nomor KK", value=str(df.iloc[5, 4]) if pd.notna(df.iloc[5, 4]) else "")
        nama_bayi = st.text_input("Nama Bayi", value=str(df.iloc[9, 4]) if pd.notna(df.iloc[9, 4]) else "")
        
        if st.button("Simpan"):
            st.info("Data berhasil dibaca.")
            
    except Exception as e:
        st.error(f"Error membaca isi sheet Excel: {e}")
