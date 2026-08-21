import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="Form Suket Lahir", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

# Fungsi untuk mencari file Excel secara otomatis (tidak sensitif huruf besar/kecil)
def find_excel_file():
    # Cari di folder data/ atau di root directory
    possible_paths = [
        "data/*.xlsx",
        "data/*.xls",
        "*.xlsx",
        "*.xls"
    ]
    for path in possible_paths:
        files = glob.glob(path)
        for f in files:
            if "suket" in f.lower() or "lahir" in f.lower():
                return f
    return None

excel_file = find_excel_file()

if not excel_file:
    st.error("⚠️ File Excel 'SUKET LAHIR.xlsx' tidak ditemukan di folder 'data/'. Pastikan file sudah di-upload ke GitHub!")
else:
    try:
        df = pd.read_excel(excel_file, sheet_name="ISIAN DATA", header=None)
        
        st.success(f"File berhasil dimuat dari: `{excel_file}`")
        st.subheader("Data Utama")
        
        no_surat = st.text_input("Nomor Surat", value=str(df.iloc[0, 0]) if pd.notna(df.iloc[0, 0]) else "")
        nama_kk = st.text_input("Nama Kepala Keluarga", value=str(df.iloc[4, 4]) if pd.notna(df.iloc[4, 4]) else "")
        no_kk = st.text_input("Nomor KK", value=str(df.iloc[5, 4]) if pd.notna(df.iloc[5, 4]) else "")
        nama_bayi = st.text_input("Nama Bayi", value=str(df.iloc[9, 4]) if pd.notna(df.iloc[9, 4]) else "")
        
        if st.button("Simpan"):
            st.info("Data berhasil dibaca.")
            
    except Exception as e:
        st.error(f"Error membaca isi sheet Excel: {e}")
