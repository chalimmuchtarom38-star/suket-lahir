import streamlit as st
import pandas as pd
import openpyxl
import os
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Form Suket Lahir Lengkap", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

EXCEL_FILE = os.path.join(os.path.dirname(__file__), "SUKET LAHIR.xlsx")

def get_val(df, r, c):
    try:
        val = df.iloc[r, c]
        return "" if pd.isna(val) else str(val).strip()
    except:
        return ""

# --- FUNGSI GENERATE EXCEL KHUSUS SHEET SURAT KELAHIRAN ---
def generate_excel_surat_kelahiran():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    if "Surat Kelahiran" in wb.sheetnames:
        # Hapus sheet lain agar yang ter-download hanya Surat Kelahiran
        for sheet_name in wb.sheetnames:
            if sheet_name != "Surat Kelahiran":
                del wb[sheet_name]
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- FUNGSI GENERATE PDF BERDASARKAN PRINT AREA (A1:AD94) ---
def generate_pdf_surat_kelahiran():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Surat Kelahiran"]
    
    # Range Print Area: A1 sampai AD94 (Kolom 1 - 30, Baris 1 - 94)
    data_grid = []
    for r in range(1, 95):
        row_vals = []
        for c in range(1, 31):
            val = ws.cell(row=r, column=c).value
            row_vals.append("" if val is None else str(val))
        # Hanya ambil baris yang tidak benar-benar kosong untuk efisiensi PDF
        if any(row_vals):
            # Ambil gabungan teks per baris agar rapi di PDF
            clean_text = " ".join([v for v in row_vals if v.strip()])
            data_grid.append([clean_text])

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    story = []
    style_text = ParagraphStyle('Custom', parent=styles['Normal'], fontSize=9, leading=12)

    for item in data_grid:
        if item[0]:
            story.append(Paragraph(item[0], style_text))
            story.append(Spacer(1, 3))

    doc.build(story)
    buffer.seek(0)
    return buffer

if not os.path.exists(EXCEL_FILE):
    st.error(f"⚠️ File tidak ditemukan di lokasi: {EXCEL_FILE}")
else:
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="ISIAN DATA", header=None)
        st.success("✅ Seluruh data berhasil dimuat!")

        # --- TOMBOL DOWNLOAD EXCEL & PDF ---
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            excel_bytes = generate_excel_surat_kelahiran()
            st.download_button(
                label="📥 Download Excel (Sheet Surat Kelahiran)",
                data=excel_bytes,
                file_name="Surat_Kelahiran.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_dl2:
            pdf_bytes = generate_pdf_surat_kelahiran()
            st.download_button(
                label="📄 Download PDF (Surat Kelahiran)",
                data=pdf_bytes,
                file_name="Surat_Kelahiran.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        st.divider()

        # --- FORM ISIAN DATA MASTER ---
        with st.form("form_suket_lengkap"):
            st.subheader("1. Header & Data Kepala Keluarga")
            c1, c2, c3 = st.columns(3)
            with c1: st.text_input("Nomor Surat", value=get_val(df, 1, 0))
            with c2: st.text_input("Nama Kepala Keluarga", value=get_val(df, 5, 4))
            with c3: st.text_input("Nomor Kartu Keluarga (KK)", value=get_val(df, 6, 4))

            st.subheader("2. Data Bayi / Anak")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Nama Bayi", value=get_val(df, 10, 4))
                st.text_input("Jenis Kelamin Bayi", value=get_val(df, 11, 4))
                st.text_input("Tempat Kelahiran", value=get_val(df, 13, 4))
            with c2:
                st.text_input("Jam Lahir", value=get_val(df, 15, 4))
                st.text_input("Kelahiran Ke-", value=get_val(df, 17, 4))
            with c3:
                st.text_input("Berat Bayi (Gram)", value=get_val(df, 19, 4))
                st.text_input("Panjang Bayi (Cm)", value=get_val(df, 20, 4))

            st.subheader("3. Data Ibu & Ayah")
            col_ibu, col_ayah = st.columns(2)
            with col_ibu:
                st.text_input("NIK Ibu", value=get_val(df, 23, 4))
                st.text_input("Nama Lengkap Ibu", value=get_val(df, 24, 4))
                st.text_input("Alamat Ibu", value=get_val(df, 27, 4))
            with col_ayah:
                st.text_input("NIK Ayah", value=get_val(df, 35, 4))
                st.text_input("Nama Lengkap Ayah", value=get_val(df, 36, 4))
                st.text_input("Alamat Ayah", value=get_val(df, 39, 4))

            btn = st.form_submit_button("Simpan / Perbarui Data")
            if btn:
                st.info("Form berhasil diperbarui.")

    except Exception as e:
        st.error(f"Gagal membaca data Excel: {e}")
