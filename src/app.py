import streamlit as st
import pandas as pd
import openpyxl
import os
import io
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Form Suket Lahir", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

EXCEL_FILE = os.path.join(os.path.dirname(__file__), "SUKET LAHIR.xlsx")
TEMPLATE_PDF = os.path.join(os.path.dirname(__file__), "template_f201.pdf")

@st.cache_data
def load_excel_data(file_path):
    return pd.read_excel(file_path, sheet_name="ISIAN DATA", header=None)

def get_val(df, r, c):
    try:
        val = df.iloc[r, c]
        return "" if pd.isna(val) else str(val).strip()
    except:
        return ""

# --- FUNGSI UPDATE DATA EXCEL KETIKA FORM DIISI ---
def update_excel_file(form_data):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    
    # 1. Update Sheet ISIAN DATA
    if "ISIAN DATA" in wb.sheetnames:
        ws_isian = wb["ISIAN DATA"]
        ws_isian.cell(row=2, column=1, value=form_data["no_surat"])       # Row 2 (index 1)
        ws_isian.cell(row=6, column=5, value=form_data["nama_kk"])        # Row 6 (index 5)
        ws_isian.cell(row=7, column=5, value=form_data["no_kk"])          # Row 7 (index 6)
        ws_isian.cell(row=11, column=5, value=form_data["nama_bayi"])     # Row 11 (index 10)
        ws_isian.cell(row=12, column=5, value=form_data["jk_bayi"])       # Row 12 (index 11)
        ws_isian.cell(row=25, column=5, value=form_data["nik_ibu"])       # Row 25 (index 24)
        ws_isian.cell(row=37, column=5, value=form_data["nik_ayah"])      # Row 37 (index 36)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- FUNGSI CETAK PDF SESUAI INPUT FORM ---
def generate_pdf_from_inputs(form_data):
    if not os.path.exists(TEMPLATE_PDF):
        return None

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 8)

    # Menuliskan data hasil inputan ke PDF Template
    can.drawString(380, 772, form_data["no_surat"])
    can.drawString(150, 745, form_data["nama_kk"])
    can.drawString(150, 730, form_data["no_kk"])
    can.drawString(150, 700, form_data["nama_bayi"])
    can.drawString(150, 685, form_data["jk_bayi"])
    can.drawString(150, 560, form_data["nik_ibu"])
    can.drawString(150, 430, form_data["nik_ayah"])

    can.save()
    packet.seek(0)

    new_pdf = PdfReader(packet)
    existing_pdf = PdfReader(TEMPLATE_PDF)
    output = PdfWriter()

    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    pdf_out = io.BytesIO()
    output.write(pdf_out)
    pdf_out.seek(0)
    return pdf_out

if not os.path.exists(EXCEL_FILE):
    st.error(f"⚠️ File Excel tidak ditemukan: {EXCEL_FILE}")
else:
    try:
        df = load_excel_data(EXCEL_FILE)

        # FORM INPUT DATA
        st.subheader("📝 Input / Edit Data Form")
        with st.form("form_suket_lengkap"):
            c1, c2, c3 = st.columns(3)
            with c1: 
                no_surat = st.text_input("Nomor Surat", value=get_val(df, 1, 0))
            with c2: 
                nama_kk = st.text_input("Nama Kepala Keluarga", value=get_val(df, 5, 4))
            with c3: 
                no_kk = st.text_input("Nomor Kartu Keluarga (KK)", value=get_val(df, 6, 4))

            c1, c2 = st.columns(2)
            with c1: 
                nama_bayi = st.text_input("Nama Bayi", value=get_val(df, 10, 4))
            with c2: 
                jk_bayi = st.text_input("Jenis Kelamin Bayi", value=get_val(df, 11, 4))

            c1, c2 = st.columns(2)
            with c1: 
                nik_ibu = st.text_input("NIK Ibu", value=get_val(df, 23, 4))
            with c2: 
                nik_ayah = st.text_input("NIK Ayah", value=get_val(df, 35, 4))

            submitted = st.form_submit_button("🔄 Perbarui Data & File")

        # Struktur data dari inputan pengguna
        form_data = {
            "no_surat": no_surat,
            "nama_kk": nama_kk,
            "no_kk": no_kk,
            "nama_bayi": nama_bayi,
            "jk_bayi": jk_bayi,
            "nik_ibu": nik_ibu,
            "nik_ayah": nik_ayah
        }

        # Mengatur Nama File Hasil Unduhan
        nama_bayi_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', nama_bayi).strip('_') if nama_bayi else "BAYI"
        nama_file_excel = f"SUKET_LAHIR_{nama_bayi_clean}.xlsx"
        nama_file_pdf = f"Surat_Kelahiran_{nama_bayi_clean}.pdf"

        st.divider()
        st.subheader("📥 Unduh File Terbaru")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label=f"📥 Download Excel Update ({nama_file_excel})",
                data=update_excel_file(form_data),
                file_name=nama_file_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_dl2:
            pdf_data = generate_pdf_from_inputs(form_data)
            if pdf_data:
                st.download_button(
                    label=f"📄 Download PDF Presisi ({nama_file_pdf})",
                    data=pdf_data,
                    file_name=nama_file_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ Upload file `template_f201.pdf` untuk mengaktifkan fungsi PDF.")

    except Exception as e:
        st.error(f"Error: {e}")
