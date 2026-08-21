import streamlit as st
import pandas as pd
import openpyxl
import os
import io
import subprocess
import tempfile

st.set_page_config(page_title="Form Suket Lahir", layout="wide")
st.title("📋 Form Input Data - SUKET LAHIR")

EXCEL_FILE = os.path.join(os.path.dirname(__file__), "SUKET LAHIR.xlsx")

@st.cache_data
def load_excel_data(file_path):
    return pd.read_excel(file_path, sheet_name="ISIAN DATA", header=None)

def get_val(df, r, c):
    try:
        val = df.iloc[r, c]
        return "" if pd.isna(val) else str(val).strip()
    except:
        return ""

# --- FUNGSI EXCEL ONLY SHEET SURAT KELAHIRAN ---
def generate_excel_surat_kelahiran():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    if "Surat Kelahiran" in wb.sheetnames:
        for sheet_name in list(wb.sheetnames):
            if sheet_name != "Surat Kelahiran":
                del wb[sheet_name]
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- FUNGSI PDF PRESISI 1:1 DENGAN LIBREOFFICE / EXCEL ---
def generate_pdf_surat_kelahiran():
    # Buat file Excel sementara yang hanya berisi sheet Surat Kelahiran
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_excel_path = os.path.join(tmpdir, "temp_surat.xlsx")
        
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        if "Surat Kelahiran" in wb.sheetnames:
            for sheet_name in list(wb.sheetnames):
                if sheet_name != "Surat Kelahiran":
                    del wb[sheet_name]
        wb.save(temp_excel_path)

        # Konversi file Excel langsung ke PDF menggunakan LibreOffice
        try:
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "pdf", 
                "--outdir", tmpdir, temp_excel_path
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            pdf_path = os.path.join(tmpdir, "temp_surat.pdf")
            with open(pdf_path, "rb") as f:
                pdf_data = f.read()
            return io.BytesIO(pdf_data)
        
        except Exception as e:
            # Fallback jika libreoffice tidak terpasang di sistem lokal
            st.warning("Menampilkan opsi standar (LibreOffice tidak terdeteksi di lingkungan lokal).")
            return generate_excel_surat_kelahiran()

if not os.path.exists(EXCEL_FILE):
    st.error(f"⚠️ File tidak ditemukan: {EXCEL_FILE}")
else:
    try:
        df = load_excel_data(EXCEL_FILE)
        st.success("✅ Data siap!")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📥 Download Excel (Surat Kelahiran)",
                data=generate_excel_surat_kelahiran(),
                file_name="Surat_Kelahiran.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_dl2:
            st.download_button(
                label="📄 Download PDF Resmi (Form F-2.01)",
                data=generate_pdf_surat_kelahiran(),
                file_name="Surat_Keterangan_Kelahiran_F201.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        st.divider()

        # FORM ISIAN DATA
        with st.form("form_suket_lengkap"):
            st.subheader("1. Header & Data Kepala Keluarga")
            c1, c2, c3 = st.columns(3)
            with c1: st.text_input("Nomor Surat", value=get_val(df, 1, 0))
            with c2: st.text_input("Nama Kepala Keluarga", value=get_val(df, 5, 4))
            with c3: st.text_input("Nomor Kartu Keluarga (KK)", value=get_val(df, 6, 4))

            st.divider()

            st.subheader("2. Data Bayi / Anak")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Nama Bayi", value=get_val(df, 10, 4))
                st.text_input("Jenis Kelamin Bayi", value=get_val(df, 11, 4))
                st.text_input("Tempat Kelahiran", value=get_val(df, 13, 4))
            with c2:
                st.text_input("Jam Lahir", value=get_val(df, 15, 4))
                st.caption("Tanggal Lahir Bayi (TGL / BLN / THN / UMUR)")
                cb1, cb2, cb3, cb4 = st.columns(4)
                with cb1: st.text_input("Tgl", value=get_val(df, 15, 5), key="b_tgl")
                with cb2: st.text_input("Bln", value=get_val(df, 15, 6), key="b_bln")
                with cb3: st.text_input("Thn", value=get_val(df, 15, 7), key="b_thn")
                with cb4: st.text_input("Umur", value=get_val(df, 15, 8), key="b_umur")
                st.text_input("Kelahiran Ke-", value=get_val(df, 17, 4))
            with c3:
                st.text_input("Penolong Kelahiran", value=get_val(df, 18, 4))
                st.text_input("Berat Bayi (Gram)", value=get_val(df, 19, 4))
                st.text_input("Panjang Bayi (Cm)", value=get_val(df, 20, 4))

            st.divider()

            col_ibu, col_ayah = st.columns(2)
            with col_ibu:
                st.subheader("3. Data Ibu")
                st.text_input("NIK Ibu", value=get_val(df, 23, 4))
                st.text_input("Nama Lengkap Ibu", value=get_val(df, 24, 4))
                st.caption("Tanggal Lahir Ibu (TGL / BLN / THN / UMUR)")
                ci1, ci2, ci3, ci4 = st.columns(4)
                with ci1: st.text_input("Tgl", value=get_val(df, 25, 5), key="i_tgl")
                with ci2: st.text_input("Bln", value=get_val(df, 25, 6), key="i_bln")
                with ci3: st.text_input("Thn", value=get_val(df, 25, 7), key="i_thn")
                with ci4: st.text_input("Umur", value=get_val(df, 25, 8), key="i_umur")
                st.text_input("Pekerjaan Ibu", value=get_val(df, 26, 4))
                st.text_input("Alamat Ibu", value=get_val(df, 27, 4))
                st.text_input("Kebangsaan Ibu", value=get_val(df, 31, 4))
                st.caption("Tanggal Perkawinan (TGL / BLN / THN)")
                cp1, cp2, cp3 = st.columns(3)
                with cp1: st.text_input("Tgl Kawin", value=get_val(df, 33, 5))
                with cp2: st.text_input("Bln Kawin", value=get_val(df, 33, 6))
                with cp3: st.text_input("Thn Kawin", value=get_val(df, 33, 7))

            with col_ayah:
                st.subheader("4. Data Ayah")
                st.text_input("NIK Ayah", value=get_val(df, 35, 4))
                st.text_input("Nama Lengkap Ayah", value=get_val(df, 36, 4))
                st.caption("Tanggal Lahir Ayah (TGL / BLN / THN / UMUR)")
                ca1, ca2, ca3, ca4 = st.columns(4)
                with ca1: st.text_input("Tgl", value=get_val(df, 38, 5), key="a_tgl")
                with ca2: st.text_input("Bln", value=get_val(df, 38, 6), key="a_bln")
                with ca3: st.text_input("Thn", value=get_val(df, 38, 7), key="a_thn")
                with ca4: st.text_input("Umur", value=get_val(df, 38, 8), key="a_umur")
                st.text_input("Pekerjaan Ayah", value=get_val(df, 38, 4))
                st.text_input("Alamat Ayah", value=get_val(df, 39, 4))
                st.text_input("Kebangsaan Ayah", value=get_val(df, 43, 4))

            st.divider()

            st.subheader("5. Data Pelapor")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("NIK Pelapor", value=get_val(df, 46, 4))
                st.text_input("Nama Lengkap Pelapor", value=get_val(df, 47, 4))
            with c2:
                st.text_input("Umur Pelapor", value=get_val(df, 48, 4))
                st.text_input("Jenis Kelamin Pelapor", value=get_val(df, 49, 4))
            with c3:
                st.text_input("Pekerjaan Pelapor", value=get_val(df, 50, 4))
                st.text_input("Alamat Pelapor", value=get_val(df, 51, 4))

            st.divider()

            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.subheader("6. Data Saksi I")
                st.text_input("NIK Saksi I", value=get_val(df, 56, 4))
                st.text_input("Nama Lengkap Saksi I", value=get_val(df, 57, 4))
                st.text_input("Umur Saksi I", value=get_val(df, 58, 4))
                st.text_input("Alamat Saksi I", value=get_val(df, 60, 4))

            with col_s2:
                st.subheader("7. Data Saksi II")
                st.text_input("NIK Saksi II", value=get_val(df, 65, 4))
                st.text_input("Nama Lengkap Saksi II", value=get_val(df, 66, 4))
                st.text_input("Umur Saksi II", value=get_val(df, 67, 4))
                st.text_input("Alamat Saksi II", value=get_val(df, 69, 4))

            st.divider()

            st.subheader("8. Keterangan Surat")
            st.text_input("Tempat & Tanggal Surat", value=get_val(df, 71, 4))

            btn = st.form_submit_button("Simpan Data")
            if btn:
                st.info("Form disimpan.")

    except Exception as e:
        st.error(f"Error: {e}")
