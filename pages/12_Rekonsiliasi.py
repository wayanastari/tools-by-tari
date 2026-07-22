import streamlit as st
import pandas as pd
import pdfplumber
import pypdf
import io
import re
from datetime import datetime

# Menggunakan PyMuPDF (fitz) untuk menambahkan highlight warna langsung ke PDF
try:
    import fitz
except ImportError:
    st.error("Silakan install PyMuPDF terlebih dahulu dengan perintah: pip install pymupdf")

st.set_page_config(page_title="Cetakita PDF Reconciler", page_icon="📄", layout="wide")

# ==========================================
# 1. NAVIGATION SUBMENU
# ==========================================
st.sidebar.title("📌 Menu Utama")
main_menu = st.sidebar.radio("Pilih Modul:", ["Finance & Accounting"])

if main_menu == "Finance & Accounting":
    submenu = st.sidebar.selectbox("Submenu:", ["Rekonsiliasi Bank", "Laporan Ringkasan"])

# ==========================================
# HELPER: BUKA & ANNOTATE PDF
# ==========================================
def highlight_pdf_text(file_bytes, highlight_rules, password=None):
    """
    Menambahkan highlight warna pada PDF berdasarkan keyword/teks tertentu.
    highlight_rules: dict, contoh {'ON202605001': (1, 0.98, 0.77), ...} -> RGB 0-1
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    if doc.is_encrypted and password:
        doc.authenticate(password)
        
    for page in doc:
        for text_pattern, color in highlight_rules.items():
            if not text_pattern or len(text_pattern) < 3:
                continue
            text_instances = page.search_for(text_pattern)
            for inst in text_instances:
                highlight = page.add_highlight_annot(inst)
                highlight.set_colors(stroke=color) # RGB Tuple
                highlight.update()
                
    output_stream = io.BytesIO()
    doc.save(output_stream)
    doc.close()
    output_stream.seek(0)
    return output_stream.getvalue()

def load_pdf_plumber(file_bytes, password=None):
    try:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted:
            if password:
                reader.decrypt(password)
            else:
                return None
        out_stream = io.BytesIO()
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.write(out_stream)
        out_stream.seek(0)
        return pdfplumber.open(out_stream)
    except Exception as e:
        return None

# ==========================================
# PARSER LAPORAN PENDAPATAN & MUTASI
# ==========================================
def parse_laporan_pendapatan(pdf_bytes):
    records = []
    pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')
        for line in lines:
            if "ON2026" in line:
                order_match = re.search(r'(ON\d+)', line)
                date_match = re.search(r'(\d{4}-\d{2}-\d{2})', line)
                
                numbers = re.findall(r'[\d\.]+', line)
                total_val = 0
                if numbers:
                    raw_amount = numbers[-1].replace('.', '')
                    if raw_amount.isdigit():
                        total_val = float(raw_amount)

                bank = "UNKNOWN"
                if "BCA" in line: bank = "BCA"
                elif "BNI" in line: bank = "BNI"
                elif "Mandiri" in line: bank = "Mandiri"

                if order_match and date_match and total_val > 0:
                    records.append({
                        'kode_order': order_match.group(1),
                        'tanggal_sys': date_match.group(1),
                        'bank_sys': bank,
                        'amount_sys': total_val,
                        'raw_line': line
                    })
    pdf.close()
    return pd.DataFrame(records)

def parse_mutasi_generic(pdf_obj, bank_name):
    records = []
    if not pdf_obj:
        return pd.DataFrame(records)
        
    for page in pdf_obj.pages:
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')
        for line in lines:
            if ("CR" in line or "KREDIT" in line or "TRANSFER" in line) and "BUNGA" not in line and "PAJAK" not in line and "BIAYA ADM" not in line:
                date_match = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})', line)
                amount_match = re.search(r'([\d\.]+)\s*(?:CR|KREDIT)?', line)
                
                if amount_match:
                    try:
                        amount = float(amount_match.group(1).replace('.', '').replace(',', '.'))
                        if amount > 0:
                            records.append({
                                'tanggal_bank': date_match.group(1) if date_match else '-',
                                'bank': bank_name,
                                'amount_bank': amount,
                                'keterangan_bank': line
                            })
                    except ValueError:
                        continue
    pdf_obj.close()
    return pd.DataFrame(records)

# RGB COLOR MAPS (0.0 - 1.0)
COLOR_BCA = (1.0, 0.98, 0.5)      # Kuning Muda
COLOR_BNI = (0.6, 0.9, 0.6)       # Hijau Muda
COLOR_MANDIRI = (0.7, 0.88, 1.0)  # Biru Muda
COLOR_PINK = (1.0, 0.7, 0.75)     # Pink / Merah Muda

# ==========================================
# 2. SUBMENU: REKONSILIASI BANK
# ==========================================
if main_menu == "Finance & Accounting" and submenu == "Rekonsiliasi Bank":
    st.title("📄 Rekonsiliasi & Auto-Highlight File PDF")
    st.caption("Upload file PDF pendapatan dan mutasi bank. Hasil PDF yang telah di-highlight warna per transaksi dapat langsung diunduh.")

    # AREA UPLOAD DI HALAMAN UTAMA
    with st.expander("📂 **Area Upload PDF (Klik untuk Membuka/Menutup)**", expanded=True):
        col_up1, col_up2 = st.columns(2)
        
        with col_up1:
            st.subheader("1. Laporan Pendapatan System")
            file_pendapatan = st.file_uploader("Upload Pendapatan Cetakita (PDF)", type=["pdf"], key="pendapatan")
            
            st.subheader("2. Mutasi Bank BCA")
            file_bca = st.file_uploader("Upload Mutasi BCA (PDF)", type=["pdf"], key="bca")

        with col_up2:
            st.subheader("3. Mutasi Bank BNI")
            file_bni = st.file_uploader("Upload Mutasi BNI (PDF)", type=["pdf"], key="bni")
            pass_bni = st.text_input("Password PDF BNI (jika ada)", type="password") if file_bni else None

            st.subheader("4. Mutasi Bank Mandiri")
            file_mandiri = st.file_uploader("Upload Mutasi Mandiri (PDF)", type=["pdf"], key="mandiri")
            pass_mandiri = st.text_input("Password PDF Mandiri (jika ada)", type="password") if file_mandiri else None

    # PROSES REKONSILIASI & PENANDAAN HIGHLIGHT PDF
    if file_pendapatan and (file_bca or file_bni or file_mandiri):
        bytes_pendapatan = file_pendapatan.read()
        df_sys = parse_laporan_pendapatan(bytes_pendapatan)
        
        dict_bank_bytes = {}
        list_df_bank = []
        
        if file_bca:
            bca_bytes = file_bca.read()
            dict_bank_bytes['BCA'] = {'bytes': bca_bytes, 'pass': None}
            pdf_bca = load_pdf_plumber(bca_bytes)
            list_df_bank.append(parse_mutasi_generic(pdf_bca, "BCA"))
            
        if file_bni:
            bni_bytes = file_bni.read()
            dict_bank_bytes['BNI'] = {'bytes': bni_bytes, 'pass': pass_bni}
            pdf_bni = load_pdf_plumber(bni_bytes, pass_bni)
            list_df_bank.append(parse_mutasi_generic(pdf_bni, "BNI"))
                
        if file_mandiri:
            mandiri_bytes = file_mandiri.read()
            dict_bank_bytes['Mandiri'] = {'bytes': mandiri_bytes, 'pass': pass_mandiri}
            pdf_mandiri = load_pdf_plumber(mandiri_bytes, pass_mandiri)
            list_df_bank.append(parse_mutasi_generic(pdf_mandiri, "Mandiri"))

        if list_df_bank:
            df_bank = pd.concat(list_df_bank, ignore_index=True)

            rules_pdf_pendapatan = {}
            rules_pdf_bank = {'BCA': {}, 'BNI': {}, 'Mandiri': {}}

            sys_matched_idx = set()
            bank_matched_idx = set()

            # PINTU MATCHING & ATUR WARNA HIGHLIGHT
            for s_idx, s_row in df_sys.iterrows():
                matched_found = False
                for b_idx, b_row in df_bank.iterrows():
                    if b_idx in bank_matched_idx:
                        continue
                    
                    if (s_row['amount_sys'] == b_row['amount_bank']) and (s_row['bank_sys'] == b_row['bank']):
                        bank_name = s_row['bank_sys']
                        color = COLOR_BCA if bank_name == 'BCA' else (COLOR_BNI if bank_name == 'BNI' else COLOR_MANDIRI)
                        
                        # Set rule highlight pendapatan system
                        rules_pdf_pendapatan[s_row['kode_order']] = color
                        
                        # Set rule highlight mutasi bank
                        # Ambil angka nominal untuk dicari & dihighlight di PDF bank
                        amt_str = f"{s_row['amount_sys']:,.0f}".replace(',', '.')
                        rules_pdf_bank[bank_name][amt_str] = color

                        sys_matched_idx.add(s_idx)
                        bank_matched_idx.add(b_idx)
                        matched_found = True
                        break
                
                if not matched_found:
                    # Jika Unmatched di System -> Highlight Pink
                    rules_pdf_pendapatan[s_row['kode_order']] = COLOR_PINK

            # Unmatched Mutasi Bank -> Highlight Pink
            unmatched_bank = df_bank[~df_bank.index.isin(bank_matched_idx)].copy()
            for b_idx, b_row in unmatched_bank.iterrows():
                bank_name = b_row['bank']
                amt_str = f"{b_row['amount_bank']:,.0f}".replace(',', '.')
                rules_pdf_bank[bank_name][amt_str] = COLOR_PINK

            st.markdown("---")
            st.subheader("📥 Download File PDF Yang Sudah Di-Highlight Warna")
            
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                st.markdown("### 1. Laporan Pendapatan Cetakita")
                pdf_pendapatan_highlighted = highlight_pdf_text(bytes_pendapatan, rules_pdf_pendapatan)
                st.download_button(
                    label="📄 Download PDF Pendapatan (Sudah Di-Highlight)",
                    data=pdf_pendapatan_highlighted,
                    file_name="Pendapatan_Cetakita_Highlighted.pdf",
                    mime="application/pdf",
                    type="primary"
                )

            with col_dl2:
                st.markdown("### 2. Rekening Koran Bank")
                for bank_name, b_info in dict_bank_bytes.items():
                    if rules_pdf_bank[bank_name]:
                        pdf_bank_highlighted = highlight_pdf_text(b_info['bytes'], rules_pdf_bank[bank_name], b_info['pass'])
                        st.download_button(
                            label=f"🏦 Download PDF Mutasi {bank_name} (Di-Highlight)",
                            data=pdf_bank_highlighted,
                            file_name=f"Mutasi_{bank_name}_Highlighted.pdf",
                            mime="application/pdf"
                        )

elif main_menu == "Finance & Accounting" and submenu == "Laporan Ringkasan":
    st.title("📈 Submenu: Laporan Ringkasan")
    st.info("Area ringkasan & grafik laporan keuangan.")