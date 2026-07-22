import streamlit as st
import pandas as pd
import pdfplumber
import pypdf
import io
import re
from datetime import datetime

st.set_page_config(page_title="Reconcile Cetakita", page_icon="📑", layout="wide")

st.title("📑 Rekonsiliasi Otomatis Pendapatan vs Mutasi Bank")
st.caption("Cetakita.com - Financial Reconciliation System")

# ==========================================
# HELPER: Buka PDF (Support Encrypted PDF BNI/Mandiri)
# ==========================================
def load_pdf_bytes(file_bytes, password=None):
    try:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted:
            if password:
                decrypted = reader.decrypt(password)
                if decrypted == 0:
                    st.error("❌ Password PDF salah! Silakan periksa kembali.")
                    return None
            else:
                st.warning("🔒 File PDF ini dilindungi sandi. Masukkan password di sidebar.")
                return None
        
        out_stream = io.BytesIO()
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.write(out_stream)
        out_stream.seek(0)
        
        return pdfplumber.open(out_stream)
    except Exception as e:
        st.error(f"Gagal membaca PDF: {e}")
        return None

# ==========================================
# 1. PARSER LAPORAN PENDAPATAN SYSTEM
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

# ==========================================
# 2. PARSER MUTASI BANK (BCA / BNI / MANDIRI)
# ==========================================
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
            # Filter hanya kredit/uang masuk
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

# ==========================================
# 3. SIDEBAR UI (UPLOAD & PASSWORD INPUT)
# ==========================================
st.sidebar.header("📂 Upload File & Sandi")

file_pendapatan = st.sidebar.file_uploader("1. Laporan Pendapatan Cetakita (PDF)", type=["pdf"])

st.sidebar.markdown("---")
st.sidebar.subheader("Mutasi Bank")

file_bca = st.sidebar.file_uploader("Mutasi BCA (PDF)", type=["pdf"])

file_bni = st.sidebar.file_uploader("Mutasi BNI (PDF)", type=["pdf"])
pass_bni = st.sidebar.text_input("Password PDF BNI", type="password") if file_bni else None

file_mandiri = st.sidebar.file_uploader("Mutasi Mandiri (PDF)", type=["pdf"])
pass_mandiri = st.sidebar.text_input("Password PDF Mandiri", type="password") if file_mandiri else None

# ==========================================
# 4. EKSEKUSI REKONSILIASI & STYLING
# ==========================================
if file_pendapatan and (file_bca or file_bni or file_mandiri):
    df_sys = parse_laporan_pendapatan(file_pendapatan.read())
    
    list_df_bank = []
    
    if file_bca:
        pdf_bca = pdfplumber.open(io.BytesIO(file_bca.read()))
        list_df_bank.append(parse_mutasi_generic(pdf_bca, "BCA"))
        
    if file_bni:
        pdf_bni = load_pdf_bytes(file_bni.read(), pass_bni)
        if pdf_bni:
            list_df_bank.append(parse_mutasi_generic(pdf_bni, "BNI"))
            
    if file_mandiri:
        pdf_mandiri = load_pdf_bytes(file_mandiri.read(), pass_mandiri)
        if pdf_mandiri:
            list_df_bank.append(parse_mutasi_generic(pdf_mandiri, "Mandiri"))

    if list_df_bank:
        df_bank = pd.concat(list_df_bank, ignore_index=True)

        matched_results = []
        sys_matched_idx = set()
        bank_matched_idx = set()

        # Match Logik 1:1
        for s_idx, s_row in df_sys.iterrows():
            for b_idx, b_row in df_bank.iterrows():
                if b_idx in bank_matched_idx:
                    continue
                
                if (s_row['amount_sys'] == b_row['amount_bank']) and (s_row['bank_sys'] == b_row['bank']):
                    matched_results.append({
                        'Status': 'MATCHED',
                        'Kode Order': s_row['kode_order'],
                        'Tgl System': s_row['tanggal_sys'],
                        'Tgl Bank': b_row['tanggal_bank'],
                        'Nominal': s_row['amount_sys'],
                        'Bank': s_row['bank_sys'],
                        'Ket Mutasi': b_row['keterangan_bank']
                    })
                    sys_matched_idx.add(s_idx)
                    bank_matched_idx.add(b_idx)
                    break

        unmatched_sys = df_sys[~df_sys.index.isin(sys_matched_idx)].copy()
        unmatched_sys['Status'] = 'UNMATCHED / BELUM MASUK BANK'

        unmatched_bank = df_bank[~df_bank.index.isin(bank_matched_idx)].copy()
        unmatched_bank['Status'] = 'UNMATCHED / TIDAK TERCATAT SYSTEM'

        # Metrics Summary
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Order Ter-match", len(matched_results))
        col2.metric("System Belum Masuk Bank", len(unmatched_sys))
        col3.metric("Mutasi Tidak Terdaftar", len(unmatched_bank))

        st.markdown("---")

        # ==========================================
        # FUNSI STYLING WARNA (HIGHLIGHTING)
        # ==========================================
        def highlight_matched(row):
            """Fungsi memberi warna baris berdasarkan Bank pada data MATCHED"""
            bank = str(row['Bank']).upper()
            if bank == 'BCA':
                return ['background-color: #FFF9C4; color: #000000;'] * len(row)  # Kuning Muda
            elif bank == 'BNI':
                return ['background-color: #C8E6C9; color: #000000;'] * len(row)  # Hijau Muda
            elif bank == 'MANDIRI':
                return ['background-color: #E1F5FE; color: #000000;'] * len(row)  # Biru Muda
            return [''] * len(row)

        def highlight_unmatched(row):
            """Fungsi memberi warna Pink pada data UNMATCHED / GA SESUAI"""
            return ['background-color: #FFCDD2; color: #000000;'] * len(row)  # Pink / Merah Muda

        # Display Tabular Results
        tab1, tab2, tab3 = st.tabs(["✅ Data Matched (Sesuai)", "⚠️ Order Belum Ada di Mutasi", "⚠️ Mutasi Tidak Terdaftar"])
        
        with tab1:
            st.subheader("Data Transaksi Cocok")
            if matched_results:
                df_matched = pd.DataFrame(matched_results)
                # Terapkan highlight berdasarkan jenis bank
                styled_matched = df_matched.style.apply(highlight_matched, axis=1)
                st.dataframe(styled_matched, use_container_width=True)
            else:
                st.info("Belum ada data matched.")

        with tab2:
            st.subheader("Pendapatan System Belum Ditemukan di Mutasi")
            if not unmatched_sys.empty:
                # Terapkan highlight Pink
                styled_unmatched_sys = unmatched_sys.style.apply(highlight_unmatched, axis=1)
                st.dataframe(styled_unmatched_sys, use_container_width=True)
            else:
                st.success("Semua order pendapatan di system sudah cocok dengan mutasi bank!")

        with tab3:
            st.subheader("Mutasi Uang Masuk Tidak Dikenali di System")
            if not unmatched_bank.empty:
                # Terapkan highlight Pink
                styled_unmatched_bank = unmatched_bank.style.apply(highlight_unmatched, axis=1)
                st.dataframe(styled_unmatched_bank, use_container_width=True)
            else:
                st.success("Tidak ada mutasi bank asing/tanpa catatan!")