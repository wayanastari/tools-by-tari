import streamlit as st
import pandas as pd
import pdfplumber
import pypdf
import io
import re
from datetime import datetime

# Menggunakan PyMuPDF (fitz) untuk mewarnai / highlight teks langsung di file PDF
try:
    import fitz
except ImportError:
    st.error("⚠️ Silakan install PyMuPDF di environment Anda: pip install pymupdf")

st.set_page_config(page_title="Rekonsiliasi Bank - Cetakita", page_icon="📄", layout="wide")

# ==========================================
# 1. HELPER & ERROR-PROOF PDF OPENER
# ==========================================
def safe_open_pdf(file_uploader, password=None):
    """
    Fungsi aman untuk membuka PDF dari BytesIOStream.
    Mencegah error 'PdfminerException' / EOF Error dengan mengelola .seek(0).
    """
    if file_uploader is None:
        return None
    try:
        file_uploader.seek(0)
        file_bytes = file_uploader.read()
        file_uploader.seek(0) # Reset pointer kembali ke awal
        
        # Buka via pypdf untuk me-resolve password/enkripsi e-Statement
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted:
            if password:
                decrypted = reader.decrypt(password)
                if decrypted == 0:
                    st.error(f"❌ Password PDF untuk {file_uploader.name} salah!")
                    return None
            else:
                st.warning(f"🔒 File {file_uploader.name} terproteksi password. Silakan masukkan password.")
                return None

        out_stream = io.BytesIO()
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.write(out_stream)
        out_stream.seek(0)
        
        return pdfplumber.open(out_stream)
    except Exception as e:
        st.error(f"❌ Gagal memproses file {file_uploader.name}: {e}")
        return None

def highlight_pdf_text(file_bytes, highlight_rules, password=None):
    """
    Menambahkan penanda stabilo (highlight) ke dalam file PDF asli.
    highlight_rules = {'TEXT_KEYWORD': (R, G, B)} -> Nilai RGB dalam rentang 0.0 s/d 1.0
    """
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        if doc.is_encrypted and password:
            doc.authenticate(password)
            
        for page in doc:
            for text_pattern, color in highlight_rules.items():
                if not text_pattern or len(str(text_pattern)) < 3:
                    continue
                text_instances = page.search_for(str(text_pattern))
                for inst in text_instances:
                    highlight = page.add_highlight_annot(inst)
                    highlight.set_colors(stroke=color)
                    highlight.update()
                    
        output_stream = io.BytesIO()
        doc.save(output_stream)
        doc.close()
        output_stream.seek(0)
        return output_stream.getvalue()
    except Exception as e:
        st.warning(f"Gagal mewarnai sebagian isi PDF: {e}")
        return file_bytes

# ==========================================
# 2. PARSERS (LAPORAN PENDAPATAN & MUTASI)
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

# WARNA RGB HIGHLIGHT (Range 0.0 - 1.0)
COLOR_BCA = (1.0, 0.98, 0.5)      # Kuning Muda
COLOR_BNI = (0.6, 0.9, 0.6)       # Hijau Muda
COLOR_MANDIRI = (0.7, 0.88, 1.0)  # Biru Muda
COLOR_PINK = (1.0, 0.7, 0.75)     # Pink / Merah Muda

# ==========================================
# 3. MAIN UI LAYOUT (MAIN AREA)
# ==========================================
st.title("📄 Rekonsiliasi Otomatis Pendapatan vs Mutasi Bank")
st.caption("Cetakita.com - Financial Reconciliation System")

# AREA UPLOAD FILE DIPINDAH TOTAL KE MAIN AREA (BUKAN SIDEBAR)
st.markdown("### 📂 Upload Dokumen PDF")
with st.expander("Klik di sini untuk Unggah Laporan Pendapatan & Rekening Koran Bank", expanded=True):
    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.subheader("1. Laporan Pendapatan System")
        file_pendapatan = st.file_uploader("Upload Pendapatan Cetakita (PDF)", type=["pdf"], key="main_pendapatan")
        
        st.subheader("2. Mutasi Bank BCA")
        file_bca = st.file_uploader("Upload Mutasi BCA (PDF)", type=["pdf"], key="main_bca")

    with col_up2:
        st.subheader("3. Mutasi Bank BNI")
        file_bni = st.file_uploader("Upload Mutasi BNI (PDF)", type=["pdf"], key="main_bni")
        pass_bni = st.text_input("Password PDF BNI (jika ada)", type="password", key="pass_bni") if file_bni else None

        st.subheader("4. Mutasi Bank Mandiri")
        file_mandiri = st.file_uploader("Upload Mutasi Mandiri (PDF)", type=["pdf"], key="main_mandiri")
        pass_mandiri = st.text_input("Password PDF Mandiri (jika ada)", type="password", key="pass_mandiri") if file_mandiri else None

# ==========================================
# 4. EKSEKUSI REKONSILIASI & PENANDAAN PDF
# ==========================================
if file_pendapatan and (file_bca or file_bni or file_mandiri):
    
    # Baca bytes laporan pendapatan
    file_pendapatan.seek(0)
    bytes_pendapatan = file_pendapatan.read()
    file_pendapatan.seek(0)
    
    df_sys = parse_laporan_pendapatan(bytes_pendapatan)
    
    dict_bank_bytes = {}
    list_df_bank = []
    
    # Process BCA
    if file_bca:
        file_bca.seek(0)
        bca_bytes = file_bca.read()
        file_bca.seek(0)
        dict_bank_bytes['BCA'] = {'bytes': bca_bytes, 'pass': None}
        pdf_bca = safe_open_pdf(file_bca)
        if pdf_bca:
            list_df_bank.append(parse_mutasi_generic(pdf_bca, "BCA"))
        
    # Process BNI
    if file_bni:
        file_bni.seek(0)
        bni_bytes = file_bni.read()
        file_bni.seek(0)
        dict_bank_bytes['BNI'] = {'bytes': bni_bytes, 'pass': pass_bni}
        pdf_bni = safe_open_pdf(file_bni, pass_bni)
        if pdf_bni:
            list_df_bank.append(parse_mutasi_generic(pdf_bni, "BNI"))
            
    # Process Mandiri
    if file_mandiri:
        file_mandiri.seek(0)
        mandiri_bytes = file_mandiri.read()
        file_mandiri.seek(0)
        dict_bank_bytes['Mandiri'] = {'bytes': mandiri_bytes, 'pass': pass_mandiri}
        pdf_mandiri = safe_open_pdf(file_mandiri, pass_mandiri)
        if pdf_mandiri:
            list_df_bank.append(parse_mutasi_generic(pdf_mandiri, "Mandiri"))

    if list_df_bank:
        df_bank = pd.concat(list_df_bank, ignore_index=True)

        rules_pdf_pendapatan = {}
        rules_pdf_bank = {'BCA': {}, 'BNI': {}, 'Mandiri': {}}

        matched_results = []
        sys_matched_idx = set()
        bank_matched_idx = set()

        # LOGIKA MATCHING 1:1 & PENENTUAN WARNA HIGHLIGHT
        for s_idx, s_row in df_sys.iterrows():
            matched_found = False
            for b_idx, b_row in df_bank.iterrows():
                if b_idx in bank_matched_idx:
                    continue
                
                if (s_row['amount_sys'] == b_row['amount_bank']) and (s_row['bank_sys'] == b_row['bank']):
                    bank_name = s_row['bank_sys']
                    color = COLOR_BCA if bank_name == 'BCA' else (COLOR_BNI if bank_name == 'BNI' else COLOR_MANDIRI)
                    
                    # Tandai kode order di PDF Pendapatan
                    rules_pdf_pendapatan[s_row['kode_order']] = color
                    
                    # Tandai nominal di PDF Mutasi Bank
                    amt_str = f"{s_row['amount_sys']:,.0f}".replace(',', '.')
                    rules_pdf_bank[bank_name][amt_str] = color

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
                    matched_found = True
                    break
            
            # Jika Pendapatan System tidak ter-match -> Highlight Pink
            if not matched_found:
                rules_pdf_pendapatan[s_row['kode_order']] = COLOR_PINK

        # Unmatched di Mutasi Bank -> Highlight Pink
        unmatched_bank = df_bank[~df_bank.index.isin(bank_matched_idx)].copy()
        for b_idx, b_row in unmatched_bank.iterrows():
            bank_name = b_row['bank']
            amt_str = f"{b_row['amount_bank']:,.0f}".replace(',', '.')
            rules_pdf_bank[bank_name][amt_str] = COLOR_PINK

        unmatched_sys = df_sys[~df_sys.index.isin(sys_matched_idx)].copy()

        # STATISTIK METRICS
        st.markdown("---")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("✅ Total Matched", len(matched_results))
        col_m2.metric("⚠️ Order System Belum Ada di Bank", len(unmatched_sys))
        col_m3.metric("⚠️ Uang Masuk Tanpa Record System", len(unmatched_bank))

        # TOMBOL DOWNLOAD PDF BERWARNA / HIGHLIGHTED
        st.markdown("---")
        st.subheader("📥 Download File PDF Ter-Highlight")
        
        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            st.markdown("#### 1. PDF Pendapatan Cetakita")
            pdf_pendapatan_highlighted = highlight_pdf_text(bytes_pendapatan, rules_pdf_pendapatan)
            st.download_button(
                label="📄 Download PDF Pendapatan (Sudah Di-Highlight)",
                data=pdf_pendapatan_highlighted,
                file_name=f"Pendapatan_Cetakita_Highlighted_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary"
            )

        with col_dl2:
            st.markdown("#### 2. PDF Rekening Koran Bank")
            for bank_name, b_info in dict_bank_bytes.items():
                if rules_pdf_bank[bank_name]:
                    pdf_bank_highlighted = highlight_pdf_text(b_info['bytes'], rules_pdf_bank[bank_name], b_info['pass'])
                    st.download_button(
                        label=f"🏦 Download PDF Mutasi {bank_name} (Di-Highlight)",
                        data=pdf_bank_highlighted,
                        file_name=f"Mutasi_{bank_name}_Highlighted_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )

        # TABEL DISPLAY DENGAN STYLING STREAMLIT
        st.markdown("---")
        def highlight_matched_tbl(row):
            bank = str(row['Bank']).upper()
            if bank == 'BCA':
                return ['background-color: #FFF9C4; color: #000000;'] * len(row)
            elif bank == 'BNI':
                return ['background-color: #C8E6C9; color: #000000;'] * len(row)
            elif bank == 'MANDIRI':
                return ['background-color: #E1F5FE; color: #000000;'] * len(row)
            return [''] * len(row)

        def highlight_unmatched_tbl(row):
            return ['background-color: #FFCDD2; color: #000000;'] * len(row)

        tab1, tab2, tab3 = st.tabs(["✅ Data Matched", "⚠️ Pendapatan Belum Masuk Bank", "⚠️ Mutasi Tidak Terdaftar"])
        
        with tab1:
            if matched_results:
                df_matched = pd.DataFrame(matched_results)
                st.dataframe(df_matched.style.apply(highlight_matched_tbl, axis=1), use_container_width=True)
            else:
                st.info("Belum ada data matched.")

        with tab2:
            if not unmatched_sys.empty:
                st.dataframe(unmatched_sys.style.apply(highlight_unmatched_tbl, axis=1), use_container_width=True)
            else:
                st.success("Semua transaksi pendapatan cocok!")

        with tab3:
            if not unmatched_bank.empty:
                st.dataframe(unmatched_bank.style.apply(highlight_unmatched_tbl, axis=1), use_container_width=True)
            else:
                st.success("Tidak ada transaksi mutasi mencurigakan!")