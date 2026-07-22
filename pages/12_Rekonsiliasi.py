import streamlit as st
import pandas as pd
import pdfplumber
import pypdf
import io
import re
from datetime import datetime

# ReportLab untuk merender ulang PDF Hasil Rekonsiliasi
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Rekonsiliasi Bank Cetakita", page_icon="📄", layout="wide")

# ==========================================
# 1. HELPER CONVERT PDF TO DATAFRAME (EXCEL/CSV LOGIC)
# ==========================================
def extract_tables_from_pdf(file_uploader, password=None):
    """Membuka PDF dan mengonversi seluruh isi tabel menjadi Pandas DataFrame"""
    if file_uploader is None:
        return None
    try:
        file_uploader.seek(0)
        file_bytes = file_uploader.read()
        file_uploader.seek(0)

        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted:
            if password:
                reader.decrypt(password)
            else:
                st.warning(f"🔒 File {file_uploader.name} terproteksi sandi!")
                return None

        out_stream = io.BytesIO()
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.write(out_stream)
        out_stream.seek(0)

        all_text = []
        with pdfplumber.open(out_stream) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text.extend(text.split('\n'))
        return all_text
    except Exception as e:
        st.error(f"Error ekstraksi PDF {file_uploader.name}: {e}")
        return []

def convert_pendapatan_to_df(lines):
    records = []
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

            if order_match and total_val > 0:
                records.append({
                    'Kode Order': order_match.group(1),
                    'Tanggal': date_match.group(1) if date_match else '-',
                    'Bank System': bank,
                    'Nominal': total_val
                })
    return pd.DataFrame(records)

def convert_mutasi_to_df(lines, bank_name):
    records = []
    for line in lines:
        if ("CR" in line or "KREDIT" in line or "TRANSFER" in line) and "BUNGA" not in line and "PAJAK" not in line and "BIAYA ADM" not in line:
            date_match = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})', line)
            amount_match = re.search(r'([\d\.]+)\s*(?:CR|KREDIT)?', line)
            
            if amount_match:
                try:
                    amount = float(amount_match.group(1).replace('.', '').replace(',', '.'))
                    if amount > 0:
                        records.append({
                            'Tanggal Mutasi': date_match.group(1) if date_match else '-',
                            'Bank': bank_name,
                            'Nominal Mutasi': amount,
                            'Keterangan': line[:50] # Ambil ringkasan
                        })
                except ValueError:
                    continue
    return pd.DataFrame(records)

# ==========================================
# 2. HELPER GENERATE NEW HIGHLIGHTED PDF
# ==========================================
def generate_reconciled_pdf(title, df_data, is_matched_pdf=True):
    """Mengubah DataFrame hasil olahan menjadi File PDF Baru ber-stabilo/highlight"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=14, leading=16, textColor=colors.HexColor('#1A237E'))
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10)

    story.append(Paragraph(f"<b>{title}</b>", title_style))
    story.append(Paragraph(f"Tanggal Cetak: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}", cell_style))
    story.append(Spacer(1, 10))

    if df_data.empty:
        story.append(Paragraph("Tidak ada data untuk ditampilkan.", cell_style))
        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    # Prepare Data Table
    headers = [Paragraph(f"<b>{col}</b>", cell_style) for col in df_data.columns]
    table_data = [headers]

    for _, row in df_data.iterrows():
        row_cells = []
        for val in row:
            if isinstance(val, (int, float)):
                formatted_val = f"Rp {val:,.0f}".replace(',', '.')
                row_cells.append(Paragraph(formatted_val, cell_style))
            else:
                row_cells.append(Paragraph(str(val), cell_style))
        table_data.append(row_cells)

    t = Table(table_data, repeatRows=1)
    
    # Base Style
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#37474F')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#B0BEC5')),
    ]

    # Dynamic Row Highlighting
    for idx, row in df_data.iterrows():
        row_idx = idx + 1 # offset for header
        if is_matched_pdf:
            bank = str(row.get('Bank System', row.get('Bank', ''))).upper()
            if 'BCA' in bank:
                bg_color = colors.HexColor('#FFF9C4') # Kuning
            elif 'BNI' in bank:
                bg_color = colors.HexColor('#C8E6C9') # Hijau
            elif 'MANDIRI' in bank:
                bg_color = colors.HexColor('#E1F5FE') # Biru
            else:
                bg_color = colors.white
        else:
            bg_color = colors.HexColor('#FFCDD2') # Pink untuk unmatched

        t_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), bg_color))

    t.setStyle(TableStyle(t_style))
    story.append(t)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 3. MAIN UI LAYOUT
# ==========================================
st.title("📄 Rekonsiliasi Bank: PDF ➔ CSV/Data ➔ PDF Highlighting")
st.caption("Cetakita Financial System - Automated Conversion & Matching")

st.markdown("### 📂 1. Upload File PDF")
with st.expander("Area Upload Laporan Pendapatan & Mutasi Bank", expanded=True):
    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.subheader("Laporan Pendapatan System")
        file_pendapatan = st.file_uploader("Upload Pendapatan Cetakita (PDF)", type=["pdf"], key="p_pdf")
        
        st.subheader("Mutasi BCA")
        file_bca = st.file_uploader("Upload Mutasi BCA (PDF)", type=["pdf"], key="bca_pdf")

    with col_up2:
        st.subheader("Mutasi BNI")
        file_bni = st.file_uploader("Upload Mutasi BNI (PDF)", type=["pdf"], key="bni_pdf")
        pass_bni = st.text_input("Password BNI", type="password") if file_bni else None

        st.subheader("Mutasi Mandiri")
        file_mandiri = st.file_uploader("Upload Mutasi Mandiri (PDF)", type=["pdf"], key="mandiri_pdf")
        pass_mandiri = st.text_input("Password Mandiri", type="password") if file_mandiri else None

# ==========================================
# 4. PROCESS: CONVERT PDF TO CSV/DATA, RECONCILE, RENDER NEW PDF
# ==========================================
if file_pendapatan and (file_bca or file_bni or file_mandiri):
    st.markdown("---")
    st.markdown("### 🔄 2. Proses Konversi Data & Rekonsiliasi")

    # Step 1: Extract Text & Convert to Dataframe
    lines_pendapatan = extract_tables_from_pdf(file_pendapatan)
    df_sys = convert_pendapatan_to_df(lines_pendapatan)

    list_df_bank = []
    if file_bca:
        lines_bca = extract_tables_from_pdf(file_bca)
        list_df_bank.append(convert_mutasi_to_df(lines_bca, "BCA"))
    if file_bni:
        lines_bni = extract_tables_from_pdf(file_bni, pass_bni)
        list_df_bank.append(convert_mutasi_to_df(lines_bni, "BNI"))
    if file_mandiri:
        lines_mandiri = extract_tables_from_pdf(file_mandiri, pass_mandiri)
        list_df_bank.append(convert_mutasi_to_df(lines_mandiri, "Mandiri"))

    if list_df_bank:
        df_bank = pd.concat(list_df_bank, ignore_index=True)

        # Step 2: Matching Logic
        matched_list = []
        sys_matched_idx = set()
        bank_matched_idx = set()

        for s_idx, s_row in df_sys.iterrows():
            for b_idx, b_row in df_bank.iterrows():
                if b_idx in bank_matched_idx:
                    continue
                if (s_row['Nominal'] == b_row['Nominal Mutasi']) and (s_row['Bank System'] == b_row['Bank']):
                    matched_list.append({
                        'Kode Order': s_row['Kode Order'],
                        'Tgl System': s_row['Tanggal'],
                        'Tgl Bank': b_row['Tanggal Mutasi'],
                        'Bank System': s_row['Bank System'],
                        'Nominal': s_row['Nominal'],
                        'Status': 'MATCHED',
                        'Keterangan Bank': b_row['Keterangan']
                    })
                    sys_matched_idx.add(s_idx)
                    bank_matched_idx.add(b_idx)
                    break

        df_matched = pd.DataFrame(matched_list)
        df_unmatched_sys = df_sys[~df_sys.index.isin(sys_matched_idx)].copy()
        if not df_unmatched_sys.empty:
            df_unmatched_sys['Status'] = 'BELUM MASUK BANK'

        df_unmatched_bank = df_bank[~df_bank.index.isin(bank_matched_idx)].copy()
        if not df_unmatched_bank.empty:
            df_unmatched_bank['Status'] = 'MUTASI TANPA ORDER'

        # Metrics
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("✅ Total Matched", len(df_matched))
        col_m2.metric("⚠️ Pending di Bank", len(df_unmatched_sys))
        col_m3.metric("⚠️ Mutasi Masuk Asing", len(df_unmatched_bank))

        # Step 3: Render Hasil Kembali ke File PDF Baru
        st.markdown("---")
        st.markdown("### 📥 3. Download Laporan PDF Hasil Rekonsiliasi (Berwarna)")

        pdf_matched_bytes = generate_reconciled_pdf("Laporan Rekonsiliasi Matched (Cocok)", df_matched, is_matched_pdf=True)
        pdf_unmatched_sys_bytes = generate_reconciled_pdf("Laporan Pendapatan Belum Masuk Rekening Bank", df_unmatched_sys, is_matched_pdf=False)
        pdf_unmatched_bank_bytes = generate_reconciled_pdf("Laporan Mutasi Uang Masuk Tanpa Record System", df_unmatched_bank, is_matched_pdf=False)

        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            st.download_button(
                label="📄 Download PDF Matched (Kuning/Hijau/Biru)",
                data=pdf_matched_bytes,
                file_name=f"Laporan_Matched_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary"
            )
        with col_dl2:
            st.download_button(
                label="📄 Download PDF Pending System (Pink)",
                data=pdf_unmatched_sys_bytes,
                file_name=f"Laporan_Pending_System_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )
        with col_dl3:
            st.download_button(
                label="📄 Download PDF Mutasi Asing (Pink)",
                data=pdf_unmatched_bank_bytes,
                file_name=f"Laporan_Mutasi_Unmatched_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )

        # Tabular Display
        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["✅ Data Matched", "⚠️ Pendapatan Belum Masuk Bank", "⚠️ Mutasi Tidak Terdaftar"])
        with tab1:
            st.dataframe(df_matched, use_container_width=True)
        with tab2:
            st.dataframe(df_unmatched_sys, use_container_width=True)
        with tab3:
            st.dataframe(df_unmatched_bank, use_container_width=True)