import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import io
import math
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor

# Konversi Satuan
MM2PT = 72.0 / 25.4  # 1 mm = 2.83465 points (ReportLab)

st.set_page_config(page_title="VDP Generator & Imposition Studio", layout="wide")

st.title("🖨️ VDP Generator & Layout Cetak Massal")
st.caption("Aplikasi Variable Data Printing dengan QR Code, Numerator, dan Penataan Lembar Cetak (PDF Vector / Corel-Ready)")

# --- SIDEBAR: UPLOAD FILE ---
with st.sidebar:
    st.header("1. Upload Asset")
    design_file = st.file_uploader("Upload Desain (PNG/JPG)", type=["png", "jpg", "jpeg"])
    data_file = st.file_uploader("Upload Data (CSV/Excel)", type=["csv", "xlsx"])
    
    st.divider()
    st.info("💡 **Tips CorelDRAW:** Setelah men-download PDF, gunakan perintah **File > Import** di CorelDRAW dan pilih **Text as Curves / Editable** untuk memisah tiap objek.")

if design_file and data_file:
    # Read Data
    if data_file.name.endswith('.csv'):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)
        
    base_img = Image.open(design_file)
    img_pixel_w, img_pixel_h = base_img.size
    aspect_ratio = img_pixel_h / img_pixel_w

    # Tab Interface
    tab1, tab2, tab3 = st.tabs(["📏 1. Ukuran & Elemen Desain", "📐 2. Layout Lembar Cetak", "👁️ 3. Preview & Generate"])

    # ----------------------------------------------------
    # TAB 1: UKURAN KARTU & POSISI ELEMEN (QR & NUMERATOR)
    # ----------------------------------------------------
    with tab1:
        st.subheader("Ukuran Fisik Desain (Satuan mm)")
        card_w_mm = st.number_input("Lebar Fisik Desain (mm)", min_value=10.0, value=90.0, step=1.0)
        card_h_mm = round(card_w_mm * aspect_ratio, 2)
        st.caption(f"Tinggi Otomatis (Sesuai Rasio Gambar): **{card_h_mm} mm**")

        col_qr, col_num = st.columns(2)

        # Setting QR Code
        with col_qr:
            st.markdown("### 📲 Pengaturan QR Code")
            qr_col = st.selectbox("Kolom Data untuk QR:", df.columns, index=0)
            qr_size_mm = st.slider("Ukuran QR (mm)", 5.0, min(card_w_mm, card_h_mm), 20.0, step=0.5)
            qr_x_mm = st.slider("Posisi QR - X dari Kiri (mm)", 0.0, card_w_mm - qr_size_mm, card_w_mm * 0.65, step=0.5)
            qr_y_mm = st.slider("Posisi QR - Y dari Atas (mm)", 0.0, card_h_mm - qr_size_mm, card_h_mm * 0.5, step=0.5)

        # Setting Numerator
        with col_num:
            st.markdown("### 🔢 Pengaturan Numerator / Teks")
            enable_num = st.checkbox("Aktifkan Numerator / Teks", value=True)
            if enable_num:
                num_col = st.selectbox("Kolom Data untuk Numerator:", df.columns, index=min(1, len(df.columns)-1))
                num_font_size = st.slider("Ukuran Font (pt)", 6, 72, 14)
                num_font_color = st.color_picker("Warna Teks", "#000000")
                num_x_mm = st.slider("Posisi Teks - X dari Kiri (mm)", 0.0, card_w_mm, card_w_mm * 0.1, step=0.5)
                num_y_mm = st.slider("Posisi Teks - Y dari Atas (mm)", 0.0, card_h_mm, card_h_mm * 0.8, step=0.5)
            else:
                num_col = None
                num_font_size, num_font_color, num_x_mm, num_y_mm = 12, "#000000", 0, 0

    # ----------------------------------------------------
    # TAB 2: LAYOUT LEMBAR CETAK (IMPOSITION / N-UP)
    # ----------------------------------------------------
    with tab2:
        st.subheader("Pengaturan Lembar Master Cetak")
        
        c1, c2 = st.columns(2)
        with c1:
            preset = st.selectbox("Preset Ukuran Kertas Master:", ["29.7 x 41.0 cm (Custom A3+)", "A3 (297 x 420 mm)", "A4 (210 x 297 mm)", "Custom"])
            if preset == "29.7 x 41.0 cm (Custom A3+)":
                sheet_w_mm, sheet_h_mm = 297.0, 410.0
            elif preset == "A3 (297 x 420 mm)":
                sheet_w_mm, sheet_h_mm = 297.0, 420.0
            elif preset == "A4 (210 x 297 mm)":
                sheet_w_mm, sheet_h_mm = 210.0, 297.0
            else:
                sheet_w_mm = st.number_input("Lebar Kertas Master (mm)", value=297.0)
                sheet_h_mm = st.number_input("Tinggi Kertas Master (mm)", value=410.0)

        with c2:
            st.markdown("### 📐 Margin & Jarak Antar Objek (Gutter)")
            gap_x_mm = st.number_input("Jarak Horisontal Antar Desain / Bleed (mm)", value=2.0, min_value=0.0, step=0.5)
            gap_y_mm = st.number_input("Jarak Vertikal Antar Desain / Bleed (mm)", value=2.0, min_value=0.0, step=0.5)
            margin_left_mm = st.number_input("Margin Kiri Kertas (mm)", value=10.0, min_value=0.0, step=1.0)
            margin_top_mm = st.number_input("Margin Atas Kertas (mm)", value=10.0, min_value=0.0, step=1.0)

        # Hitung kalkulasi kapasitas lembar
        usable_w = sheet_w_mm - margin_left_mm
        usable_h = sheet_h_mm - margin_top_mm
        
        cols_count = math.floor((usable_w + gap_x_mm) / (card_w_mm + gap_x_mm))
        rows_count = math.floor((usable_h + gap_y_mm) / (card_h_mm + gap_y_mm))
        
        cols_count = max(1, cols_count)
        rows_count = max(1, rows_count)
        items_per_sheet = cols_count * rows_count
        total_pages = math.ceil(len(df) / items_per_sheet)

        st.success(f"📊 **Kapasitas Lembar Cetak:** **{cols_count} Kolom** × **{rows_count} Baris** = **{items_per_sheet} Kartu/Lembar**. "
                   f"Total data: {len(df)} item ({total_pages} Halaman PDF).")

    # ----------------------------------------------------
    # TAB 3: PREVIEW & GENERATE PDF
    # ----------------------------------------------------
    with tab3:
        st.subheader("Pratinjau Desain Tunggal")
        
        # Rendition Preview menggunakan PIL (Skala Pixel)
        preview_scale = img_pixel_w / card_w_mm
        
        preview_img = base_img.copy().convert("RGB")
        draw = ImageDraw.Draw(preview_img)
        
        # Render Dummy QR
        dummy_qr_size_px = int(qr_size_mm * preview_scale)
        dummy_qr_x_px = int(qr_x_mm * preview_scale)
        dummy_qr_y_px = int(qr_y_mm * preview_scale)
        
        qr = qrcode.QRCode(box_size=5, border=1)
        qr.add_data("SAMPLE-QR")
        qr.make(fit=True)
        qr_img_pil = qr.make_image(fill_color="black", back_color="white").resize((dummy_qr_size_px, dummy_qr_size_px))
        preview_img.paste(qr_img_pil, (dummy_qr_x_px, dummy_qr_y_px))
        
        # Render Dummy Numerator Text
        if enable_num:
            num_x_px = int(num_x_mm * preview_scale)
            num_y_px = int(num_y_mm * preview_scale)
            font_size_px = int(num_font_size * preview_scale * 0.35) # Approx scale
            try:
                font = ImageFont.truetype("arial.ttf", font_size_px)
            except:
                font = ImageFont.load_default()
            draw.text((num_x_px, num_y_px), "INV-001 (SAMPLE)", fill=num_font_color, font=font)
        
        st.image(preview_img, width=450, caption="Preview Posisi Elemen pada Desain")

        st.divider()

        # Generate Full Master Imposition PDF
        if st.button("🚀 Generate PDF Layout Cetak Massal", type="primary"):
            with st.spinner("Menyusun layout cetak PDF... Mohon tunggu."):
                pdf_buffer = io.BytesIO()
                
                # Setup Lembar Master PDF
                sheet_w_pt = sheet_w_mm * MM2PT
                sheet_h_pt = sheet_h_mm * MM2PT
                
                c = canvas.Canvas(pdf_buffer, pagesize=(sheet_w_pt, sheet_h_pt))
                
                card_w_pt = card_w_mm * MM2PT
                card_h_pt = card_h_mm * MM2PT
                gap_x_pt = gap_x_mm * MM2PT
                gap_y_pt = gap_y_mm * MM2PT
                margin_left_pt = margin_left_mm * MM2PT
                margin_top_pt = margin_top_mm * MM2PT
                
                qr_size_pt = qr_size_mm * MM2PT
                qr_x_pt = qr_x_mm * MM2PT
                qr_y_pt = qr_y_mm * MM2PT

                num_x_pt = num_x_mm * MM2PT
                num_y_pt = num_y_mm * MM2PT

                item_idx = 0
                total_items = len(df)
                
                while item_idx < total_items:
                    # Gambar tiap grid pada lembar saat ini
                    for r in range(rows_count):
                        for col in range(cols_count):
                            if item_idx >= total_items:
                                break
                            
                            row_data = df.iloc[item_idx]
                            qr_val = str(row_data[qr_col])
                            num_val = str(row_data[num_col]) if enable_num else ""
                            
                            # Hitung Koordinat Kartu di Lembar Master (ReportLab Y dari Bawah)
                            card_left_pt = margin_left_pt + col * (card_w_pt + gap_x_pt)
                            card_top_from_top_pt = margin_top_pt + r * (card_h_pt + gap_y_pt)
                            card_bottom_pt = sheet_h_pt - card_top_from_top_pt - card_h_pt
                            
                            # 1. Gambar Desain Background
                            design_file.seek(0)
                            c.drawImage(ImageReader(design_file), card_left_pt, card_bottom_pt, width=card_w_pt, height=card_h_pt)
                            
                            # 2. Gambar QR Code
                            qr = qrcode.QRCode(box_size=10, border=1)
                            qr.add_data(qr_val)
                            qr.make(fit=True)
                            qr_obj = qr.make_image(fill_color="black", back_color="white")
                            
                            qr_mem = io.BytesIO()
                            qr_obj.save(qr_mem, format="PNG")
                            qr_mem.seek(0)
                            
                            actual_qr_y_pt = card_bottom_pt + card_h_pt - qr_y_pt - qr_size_pt
                            actual_qr_x_pt = card_left_pt + qr_x_pt
                            
                            c.drawImage(ImageReader(qr_mem), actual_qr_x_pt, actual_qr_y_pt, width=qr_size_pt, height=qr_size_pt)
                            
                            # 3. Gambar Numerator / Teks
                            if enable_num:
                                c.setFont("Helvetica-Bold", num_font_size)
                                c.setFillColor(HexColor(num_font_color))
                                actual_num_x_pt = card_left_pt + num_x_pt
                                actual_num_y_pt = card_bottom_pt + card_h_pt - num_y_pt - (num_font_size * 0.8)
                                c.drawString(actual_num_x_pt, actual_num_y_pt, num_val)
                                
                            item_idx += 1
                    
                    c.showPage() # Buat halaman baru jika sheet penuh
                
                c.save()
                pdf_buffer.seek(0)
                
                st.download_button(
                    label="📥 Download File Master PDF (Siap Impor ke CorelDRAW)",
                    data=pdf_buffer,
                    file_name="Master_Imposition_VDP.pdf",
                    mime="application/pdf"
                )
else:
    st.info("👈 Silakan upload file Desain dan File Data di sidebar untuk memulainya.")