import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import io
import math
import os
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

MM2PT = 72.0 / 25.4 

st.set_page_config(page_title="VDP & Layout Cetak Studio", layout="wide")

st.title("🖨️ VDP Generator & Imposition Studio")
st.caption("Aplikasi Variable Data Printing - Horizontal Controls & Instant Preview")

st.divider()

# ----------------------------------------------------
# 1. UPLOAD ASSETS
# ----------------------------------------------------
st.subheader("1. Upload File Utama")
col_up1, col_up2 = st.columns(2)

with col_up1:
    design_file = st.file_uploader("Upload Desain Template (PNG/JPG)", type=["png", "jpg", "jpeg"])

with col_up2:
    data_file = st.file_uploader("Upload Data (CSV/Excel)", type=["csv", "xlsx"])

if design_file and data_file:
    if data_file.name.endswith('.csv'):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)
        
    base_img = Image.open(design_file)
    img_pixel_w, img_pixel_h = base_img.size

    st.success(f"File terdeteksi! Total Data: **{len(df)} baris**. Resolusi Desain Original: **{img_pixel_w}x{img_pixel_h} px**.")
    st.divider()

    # ----------------------------------------------------
    # 2. DUMMY INITIAL VALUES FOR PREVIEW FIRST
    # ----------------------------------------------------
    # Kita inisialisasi kontainer untuk preview terlebih dahulu di paling atas
    st.subheader("2. Live Preview Real-time")
    preview_container = st.container()

    st.divider()
    st.subheader("3. Pengaturan Desain & Layout Cetak")

    # ----------------------------------------------------
    # 3. PENGATURAN BERJEJER KE SAMPING (3 KOLOM)
    # ----------------------------------------------------
    col_set1, col_set2, col_set3 = st.columns(3)

    # --- KOLOM 1: UKURAN DESAIN & QR CODE ---
    with col_set1:
        st.markdown("##### 📏 Dimensi & Orientasi Kartu")
        c_dim1, c_dim2 = st.columns(2)
        with c_dim1:
            card_w_input = st.number_input("Lebar (mm)", min_value=10.0, value=90.0, step=1.0)
        with c_dim2:
            card_h_input = st.number_input("Tinggi (mm)", min_value=10.0, value=55.0, step=1.0)
        
        orientation = st.radio("Orientasi Kartu", ["Landscape", "Portrait"], horizontal=True)

        if orientation == "Landscape":
            card_w_mm = max(card_w_input, card_h_input)
            card_h_mm = min(card_w_input, card_h_input)
        else:
            card_w_mm = min(card_w_input, card_h_input)
            card_h_mm = max(card_w_input, card_h_input)

        st.caption(f"Ukuran Kartu: **{card_w_mm} × {card_h_mm} mm**")

        st.markdown("---")
        st.markdown("##### 📲 Setting QR Code")
        qr_col = st.selectbox("Kolom Data QR:", df.columns, index=0)
        qr_size_mm = st.slider("Ukuran QR (mm)", 5.0, min(card_w_mm, card_h_mm), 20.0, step=0.5)
        qr_x_mm = st.slider("Posisi QR - X (mm)", 0.0, card_w_mm - qr_size_mm, card_w_mm * 0.65, step=0.5)
        qr_y_mm = st.slider("Posisi QR - Y (mm)", 0.0, card_h_mm - qr_size_mm, card_h_mm * 0.5, step=0.5)

    # --- KOLOM 2: NUMERATOR & CUSTOM FONT ---
    with col_set2:
        st.markdown("##### 🔢 Setting Numerator & Font")
        enable_num = st.checkbox("Aktifkan Numerator", value=True)
        if enable_num:
            num_col = st.selectbox("Kolom Data Numerator:", df.columns, index=min(1, len(df.columns)-1))
            
            font_option = st.selectbox("Pilih Jenis Font:", ["Helvetica-Bold", "Helvetica", "Courier-Bold", "Times-Bold", "Upload Custom TTF"])
            custom_font_file = None
            if font_option == "Upload Custom TTF":
                custom_font_file = st.file_uploader("Upload Font (.ttf)", type=["ttf"])

            c_n1, c_n2 = st.columns(2)
            with c_n1:
                num_font_size = st.slider("Ukuran Font (pt)", 6, 72, 14)
            with c_n2:
                num_font_color = st.color_picker("Warna Teks", "#000000")
            
            num_align = st.radio("Rata Teks", ["Kiri", "Tengah", "Kanan"], horizontal=True)
            num_x_mm = st.slider("Posisi Teks - X (mm)", 0.0, card_w_mm, card_w_mm * 0.1, step=0.5)
            num_y_mm = st.slider("Posisi Teks - Y (mm)", 0.0, card_h_mm, card_h_mm * 0.8, step=0.5)
        else:
            num_col = None
            font_option, custom_font_file = "Helvetica-Bold", None
            num_font_size, num_font_color, num_x_mm, num_y_mm, num_align = 14, "#000000", 0.0, 0.0, "Kiri"

    # --- KOLOM 3: KERTAS MASTER & GENERATE ---
    with col_set3:
        st.markdown("##### 📄 Kertas Master Cetak")
        preset = st.selectbox("Preset Lembar Master:", ["29.7 x 41.0 cm (Custom)", "A3 (297 x 420 mm)", "A4 (210 x 297 mm)", "Custom"])
        if preset == "29.7 x 41.0 cm (Custom)":
            sheet_w_mm, sheet_h_mm = 297.0, 410.0
        elif preset == "A3 (297 x 420 mm)":
            sheet_w_mm, sheet_h_mm = 297.0, 420.0
        elif preset == "A4 (210 x 297 mm)":
            sheet_w_mm, sheet_h_mm = 210.0, 297.0
        else:
            sheet_w_mm = st.number_input("Lebar Master (mm)", value=297.0)
            sheet_h_mm = st.number_input("Tinggi Master (mm)", value=410.0)

        st.markdown("---")
        st.markdown("##### ✂️ Jarak Potong & Margin")
        c_g1, c_g2 = st.columns(2)
        with c_g1:
            gap_x_mm = st.number_input("Jarak H (mm)", value=2.0, min_value=0.0, step=0.5)
            margin_left_mm = st.number_input("Margin Kiri (mm)", value=10.0, min_value=0.0, step=1.0)
        with c_g2:
            gap_y_mm = st.number_input("Jarak V (mm)", value=2.0, min_value=0.0, step=0.5)
            margin_top_mm = st.number_input("Margin Atas (mm)", value=10.0, min_value=0.0, step=1.0)

    # Kalkulasi Kapasitas Lembar Master
    usable_w = sheet_w_mm - margin_left_mm
    usable_h = sheet_h_mm - margin_top_mm
    cols_count = max(1, math.floor((usable_w + gap_x_mm) / (card_w_mm + gap_x_mm)))
    rows_count = max(1, math.floor((usable_h + gap_y_mm) / (card_h_mm + gap_y_mm)))
    items_per_sheet = cols_count * rows_count
    total_pages = math.ceil(len(df) / items_per_sheet)

    # ----------------------------------------------------
    # 4. RENDER PREVIEW KE CONTAINER PALING ATAS
    # ----------------------------------------------------
    with preview_container:
        p_col1, p_col2 = st.columns([1.2, 0.8])
        
        with p_col1:
            resized_base_img = base_img.copy().resize((int(card_w_mm * 10), int(card_h_mm * 10))).convert("RGB")
            preview_scale = (card_w_mm * 10) / card_w_mm
            
            preview_img = resized_base_img.copy()
            draw = ImageDraw.Draw(preview_img)
            
            # 1. Render QR
            dummy_qr_size_px = int(qr_size_mm * preview_scale)
            dummy_qr_x_px = int(qr_x_mm * preview_scale)
            dummy_qr_y_px = int(qr_y_mm * preview_scale)
            
            qr = qrcode.QRCode(box_size=5, border=1)
            qr.add_data("PREVIEW-QR")
            qr.make(fit=True)
            qr_img_pil = qr.make_image(fill_color="black", back_color="white").resize((dummy_qr_size_px, dummy_qr_size_px))
            preview_img.paste(qr_img_pil, (dummy_qr_x_px, dummy_qr_y_px))
            
            # 2. Render Numerator
            if enable_num:
                num_x_px = int(num_x_mm * preview_scale)
                num_y_px = int(num_y_mm * preview_scale)
                
                sample_text = str(df[num_col].iloc[0]) if num_col in df.columns else "INV-001"
                font_size_px = max(12, int(num_font_size * (preview_scale / MM2PT)))
                
                font_to_use = ImageFont.load_default()
                if font_option == "Upload Custom TTF" and custom_font_file is not None:
                    try:
                        font_to_use = ImageFont.truetype(custom_font_file, font_size_px)
                    except:
                        pass
                else:
                    for font_name in ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf"]:
                        try:
                            font_to_use = ImageFont.truetype(font_name, font_size_px)
                            break
                        except:
                            pass

                anchor = "ms" if num_align == "Tengah" else ("rs" if num_align == "Kanan" else "ls")

                try:
                    draw.text((num_x_px, num_y_px), sample_text, fill=num_font_color, font=font_to_use, anchor=anchor)
                except:
                    draw.text((num_x_px, num_y_px), sample_text, fill=num_font_color, font=font_to_use)
            
            st.image(preview_img, use_container_width=True, caption=f"Pratinjau ({card_w_mm} x {card_h_mm} mm)")

        with p_col2:
            st.info(
                f"📊 **Kapasitas Lembar Master:**\n"
                f"- Lembar Master: **{sheet_w_mm} x {sheet_h_mm} mm**\n"
                f"- Layout: **{cols_count} Kolom × {rows_count} Baris** (**{items_per_sheet} Kartu/Sheet**)\n"
                f"- Total Output: **{total_pages} Lembar Master PDF**"
            )

            if st.button("🚀 Generate PDF Master (CorelDRAW Ready)", type="primary", use_container_width=True):
                with st.spinner("Memproses layout cetak HD... Mohon tunggu sebentar."):
                    pdf_buffer = io.BytesIO()
                    sheet_w_pt, sheet_h_pt = sheet_w_mm * MM2PT, sheet_h_mm * MM2PT
                    c = canvas.Canvas(pdf_buffer, pagesize=(sheet_w_pt, sheet_h_pt))
                    
                    active_font_name = "Helvetica-Bold"
                    if enable_num:
                        if font_option == "Upload Custom TTF" and custom_font_file is not None:
                            try:
                                custom_font_file.seek(0)
                                temp_font_path = "temp_user_font.ttf"
                                with open(temp_font_path, "wb") as f:
                                    f.write(custom_font_file.read())
                                pdfmetrics.registerFont(TTFont('CustomFont', temp_font_path))
                                active_font_name = 'CustomFont'
                            except Exception as e:
                                active_font_name = "Helvetica-Bold"
                        else:
                            active_font_name = font_option

                    card_w_pt, card_h_pt = card_w_mm * MM2PT, card_h_mm * MM2PT
                    gap_x_pt, gap_y_pt = gap_x_mm * MM2PT, gap_y_mm * MM2PT
                    margin_left_pt, margin_top_pt = margin_left_mm * MM2PT, margin_top_mm * MM2PT
                    
                    qr_size_pt, qr_x_pt, qr_y_pt = qr_size_mm * MM2PT, qr_x_mm * MM2PT, qr_y_mm * MM2PT
                    num_x_pt, num_y_pt = num_x_mm * MM2PT, num_y_mm * MM2PT

                    item_idx = 0
                    total_items = len(df)
                    
                    while item_idx < total_items:
                        for r in range(rows_count):
                            for col in range(cols_count):
                                if item_idx >= total_items:
                                    break
                                
                                row_data = df.iloc[item_idx]
                                qr_val = str(row_data[qr_col])
                                num_val = str(row_data[num_col]) if enable_num else ""
                                
                                card_left_pt = margin_left_pt + col * (card_w_pt + gap_x_pt)
                                card_top_from_top_pt = margin_top_pt + r * (card_h_pt + gap_y_pt)
                                card_bottom_pt = sheet_h_pt - card_top_from_top_pt - card_h_pt
                                
                                # 1. Gambar Background HD
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
                                
                                # 3. Gambar Numerator
                                if enable_num:
                                    c.setFont(active_font_name, num_font_size)
                                    c.setFillColor(HexColor(num_font_color))
                                    actual_num_x_pt = card_left_pt + num_x_pt
                                    actual_num_y_pt = card_bottom_pt + card_h_pt - num_y_pt - (num_font_size * 0.8)
                                    
                                    if num_align == "Tengah":
                                        c.drawCentredString(actual_num_x_pt, actual_num_y_pt, num_val)
                                    elif num_align == "Kanan":
                                        c.drawRightString(actual_num_x_pt, actual_num_y_pt, num_val)
                                    else:
                                        c.drawString(actual_num_x_pt, actual_num_y_pt, num_val)
                                    
                                item_idx += 1
                        
                        c.showPage()
                    
                    c.save()
                    pdf_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Hasil PDF Master HD",
                        data=pdf_buffer,
                        file_name="Master_Cetak_VDP_HD.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
else:
    st.info("💡 Silakan upload **File Desain Template** dan **File Data CSV/Excel** di atas untuk mulai memproses.")