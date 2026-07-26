import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import io
import math
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

MM2PT = 72.0 / 25.4 

st.set_page_config(page_title="VDP & Layout Cetak Studio", layout="wide")

# CSS Compact & Lock Height Preview
st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
        max-width: 98% !important;
    }
    .stNumberInput, .stSelectbox, .stSlider, .stRadio {
        margin-bottom: -5px;
    }
    div[data-testid="stImage"] img {
        max-height: 240px !important;
        object-fit: contain;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.caption("🖨️ **VDP Generator & Imposition Studio** — Single Screen View")

# 1. AREA UPLOAD FILE
col_u1, col_u2 = st.columns(2)
with col_u1:
    design_file = st.file_uploader("1. Template Desain (PNG/JPG)", type=["png", "jpg", "jpeg"])
with col_u2:
    data_file = st.file_uploader("2. File Data (CSV/Excel)", type=["csv", "xlsx"])

if design_file and data_file:
    df = pd.read_csv(data_file) if data_file.name.endswith('.csv') else pd.read_excel(data_file)
    base_img = Image.open(design_file)

    st.markdown("---")

    col_prev_img, col_prev_info = st.columns([1.2, 0.8])

    # TABS PENGATURAN
    tab_dim, tab_qr, tab_num, tab_master = st.tabs([
        "📏 Ukuran & Orientasi", 
        "📲 Setting QR", 
        "🔢 Numerator & Font", 
        "📄 Kertas Master & Output"
    ])

    with tab_dim:
        c1, c2, c3 = st.columns(3)
        with c1:
            card_w_cm = st.number_input("Lebar Kartu (cm)", min_value=1.0, value=9.0, step=0.1)
        with c2:
            card_h_cm = st.number_input("Tinggi Kartu (cm)", min_value=1.0, value=5.5, step=0.1)
        with c3:
            orientation = st.radio("Orientasi", ["Landscape", "Portrait"], horizontal=True)

        # Konversi CM ke MM internal
        if orientation == "Landscape":
            card_w_mm, card_h_mm = max(card_w_cm, card_h_cm) * 10.0, min(card_w_cm, card_h_cm) * 10.0
        else:
            card_w_mm, card_h_mm = min(card_w_cm, card_h_cm) * 10.0, max(card_w_cm, card_h_cm) * 10.0

    with tab_qr:
        if "qr_x_mm" not in st.session_state:
            st.session_state.qr_x_mm = float(round(card_w_mm * 0.65, 1))
        if "qr_y_mm" not in st.session_state:
            st.session_state.qr_y_mm = float(round(card_h_mm * 0.5, 1))

        q1, q2, q3, q4 = st.columns(4)
        with q1:
            qr_col = st.selectbox("Kolom Data QR:", df.columns, index=0)
        with q2:
            qr_size_mm = st.slider("Ukuran QR (mm)", 5.0, min(card_w_mm, card_h_mm), 20.0, step=0.5)
        with q3:
            qr_x_mm = st.slider("Posisi X (mm)", 0.0, float(card_w_mm), key="qr_x_mm", step=0.5)
        with q4:
            qr_y_mm = st.slider("Posisi Y (mm)", 0.0, float(card_h_mm), key="qr_y_mm", step=0.5)

    with tab_num:
        enable_num = st.checkbox("Aktifkan Numerator", value=True)
        if enable_num:
            if "num_x_mm" not in st.session_state:
                st.session_state.num_x_mm = float(round(card_w_mm * 0.1, 1))
            if "num_y_mm" not in st.session_state:
                st.session_state.num_y_mm = float(round(card_h_mm * 0.8, 1))

            n1, n2, n3, n4, n5 = st.columns([1, 1.2, 1, 1, 1])
            with n1:
                num_col = st.selectbox("Kolom Numerator:", df.columns, index=min(1, len(df.columns)-1))
            with n2:
                font_family = st.selectbox("Jenis Font:", ["Helvetica", "Times-Roman", "Courier", "Upload TTF"])
                font_style = st.selectbox("Style Font:", ["Normal", "Bold", "Italic", "Bold Italic"])
                custom_font_file = st.file_uploader("Upload TTF", type=["ttf"]) if font_family == "Upload TTF" else None
            with n3:
                num_font_size = st.slider("Ukuran (pt)", 6, 72, 14)
                num_font_color = st.color_picker("Warna", "#000000")
            with n4:
                num_align = st.radio("Alignment Teks", ["Kiri", "Tengah", "Kanan"], horizontal=True)
            with n5:
                num_x_mm = st.slider("Posisi X Teks (mm)", 0.0, float(card_w_mm), key="num_x_mm", step=0.5)
                num_y_mm = st.slider("Posisi Y Teks (mm)", 0.0, float(card_h_mm), key="num_y_mm", step=0.5)

            # Map Font Name untuk ReportLab
            if font_family != "Upload TTF":
                style_map = {
                    "Normal": "",
                    "Bold": "-Bold",
                    "Italic": "-Oblique" if font_family == "Helvetica" else "-Italic",
                    "Bold Italic": "-BoldOblique" if font_family == "Helvetica" else "-BoldItalic"
                }
                active_pdf_font = f"{font_family}{style_map[font_style]}"
            else:
                active_pdf_font = "CustomFont"
        else:
            num_col, font_family, font_style, custom_font_file = None, "Helvetica", "Normal", None
            num_font_size, num_font_color, num_x_mm, num_y_mm, num_align = 14, "#000000", 0.0, 0.0, "Kiri"
            active_pdf_font = "Helvetica"

    with tab_master:
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            preset = st.selectbox("Preset Ukuran Kertas Master:", ["29.7 x 41.0 cm (Custom)", "A3 (29.7 x 42.0 cm)", "A4 (21.0 x 29.7 cm)", "Custom"])
            sheet_w_mm = 297.0 if preset != "Custom" else st.number_input("Lebar Master (mm)", value=297.0)
            sheet_h_mm = 410.0 if "41.0" in preset else (420.0 if "A3" in preset else (297.0 if "A4" in preset else st.number_input("Tinggi Master (mm)", value=410.0)))
        with m2:
            gap_x_mm = st.number_input("Jarak H / Potong (mm)", value=2.0, min_value=0.0, step=0.5)
            gap_y_mm = st.number_input("Jarak V / Potong (mm)", value=2.0, min_value=0.0, step=0.5)
        with m3:
            margin_left_mm = st.number_input("Margin Kiri Kertas (mm)", value=10.0, min_value=0.0, step=1.0)
            margin_top_mm = st.number_input("Margin Atas Kertas (mm)", value=10.0, min_value=0.0, step=1.0)
        with m4:
            usable_w, usable_h = sheet_w_mm - margin_left_mm, sheet_h_mm - margin_top_mm
            cols_count = max(1, math.floor((usable_w + gap_x_mm) / (card_w_mm + gap_x_mm)))
            rows_count = max(1, math.floor((usable_h + gap_y_mm) / (card_h_mm + gap_y_mm)))
            items_per_sheet = cols_count * rows_count
            total_pages = math.ceil(len(df) / items_per_sheet)
            st.metric("Kapasitas Cetak", f"{items_per_sheet} Kartu/Sheet", f"Total: {total_pages} Hal")

    # LIVE PREVIEW
    with col_prev_img:
        resized_base_img = base_img.copy().resize((int(card_w_mm * 10), int(card_h_mm * 10))).convert("RGB")
        preview_scale = (card_w_mm * 10) / card_w_mm
        preview_img = resized_base_img.copy()
        draw = ImageDraw.Draw(preview_img)
        
        safe_qr_x = min(qr_x_mm, card_w_mm - qr_size_mm)
        safe_qr_y = min(qr_y_mm, card_h_mm - qr_size_mm)

        # QR Preview
        dummy_qr_size_px = int(qr_size_mm * preview_scale)
        qr = qrcode.QRCode(box_size=5, border=1)
        qr.add_data("PREVIEW")
        qr.make(fit=True)
        qr_img_pil = qr.make_image(fill_color="black", back_color="white").resize((dummy_qr_size_px, dummy_qr_size_px))
        preview_img.paste(qr_img_pil, (int(safe_qr_x * preview_scale), int(safe_qr_y * preview_scale)))
        
        # Numerator Preview
        if enable_num:
            sample_text = str(df[num_col].iloc[0]) if num_col in df.columns else "INV-001"
            font_size_px = max(12, int(num_font_size * (preview_scale / MM2PT)))
            
            font_to_use = ImageFont.load_default()
            if font_family == "Upload TTF" and custom_font_file:
                try: font_to_use = ImageFont.truetype(custom_font_file, font_size_px)
                except: pass
            else:
                # Menyiapkan fallback font PIL berdasarkan style pilihan
                font_candidates = []
                if "Bold" in font_style and "Italic" in font_style: font_candidates = ["arialbi.ttf", "DejaVuSans-BoldOblique.ttf"]
                elif "Bold" in font_style: font_candidates = ["arialbd.ttf", "DejaVuSans-Bold.ttf"]
                elif "Italic" in font_style: font_candidates = ["ariali.ttf", "DejaVuSans-Oblique.ttf"]
                else: font_candidates = ["arial.ttf", "DejaVuSans.ttf"]

                for font_name in font_candidates:
                    try: font_to_use = ImageFont.truetype(font_name, font_size_px); break
                    except: pass

            anchor = "ms" if num_align == "Tengah" else ("rs" if num_align == "Kanan" else "ls")
            try: draw.text((int(num_x_mm * preview_scale), int(num_y_mm * preview_scale)), sample_text, fill=num_font_color, font=font_to_use, anchor=anchor)
            except: draw.text((int(num_x_mm * preview_scale), int(num_y_mm * preview_scale)), sample_text, fill=num_font_color, font=font_to_use)
            
        st.image(preview_img, use_container_width=True)

    with col_prev_info:
        st.caption(f"📐 Desain: **{card_w_cm} x {card_h_cm} cm** ({card_w_mm} x {card_h_mm} mm) | Master: **{sheet_w_mm/10} x {sheet_h_mm/10} cm**")
        st.caption(f"📊 Layout Grid: **{cols_count} Kolom × {rows_count} Baris** ({items_per_sheet} pcs/sheet)")
        
        if st.button("🚀 Generate PDF Master HD", type="primary", use_container_width=True):
            with st.spinner("Memproses layout PDF HD... Mohon tunggu..."):
                pdf_buffer = io.BytesIO()
                sheet_w_pt, sheet_h_pt = sheet_w_mm * MM2PT, sheet_h_mm * MM2PT
                c = canvas.Canvas(pdf_buffer, pagesize=(sheet_w_pt, sheet_h_pt))
                
                if enable_num:
                    if font_family == "Upload TTF" and custom_font_file:
                        try:
                            custom_font_file.seek(0)
                            temp_font_path = "temp_user_font.ttf"
                            with open(temp_font_path, "wb") as f: f.write(custom_font_file.read())
                            pdfmetrics.registerFont(TTFont('CustomFont', temp_font_path))
                            active_pdf_font = 'CustomFont'
                        except: active_pdf_font = "Helvetica-Bold"

                card_w_pt, card_h_pt = card_w_mm * MM2PT, card_h_mm * MM2PT
                gap_x_pt, gap_y_pt = gap_x_mm * MM2PT, gap_y_mm * MM2PT
                margin_left_pt, margin_top_pt = margin_left_mm * MM2PT, margin_top_mm * MM2PT
                qr_size_pt, qr_x_pt, qr_y_pt = qr_size_mm * MM2PT, safe_qr_x * MM2PT, safe_qr_y * MM2PT
                num_x_pt, num_y_pt = num_x_mm * MM2PT, num_y_mm * MM2PT

                item_idx, total_items = 0, len(df)
                while item_idx < total_items:
                    for r in range(rows_count):
                        for col in range(cols_count):
                            if item_idx >= total_items: break
                            row_data = df.iloc[item_idx]
                            qr_val, num_val = str(row_data[qr_col]), str(row_data[num_col]) if enable_num else ""
                            card_left_pt = margin_left_pt + col * (card_w_pt + gap_x_pt)
                            card_top_from_top_pt = margin_top_pt + r * (card_h_pt + gap_y_pt)
                            card_bottom_pt = sheet_h_pt - card_top_from_top_pt - card_h_pt
                            
                            design_file.seek(0)
                            c.drawImage(ImageReader(design_file), card_left_pt, card_bottom_pt, width=card_w_pt, height=card_h_pt)
                            
                            qr = qrcode.QRCode(box_size=10, border=1)
                            qr.add_data(qr_val)
                            qr.make(fit=True)
                            qr_mem = io.BytesIO()
                            qr.make_image(fill_color="black", back_color="white").save(qr_mem, format="PNG")
                            qr_mem.seek(0)
                            c.drawImage(ImageReader(qr_mem), card_left_pt + qr_x_pt, card_bottom_pt + card_h_pt - qr_y_pt - qr_size_pt, width=qr_size_pt, height=qr_size_pt)
                            
                            if enable_num:
                                c.setFont(active_pdf_font, num_font_size)
                                c.setFillColor(HexColor(num_font_color))
                                actual_num_x_pt = card_left_pt + num_x_pt
                                actual_num_y_pt = card_bottom_pt + card_h_pt - num_y_pt - (num_font_size * 0.8)
                                if num_align == "Tengah": c.drawCentredString(actual_num_x_pt, actual_num_y_pt, num_val)
                                elif num_align == "Kanan": c.drawRightString(actual_num_x_pt, actual_num_y_pt, num_val)
                                else: c.drawString(actual_num_x_pt, actual_num_y_pt, num_val)
                            item_idx += 1
                    c.showPage()
                c.save()
                pdf_buffer.seek(0)
                st.download_button("📥 Download PDF Master", data=pdf_buffer, file_name="Master_Cetak_VDP.pdf", mime="application/pdf", use_container_width=True)
else:
    st.info("💡 Silakan upload kedua file di atas untuk membuka workspace pratinjau & pengaturan layout.")