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

st.set_page_config(page_title="VDP & Layout Cetak Studio", layout="wide", initial_sidebar_state="expanded")

# CSS Compact agar padding & margin minim (Pas 1 Layar)
st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
        max-width: 98% !important;
    }
    .stNumberInput, .stSelectbox, .stSlider, .stRadio {
        margin-bottom: -10px;
    }
    div[data-testid="stImage"] img {
        max-height: 280px !important;
        object-fit: contain;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# SIDEBAR: UPLOAD FILE & INFORMASI RINGKAS
# ----------------------------------------------------
with st.sidebar:
    st.title("🖨️ VDP Studio")
    st.markdown("### 1. Upload File")
    design_file = st.file_uploader("Template (PNG/JPG)", type=["png", "jpg", "jpeg"])
    data_file = st.file_uploader("Data (CSV/Excel)", type=["csv", "xlsx"])

# ----------------------------------------------------
# MAIN DASHBOARD (1 LAYAR DASHBOARD)
# ----------------------------------------------------
if design_file and data_file:
    df = pd.read_csv(data_file) if data_file.name.endswith('.csv') else pd.read_excel(data_file)
    base_img = Image.open(design_file)

    # ----------------------------------------------------
    # TINGKAT 1: LIVE PREVIEW & RINGKASAN (PANEL ATAS)
    # ----------------------------------------------------
    col_prev_img, col_prev_info = st.columns([1.2, 0.8])

    # KITA BUTUH DUMMY INPUT DULU SEBELUM DITERAPKAN KE PREVIEW
    # Didefinisikan lewat Session State atau Tab Pengaturan di Bawah
    
    # ----------------------------------------------------
    # TINGKAT 2: PENGATURAN COMPACT TABS (PANEL BAWAH)
    # ----------------------------------------------------
    tab_dim, tab_qr, tab_num, tab_master = st.tabs([
        "📏 Ukuran & Orientasi", 
        "📲 Setting QR", 
        "🔢 Numerator & Font", 
        "📄 Kertas Master & Output"
    ])

    with tab_dim:
        c1, c2, c3 = st.columns(3)
        with c1:
            card_w_input = st.number_input("Lebar (mm)", min_value=10.0, value=90.0, step=1.0)
        with c2:
            card_h_input = st.number_input("Tinggi (mm)", min_value=10.0, value=55.0, step=1.0)
        with c3:
            orientation = st.radio("Orientasi", ["Landscape", "Portrait"], horizontal=True)

        if orientation == "Landscape":
            card_w_mm, card_h_mm = max(card_w_input, card_h_input), min(card_w_input, card_h_input)
        else:
            card_w_mm, card_h_mm = min(card_w_input, card_h_input), max(card_w_input, card_h_input)

    with tab_qr:
        q1, q2, q3, q4 = st.columns(4)
        with q1:
            qr_col = st.selectbox("Kolom QR:", df.columns, index=0)
        with q2:
            qr_size_mm = st.slider("Ukuran QR (mm)", 5.0, min(card_w_mm, card_h_mm), 20.0, step=0.5)
        with q3:
            qr_x_mm = st.slider("Posisi X (mm)", 0.0, card_w_mm - qr_size_mm, card_w_mm * 0.65, step=0.5)
        with q4:
            qr_y_mm = st.slider("Posisi Y (mm)", 0.0, card_h_mm - qr_size_mm, card_h_mm * 0.5, step=0.5)

    with tab_num:
        enable_num = st.checkbox("Aktifkan Numerator", value=True)
        if enable_num:
            n1, n2, n3, n4, n5 = st.columns([1, 1, 1, 1, 1])
            with n1:
                num_col = st.selectbox("Kolom Data:", df.columns, index=min(1, len(df.columns)-1))
            with n2:
                font_option = st.selectbox("Font:", ["Helvetica-Bold", "Helvetica", "Courier-Bold", "Times-Bold", "Upload TTF"])
                custom_font_file = st.file_uploader("File TTF", type=["ttf"]) if font_option == "Upload TTF" else None
            with n3:
                num_font_size = st.slider("Ukuran (pt)", 6, 72, 14)
                num_font_color = st.color_picker("Warna", "#000000")
            with n4:
                num_align = st.radio("Alignment", ["Kiri", "Tengah", "Kanan"], horizontal=True)
            with n5:
                num_x_mm = st.slider("Posisi X (mm)", 0.0, card_w_mm, card_w_mm * 0.1, step=0.5)
                num_y_mm = st.slider("Posisi Y (mm)", 0.0, card_h_mm, card_h_mm * 0.8, step=0.5)
        else:
            num_col, font_option, custom_font_file = None, "Helvetica-Bold", None
            num_font_size, num_font_color, num_x_mm, num_y_mm, num_align = 14, "#000000", 0.0, 0.0, "Kiri"

    with tab_master:
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            preset = st.selectbox("Preset Lembar Master:", ["29.7 x 41.0 cm (Custom)", "A3 (297 x 420 mm)", "A4 (210 x 297 mm)", "Custom"])
            sheet_w_mm = 297.0 if preset != "Custom" else st.number_input("Lebar Master (mm)", value=297.0)
            sheet_h_mm = 410.0 if "41.0" in preset else (420.0 if "A3" in preset else (297.0 if "A4" in preset else st.number_input("Tinggi Master (mm)", value=410.0)))
        with m2:
            gap_x_mm = st.number_input("Jarak H / Potong (mm)", value=2.0, min_value=0.0, step=0.5)
            gap_y_mm = st.number_input("Jarak V / Potong (mm)", value=2.0, min_value=0.0, step=0.5)
        with m3:
            margin_left_mm = st.number_input("Margin Kiri (mm)", value=10.0, min_value=0.0, step=1.0)
            margin_top_mm = st.number_input("Margin Atas (mm)", value=10.0, min_value=0.0, step=1.0)
        with m4:
            usable_w, usable_h = sheet_w_mm - margin_left_mm, sheet_h_mm - margin_top_mm
            cols_count = max(1, math.floor((usable_w + gap_x_mm) / (card_w_mm + gap_x_mm)))
            rows_count = max(1, math.floor((usable_h + gap_y_mm) / (card_h_mm + gap_y_mm)))
            items_per_sheet = cols_count * rows_count
            total_pages = math.ceil(len(df) / items_per_sheet)
            st.metric("Kapasitas Sheet", f"{items_per_sheet} Kartu/Sheet", f"Total: {total_pages} Hal")

    # ----------------------------------------------------
    # RENDER PREVIEW KE ATAS (SEKARANG SUDAH ADA INPUT DARI TAB)
    # ----------------------------------------------------
    with col_prev_img:
        resized_base_img = base_img.copy().resize((int(card_w_mm * 10), int(card_h_mm * 10))).convert("RGB")
        preview_scale = (card_w_mm * 10) / card_w_mm
        preview_img = resized_base_img.copy()
        draw = ImageDraw.Draw(preview_img)
        
        # QR Preview
        dummy_qr_size_px = int(qr_size_mm * preview_scale)
        qr = qrcode.QRCode(box_size=5, border=1)
        qr.add_data("PREVIEW")
        qr.make(fit=True)
        qr_img_pil = qr.make_image(fill_color="black", back_color="white").resize((dummy_qr_size_px, dummy_qr_size_px))
        preview_img.paste(qr_img_pil, (int(qr_x_mm * preview_scale), int(qr_y_mm * preview_scale)))
        
        # Numerator Preview
        if enable_num:
            sample_text = str(df[num_col].iloc[0]) if num_col in df.columns else "INV-001"
            font_size_px = max(12, int(num_font_size * (preview_scale / MM2PT)))
            
            font_to_use = ImageFont.load_default()
            if font_option == "Upload TTF" and custom_font_file:
                try: font_to_use = ImageFont.truetype(custom_font_file, font_size_px)
                except: pass
            else:
                for font_name in ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf"]:
                    try: font_to_use = ImageFont.truetype(font_name, font_size_px); break
                    except: pass

            anchor = "ms" if num_align == "Tengah" else ("rs" if num_align == "Kanan" else "ls")
            try: draw.text((int(num_x_mm * preview_scale), int(num_y_mm * preview_scale)), sample_text, fill=num_font_color, font=font_to_use, anchor=anchor)
            except: draw.text((int(num_x_mm * preview_scale), int(num_y_mm * preview_scale)), sample_text, fill=num_font_color, font=font_to_use)
            
        st.image(preview_img, use_container_width=True)

    with col_prev_info:
        st.caption(f"📐 Desain: **{card_w_mm} x {card_h_mm} mm** | Master: **{sheet_w_mm} x {sheet_h_mm} mm**")
        st.caption(f"📊 Layout Grid: **{cols_count} Kolom × {rows_count} Baris** ({items_per_sheet} pcs/sheet)")
        
        if st.button("🚀 Generate PDF Master HD", type="primary", use_container_width=True):
            with st.spinner("Memproses layout PDF..."):
                pdf_buffer = io.BytesIO()
                sheet_w_pt, sheet_h_pt = sheet_w_mm * MM2PT, sheet_h_mm * MM2PT
                c = canvas.Canvas(pdf_buffer, pagesize=(sheet_w_pt, sheet_h_pt))
                
                active_font_name = "Helvetica-Bold"
                if enable_num:
                    if font_option == "Upload TTF" and custom_font_file:
                        try:
                            custom_font_file.seek(0)
                            temp_font_path = "temp_user_font.ttf"
                            with open(temp_font_path, "wb") as f: f.write(custom_font_file.read())
                            pdfmetrics.registerFont(TTFont('CustomFont', temp_font_path))
                            active_font_name = 'CustomFont'
                        except: active_font_name = "Helvetica-Bold"
                    else: active_font_name = font_option

                card_w_pt, card_h_pt = card_w_mm * MM2PT, card_h_mm * MM2PT
                gap_x_pt, gap_y_pt = gap_x_mm * MM2PT, gap_y_mm * MM2PT
                margin_left_pt, margin_top_pt = margin_left_mm * MM2PT, margin_top_mm * MM2PT
                qr_size_pt, qr_x_pt, qr_y_pt = qr_size_mm * MM2PT, qr_x_mm * MM2PT, qr_y_mm * MM2PT
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
                                c.setFont(active_font_name, num_font_size)
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
    st.info("💡 Upload file desain & data di **Sidebar (Kiri)** untuk mulai.")