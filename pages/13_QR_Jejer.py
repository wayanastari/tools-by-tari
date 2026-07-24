import streamlit as st
import pandas as pd
import qrcode
from PIL import Image
import io
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import zipfile

st.set_page_config(page_title="VDP QR & Numerator Generator", layout="wide")

st.title("VT-Print: QR Code & Numerator Generator")
st.caption("Generate QR Code / Nomor Seri Massal untuk CorelDRAW (PDF Vector)")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("1. Upload Assets")
    design_file = st.file_uploader("Upload Desain (PNG/JPG)", type=["png", "jpg", "jpeg"])
    data_file = st.file_uploader("Upload Data (CSV/Excel)", type=["csv", "xlsx"])

    if design_file and data_file:
        # Load Data
        if data_file.name.endswith('.csv'):
            df = pd.read_csv(data_file)
        else:
            df = pd.read_excel(data_file)
        
        st.success(f"Data terbaca: {len(df)} baris")
        selected_col = st.selectbox("Pilih Kolom untuk QR Code:", df.columns)
        
        # Load Image Preview
        base_img = Image.open(design_file)
        img_w, img_h = base_img.size
        
        st.subheader("2. Pengaturan Posisi QR")
        pos_x = st.slider("Posisi X (Pixel)", 0, img_w, int(img_w * 0.7))
        pos_y = st.slider("Posisi Y (Pixel)", 0, img_h, int(img_h * 0.7))
        qr_size = st.slider("Ukuran QR (Pixel)", 50, min(img_w, img_h), 150)

with col2:
    if design_file and data_file:
        st.subheader("Preview Posisi")
        
        # Buat dummy QR untuk preview
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data("PREVIEW123")
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").resize((qr_size, qr_size))
        
        # Overlap ke preview
        preview_img = base_img.copy()
        preview_img.paste(qr_img, (pos_x, pos_y))
        st.image(preview_img, use_container_width=True, caption="Pratinjau Tata Letak")

        if st.button("Generate Semua File (PDF / Corel Ready)", type="primary"):
            pdf_buffer = io.BytesIO()
            
            # Setup Canvas PDF sesuai ukuran piksel gambar (72 dpi conversion)
            c = canvas.Canvas(pdf_buffer, pagesize=(img_w, img_h))
            
            for idx, row in df.iterrows():
                val = str(row[selected_col])
                
                # 1. Draw Background Design
                design_file.seek(0)
                c.drawImage(ImageReader(design_file), 0, 0, width=img_w, height=img_h)
                
                # 2. Generate QR Code
                qr = qrcode.QRCode(box_size=10, border=1)
                qr.add_data(val)
                qr.make(fit=True)
                qr_obj = qr.make_image(fill_color="black", back_color="white")
                
                qr_mem = io.BytesIO()
                qr_obj.save(qr_mem, format="PNG")
                qr_mem.seek(0)
                
                # ReportLab mengukur Y dari bawah ke atas, balik koordinat Y
                pdf_y = img_h - pos_y - qr_size
                c.drawImage(ImageReader(qr_mem), pos_x, pdf_y, width=qr_size, height=qr_size)
                
                c.showPage()
                
            c.save()
            pdf_buffer.seek(0)
            
            st.download_button(
                label="Download File PDF (Bisa Dibuka di CorelDRAW)",
                data=pdf_buffer,
                file_name="hasil_vdp_corel_ready.pdf",
                mime="application/pdf"
            )