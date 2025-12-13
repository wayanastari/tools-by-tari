import streamlit as st
import pandas as pd
import qrcode
import zipfile
import io
import re
from urllib.parse import urlparse

st.set_page_config(page_title="QR Generator Massal", layout="centered")

st.title("📦 QR Code Generator Massal (Excel → ZIP)")
st.write("Upload file Excel berisi link, hasil QR akan di-download sebagai ZIP (PNG).")

uploaded_file = st.file_uploader(
    "Upload file Excel (.xlsx)",
    type=["xlsx"]
)

def safe_filename(text, max_length=80):
    text = re.sub(r"https?://", "", text)
    text = re.sub(r"[^\w\-_.]", "_", text)
    return text[:max_length]

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success(f"File terbaca: {len(df)} baris")

    link_column = st.selectbox(
        "Pilih kolom berisi link:",
        df.columns
    )

    size = st.slider("Ukuran QR (px)", 200, 800, 400, 50)

    if st.button("🚀 Generate QR & Download ZIP"):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, link in enumerate(df[link_column], start=1):
                if pd.isna(link):
                    continue

                link = str(link).strip()

                qr = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_Q,
                    box_size=10,
                    border=4
                )
                qr.add_data(link)
                qr.make(fit=True)

                img = qr.make_image(fill_color="black", back_color="white")
                img = img.resize((size, size))

                filename = f"{idx:04d}_{safe_filename(link)}.png"

                img_bytes = io.BytesIO()
                img.save(img_bytes, format="PNG")

                zipf.writestr(filename, img_bytes.getvalue())

        zip_buffer.seek(0)

        st.success("QR berhasil dibuat 🎉")

        st.download_button(
            label="⬇️ Download ZIP QR Code",
            data=zip_buffer,
            file_name="QR_CODES.zip",
            mime="application/zip"
        )
