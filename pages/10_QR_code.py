import streamlit as st
import pandas as pd
import qrcode
import zipfile
import io
import re

st.set_page_config(page_title="QR Generator Massal", layout="centered")

st.title("📦 QR Code Generator Massal (Excel → ZIP)")
st.write("Kolom A = nama file | Kolom B = link QR")

uploaded_file = st.file_uploader(
    "Upload file Excel (.xlsx)",
    type=["xlsx"]
)

def safe_filename(text):
    text = str(text)
    text = re.sub(r"[^\w\-_.]", "_", text)
    return text.strip("_")

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=0)

    if df.shape[1] < 2:
        st.error("Excel harus punya minimal 2 kolom (A: nama, B: link)")
        st.stop()

    st.success(f"File terbaca: {len(df)} baris")

    size = st.slider("Ukuran QR (px)", 200, 800, 400, 50)

    if st.button("🚀 Generate QR & Download ZIP"):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                filename_raw = row.iloc[0]   # kolom A
                link = row.iloc[1]           # kolom B

                if pd.isna(filename_raw) or pd.isna(link):
                    continue

                filename = safe_filename(filename_raw) + ".png"
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
