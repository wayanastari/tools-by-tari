import streamlit as st
import pandas as pd
import qrcode
import zipfile
import io
import re

st.set_page_config(
    page_title="QR Generator Massal",
    layout="centered"
)

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

def format_number(n):
    if n < 100:
        return f"{n:02d}"
    return f"0{n}"

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=0)

    if df.shape[1] < 2:
        st.error("Excel harus punya minimal 2 kolom (A: nama, B: link)")
        st.stop()

    st.success(f"File terbaca: {len(df)} baris")

    # --- UPDATE: Minimal 1.8 cm ---
    st.info("💡 Disarankan minimal 1.8 cm agar QR tetap mudah dipindai oleh kamera HP.")
    size_cm = st.slider(
        "Pilih Ukuran QR (cm)",
        min_value=1.8,  # Berubah dari 1.6 ke 1.8
        max_value=10.0,
        value=1.8,
        step=0.1,
        format="%.1f cm"
    )

    # Konversi CM ke Pixel (DPI 300)
    # (1.8 / 2.54) * 300 = ~212 pixel
    pixel_size = int((size_cm / 2.54) * 300)

    if st.button("🚀 Generate QR & Download ZIP"):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                filename_raw = row.iloc[0]
                link = row.iloc[1]

                if pd.isna(filename_raw) or pd.isna(link):
                    continue

                nomor = format_number(idx + 1)
                nama = safe_filename(filename_raw)
                filename = f"{nomor}_{nama}.png"

                qr = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_Q, # High quality
                    box_size=10,
                    border=4
                )
                qr.add_data(str(link).strip())
                qr.make(fit=True)

                img = qr.make_image(
                    fill_color="black",
                    back_color="white"
                ).resize((pixel_size, pixel_size))

                img_bytes = io.BytesIO()
                # Sertakan info DPI agar ukuran fisik terbaca di Word/Photoshop
                img.save(img_bytes, format="PNG", dpi=(300, 300))

                zipf.writestr(filename, img_bytes.getvalue())

        zip_buffer.seek(0)

        st.success(f"QR berhasil dibuat (Ukuran: {size_cm} cm) 🎉")

        st.download_button(
            label=f"⬇️ Download ZIP QR ({size_cm}cm)",
            data=zip_buffer,
            file_name=f"QR_CODES_{size_cm}cm.zip",
            mime="application/zip"
        )