import streamlit as st
from pikepdf import Pdf, Encryption
from io import BytesIO

st.title("🔐 Strong PDF Encryption (AES-256)")
st.write("Upload PDF → Kunci AES-256 → Blok Print/Copy/Edit → Download")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    st.success("PDF berhasil diupload!")

    owner_pass = st.text_input("Owner Password (untuk proteksi)", type="password")

    if owner_pass and st.button("🔒 Buat PDF Strong Lock (AES-256)"):

        # 1. Baca PDF sebagai reader
        original = Pdf.open(BytesIO(uploaded_file.read()))

        # 2. Buat PDF baru (ini kunci utama agar tidak error)
        new_pdf = Pdf.new()

        # 3. Copy semua halaman manual
        for page in original.pages:
            new_pdf.pages.append(page)

        # 4. AES-256 strong encryption
        encryption = Encryption(
            owner=owner_pass,
            user="",        # buka tanpa password
            allow=[]        # semua izin diblok total
        )

        # 5. Save ke buffer
        output = BytesIO()
        new_pdf.save(output, encryption=encryption)
        output.seek(0)

        # 6. Download
        st.download_button(
            label="⬇️ Download PDF Terkunci (AES-256)",
            data=output,
            file_name="strong_locked.pdf",
            mime="application/pdf"
        )

        st.success("PDF berhasil dikunci dengan AES-256!")
