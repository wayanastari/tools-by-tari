import streamlit as st
from pikepdf import Pdf, Encryption
from io import BytesIO

st.title("🔐 Strong PDF Encryption (AES-256)")
st.write("Upload PDF → Kunci AES-256 → Blok Print/Copy/Edit → Download")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    st.success("PDF berhasil diupload!")

    owner_pass = st.text_input("Owner Password (untuk proteksi)", type="password")

    st.write("Semua permission akan diblok secara total:")
    st.write("- ❌ Print")
    st.write("- ❌ Copy")
    st.write("- ❌ Edit / Modify")
    st.write("- ❌ Fill Forms")
    st.write("- ❌ Extraction")
    st.write("- ❌ Comment / Annotation")

    if owner_pass and st.button("🔒 Buat PDF Strong Lock (AES-256)"):
        # Read file
        input_pdf = BytesIO(uploaded_file.read())
        pdf = Pdf.open(input_pdf)

        # AES-256 strong encryption (pikepdf will auto apply)
        encryption = Encryption(
            owner=owner_pass,
            user="",          # kosong → user buka tanpa password
            allow=[]          # tidak ada izin sama sekali
        )

        output = BytesIO()
        pdf.save(output, encryption=encryption)
        output.seek(0)

        st.download_button(
            label="⬇️ Download PDF Terkunci (AES-256)",
            data=output,
            file_name="strong_locked.pdf",
            mime="application/pdf"
        )

        st.success("PDF berhasil dikunci dengan AES-256!")
