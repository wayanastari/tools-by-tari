import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import io
import os

def create_a4_grid_pdf(uploaded_files_data):
    if len(uploaded_files_data) != 9:
        st.error("Mohon unggah tepat 9 gambar.")
        return None

    a4_width_px = int(8.27 * 300)
    a4_height_px = int(11.69 * 300)

    a4_canvas = Image.new('RGB', (a4_width_px, a4_height_px), 'white')
    draw = ImageDraw.Draw(a4_canvas)

    margin = 50
    padding = 20
    border_width = 3
    border_color = (0, 0, 0)
    text_color = (0, 0, 0) # Warna teks (hitam RGB)

    # Coba memuat font default. Jika tidak ada, gunakan font bawaan Pillow
    try:
        # Coba muat versi bold dari font
        font_path = "arialbd.ttf"
        font_size = 40 # Ukuran font yang lebih besar
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        st.warning("Font bold tidak ditemukan ('arialbd.ttf'). Menggunakan font default Pillow atau non-bold.")
        font_size = 30 # Ukuran font yang lebih besar untuk default
        font = ImageFont.load_default() 
        text_color = (100, 100, 100) 
    # --- Akhir perubahan ---

    text_height_estimate = font_size + 10

    processed_items = []
    for uploaded_file_obj in uploaded_files_data:
        img = Image.open(uploaded_file_obj)
        
        base_filename = os.path.basename(uploaded_file_obj.name)
        filename_without_ext = os.path.splitext(base_filename)[0]

        if img.width > img.height:
            img = img.rotate(90, expand=True)

        img_max_width = (a4_width_px - (2 * margin) - (2 * padding)) // 3 - (2 * border_width)
        img_max_height = (a4_height_px - (2 * margin) - (2 * padding)) // 3 - (2 * border_width) - text_height_estimate

        if img.width / img.height > img_max_width / img_max_height:
            new_width = img_max_width
            new_height = int(img.height * (img_max_width / img.width))
        else:
            new_height = img_max_height
            new_width = int(img.width * (img_max_height / img.height))

        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        processed_items.append((img, filename_without_ext))

    slot_width = (a4_width_px - (2 * margin) - (2 * padding)) // 3
    slot_height = (a4_height_px - (2 * margin) - (2 * padding)) // 3


    for i in range(3):
        for j in range(3):
            img_index = i * 3 + j
            if img_index < len(processed_items):
                current_img, filename = processed_items[img_index]

                slot_x_start = margin + j * (slot_width + padding)
                slot_y_start = margin + i * (slot_height + padding)

                border_coords = (
                    slot_x_start,
                    slot_y_start,
                    slot_x_start + slot_width,
                    slot_y_start + slot_height
                )
                draw.rectangle(border_coords, outline=border_color, width=border_width)

                bbox = draw.textbbox((0,0), filename, font=font)
                text_width = bbox[2] - bbox[0]
                
                text_x = slot_x_start + (slot_width - text_width) // 2
                text_y = slot_y_start + border_width + 5

                draw.text((text_x, text_y), filename, fill=text_color, font=font)

                center_x = slot_x_start + (slot_width - current_img.width) // 2
                center_y = slot_y_start + text_height_estimate + border_width + 5

                a4_canvas.paste(current_img, (center_x, center_y))

    pdf_buffer = io.BytesIO()
    a4_canvas.save(pdf_buffer, format="PDF")
    pdf_bytes = pdf_buffer.getvalue()

    return pdf_bytes

st.set_page_config(layout="centered", page_title="Jejer Gambar ke A4")

st.title("Jejer 9 Gambar ke 1 Lembar A4 ")
st.write("Unggah sembilan gambar Kamu. Gambar landscape akan otomatis diputar, gambar kecil akan di-auto-scale, dan **nama file (tanpa ekstensi) akan ditampilkan di atas setiap gambar dalam huruf tebal dan lebih besar**, lalu bisa diunduh dalam format PDF.")

uploaded_files = st.file_uploader(
    "Pilih hingga 9 gambar (JPG, PNG)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 9:
        st.warning(f"Kamu mengunggah {len(uploaded_files)} gambar. Hanya 9 gambar pertama yang akan diproses.")
        uploaded_files = uploaded_files[:9]

    st.subheader("Pratinjau Gambar yang Diunggah:")
    cols = st.columns(3)
    for idx, file in enumerate(uploaded_files):
        base_filename_preview = os.path.basename(file.name)
        filename_without_ext_preview = os.path.splitext(base_filename_preview)[0]
        with cols[idx % 3]:
            st.image(file, caption=f"Gambar {idx+1}: {filename_without_ext_preview}", width=150)

    if len(uploaded_files) == 9:
        if st.button("Proses dan Buat PDF"):
            with st.spinner("Memproses gambar dan membuat PDF..."):
                pdf_data = create_a4_grid_pdf(uploaded_files)
                if pdf_data:
                    st.success("PDF berhasil dibuat!")
                    st.download_button(
                        label="Unduh PDF A4",
                        data=pdf_data,
                        file_name="gambar_a4_grid_bold_filenames.pdf",
                        mime="application/pdf"
                    )
                    st.info("Untuk mencetak, download PDF lalu buka filenya. Setelah itu print seperti biasa. ")
    else:
        st.info(f"Mohon unggah {9 - len(uploaded_files)} gambar lagi untuk memulai proses.")
else:
    st.info("Unggah gambar Kamu untuk memulai!")

st.markdown("---")
st.write("❤ Dibuat oleh Tari ❤")
