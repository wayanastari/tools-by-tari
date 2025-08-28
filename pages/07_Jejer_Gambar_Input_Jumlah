import streamlit as st
from PIL import Image
import io

def create_a4_canvas():
    """Membuat kanvas A4 baru dengan ukuran potret (300 DPI)."""
    # Ukuran A4 dalam piksel pada 300 DPI
    a4_width, a4_height = 2480, 3508
    return Image.new('RGB', (a4_width, a4_height), 'white')

def process_and_save_images(image_data):
    """
    Memproses dan menyimpan gambar ke dalam satu atau lebih lembar A4.
    
    Args:
        image_data (list): Daftar tuple berisi (object_gambar, jumlah_cetak).
    """
    all_images_to_print = []
    
    for img_obj, count in image_data:
        # Tambahkan objek gambar ke daftar sebanyak jumlah yang diminta
        all_images_to_print.extend([img_obj] * count)

    if not all_images_to_print:
        st.warning("Tidak ada gambar yang diunggah atau dipilih untuk dicetak.")
        return None

    # Ukuran gambar optimal untuk 6 gambar per A4 (2x3 grid)
    cols, rows = 2, 3
    a4_width, a4_height = 2480, 3508
    margin = 50
    img_width = (a4_width - (cols + 1) * margin) // cols
    img_height = (a4_height - (rows + 1) * margin) // rows
    
    # Cetak gambar ke lembar-lembar A4
    page_number = 1
    current_page = create_a4_canvas()
    img_on_page = 0
    
    output_images = []

    for img in all_images_to_print:
        if img_on_page >= 6:
            output_images.append(current_page)
            page_number += 1
            current_page = create_a4_canvas()
            img_on_page = 0

        # Ukuran ulang gambar dan tempelkan
        resized_img = img.resize((img_width, img_height), Image.LANCZOS)
        
        row_idx = img_on_page // cols
        col_idx = img_on_page % cols
        
        x_pos = col_idx * (img_width + margin) + margin
        y_pos = row_idx * (img_height + margin) + margin
        
        current_page.paste(resized_img, (x_pos, y_pos))
        img_on_page += 1
        
    # Simpan halaman terakhir
    if img_on_page > 0:
        output_images.append(current_page)
    
    return output_images

st.set_page_config(layout="wide")
st.title("🖨️ Jejer Gambar Otomatis ke Kertas A4")

st.markdown("""
Aplikasi ini membantu Anda menjejerkan beberapa gambar ke dalam kertas A4.
Cukup unggah gambar dan tentukan berapa kali setiap gambar akan dicetak.
""")

uploaded_files = st.file_uploader("Pilih beberapa file gambar", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    image_data_input = []
    
    # Menampilkan file yang diunggah dan meminta input jumlah cetak
    st.subheader("Detail Cetak")
    for file in uploaded_files:
        try:
            # Buka file yang diunggah sebagai objek gambar PIL
            image = Image.open(file)
            col1, col2 = st.columns([1, 2])
            with col1:
                st.image(image, caption=f"Gambar: {file.name}", use_column_width=True)
            with col2:
                count = st.number_input(f"Berapa kali '{file.name}' akan dicetak?", min_value=1, value=1, key=file.name)
            image_data_input.append((image, count))
        except Exception as e:
            st.error(f"Gagal memuat {file.name}. Pastikan file adalah gambar yang valid.")
    
    st.markdown("---")
    
    if st.button("Proses dan Cetak"):
        with st.spinner("Sedang memproses gambar..."):
            result_pages = process_and_save_images(image_data_input)
            
            if result_pages:
                st.subheader("Hasil Akhir")
                for i, page in enumerate(result_pages):
                    st.success(f"Halaman {i+1} berhasil dibuat!")
                    
                    # Konversi gambar ke byte untuk ditampilkan dan diunduh
                    buf = io.BytesIO()
                    page.save(buf, format="PNG")
                    byte_im = buf.getvalue()
                    
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        st.image(byte_im, caption=f"Halaman {i+1}", use_column_width=True)
                    with col2:
                        st.download_button(
                            label=f"📥 Unduh Halaman {i+1}",
                            data=byte_im,
                            file_name=f"halaman_{i+1}.png",
                            mime="image/png"
                        )
