import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import os
import io
import zipfile # Import library zipfile

FONT_FILE_NAME = "Poppins-Bold.ttf" 
# Fallback font jika font kustom tidak ditemukan (misalnya "arial.ttf" atau "DejaVuSans-Bold.ttf")
FALLBACK_FONT_NAME = "arial.ttf" 

# Fungsi untuk memproses gambar
def process_single_image(background_img, product_img_file, watermark_img, target_size_bytes):
    """
    Menggabungkan background, produk, dan watermark, lalu menyesuaikan kualitas JPG
    agar ukuran file di bawah target_size_bytes.
    Mengembalikan gambar PIL yang sudah diproses, data biner, dan nama file.
    """
    try:
        # Buka gambar produk dari file upload (Streamlit menggunakan BytesIO)
        product_image = Image.open(product_img_file).convert("RGBA")

        # Salin background untuk setiap produk agar tidak menimpa
        combined_image = background_img.copy()

        # Atur ukuran produk agar sesuai (misal, 60% dari lebar background)
        bg_width, bg_height = combined_image.size
        product_width_target = int(bg_width * 0.7)
        product_height_target = int(product_image.size[1] * (product_width_target / product_image.size[0]))
        product_image = product_image.resize((product_width_target, product_height_target), Image.Resampling.LANCZOS)

        # Posisi produk di tengah background (geser sedikit ke atas untuk nama produk)
        product_x = (bg_width - product_width_target) // 2
        product_y = (bg_height - product_height_target) // 2 - int(bg_height * 0.05) 

        # Tempel produk ke background
        combined_image.paste(product_image, (product_x, product_y), product_image)

        # --- Modifikasi: Tambahkan nama produk (Capslock, Font Poppins Bold) ---
        product_name_raw = os.path.splitext(product_img_file.name)[0] 
        product_name_caps = product_name_raw.upper() # Otomatis Capslock

        draw = ImageDraw.Draw(combined_image)
        try:
            font_size = int(bg_height * 0.08) 
            # Coba load font Poppins Bold dari file yang di-upload
            font = ImageFont.truetype(FONT_FILE_NAME, font_size)
            # st.info(f"Menggunakan font: {FONT_FILE_NAME}") # Komentar ini karena akan sering muncul
        except IOError:
            # st.warning(f"Font '{FONT_FILE_NAME}' tidak ditemukan. Mencoba font fallback '{FALLBACK_FONT_NAME}'. Pastikan '{FONT_FILE_NAME}' ada di repositori Anda.") # Komentar ini
            try:
                font = ImageFont.truetype(FALLBACK_FONT_NAME, font_size)
            except IOError:
                # Jika fallback pun gagal, gunakan font default Pillow
                st.error("Gagal memuat font fallback. Menggunakan font default Pillow.")
                font = ImageFont.load_default()
        except Exception as e:
            # Tangani error lain yang mungkin terjadi saat memuat font
            st.error(f"Gagal memuat font, menggunakan font default Pillow. Error: {e}")
            font = ImageFont.load_default()

        text_color = (0, 0, 0, 255)  # Hitam
        text_bbox = draw.textbbox((0,0), product_name_caps, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        text_x = (bg_width - text_width) // 2
        text_y = product_y + product_height_target + int(bg_height * 0.05) 

        draw.text((text_x, text_y), product_name_caps, font=font, fill=text_color)

        # --- Watermark di Tengah, Besar, Sangat Transparan ---
        wm_width, wm_height = watermark_img.size
        target_wm_width = int(bg_width * 0.7) 
        target_wm_height = int(wm_height * (target_wm_width / wm_width))
        resized_watermark = watermark_img.resize((target_wm_width, target_wm_height), Image.Resampling.LANCZOS)

        # Ubah opasitas watermark
        if resized_watermark.mode != 'RGBA':
            resized_watermark = resized_watermark.convert('RGBA')

        alpha = resized_watermark.split()[3] 
        alpha = Image.eval(alpha, lambda x: x * 0.1) # 10% dari opasitas asli
        resized_watermark.putalpha(alpha)

        # Posisi watermark di tengah gambar
        wm_x = (bg_width - resized_watermark.size[0]) // 2
        wm_y = (bg_height - resized_watermark.size[1]) // 2

        combined_image.paste(resized_watermark, (wm_x, wm_y), resized_watermark)

        # --- Bagian Penyimpanan Gambar Otomatis di Bawah 1 MB ---
        if combined_image.mode == 'RGBA':
            combined_image = combined_image.convert('RGB')

        target_size_bytes = 1 * 1024 * 1024
        current_quality = 90
        quality_step = 5

        img_byte_arr = io.BytesIO()
        while current_quality >= 20: 
            img_byte_arr.seek(0)
            img_byte_arr.truncate(0)

            combined_image.save(img_byte_arr, format='JPEG', quality=current_quality, optimize=True)
            file_size = img_byte_arr.tell()

            if file_size <= target_size_bytes:
                # st.info(f"Produk '{product_name_raw}': Ukuran {file_size / (1024*1024):.2f} MB (Kualitas: {current_quality}).") # Komentar ini
                return combined_image, img_byte_arr.getvalue(), product_name_raw + ".jpg"
            else:
                current_quality -= quality_step
                # st.info(f"Produk '{product_name_raw}': Ukuran {file_size / (1024*1024):.2f} MB, mencoba kualitas {current_quality}...") # Komentar ini

        # st.warning(f"Produk '{product_name_raw}' masih di atas 1MB ({file_size / (1024*1024):.2f} MB) meskipun kualitas sudah diturunkan maksimal. Pertimbangkan resolusi gambar sumber.") # Komentar ini
        return combined_image, img_byte_arr.getvalue(), product_name_raw + ".jpg"

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses gambar: {e}")
        return None, None, None


# --- Bagian Antarmuka Streamlit ---
st.set_page_config(page_title="Penggabung Gambar Produk Otomatis", layout="centered")

st.title("🖼️ Penggabung Gambar Produk Otomatis")
st.markdown(f"""
Aplikasi ini menggabungkan **background**, **foto produk**, dan **watermark**.
**Nama produk akan otomatis menjadi CAPSLOCK dengan font Poppins Bold**, dan **watermark akan ditempatkan di tengah, lebih besar, dan sangat transparan**.
Hasil gambar otomatis dioptimasi menjadi **JPEG di bawah 1 MB**.
""")

# Input Background
st.header("1. Unggah Background (1 file)")
uploaded_background = st.file_uploader("Pilih file gambar background (JPG/PNG)", type=["jpg", "jpeg", "png"])
background_image_pil = None
if uploaded_background:
    background_image_pil = Image.open(uploaded_background).convert("RGBA")
    st.image(background_image_pil, caption="Background Terpilih", use_column_width=True)

# Input Foto Produk (banyak)
st.header("2. Unggah Foto Produk (Banyak File)")
uploaded_products = st.file_uploader("Pilih satu atau lebih file gambar produk (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Input Watermark
st.header("3. Unggah Watermark (1 file)")
uploaded_watermark = st.file_uploader("Pilih file gambar watermark (JPG/PNG, disarankan PNG transparan)", type=["jpg", "jpeg", "png"])
watermark_image_pil = None
if uploaded_watermark:
    watermark_image_pil = Image.open(uploaded_watermark).convert("RGBA")
    st.image(watermark_image_pil, caption="Watermark Terpilih", use_column_width=True)

st.markdown("---")

# Tombol Proses
if st.button("Mulai Penggabungan dan Optimasi"):
    if not uploaded_background:
        st.error("Mohon unggah file background terlebih dahulu.")
    elif not uploaded_products:
        st.error("Mohon unggah setidaknya satu foto produk.")
    elif not uploaded_watermark:
        st.error("Mohon unggah file watermark terlebih dahulu.")
    else:
        st.write("Memulai pemrosesan gambar... Ini mungkin memerlukan waktu beberapa saat jika ada banyak gambar.")
        
        # Buffer untuk menyimpan file ZIP
        zip_buffer = io.BytesIO()
        
        # Buat objek ZipFile
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
            with st.spinner("Sedang memproses gambar dan mengemasnya dalam file ZIP, harap tunggu..."):
                target_size_bytes = 1 * 1024 * 1024 # 1 MB

                # List untuk menyimpan informasi status pemrosesan
                status_messages = []

                for i, product_file in enumerate(uploaded_products):
                    st.subheader(f"Memproses Produk: {product_file.name}")
                    processed_pil_image, processed_bytes, output_filename = process_single_image(
                        background_image_pil, product_file, watermark_image_pil, target_size_bytes
                    )

                    if processed_pil_image and processed_bytes:
                       
                        st.image(processed_pil_image, caption=f"Hasil untuk {output_filename}", use_column_width=True)
                        
                        # Tambahkan gambar ke dalam ZIP
                        zf.writestr(output_filename, processed_bytes)
                        
                        # Simpan pesan status
                        status_messages.append(f"✅ '{output_filename}' berhasil diproses.")
                    else:
                        status_messages.append(f"❌ Gagal memproses '{product_file.name}'.")

            # Tampilkan ringkasan status setelah semua selesai
            for msg in status_messages:
                st.write(msg)

        # Setelah loop selesai dan ZIP file sudah ditutup
        if status_messages: # Pastikan ada sesuatu yang diproses
            zip_buffer.seek(0) # Kembali ke awal buffer
            st.success("Semua gambar berhasil diproses dan dikemas dalam file ZIP!")
            
            st.download_button(
                label="Unduh Semua Gambar (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="hasil_gambar_produk.zip",
                mime="application/zip"
            )
        else:
            st.warning("Tidak ada gambar yang berhasil diproses.")
