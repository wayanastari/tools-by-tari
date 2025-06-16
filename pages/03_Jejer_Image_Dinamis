import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import io
import os

# Fungsi pembantu untuk memecah teks menjadi beberapa baris agar muat dalam lebar tertentu
def wrap_text(text, font, max_width_px):
    """
    Memecah teks menjadi beberapa baris jika melebihi lebar maksimum.
    """
    lines = []
    if not text:
        return [""] # Mengembalikan baris kosong jika teks kosong
    
    words = text.split(' ')
    current_line_words = []
    
    for word in words:
        # Buat baris percobaan dengan kata berikutnya
        test_line = ' '.join(current_line_words + [word])
        
        # Dapatkan ukuran bounding box dari baris percobaan
        # getbbox() mengembalikan (left, top, right, bottom)
        bbox = font.getbbox(test_line) 
        text_width = bbox[2] - bbox[0] # Lebar teks = right - left

        # Jika baris percobaan tidak melebihi lebar maksimum, tambahkan kata
        if text_width <= max_width_px:
            current_line_words.append(word)
        else:
            # Jika melebihi, simpan baris saat ini dan mulai baris baru dengan kata ini
            if current_line_words: # Pastikan ada sesuatu untuk ditambahkan sebelum memulai baris baru
                lines.append(' '.join(current_line_words))
            current_line_words = [word] # Kata saat ini menjadi awal baris baru
            
            # Khusus untuk kata yang sangat panjang sehingga melebihi max_width_px sendirian
            # Ini mungkin masih memanjang jika satu kata terlalu panjang, tapi ini trade-off
            # Jika sangat penting, perlu logika pemotongan kata.
            bbox = font.getbbox(word)
            if bbox[2] - bbox[0] > max_width_px:
                # Jika satu kata saja sudah terlalu panjang, pecah kata jika memungkinkan
                # Untuk penyederhanaan, kita biarkan kata panjang ini apa adanya
                # Sebagai alternatif, bisa dipotong atau ditambahkan elipsis.
                pass 

    # Tambahkan baris terakhir yang tersisa setelah loop
    if current_line_words:
        lines.append(' '.join(current_line_words))
    
    return "\n".join(lines)


def create_a4_grid_pdf(uploaded_files_data):
    """
    Memproses daftar file gambar, menatanya secara dinamis ke halaman A4,
    dan mengembalikan PDF multi-halaman dalam bentuk byte.
    Gambar landscape akan diputar, gambar akan diskalakan agar muat,
    dan nama file akan ditampilkan di atas setiap gambar dengan wrap text.
    """
    if not uploaded_files_data:
        st.error("Mohon unggah gambar.")
        return None

    # Dimensi A4 dalam piksel (pada 300 DPI untuk kualitas cetak yang baik)
    A4_WIDTH_PX = int(8.27 * 300)  # Lebar A4: 2481 piksel
    A4_HEIGHT_PX = int(11.69 * 300) # Tinggi A4: 3507 piksel

    # Margin dan pengaturan spasi untuk tata letak
    MARGIN = 50  # Margin di sekitar seluruh konten di halaman
    ITEM_SPACING = 30 # Jarak antar item gambar (horizontal dan vertikal)
    BORDER_WIDTH = 3 # Ketebalan bingkai di sekitar setiap slot gambar
    BORDER_COLOR = (0, 0, 0) # Warna bingkai: Hitam
    TEXT_COLOR = (0, 0, 0)   # Warna teks: Hitam

    # Pengaturan font untuk nama file
    try:
        font_path = "arialbd.ttf"
        font_size = 40
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        st.warning("Font bold ('arialbd.ttf') tidak ditemukan. Menggunakan font default Pillow.")
        font_size = 30
        font = ImageFont.load_default()
        TEXT_COLOR = (100, 100, 100) # Menggunakan warna abu-abu untuk font default

    all_pdf_pages = [] # List untuk menyimpan setiap objek gambar A4 yang dihasilkan
    current_page = None # Objek gambar untuk halaman A4 saat ini
    draw = None         # Objek ImageDraw untuk menggambar di halaman saat ini

    # Posisi awal untuk menempatkan item di halaman saat ini
    current_x = MARGIN
    current_y = MARGIN
    current_row_max_height = 0 # Menyimpan tinggi maksimum item di baris saat ini

    # Hitung area konten yang tersedia di dalam margin halaman
    page_content_width = A4_WIDTH_PX - (2 * MARGIN)
    page_content_height = A4_HEIGHT_PX - (2 * MARGIN)

    # Target kolom per baris untuk penskalaan awal, mencoba minimal 3 gambar per baris
    TARGET_COLS_PER_ROW = 3 
    # Lebar maksimum yang diizinkan untuk setiap item (termasuk border dan padding)
    # Ini menentukan seberapa besar gambar bisa diskalakan agar muat 3 per baris
    MAX_ITEM_WIDTH_PER_COL = (page_content_width - (TARGET_COLS_PER_ROW - 1) * ITEM_SPACING) // TARGET_COLS_PER_ROW

    # Loop melalui setiap file gambar yang diunggah
    for uploaded_file_obj in uploaded_files_data:
        try:
            img_original = Image.open(uploaded_file_obj)
        except Exception as e:
            st.warning(f"Tidak dapat membuka file {uploaded_file_obj.name}. Melompati. Error: {e}")
            continue
        
        # Ekstrak nama file tanpa ekstensi untuk ditampilkan
        base_filename = os.path.basename(uploaded_file_obj.name)
        filename_without_ext = os.path.splitext(base_filename)[0]

        # Putar gambar landscape ke potret (jika lebar > tinggi)
        if img_original.width > img_original.height:
            img_original = img_original.rotate(90, expand=True)

        scaled_img = img_original.copy()
        # Skalakan gambar agar lebar maksimumnya sesuai dengan MAX_ITEM_WIDTH_PER_COL
        # Tinggi akan disesuaikan secara proporsional.
        # Jika gambar sudah lebih kecil, biarkan saja kecuali terlalu besar
        if scaled_img.width > MAX_ITEM_WIDTH_PER_COL:
            scaled_img.thumbnail((MAX_ITEM_WIDTH_PER_COL, A4_HEIGHT_PX), Image.Resampling.LANCZOS)
        
        # Wrap teks nama file
        # Lebar maksimum untuk teks adalah lebar item dikurangi padding dan border
        wrapped_filename = wrap_text(filename_without_ext, font, scaled_img.width - (2 * BORDER_WIDTH) + 10) # +10 untuk toleransi

        # Hitung tinggi teks yang sebenarnya setelah di-wrap
        # Gunakan multiline_textbbox untuk teks multi-baris
        bbox_text = draw.multiline_textbbox((0,0), wrapped_filename, font=font) if current_page else font.getbbox(wrapped_filename)
        text_actual_height = bbox_text[3] - bbox_text[1] + 10 # Tambah sedikit padding

        # Hitung total ruang yang akan ditempati oleh item ini (gambar + teks + bingkai)
        item_total_width = scaled_img.width + (2 * BORDER_WIDTH) 
        item_total_height = scaled_img.height + text_actual_height + (2 * BORDER_WIDTH) 

        # --- Logika Penempatan ---
        # 1. Cek apakah item muat secara horizontal di baris saat ini
        # current_x + lebar_item_ini + jarak_antar_item > batas_kanan_halaman
        if current_x + item_total_width + ITEM_SPACING > A4_WIDTH_PX - MARGIN:
            # Tidak cukup ruang di baris saat ini, pindah ke baris berikutnya
            current_x = MARGIN # Kembali ke margin kiri
            current_y += current_row_max_height + ITEM_SPACING # Turunkan Y sebesar tinggi maks baris sebelumnya + jarak
            current_row_max_height = 0 # Reset tinggi maks untuk baris baru

        # 2. Cek apakah item muat secara vertikal di halaman saat ini (setelah potensi pindah baris)
        # current_y + tinggi_item_ini + jarak_antar_item > batas_bawah_halaman
        if current_y + item_total_height + ITEM_SPACING > A4_HEIGHT_PX - MARGIN:
            # Tidak cukup ruang di halaman saat ini, mulai halaman baru
            if current_page: # Hanya tambahkan halaman jika sudah ada konten di dalamnya
                all_pdf_pages.append(current_page)
            # Inisialisasi halaman baru
            current_page = None 
            current_x = MARGIN
            current_y = MARGIN
            current_row_max_height = 0
        
        # Jika belum ada halaman atau sudah dipaksa untuk membuat halaman baru:
        if current_page is None:
            current_page = Image.new('RGB', (A4_WIDTH_PX, A4_HEIGHT_PX), 'white')
            draw = ImageDraw.Draw(current_page)

        # Gambar bingkai (slot) untuk item saat ini
        # Koordinat bingkai mendefinisikan batas luar item, termasuk bingkai itu sendiri.
        border_x0 = current_x
        border_y0 = current_y
        border_x1 = border_x0 + item_total_width
        border_y1 = border_y0 + item_total_height
        draw.rectangle((border_x0, border_y0, border_x1, border_y1), outline=BORDER_COLOR, width=BORDER_WIDTH)

        # Tempatkan teks di dalam bingkai, di tengah horizontal di bagian atas
        # Gunakan draw.multiline_text untuk teks multi-baris
        text_x = border_x0 + BORDER_WIDTH + (item_total_width - (2 * BORDER_WIDTH) - (bbox_text[2] - bbox_text[0])) // 2
        text_y = border_y0 + BORDER_WIDTH + 5 # Sedikit offset dari bingkai atas
        draw.multiline_text((text_x, text_y), wrapped_filename, fill=TEXT_COLOR, font=font, align="center")

        # Tempatkan gambar di dalam bingkai, di bawah teks, di tengah horizontal
        img_x = border_x0 + BORDER_WIDTH + (item_total_width - (2 * BORDER_WIDTH) - scaled_img.width) // 2
        img_y = border_y0 + BORDER_WIDTH + text_actual_height + 5 # Offset dari bingkai atas dan teks
        current_page.paste(scaled_img, (img_x, img_y))

        # Perbarui posisi X untuk item berikutnya di baris yang sama
        current_x += item_total_width + ITEM_SPACING
        # Perbarui tinggi maksimum untuk item di baris saat ini (penting untuk menghitung baris berikutnya)
        current_row_max_height = max(current_row_max_height, item_total_height)
    
    # Setelah loop selesai, tambahkan halaman terakhir jika ada konten di dalamnya
    if current_page:
        all_pdf_pages.append(current_page)

    # Simpan semua halaman A4 yang dihasilkan sebagai satu PDF multi-halaman
    pdf_buffer = io.BytesIO()
    if all_pdf_pages:
        first_page = all_pdf_pages[0]
        remaining_pages = all_pdf_pages[1:] if len(all_pdf_pages) > 1 else []
        first_page.save(pdf_buffer, format="PDF", save_all=True, append_images=remaining_pages)
    
    pdf_bytes = pdf_buffer.getvalue()
    return pdf_bytes

# --- Streamlit UI ---
st.set_page_config(layout="centered", page_title="Jejer Gambar ke A4 Multi-Page")

st.title("Jejer Gambar ke Lembar A4 (Otomatis & Fleksibel)")
st.write("""
Unggah gambar Kamu. Gambar akan secara otomatis diatur pada halaman A4. Jika gambar terlalu besar untuk satu baris, akan pindah ke baris baru. Jika tidak muat di satu halaman, akan dilanjutkan ke halaman berikutnya.
Gambar landscape akan otomatis diputar, gambar akan diskalakan agar muat di halaman, dan **nama file (tanpa ekstensi) akan ditampilkan di atas setiap gambar, memanjang ke bawah jika terlalu panjang**.
Setelah diproses, Kamu bisa mengunduh hasilnya dalam format PDF multi-halaman.
""")

uploaded_files = st.file_uploader(
    "Pilih gambar (JPG, PNG)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    st.subheader("Pratinjau Gambar yang Diunggah:")
    # Tampilkan pratinjau dalam kolom
    cols = st.columns(3)
    for idx, file in enumerate(uploaded_files):
        base_filename_preview = os.path.basename(file.name)
        filename_without_ext_preview = os.path.splitext(base_filename_preview)[0]
        with cols[idx % 3]: # Berputar melalui 3 kolom
            st.image(file, caption=f"Gambar {idx+1}: {filename_without_ext_preview}", width=150)

    # Tombol proses dan unduh
    if st.button("Proses dan Buat PDF"):
        with st.spinner("Memproses gambar dan membuat PDF..."):
            pdf_data = create_a4_grid_pdf(uploaded_files)
            if pdf_data:
                st.success("PDF berhasil dibuat!")
                st.download_button(
                    label="Unduh PDF A4 (Multi-Halaman)",
                    data=pdf_data,
                    file_name="gambar_a4_otomatis_multi_page.pdf",
                    mime="application/pdf"
                )
                st.info("Untuk mencetak, unduh PDF lalu buka filenya. Setelah itu cetak seperti biasa.")
            else:
                st.error("Gagal membuat PDF. Pastikan format gambar benar atau coba lagi.")
else:
    st.info("Unggah gambar Kamu untuk memulai!")

st.markdown("---")
st.write("❤ Dibuat oleh Tari ❤")
