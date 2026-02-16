import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import io
import os

# Function to wrap text
def wrap_text(text, font, max_width):
    lines = []
    if not text:
        return lines
    
    words = text.split(' ')
    current_line = []
    
    for word in words:
        test_line = ' '.join(current_line + [word])
        # Gunakan dummy image untuk kalkulasi bbox
        draw = ImageDraw.Draw(Image.new('RGB', (1,1)))
        bbox = draw.textbbox((0,0), test_line, font=font)
        test_width = bbox[2] - bbox[0]
        
        if test_width <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
            
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines

def create_a4_grid_pdf(uploaded_files_data):
    if not uploaded_files_data:
        st.error("Mohon unggah gambar.")
        return None

    # A4 dimensions (300 DPI)
    a4_width_px = int(8.27 * 300)
    a4_height_px = int(11.69 * 300)

    margin = 60 
    padding = 40 
    border_width = 2 
    border_color = (180, 180, 180) # Abu-abu tipis agar elegan
    text_color = (0, 0, 0)  

    # --- PENYESUAIAN UKURAN FONT ---
    # Ukuran ditingkatkan ke 45-50 agar terbaca jelas di kertas A4
    font_size = 45 
    try:
        # Mencoba mencari font tebal agar lebih terbaca
        paths = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "arialbd.ttf", 
            "Arial.ttf",
            "/Library/Fonts/Arial.ttf", # Mac
            "C:\\Windows\\Fonts\\arial.ttf" # Windows
        ]
        font = None
        for path in paths:
            if os.path.exists(path):
                font = ImageFont.truetype(path, font_size)
                break
        if font is None:
            font = ImageFont.load_default()
    except:
        font = ImageFont.load_default()

    max_image_width_per_col = (a4_width_px - (2 * margin) - (2 * padding)) // 3
    max_image_height_per_row_area = (a4_height_px - (2 * margin) - (1 * padding)) // 2
    target_aspect_ratio = 9 / 16.0 

    img_max_width = max_image_width_per_col - (2 * border_width)
    
    # Estimasi tinggi teks (Line height + spacing)
    text_line_height_estimate = font_size + 15
    estimated_total_text_height = (2 * text_line_height_estimate) + 20

    img_max_height = max_image_height_per_row_area - (2 * border_width) - estimated_total_text_height
    img_max_width = int(img_max_height * target_aspect_ratio)

    actual_drawn_slot_width = img_max_width + (2 * border_width)
    actual_drawn_slot_height_base = img_max_height + (2 * border_width) + estimated_total_text_height 

    all_pdf_pages = [] 

    for page_idx in range(0, len(uploaded_files_data), 6):
        a4_canvas = Image.new('RGB', (a4_width_px, a4_height_px), 'white')
        draw = ImageDraw.Draw(a4_canvas)
        current_batch_files = uploaded_files_data[page_idx : page_idx + 6]
        
        processed_items_for_page = []
        for uploaded_file_obj in current_batch_files:
            img = Image.open(uploaded_file_obj).convert("RGB")
            base_filename = os.path.basename(uploaded_file_obj.name)
            filename_without_ext = os.path.splitext(base_filename)[0]

            if img.width > img.height:
                img = img.rotate(90, expand=True)

            img.thumbnail((img_max_width, img_max_height), Image.Resampling.LANCZOS)
            processed_items_for_page.append((img, filename_without_ext))

        for i in range(2): 
            for j in range(3): 
                img_index_in_batch = i * 3 + j 
                if img_index_in_batch < len(processed_items_for_page):
                    current_img, filename = processed_items_for_page[img_index_in_batch]

                    grid_cell_x_start = margin + j * (actual_drawn_slot_width + padding)
                    grid_cell_y_start = margin + i * (actual_drawn_slot_height_base + padding)

                    img_x = grid_cell_x_start + (actual_drawn_slot_width - current_img.width) // 2
                    img_y = grid_cell_y_start + border_width 

                    a4_canvas.paste(current_img, (img_x, img_y))
                    
                    max_text_width = actual_drawn_slot_width - 10 
                    wrapped_lines = wrap_text(filename, font, max_text_width)
                    
                    current_text_y = img_y + current_img.height + 15 
                    
                    for line in wrapped_lines[:2]: # Batasi 2 baris agar tidak tumpang tindih
                        bbox = draw.textbbox((0,0), line, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_x = grid_cell_x_start + (actual_drawn_slot_width - text_width) // 2
                        draw.text((text_x, current_text_y), line, fill=text_color, font=font)
                        current_text_y += (bbox[3] - bbox[1]) + 10

                    # Gambar border kotak
                    draw.rectangle(
                        (grid_cell_x_start, grid_cell_y_start, 
                         grid_cell_x_start + actual_drawn_slot_width, 
                         grid_cell_y_start + actual_drawn_slot_height_base), 
                        outline=border_color, width=border_width
                    )

        all_pdf_pages.append(a4_canvas)

    pdf_buffer = io.BytesIO()
    if all_pdf_pages:
        first_page = all_pdf_pages[0]
        remaining_pages = all_pdf_pages[1:]
        first_page.save(pdf_buffer, format="PDF", save_all=True, append_images=remaining_pages)
        
    return pdf_buffer.getvalue()

# --- Streamlit UI ---
st.set_page_config(layout="centered", page_title="Bukti Transfer ke A4")

st.title("Jejer Gambar ke A4 (3x2)")
st.write("Urutan gambar akan sesuai dengan **urutan saat kamu memilih/mengunggah file**.")

uploaded_files = st.file_uploader(
    "Pilih gambar (JPG, PNG)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    # --- LOGIKA URUTAN: Kita hapus .sort() agar sesuai urutan upload ---
    # Jika ingin benar-benar yakin sesuai urutan klik, user harus pilih satu per satu 
    # atau mengandalkan urutan selection dari file explorer mereka.

    st.subheader("Pratinjau Urutan Cetak:")
    cols = st.columns(3)
    for idx, file in enumerate(uploaded_files):
        filename_preview = os.path.splitext(file.name)[0]
        with cols[idx % 3]:
            st.image(file, caption=f"Urutan {idx+1}: {filename_preview}", use_container_width=True)

    if st.button("Proses dan Buat PDF"):
        with st.spinner("Sedang menyusun gambar..."):
            pdf_data = create_a4_grid_pdf(uploaded_files)
            if pdf_data:
                st.success("PDF siap diunduh!")
                st.download_button(
                    label="📥 Unduh PDF Sekarang",
                    data=pdf_data,
                    file_name="bukti_transfer_urut_upload.pdf",
                    mime="application/pdf"
                )
else:
    st.info("💡 Tips: Pilih file secara berurutan saat mengunggah agar hasilnya pas!")

st.markdown("---")
st.write("❤ Dibuat oleh Tari ❤")