import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import io
import os

def create_a4_grid_pdf(uploaded_files_data):
    if not uploaded_files_data:
        st.error("Mohon unggah gambar.")
        return None

    # A4 dimensions in pixels (at 300 DPI) for PORTRAIT
    a4_width_px = int(8.27 * 300)
    a4_height_px = int(11.69 * 300)

    # Margins and padding for the grid layout
    margin = 50  # Margin around the entire grid on the page
    padding = 20 # Padding between image slots
    border_width = 3 # Thickness of the border around each image slot
    border_color = (0, 0, 0) # Black border
    text_color = (0, 0, 0)   # Black text

    # Attempt to load a bold font (DejaVuSans-Bold is common on Linux/WSL, adjust for other OS)
    # Fallback to default Pillow font if not found.
    try:
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
        if not os.path.exists(font_path):
            font_path = "arialbd.ttf" # Common Arial Bold on Windows
            if not os.path.exists(font_path): # Fallback for other systems
                st.warning("Font bold ('DejaVuSans-Bold.ttf' or 'arialbd.ttf') tidak ditemukan. Menggunakan font default Pillow.")
                font_size = 30
                font = ImageFont.load_default()
                text_color = (100, 100, 100) # Slightly lighter grey for default font
            else:
                font_size = 40
                font = ImageFont.truetype(font_path, font_size)
        else:
            font_size = 40
            font = ImageFont.truetype(font_path, font_size)
    except IOError:
        st.warning("Terjadi masalah saat memuat font. Menggunakan font default Pillow.")
        font_size = 30
        font = ImageFont.load_default()
        text_color = (100, 100, 100) # Slightly lighter grey for default font


    # Estimate height needed for text (filename)
    text_height_estimate = font_size + 10

    # Calculate dimensions for each image slot in the 3x2 grid (3 rows, 2 columns)
    slot_width = (a4_width_px - (2 * margin) - (1 * padding)) // 2 # 2 columns
    slot_height = (a4_height_px - (2 * margin) - (2 * padding)) // 3 # 3 rows

    # Calculate max dimensions for the image *within* its slot, accounting for border and text
    # Text is below the image, so it consumes height from the slot
    img_max_width = slot_width - (2 * border_width)
    img_max_height = slot_height - (2 * border_width) - text_height_estimate

    all_pdf_pages = [] # List to store each generated A4 Image object

    # Process images in batches of 6 for each A4 page (3 rows x 2 columns)
    for page_idx in range(0, len(uploaded_files_data), 6):
        # Create a new blank A4 canvas for the current page (portrait)
        a4_canvas = Image.new('RGB', (a4_width_px, a4_height_px), 'white')
        draw = ImageDraw.Draw(a4_canvas)

        # Get the batch of up to 6 files for the current page
        current_batch_files = uploaded_files_data[page_idx : page_idx + 6]
        
        processed_items_for_page = []
        for uploaded_file_obj in current_batch_files:
            img = Image.open(uploaded_file_obj)
            
            # Extract filename without extension for display
            base_filename = os.path.basename(uploaded_file_obj.name)
            filename_without_ext = os.path.splitext(base_filename)[0]

            # Rotate landscape images to portrait if the image itself is landscape
            if img.width > img.height:
                img = img.rotate(90, expand=True)

            # Resize image to fit within the calculated max dimensions while maintaining aspect ratio
            if img.width / img.height > img_max_width / img_max_height:
                new_width = img_max_width
                new_height = int(img.height * (img_max_width / img.width))
            else:
                new_height = img_max_height
                new_width = int(img.width * (img_max_height / img.height))

            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            processed_items_for_page.append((img, filename_without_ext))

        # Place images and text on the current A4 canvas
        # Grid is 3 rows (i), 2 columns (j)
        for i in range(3): # Row (0, 1, or 2)
            for j in range(2): # Column (0 or 1)
                img_index_in_batch = i * 2 + j # Now 2 columns per row
                if img_index_in_batch < len(processed_items_for_page):
                    current_img, filename = processed_items_for_page[img_index_in_batch]

                    # Calculate top-left corner of the current slot
                    slot_x_start = margin + j * (slot_width + padding)
                    slot_y_start = margin + i * (slot_height + padding)

                    # Draw border around the slot
                    border_coords = (
                        slot_x_start,
                        slot_y_start,
                        slot_x_start + slot_width,
                        slot_y_start + slot_height
                    )
                    draw.rectangle(border_coords, outline=border_color, width=border_width)

                    # Calculate image position (centered horizontally within the slot, at the top)
                    img_x = slot_x_start + (slot_width - current_img.width) // 2
                    img_y = slot_y_start + border_width # Small offset from border

                    a4_canvas.paste(current_img, (img_x, img_y))
                    
                    # Calculate text position (centered horizontally within the slot, below the image)
                    bbox = draw.textbbox((0,0), filename, font=font)
                    text_width = bbox[2] - bbox[0]
                    
                    text_x = slot_x_start + (slot_width - text_width) // 2
                    text_y = img_y + current_img.height + 5 # 5 pixels below image

                    draw.text((text_x, text_y), filename, fill=text_color, font=font)
        
        all_pdf_pages.append(a4_canvas) # Add the completed A4 page to the list

    # Save all generated A4 pages as a single multi-page PDF
    pdf_buffer = io.BytesIO()
    if all_pdf_pages:
        first_page = all_pdf_pages[0]
        remaining_pages = all_pdf_pages[1:] if len(all_pdf_pages) > 1 else []
        first_page.save(pdf_buffer, format="PDF", save_all=True, append_images=remaining_pages)
        
    pdf_bytes = pdf_buffer.getvalue()
    return pdf_bytes

# --- Streamlit UI ---
st.set_page_config(layout="centered", page_title="Jejer Gambar Bukti Transfer ke A4")

st.title("Jejer Gambar Bukti Transfer ke Lembar A4 (Portrait, 3x2 Grid per Halaman)")
st.write("""
Unggah gambar bukti transfer Kamu. Setiap 6 gambar akan ditempatkan dalam satu halaman A4 dengan tata letak grid **3x2** (tiga baris, dua kolom) dalam orientasi **portrait**.
Gambar landscape (seperti bukti transfer) akan otomatis diputar agar pas, gambar akan di-auto-scale, dan **nama file (tanpa ekstensi) akan ditampilkan di bawah setiap gambar dalam huruf tebal dan lebih besar**.
Setelah diproses, Kamu bisa mengunduh hasilnya dalam format PDF multi-halaman.
""")

uploaded_files = st.file_uploader(
    "Pilih gambar (JPG, PNG)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    st.subheader("Pratinjau Gambar yang Diunggah:")
    # Display preview in columns
    cols = st.columns(3) # Still use 3 columns for cleaner preview layout
    for idx, file in enumerate(uploaded_files):
        base_filename_preview = os.path.basename(file.name)
        filename_without_ext_preview = os.path.splitext(base_filename_preview)[0]
        with cols[idx % 3]: # Cycle through 3 columns
            st.image(file, caption=f"Gambar {idx+1}: {filename_without_ext_preview}", width=150)

    # Process and download button
    if st.button("Proses dan Buat PDF"):
        with st.spinner("Memproses gambar dan membuat PDF..."):
            pdf_data = create_a4_grid_pdf(uploaded_files)
            if pdf_data:
                st.success("PDF berhasil dibuat!")
                st.download_button(
                    label="Unduh PDF A4 (Multi-Halaman)",
                    data=pdf_data,
                    file_name="bukti_transfer_a4_grid_multi_page.pdf",
                    mime="application/pdf"
                )
                st.info("Untuk mencetak, unduh PDF lalu buka filenya. Setelah itu cetak seperti biasa.")
            else:
                st.error("Gagal membuat PDF. Pastikan format gambar benar.")
else:
    st.info("Unggah gambar Kamu untuk memulai!")

st.markdown("---")
st.write("❤ Dibuat oleh Tari ❤")
