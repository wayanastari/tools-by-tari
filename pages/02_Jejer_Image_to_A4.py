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

    # Attempt to load a bold font
    try:
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
        if not os.path.exists(font_path):
            font_path = "arialbd.ttf" # Common Arial Bold on Windows
            if not os.path.exists(font_path): # Fallback if specific paths fail
                st.warning("Font bold ('DejaVuSans-Bold.ttf' or 'arialbd.ttf') tidak ditemukan. Menggunakan font default Pillow.")
                font_size = 25 # Slightly smaller font for 3 columns
                font = ImageFont.load_default()
                text_color = (100, 100, 100) # Slightly lighter grey for default font
            else:
                font_size = 35 # Adjusted font size for 3 columns
                font = ImageFont.truetype(font_path, font_size)
        else:
            font_size = 35 # Adjusted font size for 3 columns
            font = ImageFont.truetype(font_path, font_size)
    except IOError:
        st.warning("Terjadi masalah saat memuat font. Menggunakan font default Pillow.")
        font_size = 25
        font = ImageFont.load_default()
        text_color = (100, 100, 100) # Slightly lighter grey for default font

    # Estimate height needed for text (filename)
    text_area_height_needed = font_size + 10 

    # --- Calculation for 3 Columns x 2 Rows with TALL Rectangular Slots on PORTRAIT A4 ---
    # We need 3 columns. The available width for images (minus margins and padding between columns)
    available_width_for_images = a4_width_px - (2 * margin) - (2 * padding)
    
    # Max width for an image in one of the 3 columns
    max_image_width_per_col = available_width_for_images // 3

    # We want the slots to be taller than wide. Let's aim for a common mobile screenshot aspect ratio (e.g., 9:16 or similar tall ratio)
    # We will maximize the height of the image based on its aspect ratio and the available height,
    # ensuring its width fits within max_image_width_per_col.

    # Total vertical space available for 2 rows of images + text
    available_height_for_images = a4_height_px - (2 * margin) - (1 * padding)
    max_image_height_per_row_area = available_height_for_images // 2

    # Calculate actual max dimensions for the image *within* its slot
    # We'll use the max_image_width_per_col as the constraining factor first
    img_max_width = max_image_width_per_col - (2 * border_width) 
    
    # Assume a target aspect ratio for height relative to width (e.g., 16:9 for portrait phone screens)
    # This means height = width * (16/9)
    target_height_from_width = int(img_max_width * (16/9.0)) # Make it significantly taller

    # The effective height of the image slot includes the image itself, border, and text area
    effective_slot_height_needed = target_height_from_width + (2 * border_width) + text_area_height_needed

    # If the calculated effective slot height is too big for the available row height, scale down
    if effective_slot_height_needed > max_image_height_per_row_area:
        # Scale down the image height to fit the row, then recalculate width
        img_max_height = max_image_height_per_row_area - (2 * border_width) - text_area_height_needed
        img_max_width = int(img_max_height * (9.0/16.0)) # Recalculate width based on new height and aspect ratio
        
        # Recalculate effective slot height with new img_max_height
        effective_slot_height_needed = img_max_height + (2 * border_width) + text_area_height_needed
    else:
        img_max_height = target_height_from_width

    # The actual 'drawn' slot width will be based on img_max_width plus borders
    actual_drawn_slot_width = img_max_width + (2 * border_width)
    actual_drawn_slot_height = img_max_height + (2 * border_width) + text_area_height_needed


    all_pdf_pages = [] # List to store each generated A4 Image object

    # Process images in batches of 6 for each A4 page (3 columns x 2 rows = 6 images)
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

            # Rotate landscape images to portrait to fit rectangular slots well
            if img.width > img.height:
                img = img.rotate(90, expand=True)

            # Resize image to fit within the calculated img_max_width/height while maintaining aspect ratio
            if img.width / img.height > img_max_width / img_max_height:
                new_width = img_max_width
                new_height = int(img.height * (img_max_width / img.width))
            else:
                new_height = img_max_height
                new_width = int(img.width * (img_max_height / img.height))

            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            processed_items_for_page.append((img, filename_without_ext))

        # Place images and text on the current A4 canvas
        # Grid is 2 rows (i), 3 columns (j)
        for i in range(2): # Row (0 or 1)
            for j in range(3): # Column (0, 1, or 2)
                img_index_in_batch = i * 3 + j # Correct index for 2 rows, 3 columns
                if img_index_in_batch < len(processed_items_for_page):
                    current_img, filename = processed_items_for_page[img_index_in_batch]

                    # Calculate top-left corner of the "grid cell"
                    grid_cell_x_start = margin + j * (actual_drawn_slot_width + padding)
                    grid_cell_y_start = margin + i * (actual_drawn_slot_height + padding)

                    # Ensure we stay within A4 bounds, especially for the last item/row/column
                    if grid_cell_x_start + actual_drawn_slot_width > a4_width_px - margin:
                        grid_cell_x_start = a4_width_px - margin - actual_drawn_slot_width
                    if grid_cell_y_start + actual_drawn_slot_height > a4_height_px - margin:
                        grid_cell_y_start = a4_height_px - margin - actual_drawn_slot_height


                    # Draw border around the actual rectangular content area
                    border_coords = (
                        grid_cell_x_start,
                        grid_cell_y_start,
                        grid_cell_x_start + actual_drawn_slot_width,
                        grid_cell_y_start + actual_drawn_slot_height
                    )
                    draw.rectangle(border_coords, outline=border_color, width=border_width)

                    # Calculate image position (centered horizontally within its drawn slot)
                    img_x = grid_cell_x_start + (actual_drawn_slot_width - current_img.width) // 2
                    img_y = grid_cell_y_start + border_width # Small offset from top border of drawn slot

                    a4_canvas.paste(current_img, (img_x, img_y))
                    
                    # Calculate text position (centered horizontally, directly below the image)
                    bbox = draw.textbbox((0,0), filename, font=font)
                    text_width = bbox[2] - bbox[0]

                    text_x = grid_cell_x_start + (actual_drawn_slot_width - text_width) // 2
                    text_y = img_y + current_img.height + 5 

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
st.set_page_config(layout="centered", page_title="Jejer Bukti Transfer ke A4")

st.title("Jejer Gambar Bukti Transfer ke Lembar A4 (Portrait, 3 Kolom x 2 Baris)")
st.write("""
Unggah gambar bukti transfer Kamu. Setiap 6 gambar akan ditempatkan dalam satu halaman A4 dengan tata letak grid **3 kolom x 2 baris** (total 6 gambar) dalam orientasi **portrait**.
**Setiap gambar akan berada dalam kotak persegi panjang (lebih tinggi dari lebar)**. Gambar landscape akan otomatis diputar agar pas, gambar akan di-auto-scale, dan **nama file (tanpa ekstensi) akan ditampilkan di bawah setiap gambar**.
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
    cols = st.columns(3) # Use 3 columns for preview
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
                    file_name="bukti_transfer_a4_grid_portrait_3col_2row.pdf",
                    mime="application/pdf"
                )
                st.info("Untuk mencetak, unduh PDF lalu buka filenya. Setelah itu cetak seperti biasa.")
            else:
                st.error("Gagal membuat PDF. Pastikan format gambar benar.")
else:
    st.info("Unggah gambar Kamu untuk memulai!")

st.markdown("---")
st.write("❤ Dibuat oleh Tari ❤")
