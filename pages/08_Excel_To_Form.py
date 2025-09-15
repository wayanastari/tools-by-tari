import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def create_excel_with_forms(data_file, form_file):
    """
    Menggabungkan data dari data_file ke dalam form dari form_file
    dan menghasilkan satu file Excel.
    """
    try:
        # Membaca data dari file CSV
        data_df = pd.read_csv(data_file)
        
        # Membaca form template dari file Excel
        # Gunakan openpyxl untuk membaca template
        form_wb = openpyxl.load_workbook(form_file)
        form_ws = form_wb.active

        # Membuat workbook baru untuk hasil
        result_wb = openpyxl.Workbook()
        result_ws = result_wb.active

        # Mendapatkan mapping dari header data ke sel di form
        # Asumsi header data berada di baris pertama
        # Tentukan sel mana yang akan diisi
        
        # --- Lakukan pemetaan data ke sel form ---
        # Anda perlu menyesuaikan pemetaan ini sesuai dengan struktur form Anda
        # Berdasarkan contoh form yang Anda berikan, pemetaannya adalah:
        # Nama Bayi/Balita -> sel C2
        # NIK -> sel C3
        # Tanggal Lahir -> sel C4
        # Jenis Kelamin -> sel C2 (ditambahkan)
        # BB -> sel C5
        # TB -> sel C6
        # Nama Ayah/Ibu -> sel C7
        
        # Jarak antara setiap form di baris (misal 10 baris)
        form_height = 10 

        # Loop setiap baris data dan isi form
        for index, row in data_df.iterrows():
            # Mengisi form baru di lembar hasil
            start_row = index * form_height + 1
            
            # Salin isi form template
            for r in range(1, form_ws.max_row + 1):
                for c in range(1, form_ws.max_column + 1):
                    cell = form_ws.cell(row=r, column=c)
                    new_cell = result_ws.cell(row=start_row + r - 1, column=c)
                    if cell.value is not None:
                        new_cell.value = cell.value

            # Mengisi data ke sel yang sesuai
            result_ws.cell(row=start_row + 1, column=3).value = str(row['Nama Bayi/Balita'])
            result_ws.cell(row=start_row + 2, column=3).value = str(row['NIK'])
            result_ws.cell(row=start_row + 3, column=3).value = str(row['TANGGAL LAHIR'])
            result_ws.cell(row=start_row + 4, column=3).value = str(row['BB']) + ' Kg' if pd.notna(row['BB']) else ''
            result_ws.cell(row=start_row + 5, column=3).value = str(row['TB']) + ' Cm' if pd.notna(row['TB']) else ''
            result_ws.cell(row=start_row + 6, column=3).value = f"{row['AYAH']} / {row['IBU']}"
            
            # Mengisi Jenis Kelamin di sel C2
            gender = ''
            if pd.notna(row['L']) and row['L'] == 1:
                gender = 'Laki-laki'
            elif pd.notna(row['P']) and row['P'] == 1:
                gender = 'Perempuan'
            
            # Gabungkan Nama Bayi/Balita dan Jenis Kelamin
            result_ws.cell(row=start_row + 1, column=3).value = f"{row['Nama Bayi/Balita']} ({gender})"

        # Simpan workbook ke dalam buffer memori
        buffer = BytesIO()
        result_wb.save(buffer)
        buffer.seek(0)
        
        return buffer

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        return None

# --- Antarmuka Streamlit ---

st.title('Generator Form Excel Otomatis')
st.markdown("Unggah file data dan template form Anda. Program akan mengisi form secara otomatis.")

# Unggah file data
uploaded_data_file = st.file_uploader("Unggah File Data (CSV/Excel)", type=['csv', 'xlsx'])

# Unggah file form
uploaded_form_file = st.file_uploader("Unggah File Template Form (Excel)", type=['xlsx'])

if uploaded_data_file and uploaded_form_file:
    if st.button('Generate Excel File'):
        with st.spinner('Sedang memproses...'):
            try:
                # Menggunakan Pandas untuk membaca file data
                if uploaded_data_file.name.endswith('.csv'):
                    data_df = pd.read_csv(uploaded_data_file)
                else:
                    data_df = pd.read_excel(uploaded_data_file)
                
                # Mengisi form
                excel_buffer = create_excel_with_forms(uploaded_data_file, uploaded_form_file)
                
                if excel_buffer:
                    st.success("File Excel berhasil dibuat!")
                    
                    # Tombol unduh
                    st.download_button(
                        label="Unduh File Excel",
                        data=excel_buffer,
                        file_name="Kartu_Bayi_Balita_Terisi.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
                st.info("Pastikan file data dan form memiliki format yang benar dan sesuai.")