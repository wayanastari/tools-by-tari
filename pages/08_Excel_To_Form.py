import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Generator Form Excel Otomatis", layout="wide")

def get_filled_workbook(data_df, form_wb):
    """Mengisi form template dengan data dari DataFrame."""
    try:
        form_ws = form_wb.active
        result_wb = openpyxl.Workbook()
        result_ws = result_wb.active
        
        # Jarak antara setiap form di baris (misal 10 baris)
        form_height = 10 

        # Loop setiap baris data dan isi form
        for index, row in data_df.iterrows():
            start_row = index * form_height + 1
            
            # Salin isi form template, termasuk format dasar
            for r in range(1, form_ws.max_row + 1):
                for c in range(1, form_ws.max_column + 1):
                    cell = form_ws.cell(row=r, column=c)
                    new_cell = result_ws.cell(row=start_row + r - 1, column=c)
                    new_cell.value = cell.value

            # Mengisi data ke sel yang sesuai
            nama_bayi = str(row.get('Nama Bayi/Balita', ''))
            nik = str(row.get('NIK', ''))
            tgl_lahir = str(row.get('TANGGAL LAHIR', ''))
            bb = str(row.get('BB', ''))
            tb = str(row.get('TB', ''))
            ayah_ibu = f"{row.get('AYAH', '')} / {row.get('IBU', '')}"
            
            # Mengisi jenis kelamin
            gender = ''
            if pd.notna(row.get('L')) and row.get('L') == 1:
                gender = 'Laki-laki'
            elif pd.notna(row.get('P')) and row.get('P') == 1:
                gender = 'Perempuan'

            # Update nilai sel pada worksheet hasil
            # Kolom-kolom ini harus disesuaikan dengan posisi yang benar di form template Anda
            result_ws.cell(row=start_row + 1, column=3, value=nama_bayi)
            result_ws.cell(row=start_row + 1, column=9, value=gender)
            result_ws.cell(row=start_row + 2, column=3, value=nik)
            result_ws.cell(row=start_row + 3, column=3, value=tgl_lahir)
            result_ws.cell(row=start_row + 4, column=3, value=f"{bb} Kg" if bb else "")
            result_ws.cell(row=start_row + 5, column=3, value=f"{tb} Cm" if tb else "")
            result_ws.cell(row=start_row + 6, column=3, value=ayah_ibu)

        return result_wb

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        return None

# --- UI Streamlit ---
st.title('Generator Form Bayi/Balita')
st.markdown("Unggah file data dan template form Excel, lalu unduh file yang sudah terisi.")

# Unggah file data
data_file_upload = st.file_uploader("1. Unggah File Data (.csv atau .xlsx)", type=['csv', 'xlsx'])

# Unggah file form
form_file_upload = st.file_uploader("2. Unggah File Template Form (.xlsx)", type=['xlsx'])

if data_file_upload and form_file_upload:
    if st.button('3. Generate dan Unduh File Excel'):
        with st.spinner('Memproses, mohon tunggu...'):
            try:
                # Baca data
                if data_file_upload.name.endswith('.csv'):
                    data_df = pd.read_csv(data_file_upload, encoding='latin1')
                else:
                    data_df = pd.read_excel(data_file_upload)
                
                # Baca template form
                form_wb = openpyxl.load_workbook(form_file_upload)
                
                # Buat workbook yang sudah terisi
                filled_wb = get_filled_workbook(data_df, form_wb)

                if filled_wb:
                    # Simpan workbook ke buffer
                    output = BytesIO()
                    filled_wb.save(output)
                    output.seek(0)
                    
                    st.success("File Excel berhasil dibuat!")
                    
                    # Tampilkan tombol unduh
                    st.download_button(
                        label="Klik untuk Unduh",
                        data=output,
                        file_name="Kartu_Bayi_Balita_Terisi.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Gagal memproses file. Pastikan format file data dan template sudah benar. Detail error: {e}")