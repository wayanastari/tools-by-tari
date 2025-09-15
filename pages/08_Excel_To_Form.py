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
        
        # Asumsi header berada di baris pertama
        headers = ["Nama Bayi/Balita", "NIK", "Tanggal Lahir", "Berat Badan Lahir", "Panjang Badan Lahir", "Nama Ayah/Ibu", "Alamat", "No. Hp"]
        result_ws.append(headers)

        # Loop setiap baris data dan isi form
        for index, row in data_df.iterrows():
            nama_bayi = str(row.get('Nama Bayi/Balita', ''))
            nik = str(row.get('NIK', ''))
            tgl_lahir = str(row.get('TANGGAL LAHIR', ''))
            bb = str(row.get('BB', ''))
            tb = str(row.get('TB', ''))
            ayah_ibu = f"{str(row.get('AYAH', ''))} / {str(row.get('IBU', ''))}"

            # Mengisi jenis kelamin
            gender = ''
            if pd.notna(row.get('L')) and row.get('L') == 1:
                gender = 'Laki-laki'
            elif pd.notna(row.get('P')) and row.get('P') == 1:
                gender = 'Perempuan'

            # Gabungkan Nama Bayi/Balita dengan Jenis Kelamin
            nama_bayi_lengkap = f"{nama_bayi} ({gender})"
            
            # Buat list data untuk satu baris
            data_row = [
                nama_bayi_lengkap,
                nik,
                tgl_lahir,
                f"{bb} Kg" if bb else "",
                f"{tb} Cm" if tb else "",
                ayah_ibu,
                # Asumsi tidak ada data Alamat dan No. Hp
                "",
                ""
            ]

            # Tambahkan baris data ke worksheet hasil
            result_ws.append(data_row)
        
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