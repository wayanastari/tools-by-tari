import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Generator Form Excel Otomatis", layout="wide")

def get_filled_workbook(data_df, form_wb):
    """Mengisi form template dengan data dari DataFrame."""
    try:
        # Baca template form
        form_ws = form_wb.active
        
        # Tentukan tinggi form template (jumlah baris yang ingin disalin)
        form_height = 9 

        # Buat workbook baru untuk hasil
        result_wb = openpyxl.Workbook()
        result_ws = result_wb.active

        # Mendapatkan nama kolom yang tepat dari DataFrame
        def find_column_name(df, possible_names):
            for name in possible_names:
                if name in df.columns:
                    return name
            return None

        col_alamat = find_column_name(data_df, ['Alamat', 'ALAMAT', 'alamat'])
        col_no_hp = find_column_name(data_df, ['No. Hp', 'NO. HP', 'no. hp', 'No.hp'])


        # Loop setiap baris data dan isi form
        for index, row in data_df.iterrows():
            # Tentukan baris awal untuk form baru
            start_row = index * (form_height + 1) + 1  # +1 untuk baris kosong

            # Salin isi form template
            for r in range(1, form_ws.max_row + 1):
                for c in range(1, form_ws.max_column + 1):
                    cell = form_ws.cell(row=r, column=c)
                    new_cell = result_ws.cell(row=start_row + r - 1, column=c)
                    new_cell.value = cell.value

            # Mengisi data ke sel yang sesuai berdasarkan posisi yang diindikasikan
            nama_bayi = str(row.get('Nama Bayi/Balita', ''))
            
            # Perbaikan untuk NIK
            nik_val = row.get('NIK', '')
            nik = str(int(nik_val)) if pd.notna(nik_val) else ''

            # Perbaikan untuk Tanggal Lahir
            tgl_lahir_val = row.get('TANGGAL LAHIR', '')
            if pd.notna(tgl_lahir_val):
                try:
                    tgl_lahir = pd.to_datetime(tgl_lahir_val).strftime('%Y-%m-%d')
                except:
                    tgl_lahir = str(tgl_lahir_val)
            else:
                tgl_lahir = ''

            bb = str(row.get('BB', ''))
            tb = str(row.get('TB', ''))
            nama_ayah = str(row.get('AYAH', ''))
            nama_ibu = str(row.get('IBU', ''))
            alamat = str(row.get(col_alamat, '')) if col_alamat else ""
            no_hp = str(row.get(col_no_hp, '')) if col_no_hp else ""
            
            # Mengisi jenis kelamin dan menghapus yang tidak relevan
            gender_text = '(Perempuan)' if pd.notna(row.get('P')) and row.get('P') == 1 else '(Laki-laki)'
            result_ws.cell(row=start_row + 1, column=11).value = gender_text

            # Update nilai sel pada worksheet hasil
            result_ws.cell(row=start_row + 1, column=3, value=nama_bayi)
            result_ws.cell(row=start_row + 2, column=3, value=nik)
            result_ws.cell(row=start_row + 3, column=3, value=tgl_lahir)
            result_ws.cell(row=start_row + 4, column=3, value=f"{bb} Kg" if bb else "")
            result_ws.cell(row=start_row + 5, column=3, value=f"{tb} Cm" if tb else "")
            result_ws.cell(row=start_row + 6, column=3, value=nama_ayah)
            result_ws.cell(row=start_row + 7, column=3, value=nama_ibu)
            result_ws.cell(row=start_row + 8, column=3, value=alamat)
            result_ws.cell(row=start_row + 9, column=3, value=no_hp)
            
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