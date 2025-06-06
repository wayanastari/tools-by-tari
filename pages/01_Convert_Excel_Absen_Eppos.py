import pandas as pd
import streamlit as st
import io
from datetime import datetime, time, date, timedelta
import re
import math
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Protection # Masih diimpor tapi tidak akan digunakan untuk mengunci sel

# --- Inisialisasi variabel di awal skrip untuk menghindari NameError ---
df_processed = None
bulan_laporan_val = None
tahun_laporan_val = None

# Mengatur konfigurasi halaman
st.set_page_config(
    layout="wide",
    page_title="Dashboard Tools Ni Wayan Astari"
)

# Menampilkan judul dashboard
st.title("Convert Log Absensi Mesin Eppos")
st.write("Unggah file Excel laporan sidik jari Kamu di sini untuk diproses.")

def process_attendance_log(uploaded_file):
    """
    Fungsi utama untuk memproses file log absensi.
    Menerima file yang diunggah Streamlit.
    """
    st.info("Memulai pengolahan data log karyawan dari laporan sidik jari...")

    try:
        df_raw = pd.read_excel(io.BytesIO(uploaded_file.read()), header=None)
        st.success(f"File '{uploaded_file.name}' berhasil dibaca secara mentah.")
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        st.warning("Pastikan file yang diunggah adalah file Excel (.xlsx atau .xls) yang valid.")
        return None, None, None

    if df_raw.empty:
        st.warning("Sheet yang dibaca kosong atau tidak dapat dibaca. Tidak ada data untuk diolah.")
        return None, None, None

    st.subheader("DEBUG INFO: 10 Baris Pertama dari Sheet Mentah")
    st.dataframe(df_raw.head(10))
    st.info(f"Ukuran DataFrame Mentah: Rows: {df_raw.shape[0]}, Columns: {df_raw.shape[1]}")

    st.info("Mengekstraksi data log dari format laporan...")

    tahun_laporan = None
    bulan_laporan = None
    periode_row_index = -1

    for r_idx, row_series in df_raw.iterrows():
        row_str = ' '.join(row_series.dropna().astype(str).tolist())
        
        match = re.search(r'Periode\s*:\s*(\d{4})/(\d{2})/\d{2}\s*~\s*\d{2}/\d{2}', row_str)
        if match:
            tahun_laporan = int(match.group(1))
            bulan_laporan = int(match.group(2))
            periode_row_index = r_idx
            st.info(f"Periode laporan teridentifikasi: Tahun={tahun_laporan}, Bulan={bulan_laporan} dari baris {r_idx + 1}")
            break

    if tahun_laporan is None or bulan_laporan is None:
        st.error("Error: Informasi periode (tahun/bulan) tidak ditemukan di file Excel.")
        st.warning("Pastikan ada baris yang mengandung 'Periode :YYYY/MM/DD ~ MM/DD'.")
        return None, None, None

    days_row_idx_global = periode_row_index + 1
    st.info(f"Baris nomor hari teridentifikasi di baris {days_row_idx_global + 1}")

    col_day_mapping_global = {}
    for c_idx in range(df_raw.shape[1]):
        try:
            cell_val_day = df_raw.iloc[days_row_idx_global, c_idx]
            if pd.isna(cell_val_day):
                continue
            day_num = int(float(cell_val_day))
            if 1 <= day_num <= 31:
                col_day_mapping_global[day_num] = c_idx
        except (ValueError, TypeError):
            continue

    if not col_day_mapping_global:
        st.error("Tidak dapat menemukan nomor hari (1-31) di baris yang diharapkan.")
        st.warning("Pastikan baris nomor hari di bawah baris periode memiliki format yang benar.")
        return None, None, None

    st.info(f"Ditemukan {len(col_day_mapping_global)} nomor hari.")

    raw_logs = []
    employee_blocks_indices = []
    for r_idx in range(df_raw.shape[0]):
        cell_val_col_0 = str(df_raw.iloc[r_idx, 0]).strip() if df_raw.shape[1] > 0 else ''
        cell_val_col_1 = str(df_raw.iloc[r_idx, 1]).strip() if df_raw.shape[1] > 1 else ''

        if 'No :' in cell_val_col_0 or 'No :' in cell_val_col_1:
            employee_blocks_indices.append(r_idx)

    if not employee_blocks_indices:
        st.error("Tidak ada blok karyawan yang ditemukan (tidak ada baris yang mengandung 'No :').")
        st.warning("Pastikan format laporan sidik jari konsisten.")
        return None, None, None

    st.info(f"Ditemukan {len(employee_blocks_indices)} blok karyawan.")

    for i, start_row_idx in enumerate(employee_blocks_indices):
        nama_karyawan = None
        name_label_col_idx = -1
        for col_idx in range(df_raw.shape[1]):
            cell_val = str(df_raw.iloc[start_row_idx, col_idx]).strip()
            if 'Nama :' in cell_val:
                name_label_col_idx = col_idx
                break
        
        if name_label_col_idx != -1:
            for scan_col_idx in range(name_label_col_idx + 1, df_raw.shape[1]):
                candidate_name = str(df_raw.iloc[start_row_idx, scan_col_idx]).strip()
                if candidate_name and candidate_name != 'nan' and not candidate_name.startswith('Dept :'):
                    nama_karyawan = candidate_name
                    break
        
        if nama_karyawan is None or nama_karyawan == 'nan' or nama_karyawan == '':
            st.warning(f"Peringatan: Nama karyawan tidak ditemukan atau kosong untuk blok dimulai dari baris {start_row_idx + 1}. Melewati blok ini.")
            continue

        logs_start_row_for_this_block = start_row_idx + 1 
        sorted_day_nums = sorted(col_day_mapping_global.keys())

        for day_num in sorted_day_nums:
            col_idx = col_day_mapping_global[day_num]
            try:
                current_date = date(tahun_laporan, bulan_laporan, day_num)
            except ValueError as ve:
                st.warning(f"Peringatan: Tanggal tidak valid ({tahun_laporan}/{bulan_laporan}/{day_num}) untuk '{nama_karyawan}'. Error: {ve}. Melewati hari ini.")
                continue

            for r_idx_log in range(logs_start_row_for_this_block, df_raw.shape[0]):
                if r_idx_log in employee_blocks_indices and r_idx_log != start_row_idx:
                    break
                
                if col_idx >= df_raw.shape[1]:
                    continue

                log_cell_value = str(df_raw.iloc[r_idx_log, col_idx]).strip()

                if pd.isna(log_cell_value) or log_cell_value == '' or log_cell_value == 'nan':
                    if r_idx_log + 2 < df_raw.shape[0]:
                        next_cell_val_1 = str(df_raw.iloc[r_idx_log + 1, col_idx]).strip() if col_idx < df_raw.shape[1] else ''
                        next_cell_val_2 = str(df_raw.iloc[r_idx_log + 2, col_idx]).strip() if col_idx < df_raw.shape[1] else ''

                        if (pd.isna(next_cell_val_1) or next_cell_val_1 == '' or next_cell_val_1 == 'nan') and \
                           (pd.isna(next_cell_val_2) or next_cell_val_2 == '' or next_cell_val_2 == 'nan'):
                            break
                    else:
                        break
                    continue

                potential_times = re.split(r'\s+|\n', log_cell_value)
                
                for time_str in potential_times:
                    time_str = time_str.strip()
                    if not time_str:
                        continue

                    parsed_time = None
                    try:
                        parsed_time = datetime.strptime(time_str, '%H:%M:%S').time()
                    except ValueError:
                        try:
                            parsed_time = datetime.strptime(time_str, '%H:%M').time()
                        except ValueError:
                            try:
                                if isinstance(df_raw.iloc[r_idx_log, col_idx], (float, int)):
                                    excel_time_float = df_raw.iloc[r_idx_log, col_idx]
                                    total_seconds = excel_time_float * 24 * 60 * 60
                                    hours, remainder = divmod(total_seconds, 3600)
                                    minutes, seconds = divmod(remainder, 60)
                                    parsed_time = time(int(hours), int(minutes), int(seconds))
                                else:
                                    dt_obj = pd.to_datetime(time_str)
                                    parsed_time = dt_obj.time()
                            except Exception:
                                pass
                    
                    if parsed_time:
                        raw_logs.append({
                            'Nama': nama_karyawan,
                            'Tanggal': current_date,
                            'Jam': parsed_time
                        })

    if not raw_logs:
        st.error("Tidak ada log absensi yang berhasil diekstrak dari file.")
        st.warning("Pastikan format data konsisten dan tidak ada masalah parsing.")
        return None, None, None

    df_log_mentah = pd.DataFrame(raw_logs)
    df_log_mentah.sort_values(by=['Nama', 'Tanggal', 'Jam'], inplace=True)
    st.success(f"Berhasil mengekstrak {len(df_log_mentah)} log absensi mentah.")
    st.subheader("DEBUG INFO: 5 Baris Pertama dari Log Mentah yang Diekstrak")
    st.dataframe(df_log_mentah.head())

    st.info("Mengolah log mentah ke format output yang diinginkan (dengan aturan shift)...")

    kolom_final = ['No', 'Nama', 'Tanggal', 'Jam Datang', 'Jam Pulang', 
                   'Jam Istirahat Mulai', 'Jam Istirahat Selesai', 
                   'Durasi Jam Kerja', 'Durasi Kerja Pembulatan', 'Durasi Istirahat',
                   'Keterangan Tambahan 1', 'Keterangan Tambahan 2'] 
    
    log_karyawan_harian = {}

    for index, row in df_log_mentah.iterrows():
        nama = row['Nama']
        tanggal = row['Tanggal']
        jam_obj = row['Jam']

        key = (nama, tanggal)

        if key not in log_karyawan_harian:
            log_karyawan_harian[key] = {
                'No': 0, 
                'Nama': nama,
                'Tanggal': tanggal,
                'Jam Datang': None,
                'Jam Pulang': None,
                'Jam Istirahat Mulai': None,
                'Jam Istirahat Selesai': None,
                'Durasi Jam Kerja': None,
                'Durasi Kerja Pembulatan': None,
                'Durasi Istirahat': None,
                'Keterangan Tambahan 1': '', 
                'Keterangan Tambahan 2': '', 
                'log_all_times': []
            }
        log_karyawan_harian[key]['log_all_times'].append(jam_obj)


    for key, data in log_karyawan_harian.items():
        all_times = sorted(data['log_all_times'])

        if not all_times:
            continue

        data['Jam Datang'] = None
        data['Jam Pulang'] = None
        data['Jam Istirahat Mulai'] = None
        data['Jam Istirahat Selesai'] = None

        for t_log in all_times:
            if time(6, 0, 0) <= t_log <= time(9, 0, 0):
                if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                    data['Jam Datang'] = t_log
            elif time(14, 30, 0) <= t_log <= time(15, 0, 0):
                if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                    data['Jam Datang'] = t_log

        if data['Jam Datang'] is None and all_times:
             data['Jam Datang'] = all_times[0]

        for t_log in reversed(all_times):
            if time(23, 0, 0) <= t_log <= time(23, 59, 59):
                if data['Jam Pulang'] is None or t_log > data['Jam Pulang']:
                    data['Jam Pulang'] = t_log
                break
            elif time(16, 0, 0) <= t_log <= time(18, 0, 0):
                if data['Jam Pulang'] is None or t_log > data['Jam Pulang']:
                    data['Jam Pulang'] = t_log

        if data['Jam Pulang'] is None and all_times:
            data['Jam Pulang'] = all_times[-1]

        potential_break_times = []
        for t_log in all_times:
            if time(11, 0, 0) <= t_log <= time(14, 0, 0):
                potential_break_times.append(t_log)
            elif time(19, 0, 0) <= t_log <= time(22, 0, 0):
                potential_break_times.append(t_log)
        
        if len(potential_break_times) >= 2:
            sorted_breaks = sorted(potential_break_times)
            data['Jam Istirahat Mulai'] = sorted_breaks[0]
            data['Jam Istirahat Selesai'] = sorted_breaks[-1]
        elif len(potential_break_times) == 1:
            data['Jam Istirahat Mulai'] = potential_break_times[0]
            data['Jam Istirahat Selesai'] = potential_break_times[0]

        if data['Jam Datang'] and data['Jam Pulang']:
            dt_datang = datetime.combine(data['Tanggal'], data['Jam Datang'])
            dt_pulang = datetime.combine(data['Tanggal'], data['Jam Pulang'])
            durasi_kerja_td = dt_pulang - dt_datang
            
            total_seconds = durasi_kerja_td.total_seconds()
            hours = int(total_seconds // 3600)
            minutes = int((total_seconds % 3600) // 60)
            
            durasi_str = []
            if hours > 0:
                durasi_str.append(f"{hours} jam")
            if minutes > 0:
                durasi_str.append(f"{minutes} menit")
            
            data['Durasi Jam Kerja'] = ' '.join(durasi_str) if durasi_str else '0 menit'

            total_minutes_for_rounding = total_seconds / 60
            full_hours = math.floor(total_minutes_for_rounding / 60)
            remaining_minutes = total_minutes_for_rounding % 60

            if remaining_minutes >= 30:
                data['Durasi Kerja Pembulatan'] = f"{int(full_hours + 1)} jam"
            else:
                data['Durasi Kerja Pembulatan'] = f"{int(full_hours)} jam"

        else:
            data['Durasi Jam Kerja'] = ''
            data['Durasi Kerja Pembulatan'] = ''

        if data['Jam Istirahat Mulai'] and data['Jam Istirahat Selesai']:
            dt_istirahat_mulai = datetime.combine(data['Tanggal'], data['Jam Istirahat Mulai'])
            dt_istirahat_selesai = datetime.combine(data['Tanggal'], data['Jam Istirahat Selesai'])
            durasi_istirahat_td = dt_istirahat_selesai - dt_istirahat_mulai
            total_minutes_istirahat = durasi_istirahat_td.total_seconds() / 60
            data['Durasi Istirahat'] = f"{int(total_minutes_istirahat)} menit"
        else:
            data['Durasi Istirahat'] = ''

    df_hasil_list = []
    for key, data in log_karyawan_harian.items():
        row_data = {col: data.get(col) for col in kolom_final}
        df_hasil_list.append(row_data)

    df_hasil = pd.DataFrame(df_hasil_list)

    df_hasil['No'] = range(1, len(df_hasil) + 1)

    st.success("Pengolahan data selesai.")
    st.subheader("Data Hasil Akhir")
    st.dataframe(df_hasil)
    st.info(f"Total baris: {len(df_hasil)}")
    
    return df_hasil, bulan_laporan, tahun_laporan

uploaded_file = st.file_uploader("Pilih file Excel (.xlsx atau .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.subheader("File yang Diunggah:")
    st.write(uploaded_file.name)

    df_processed, bulan_laporan_val, tahun_laporan_val = process_attendance_log(uploaded_file)

    if df_processed is not None:
        st.subheader("Download Hasil Pengolahan")
        
        df_processed_for_excel = df_processed.copy()
        if 'Keterangan Tambahan 1' not in df_processed_for_excel.columns:
            df_processed_for_excel['Keterangan Tambahan 1'] = ''
        if 'Keterangan Tambahan 2' not in df_processed_for_excel.columns:
            df_processed_for_excel['Keterangan Tambahan 2'] = ''

        for col in ['Jam Datang', 'Jam Pulang', 'Jam Istirahat Mulai', 'Jam Istirahat Selesai']:
            df_processed_for_excel[col] = df_processed_for_excel[col].apply(lambda x: x.strftime('%H:%M:%S') if isinstance(x, time) else '')

        output_buffer = io.BytesIO()
        
        if bulan_laporan_val is not None and tahun_laporan_val is not None:
            output_file_name = f'Rekap_Absensi_Bulan_{bulan_laporan_val:02d}_{tahun_laporan_val}.xlsx'
        else:
            output_file_name = 'Rekap_Absensi_Tanpa_Periode.xlsx'

        wb = Workbook()
        ws = wb.active
        ws.title = 'Rekap Absensi'

        # --- PERUBAHAN DI SINI: columns_to_lock menjadi kosong atau dihapus ---
        # Karena kita ingin mematikan proteksi sheet, daftar ini tidak lagi relevan
        # atau bisa dibiarkan kosong
        columns_to_lock = [] # Daftar kolom yang dikunci sekarang kosong
        
        header_names = list(df_processed_for_excel.columns)
        
        # --- Bagian Kritis: Menentukan Rentang Maksimum Worksheet ---
        max_populated_row = df_processed_for_excel.shape[0] + 1 
        max_rows_for_protection = max(max_populated_row + 50, 1000) 
        num_cols = len(header_names)

        # --- Iterasi untuk mengatur properti proteksi SETIAP sel yang mungkin digunakan ---
        # Karena columns_to_lock kosong, semua sel akan diatur locked=False
        for r_idx_excel in range(1, max_rows_for_protection + 1):
            for c_idx in range(num_cols):
                column_name = header_names[c_idx]
                cell = ws.cell(row=r_idx_excel, column=c_idx + 1)
                
                if r_idx_excel == 1:
                    cell.value = column_name
                
                # --- PERUBAHAN DI SINI: Selalu set locked=False ---
                # Karena columns_to_lock sekarang kosong, kondisi ini akan selalu True
                # jika Anda tetap ingin menggunakan struktur if, atau bisa disederhanakan
                cell.protection = Protection(locked=False) 
        
        # --- Isi data dari DataFrame ke sel yang sudah diatur proteksinya ---
        for r_idx, row_data in enumerate(dataframe_to_rows(df_processed_for_excel, index=False, header=False)):
            row_num_excel = r_idx + 2 
            for c_idx, cell_value in enumerate(row_data):
                cell = ws.cell(row=row_num_excel, column=c_idx + 1)
                cell.value = cell_value 

        # --- Mengatur Lebar Kolom ---
        DEFAULT_COLUMN_WIDTH = 25
        for i, col_name in enumerate(header_names):
            column_letter = chr(ord('A') + i)
            ws.column_dimensions[column_letter].width = DEFAULT_COLUMN_WIDTH

        # --- PERUBAHAN DI SINI: Baris ini DIHAPUS atau DIKOMEN ---
        # ws.protection.sheet = True 
        # ws.protection.password = 'YourStrongPassword'

        wb.save(output_buffer)
        output_buffer.seek(0)

        st.download_button(
            label="Unduh File Excel Hasil",
            data=output_buffer.getvalue(),
            file_name=output_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"File '{output_file_name}' siap diunduh. Semua kolom bisa diedit dan lebar kolom diatur.")

else:
    col_main, col_info = st.columns([3, 1])

    with col_main:
        st.info("Silakan unggah file Excel Kamu untuk memulai.")

    with col_info:
        st.markdown("""
        ---
        **Format File yang Diharapkan:**
        * File Excel (.xlsx atau .xls) dari mesin sidik jari pertama.
        * Harus mengandung baris "Periode :YYYY/MM/DD ~ MM/DD" untuk menentukan tahun dan bulan.
        * Harus ada baris nomor hari (1-31) di bawah baris periode.
        * Harus ada blok karyawan yang diawali dengan "No :", diikuti nama karyawan.
        * Log absensi per hari harus ada di kolom-kolom di bawah nomor hari.
        """)
