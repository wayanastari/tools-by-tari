import pandas as pd
import streamlit as st
import io
from datetime import datetime, time, date, timedelta
import re
import math
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Protection

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
    start_date_periode = None # New variable
    end_date_periode = None   # New variable

    for r_idx, row_series in df_raw.iterrows():
        row_str = ' '.join(row_series.dropna().astype(str).tolist())
        
        # Modified regex to capture both start and end dates
        match = re.search(r'Periode\s*:\s*(\d{4})/(\d{2})/(\d{2})\s*~\s*(\d{2})/(\d{2})', row_str)
        if match:
            tahun_laporan = int(match.group(1))
            bulan_laporan = int(match.group(2))
            day_start = int(match.group(3))
            month_end = int(match.group(4))
            day_end = int(match.group(5))

            try:
                start_date_periode = date(tahun_laporan, bulan_laporan, day_start)
                # Assume end month is same year as start month, unless month_end implies next year (e.g. Dec -> Jan)
                # For simplicity, if month_end < bulan_laporan, assume it's next year (this might need adjustment for cross-year reports)
                if month_end < bulan_laporan:
                    end_date_periode = date(tahun_laporan + 1, month_end, day_end)
                else:
                    end_date_periode = date(tahun_laporan, month_end, day_end)

                st.info(f"Periode laporan teridentifikasi: {start_date_periode.strftime('%Y/%m/%d')} ~ {end_date_periode.strftime('%Y/%m/%d')} dari baris {r_idx + 1}")
                periode_row_index = r_idx
                break
            except ValueError as ve:
                st.error(f"Error parsing date from 'Periode' string: {ve}")
                st.warning("Pastikan format tanggal di baris 'Periode' adalah YYYY/MM/DD ~ MM/DD.")
                return None, None, None

    if start_date_periode is None or end_date_periode is None:
        st.error("Error: Informasi periode (tanggal mulai/akhir) tidak ditemukan atau tidak valid di file Excel.")
        st.warning("Pastikan ada baris yang mengandung 'Periode :YYYY/MM/DD ~ MM/DD' dengan format tanggal yang benar.")
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

    # Collect all employee names and their first row index
    employee_names_and_indices = []
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
        employee_names_and_indices.append((nama_karyawan, start_row_idx))

    # --- GENERATE DATES BASED ON PERIODE INSTEAD OF WHOLE MONTH ---
    all_dates_in_month = []
    current_date_iter = start_date_periode
    while current_date_iter <= end_date_periode:
        all_dates_in_month.append(current_date_iter)
        current_date_iter += timedelta(days=1)
    
    st.info(f"Rentang tanggal yang akan diproses: {start_date_periode.strftime('%Y-%m-%d')} sampai {end_date_periode.strftime('%Y-%m-%d')}")

    # --- Initialize log_karyawan_harian with all employees and all dates in the extracted period ---
    log_karyawan_harian = {}
    for nama_karyawan, _ in employee_names_and_indices:
        for current_date_in_period in all_dates_in_month:
            key = (nama_karyawan, current_date_in_period)
            log_karyawan_harian[key] = {
                'No': 0, 
                'Nama': nama_karyawan,
                'Tanggal': current_date_in_period,
                'Log Jam Mentah': '', # New column to store raw times as string
                'Jam Datang': None,
                'Jam Pulang': None,
                'Jam Istirahat Mulai': None,
                'Jam Istirahat Selesai': None,
                'Durasi Jam Kerja': None,
                'Durasi Kerja Pembulatan': None,
                'Durasi Istirahat': None,
                'Keterangan Tambahan 1': '', 
                'Keterangan Tambahan 2': '', 
                'log_all_times': [] # For internal processing (time objects)
            }


    # Populate raw_logs list with parsed times and raw string for the new column
    for i, (nama_karyawan, start_row_idx) in enumerate(employee_names_and_indices):
        logs_start_row_for_this_block = start_row_idx + 1 
        # Only process days that are within the specified period
        sorted_day_nums = sorted([d for d in col_day_mapping_global.keys() if date(tahun_laporan, bulan_laporan, d) >= start_date_periode and date(tahun_laporan, bulan_laporan, d) <= end_date_periode])

        for day_num in sorted_day_nums:
            col_idx = col_day_mapping_global[day_num]
            try:
                # Use the year/month from the Periode line, and the day from the column header
                current_date = date(tahun_laporan, bulan_laporan, day_num)
            except ValueError as ve:
                st.warning(f"Peringatan: Tanggal tidak valid ({tahun_laporan}/{bulan_laporan}/{day_num}) untuk '{nama_karyawan}'. Error: {ve}. Melewati hari ini.")
                continue
            
            key = (nama_karyawan, current_date)
            
            raw_cell_times_strings = [] 

            for r_idx_log in range(logs_start_row_for_this_block, df_raw.shape[0]):
                if r_idx_log in employee_blocks_indices and r_idx_log != start_row_idx:
                    break # Reached the next employee block
                
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
                
                potential_times_in_cell = re.split(r'\s+|\n', log_cell_value)
                parsed_times_from_this_cell = [] 

                for time_str_raw in potential_times_in_cell:
                    time_str = time_str_raw.strip()
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
                        # Only add to log_all_times if the date is within the specified period
                        if start_date_periode <= current_date <= end_date_periode:
                            log_karyawan_harian[key]['log_all_times'].append(parsed_time)
                            parsed_times_from_this_cell.append(parsed_time.strftime('%H:%M:%S')) 

                if parsed_times_from_this_cell:
                    raw_cell_times_strings.extend(parsed_times_from_this_cell)
            
            # --- IMPORTANT: Ensure 'Log Jam Mentah' is only updated if the date is within the period ---
            if start_date_periode <= current_date <= end_date_periode and raw_cell_times_strings:
                # The sorted() here ensures consistent order, which is generally good practice
                # If you want the EXACT order as in the original cell, remove sorted()
                log_karyawan_harian[key]['Log Jam Mentah'] = '\n'.join(sorted(raw_cell_times_strings))
            # Else, it remains empty as initialized for dates outside the actual scanned range
            

    # If no logs were extracted at all, return None
    # This check is now better placed after attempting to parse all logs
    if all(not data['log_all_times'] and not data['Log Jam Mentah'] for data in log_karyawan_harian.values()):
        st.warning("Tidak ada log absensi yang berhasil diekstrak dari file untuk karyawan manapun.")
        # We might still want to return an empty DataFrame with correct dates if no logs
        # but for now, if truly no logs, we return None as before.
        return None, None, None

    st.success(f"Berhasil mengekstrak log absensi mentah dan menginisialisasi semua tanggal.")
    
    # Debug raw_logs for some entries
    # Sample only from keys that actually have logs to avoid showing empty entries here
    sample_raw_logs_df = pd.DataFrame([
        {'Nama': data['Nama'], 'Tanggal': data['Tanggal'], 'Jam': t} 
        for key, data in log_karyawan_harian.items() 
        if data['log_all_times'] # Only include entries that actually had times
        for t in data['log_all_times']
    ])
    sample_raw_logs_df.sort_values(by=['Nama', 'Tanggal', 'Jam'], inplace=True)
    st.subheader("DEBUG INFO: 5 Baris Pertama dari Log Mentah yang Diekstrak (untuk pemrosesan)")
    st.dataframe(sample_raw_logs_df.head())


    st.info("Mengolah log mentah ke format output yang diinginkan (dengan aturan shift)...")

    # Update kolom_final to include the new 'Log Jam Mentah' column
    kolom_final = ['No', 'Nama', 'Tanggal', 'Log Jam Mentah', 'Jam Datang', 'Jam Pulang', 
                   'Jam Istirahat Mulai', 'Jam Istirahat Selesai', 
                   'Durasi Jam Kerja', 'Durasi Kerja Pembulatan', 'Durasi Istirahat',
                   'Keterangan Tambahan 1', 'Keterangan Tambahan 2'] 
    
    # Process the aggregated daily logs
    for key, data in log_karyawan_harian.items():
        all_times = sorted(data['log_all_times'])

        # Skip if no logs for this day/employee (keep row with empty times)
        if not all_times:
            continue

        # --- Tentukan Jam Datang ---
        data['Jam Datang'] = None
        for t_log in all_times:
            if time(6, 0, 0) <= t_log <= time(9, 0, 0): # Pagi
                if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                    data['Jam Datang'] = t_log
            elif time(14, 30, 0) <= t_log <= time(15, 0, 0): # Siang (untuk shift siang mungkin)
                # Hanya jika belum ada jam datang yang lebih pagi
                if data['Jam Datang'] is None or t_log < data['Jam Datang']: 
                    data['Jam Datang'] = t_log
        
        # Fallback for Jam Datang if not found in specific windows
        if data['Jam Datang'] is None and all_times:
            data['Jam Datang'] = all_times[0]


        # --- Tentukan Jam Pulang ---
        data['Jam Pulang'] = None
        # Check if it's Saturday (5) or Sunday (6)
        is_weekend = data['Tanggal'].weekday() in [5, 6] 

        if is_weekend:
            # Prioritize 19:00 - 20:00 for weekends
            for t_log in reversed(all_times): # Iterate in reverse to find the latest
                if time(19, 0, 0) <= t_log <= time(20, 0, 0):
                    data['Jam Pulang'] = t_log
                    break # Found the desired weekend pulang, stop searching
            
            # If not found in 19:00-20:00, check other general evening/late night hours
            if data['Jam Pulang'] is None:
                for t_log in reversed(all_times):
                    if time(23, 0, 0) <= t_log <= time(23, 59, 59):
                        data['Jam Pulang'] = t_log
                        break
                    elif time(16, 0, 0) <= t_log <= time(18, 0, 0):
                        if data['Jam Pulang'] is None or t_log > data['Jam Pulang']: # Ensure we get the latest within range
                            data['Jam Pulang'] = t_log
        else: # Weekdays
            for t_log in reversed(all_times):
                if time(23, 0, 0) <= t_log <= time(23, 59, 59):
                    data['Jam Pulang'] = t_log
                    break
                elif time(16, 0, 0) <= t_log <= time(18, 0, 0):
                    if data['Jam Pulang'] is None or t_log > data['Jam Pulang']:
                        data['Jam Pulang'] = t_log

        # Fallback for Jam Pulang if not found in specific windows
        if data['Jam Pulang'] is None and all_times:
            data['Jam Pulang'] = all_times[-1]


        # --- Tentukan Jam Istirahat ---
        potential_break_times = []
        if data['Jam Datang'] and data['Jam Pulang']:
            # Consider only logs between Jam Datang and Jam Pulang for breaks
            # This helps to exclude Jam Datang/Pulang if they fall into break windows
            internal_logs = [t_log for t_log in all_times 
                             if data['Jam Datang'] < t_log < data['Jam Pulang']]
        else:
            internal_logs = all_times # If no definite start/end, consider all logs

        for t_log in internal_logs:
            if time(11, 0, 0) <= t_log <= time(14, 0, 0):
                potential_break_times.append(t_log)
            elif time(19, 0, 0) <= t_log <= time(22, 0, 0):
                potential_break_times.append(t_log)
        
        if len(potential_break_times) >= 2:
            sorted_breaks = sorted(potential_break_times)
            data['Jam Istirahat Mulai'] = sorted_breaks[0]
            data['Jam Istirahat Selesai'] = sorted_breaks[-1]
        elif len(potential_break_times) == 1:
            # If only one break log, it's both start and end (break duration will be 0)
            data['Jam Istirahat Mulai'] = potential_break_times[0]
            data['Jam Istirahat Selesai'] = potential_break_times[0]
        else:
            data['Jam Istirahat Mulai'] = None
            data['Jam Istirahat Selesai'] = None

        # --- Hitung Durasi Jam Kerja ---
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

        # --- Hitung Durasi Istirahat ---
        if data['Jam Istirahat Mulai'] and data['Jam Istirahat Selesai']:
            dt_istirahat_mulai = datetime.combine(data['Tanggal'], data['Jam Istirahat Mulai'])
            dt_istirahat_selesai = datetime.combine(data['Tanggal'], data['Jam Istirahat Selesai'])
            durasi_istirahat_td = dt_istirahat_selesai - dt_istirahat_mulai
            total_minutes_istirahat = durasi_istirahat_td.total_seconds() / 60
            data['Durasi Istirahat'] = f"{int(total_minutes_istirahat)} menit"
        else:
            data['Durasi Istirahat'] = ''

    # Convert dictionary to list of dictionaries for DataFrame creation
    df_hasil_list = []
    # Sort keys to ensure consistent order (by employee name, then date)
    sorted_keys = sorted(log_karyawan_harian.keys())
    for key in sorted_keys:
        data = log_karyawan_harian[key]
        row_data = {col: data.get(col) for col in kolom_final}
        df_hasil_list.append(row_data)

    df_hasil = pd.DataFrame(df_hasil_list)

    df_hasil['No'] = range(1, len(df_hasil) + 1)

    st.success("Pengolahan data selesai.")
    st.subheader("Data Hasil Akhir")
    st.dataframe(df_hasil)
    st.info(f"Total baris: {len(df_hasil)}")
    
    return df_hasil, bulan_laporan, tahun_laporan # Still return bulan_laporan, tahun_laporan for file name consistency

uploaded_file = st.file_uploader("Pilih file Excel (.xlsx atau .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.subheader("File yang Diunggah:")
    st.write(uploaded_file.name)

    df_processed, bulan_laporan_val, tahun_laporan_val = process_attendance_log(uploaded_file)

    if df_processed is not None:
        st.subheader("Download Hasil Pengolahan")
        
        df_processed_for_excel = df_processed.copy()
        
        # Ensure all columns exist before processing
        required_cols = ['Keterangan Tambahan 1', 'Keterangan Tambahan 2']
        for col in required_cols:
            if col not in df_processed_for_excel.columns:
                df_processed_for_excel[col] = ''

        # Format time columns for Excel output
        for col in ['Jam Datang', 'Jam Pulang', 'Jam Istirahat Mulai', 'Jam Istirahat Selesai']:
            df_processed_for_excel[col] = df_processed_for_excel[col].apply(lambda x: x.strftime('%H:%M:%S') if isinstance(x, time) else '')

        output_buffer = io.BytesIO()
        
        if bulan_laporan_val is not None and tahun_laporan_val is not None:
            # Use a more generic name as the period is now specific
            output_file_name = f'Rekap_Absensi_Periode_{bulan_laporan_val:02d}_{tahun_laporan_val}.xlsx'
        else:
            output_file_name = 'Rekap_Absensi_Tanpa_Periode.xlsx'

        wb = Workbook()
        ws = wb.active
        ws.title = 'Rekap Absensi'

        columns_to_lock = [] 
        
        header_names = list(df_processed_for_excel.columns)
        
        max_populated_row = df_processed_for_excel.shape[0] + 1 
        max_rows_for_protection = max(max_populated_row + 50, 1000) 
        num_cols = len(header_names)

        for r_idx_excel in range(1, max_rows_for_protection + 1):
            for c_idx in range(num_cols):
                column_name = header_names[c_idx]
                cell = ws.cell(row=r_idx_excel, column=c_idx + 1)
                
                if r_idx_excel == 1:
                    cell.value = column_name
                
                cell.protection = Protection(locked=False) 
                
                # Tambahkan fitur wrap text untuk kolom 'Log Jam Mentah'
                if column_name == 'Log Jam Mentah':
                    cell.alignment = cell.alignment.copy(wrapText=True)
        
        for r_idx, row_data in enumerate(dataframe_to_rows(df_processed_for_excel, index=False, header=False)):
            row_num_excel = r_idx + 2 
            for c_idx, cell_value in enumerate(row_data):
                cell = ws.cell(row=row_num_excel, column=c_idx + 1)
                cell.value = cell_value 
                # Pastikan wrap text tetap aktif untuk sel data di kolom 'Log Jam Mentah'
                if header_names[c_idx] == 'Log Jam Mentah':
                    cell.alignment = cell.alignment.copy(wrapText=True)


        DEFAULT_COLUMN_WIDTH = 25
        for i, col_name in enumerate(header_names):
            column_letter = chr(ord('A') + i)
            ws.column_dimensions[column_letter].width = DEFAULT_COLUMN_WIDTH

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
