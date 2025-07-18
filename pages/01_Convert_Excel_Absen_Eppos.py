import pandas as pd
import streamlit as st
import io
from datetime import datetime, time, date, timedelta
import re
import math
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Protection, Alignment # Import Alignment

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
    start_date_periode = None
    end_date_periode = None

    for r_idx, row_series in df_raw.iterrows():
        row_str = ' '.join(row_series.dropna().astype(str).tolist())
        
        # Regex updated to capture the year for the end date as well, if available.
        # This makes the date parsing more robust for cross-year periods.
        # It tries to match YYYY/MM/DD ~ YYYY/MM/DD first, then falls back to YYYY/MM/DD ~ MM/DD
        match = re.search(r'Periode\s*:\s*(\d{4})/(\d{2})/(\d{2})\s*~\s*(?:(\d{4})/)?(\d{2})/(\d{2})', row_str)
        if match:
            start_year = int(match.group(1))
            start_month = int(match.group(2))
            start_day = int(match.group(3))
            
            end_year_str = match.group(4)
            end_month = int(match.group(5))
            end_day = int(match.group(6))

            try:
                start_date_periode = date(start_year, start_month, start_day)
                
                if end_year_str: # If end year is explicitly provided in the regex
                    end_year = int(end_year_str)
                else: # Infer end year if not explicitly provided (e.g., YYYY/MM/DD ~ MM/DD)
                    if end_month < start_month:
                        end_year = start_year + 1
                    else:
                        end_year = start_year
                
                end_date_periode = date(end_year, end_month, end_day)

                # Set bulan_laporan and tahun_laporan to the start of the period for file naming consistency
                bulan_laporan = start_month
                tahun_laporan = start_year

                st.info(f"Periode laporan teridentifikasi: {start_date_periode.strftime('%Y/%m/%d')} ~ {end_date_periode.strftime('%Y/%m/%d')} dari baris {r_idx + 1}")
                periode_row_index = r_idx
                break
            except ValueError as ve:
                st.error(f"Error parsing date from 'Periode' string: {ve}")
                st.warning("Pastikan format tanggal di baris 'Periode' adalah YYYY/MM/DD ~ [YYYY/][MM]/DD.")
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

    all_dates_in_period = []
    current_date_iter = start_date_periode
    while current_date_iter <= end_date_periode:
        all_dates_in_period.append(current_date_iter)
        current_date_iter += timedelta(days=1)
    
    st.info(f"Rentang tanggal yang akan diproses: {start_date_periode.strftime('%Y-%m-%d')} sampai {end_date_periode.strftime('%Y-%m-%d')}")

    log_karyawan_harian = {}
    for nama_karyawan, _ in employee_names_and_indices:
        for current_date_in_period in all_dates_in_period:
            key = (nama_karyawan, current_date_in_period)
            log_karyawan_harian[key] = {
                'No': 0, 
                'Nama': nama_karyawan,
                'Tanggal': current_date_in_period,
                'Log Jam Mentah': '',
                'Jam Datang': None,
                'Jam Pulang': None,
                'Jam Istirahat Mulai': None,
                'Jam Istirahat Selesai': None,
                'Durasi Jam Kerja': None,
                'Durasi Istirahat': None,
                'Keterangan Tambahan 1': '', 
                'Keterangan Tambahan 2': '', 
                'log_all_times': []
            }

    # Populate log_karyawan_harian by iterating through all_dates_in_period
    for nama_karyawan, start_row_idx in employee_names_and_indices:
        logs_start_row_for_this_block = start_row_idx + 1
        
        for current_date in all_dates_in_period: # Iterate through the correctly generated dates
            day_num = current_date.day # Get the day from the actual date object
            
            if day_num not in col_day_mapping_global:
                # This handles cases where a day in the period might not have a corresponding column in the raw data
                st.warning(f"Peringatan: Kolom untuk hari {day_num} tidak ditemukan di baris nomor hari. Melewati hari ini untuk '{nama_karyawan}'.")
                continue

            col_idx = col_day_mapping_global[day_num]
            
            key = (nama_karyawan, current_date) # Use the correct current_date object
            
            raw_cell_times_strings = [] 
            parsed_times_for_this_day = []

            for r_idx_log in range(logs_start_row_for_this_block, df_raw.shape[0]):
                if r_idx_log in employee_blocks_indices and r_idx_log != start_row_idx:
                    break # Reached the next employee block
                
                if col_idx >= df_raw.shape[1]:
                    continue

                log_cell_value = df_raw.iloc[r_idx_log, col_idx] # Get the raw cell value first

                # Heuristic to stop reading if three consecutive empty cells are found
                # Moved this check here to apply before attempting to process the cell value
                if pd.isna(log_cell_value) or str(log_cell_value).strip() == '' or str(log_cell_value).strip() == 'nan':
                    if r_idx_log + 2 < df_raw.shape[0]:
                        next_cell_val_1 = str(df_raw.iloc[r_idx_log + 1, col_idx]).strip() if col_idx < df_raw.shape[1] else ''
                        next_cell_val_2 = str(df_raw.iloc[r_idx_log + 2, col_idx]).strip() if col_idx < df_raw.shape[1] else ''

                        if (pd.isna(next_cell_val_1) or next_cell_val_1 == '' or next_cell_val_1 == 'nan') and \
                           (pd.isna(next_cell_val_2) or next_cell_val_2 == '' or next_cell_val_2 == 'nan'):
                            break 
                    else:
                        break # End of file or end of block
                    
                    continue # Continue to next row if current cell is empty but not enough consecutive empties to stop
                
                # Now convert to string for regex splitting
                time_str_full_cell = str(log_cell_value).strip() 
                potential_times_in_cell = re.split(r'\s+|\n', time_str_full_cell)

                for time_str_raw in potential_times_in_cell:
                    time_str = time_str_raw.strip()
                    if not time_str:
                        continue

                    parsed_time = None
                    try:
                        # Attempt 1: HH:MM:SS
                        parsed_time = datetime.strptime(time_str, '%H:%M:%S').time()
                    except ValueError:
                        try:
                            # Attempt 2: HH:MM
                            parsed_time = datetime.strptime(time_str, '%H:%M').time()
                        except ValueError:
                            # Attempt 3: Excel float time
                            # Only try to parse as float if the raw cell value is a float/int
                            # AND it's within a reasonable range for an Excel time (0 to <1)
                            if isinstance(log_cell_value, (float, int)) and 0 <= log_cell_value < 1:
                                excel_time_float = float(log_cell_value)
                                total_seconds = excel_time_float * 24 * 60 * 60
                                hours = int(total_seconds // 3600)
                                remainder = total_seconds % 3600
                                minutes = int(remainder // 60)
                                seconds = round(remainder % 60)
                                
                                # Ensure hours are within 0-23 range before creating time object
                                if 0 <= hours <= 23:
                                    parsed_time = time(hours, minutes, seconds)
                                else:
                                    # This indicates it was a number but not a valid time float
                                    st.warning(f"Peringatan: Nilai numerik '{log_cell_value}' di baris {r_idx_log + 1}, kolom {col_idx + 1} tidak dapat diinterpretasikan sebagai waktu yang valid (jam di luar 0-23).")

                            else:
                                # Attempt 4: Fallback to pandas to_datetime for other general formats
                                try:
                                    # pd.to_datetime is more flexible but can sometimes misinterpret numbers
                                    dt_obj = pd.to_datetime(time_str, errors='coerce') # Use errors='coerce' to turn unparsable into NaT
                                    if pd.notna(dt_obj):
                                        parsed_time = dt_obj.time()
                                except Exception:
                                    pass # Could not parse this time string
                    
                    if parsed_time:
                        parsed_times_for_this_day.append(parsed_time)
                        raw_cell_times_strings.append(parsed_time.strftime('%H:%M:%S')) 

            # Assign collected times to the correct log_karyawan_harian entry
            if parsed_times_for_this_day:
                log_karyawan_harian[key]['log_all_times'].extend(parsed_times_for_this_day)
                # Sort the raw string times for consistent output
                log_karyawan_harian[key]['Log Jam Mentah'] = '\n'.join(sorted(raw_cell_times_strings))
            # If no times for this day, 'Log Jam Mentah' remains empty as initialized

    if all(not data['log_all_times'] and not data['Log Jam Mentah'] for data in log_karyawan_harian.values()):
        st.warning("Tidak ada log absensi yang berhasil diekstrak dari file untuk karyawan manapun.")
        return None, None, None

    st.success(f"Berhasil mengekstrak log absensi mentah dan menginisialisasi semua tanggal.")
    
    sample_raw_logs_df = pd.DataFrame([
        {'Nama': data['Nama'], 'Tanggal': data['Tanggal'], 'Jam': t} 
        for key, data in log_karyawan_harian.items() 
        if data['log_all_times']
        for t in data['log_all_times']
    ])
    sample_raw_logs_df.sort_values(by=['Nama', 'Tanggal', 'Jam'], inplace=True)
    st.subheader("DEBUG INFO: 5 Baris Pertama dari Log Mentah yang Diekstrak (untuk pemrosesan)")
    st.dataframe(sample_raw_logs_df.head())


    st.info("Mengolah log mentah ke format output yang diinginkan (dengan aturan shift)...")

    kolom_final = ['No', 'Nama', 'Tanggal', 'Log Jam Mentah', 'Jam Datang', 'Jam Pulang', 
                   'Jam Istirahat Mulai', 'Jam Istirahat Selesai', 
                   'Durasi Jam Kerja', 'Durasi Istirahat',
                   'Keterangan Tambahan 1', 'Keterangan Tambahan 2'] 
    
    # Process the aggregated daily logs
    for key, data in log_karyawan_harian.items():
        all_times = sorted(data['log_all_times'])

        if not all_times:
            continue

        # --- Definisi Jendela Waktu Shift ---
        # Untuk Jam Datang:
        SHIFT_PAGI_DATANG_START = time(6, 0, 0)
        SHIFT_PAGI_DATANG_END = time(9, 30, 0) # Diperluas sedikit untuk fleksibilitas

        SHIFT_SIANG_DATANG_START = time(13, 0, 0)
        SHIFT_SIANG_DATANG_END = time(16, 0, 0)

        SHIFT_MIDDLE_DATANG_START = time(10, 0, 0)
        SHIFT_MIDDLE_DATANG_END = time(12, 30, 0)

        # Untuk Jam Pulang:
        SHIFT_PAGI_PULANG_START = time(16, 0, 0)
        SHIFT_PAGI_PULANG_END = time(19, 0, 0) # Diperluas ke 7 malam

        SHIFT_SIANG_PULANG_START = time(21, 0, 0) # Dimulai lebih awal
        SHIFT_SIANG_PULANG_END = time(23, 59, 59) # Bisa juga extend ke 00:XX keesokan hari

        SHIFT_MIDDLE_PULANG_START = time(19, 0, 0) # Dimulai dari jam 7 malam
        SHIFT_MIDDLE_PULANG_END = time(21, 0, 0) # Hingga jam 9 malam

        # Jendela Waktu Istirahat
        ISTIRAHAT_SIANG_START = time(11, 30, 0)
        ISTIRAHAT_SIANG_END = time(14, 30, 0)

        ISTIRAHAT_MALAM_START = time(18, 0, 0) # Diperluas ke 6 sore
        ISTIRAHAT_MALAM_END = time(22, 0, 0) # Diperluas ke 10 malam

        # --- Tentukan Jam Datang ---
        data['Jam Datang'] = None
        
        # Prioritas untuk Jam Datang (Shift Pagi -> Shift Middle -> Shift Siang)
        for t_log in all_times:
            if SHIFT_PAGI_DATANG_START <= t_log <= SHIFT_PAGI_DATANG_END:
                if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                    data['Jam Datang'] = t_log
                break # Ambil yang paling awal di jendela ini, dan asumsikan ini Jam Datang

        if data['Jam Datang'] is None: # Jika belum ditemukan di shift pagi
            for t_log in all_times:
                if SHIFT_MIDDLE_DATANG_START <= t_log <= SHIFT_MIDDLE_DATANG_END:
                    if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                        data['Jam Datang'] = t_log
                    break # Ambil yang paling awal di jendela ini

        if data['Jam Datang'] is None: # Jika belum ditemukan di shift middle
            for t_log in all_times:
                if SHIFT_SIANG_DATANG_START <= t_log <= SHIFT_SIANG_DATANG_END:
                    if data['Jam Datang'] is None or t_log < data['Jam Datang']:
                        data['Jam Datang'] = t_log
                    break # Ambil yang paling awal di jendela ini
        
        # Fallback: jika tidak ada yang cocok di jendela shift, ambil log pertama keseluruhan
        if data['Jam Datang'] is None and all_times:
            data['Jam Datang'] = all_times[0]


        # --- Tentukan Jam Pulang ---
        data['Jam Pulang'] = None
        
        # Prioritas untuk Jam Pulang (Shift Siang -> Shift Middle -> Shift Pagi) - mencari log TERAKHIR
        for t_log in reversed(all_times): # Iterasi terbalik untuk mencari log terakhir
            if SHIFT_SIANG_PULANG_START <= t_log <= SHIFT_SIANG_PULANG_END:
                data['Jam Pulang'] = t_log
                break # Ambil yang paling akhir di jendela ini

        if data['Jam Pulang'] is None:
            for t_log in reversed(all_times):
                if SHIFT_MIDDLE_PULANG_START <= t_log <= SHIFT_MIDDLE_PULANG_END:
                    data['Jam Pulang'] = t_log
                    break

        if data['Jam Pulang'] is None:
            for t_log in reversed(all_times):
                if SHIFT_PAGI_PULANG_START <= t_log <= SHIFT_PAGI_PULANG_END:
                    data['Jam Pulang'] = t_log
                    break
        
        # Fallback: jika tidak ada yang cocok di jendela shift, ambil log terakhir keseluruhan
        if data['Jam Pulang'] is None and all_times:
            data['Jam Pulang'] = all_times[-1]

        # --- Tentukan Jam Istirahat ---
        potential_break_times = []
        
        # Kumpulkan semua log yang berada di antara Jam Datang dan Jam Pulang
        # dan juga dalam jendela waktu istirahat yang umum
        
        # Filter logs that are strictly between Jam Datang and Jam Pulang
        internal_logs = []
        if data['Jam Datang'] and data['Jam Pulang']:
            dt_datang_for_comparison = datetime.combine(data['Tanggal'], data['Jam Datang'])
            dt_pulang_for_comparison = datetime.combine(data['Tanggal'], data['Jam Pulang'])
            if dt_pulang_for_comparison < dt_datang_for_comparison:
                dt_pulang_for_comparison += timedelta(days=1) # Adjust for overnight shifts

            for t_log in all_times:
                dt_log_for_comparison = datetime.combine(data['Tanggal'], t_log)
                if t_log < data['Jam Datang']: # Log before check-in is not relevant for break
                    continue
                
                # If it's an overnight shift, handle logs potentially on the next day for comparison
                if dt_log_for_comparison < dt_datang_for_comparison and t_log < time(5,0,0): # Heuristic for early morning next day logs
                     dt_log_for_comparison += timedelta(days=1)
                
                # Ensure the log is strictly between Jam Datang and Jam Pulang (excluding endpoints)
                if dt_datang_for_comparison < dt_log_for_comparison < dt_pulang_for_comparison:
                    internal_logs.append(t_log)
        else: # If Jam Datang or Jam Pulang not found, all_times are potential (less accurate)
            internal_logs = all_times # Consider all times as internal logs if no clear check-in/out

        # From internal logs, find those within typical break windows
        for t_log in internal_logs:
            if (ISTIRAHAT_SIANG_START <= t_log <= ISTIRAHAT_SIANG_END) or \
               (ISTIRAHAT_MALAM_START <= t_log <= ISTIRAHAT_MALAM_END):
                potential_break_times.append(t_log)
        
        # Sort these potential break times
        sorted_potential_break_times = sorted(potential_break_times)

        if len(sorted_potential_break_times) >= 2:
            data['Jam Istirahat Mulai'] = sorted_potential_break_times[0]
            data['Jam Istirahat Selesai'] = sorted_potential_break_times[-1]
        elif len(sorted_potential_break_times) == 1:
            # Jika hanya ada satu log istirahat, asumsikan ini adalah mulai dan selesai
            data['Jam Istirahat Mulai'] = sorted_potential_break_times[0]
            data['Jam Istirahat Selesai'] = sorted_potential_break_times[0]
        else:
            data['Jam Istirahat Mulai'] = None
            data['Jam Istirahat Selesai'] = None

        # --- Hitung Durasi Jam Kerja ---
        if data['Jam Datang'] and data['Jam Pulang']:
            dt_datang = datetime.combine(data['Tanggal'], data['Jam Datang'])
            dt_pulang = datetime.combine(data['Tanggal'], data['Jam Pulang'])
            
            # Handle overnight shifts: if pulang time is earlier than datang, assume it's next day
            if dt_pulang < dt_datang:
                dt_pulang += timedelta(days=1)

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

        else:
            data['Durasi Jam Kerja'] = ''

        # --- Hitung Durasi Istirahat ---
        if data['Jam Istirahat Mulai'] and data['Jam Istirahat Selesai']:
            dt_istirahat_mulai = datetime.combine(data['Tanggal'], data['Jam Istirahat Mulai'])
            dt_istirahat_selesai = datetime.combine(data['Tanggal'], data['Jam Istirahat Selesai'])
            
            # Handle cases where break might span past midnight (unlikely for a typical break, but good for robustness)
            if dt_istirahat_selesai < dt_istirahat_mulai:
                dt_istirahat_selesai += timedelta(days=1)

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
    
    return df_hasil, bulan_laporan, tahun_laporan

uploaded_file = st.file_uploader("Pilih file Excel (.xlsx atau .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.subheader("File yang Diunggah:")
    st.write(uploaded_file.name)

    df_processed, bulan_laporan_val, tahun_laporan_val = process_attendance_log(uploaded_file)

    if df_processed is not None:
        st.subheader("Download Hasil Pengolahan")
        
        df_processed_for_excel = df_processed.copy()
        
        required_cols = ['Keterangan Tambahan 1', 'Keterangan Tambahan 2']
        for col in required_cols:
            if col not in df_processed_for_excel.columns:
                df_processed_for_excel[col] = ''

        # Format time columns for Excel output
        for col in ['Jam Datang', 'Jam Pulang', 'Jam Istirahat Mulai', 'Jam Istirahat Selesai']:
            df_processed_for_excel[col] = df_processed_for_excel[col].apply(lambda x: x.strftime('%H:%M:%S') if isinstance(x, time) else '')

        output_buffer = io.BytesIO()
        
        if bulan_laporan_val is not None and tahun_laporan_val is not None:
            output_file_name = f'Rekap_Absensi_Periode_{bulan_laporan_val:02d}_{tahun_laporan_val}.xlsx'
        else:
            output_file_name = 'Rekap_Absensi_Tanpa_Periode.xlsx'

        wb = Workbook()
        ws = wb.active
        ws.title = 'Rekap Absensi'

        header_names = list(df_processed_for_excel.columns)
        
        for r_idx, row_data in enumerate(dataframe_to_rows(df_processed_for_excel, index=False, header=True)): # Set header=True to get headers
            row_num_excel = r_idx + 1 # Start from row 1 for headers, row 2 for data
            for c_idx, cell_value in enumerate(row_data):
                cell = ws.cell(row=row_num_excel, column=c_idx + 1)
                cell.value = cell_value
                # Apply wrap text to 'Log Jam Mentah' column cells (both header and data)
                if header_names[c_idx] == 'Log Jam Mentah':
                    cell.alignment = Alignment(wrapText=True) # Correct way to set alignment
                
                # Menghapus proteksi sel, agar semua sel bisa diedit/dihapus/disembunyikan
                cell.protection = Protection(locked=False) 

        DEFAULT_COLUMN_WIDTH = 25
        for i, col_name in enumerate(header_names):
            column_letter = chr(ord('A') + i)
            ws.column_dimensions[column_letter].width = DEFAULT_COLUMN_WIDTH

        # Menonaktifkan proteksi sheet agar bisa dihide/dihapus
        ws.protection.sheet = False 

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
        * Harus mengandung baris "Periode :YYYY/MM/DD ~ MM/DD" atau "Periode :YYYY/MM/DD ~ YYYY/MM/DD" untuk menentukan tahun dan bulan.
        * Harus ada baris nomor hari (1-31) di bawah baris periode.
        * Harus ada blok karyawan yang diawali dengan "No :", diikuti nama karyawan.
        * Log absensi per hari harus ada di kolom-kolom di bawah nomor hari.
        """)
