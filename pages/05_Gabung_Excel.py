import streamlit as st
import pandas as pd
import io

def app():
    st.title("Penggabung File Excel (Header Hanya dari File Pertama)")
    st.write("Unggah beberapa file Excel (.xlsx atau .xls) Kamu di sini. Program akan menggabungkan semua sheet dari setiap file menjadi satu DataFrame tunggal. Header hanya akan diambil dari file Excel pertama yang diunggah.")

    uploaded_files = st.file_uploader(
        "Pilih beberapa file Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        all_data = pd.DataFrame()
        first_file_processed = False

        for file in uploaded_files:
            try:
                # Membaca semua sheet dari file Excel
                xls = pd.ExcelFile(file)
                for sheet_name in xls.sheet_names:
                    if not first_file_processed:
                        # Untuk file pertama, baca dengan header
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        first_file_processed = True
                    else:
                        # Untuk file berikutnya, baca tanpa header dan tetapkan nama kolom dari file pertama
                        df_no_header = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                        if not all_data.empty and df_no_header.shape[1] == all_data.shape[1]:
                            df_no_header.columns = all_data.columns
                            df = df_no_header
                        else:
                            st.warning(f"Jumlah kolom di '{file.name}' sheet '{sheet_name}' tidak cocok dengan file pertama. Sheet ini mungkin tidak digabungkan dengan benar atau memerlukan penyesuaian manual.")
                            continue # Lewati sheet ini jika kolom tidak cocok
                    all_data = pd.concat([all_data, df], ignore_index=True)

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file '{file.name}': {e}")
                continue # Lanjut ke file berikutnya jika ada error

        if not all_data.empty:
            st.subheader("Hasil Penggabungan")
            st.dataframe(all_data)

            # Tombol untuk mengunduh hasil
            csv_buffer = io.StringIO()
            all_data.to_csv(csv_buffer, index=False)
            st.download_button(
                label="Unduh sebagai CSV",
                data=csv_buffer.getvalue(),
                file_name="gabungan_excel.csv",
                mime="text/csv"
            )

            excel_buffer = io.BytesIO()
            all_data.to_excel(excel_buffer, index=False, engine='xlsxwriter')
            st.download_button(
                label="Unduh sebagai Excel",
                data=excel_buffer.getvalue(),
                file_name="gabungan_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Tidak ada data yang berhasil digabungkan. Pastikan file Excel yang diunggah memiliki format yang benar dan struktur kolom yang konsisten.")

if __name__ == "__main__":
    app()

st.write("❤ Dibuat oleh Tari ❤")
