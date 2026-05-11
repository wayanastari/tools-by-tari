import streamlit as st
import yt_dlp
import os
import glob

# Konfigurasi halaman
st.set_page_config(page_title="YouTube to MP3 Converter", page_icon="🎵")

st.title("🎵 YouTube to MP3 Downloader")
st.markdown("Masukkan link video YouTube di bawah ini untuk mengambil audionya saja.")

# Input URL dari user
url = st.text_input("Link YouTube:", placeholder="https://www.youtube.com/watch?v=...")

def download_audio(link):
    # Opsi yt-dlp
    ydl_opts = {
        'format': 'bestaudio/best',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'outtmpl': 'downloads/%(title)s.%(ext)s', # Menyimpan di folder 'downloads'
        'quiet': True
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(link, download=True)
        # Mendapatkan path file yang baru saja diunduh
        file_path = ydl.prepare_filename(info).replace('.webm', '.mp3').replace('.m4a', '.mp3')
        return file_path

if st.button("Convert to MP3"):
    if url:
        try:
            with st.spinner("Sedang memproses... Mohon tunggu."):
                # Folder sementara untuk download
                if not os.path.exists("downloads"):
                    os.makedirs("downloads")
                
                final_file = download_audio(url)
                
                # Menampilkan tombol download di UI
                with open(final_file, "rb") as f:
                    st.audio(f.read(), format="audio/mp3")
                    st.download_button(
                        label="Unduh File MP3",
                        data=f,
                        file_name=os.path.basename(final_file),
                        mime="audio/mpeg"
                    )
                st.success(f"Berhasil mengonversi: {os.path.basename(final_file)}")
                
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
    else:
        st.warning("Silakan masukkan URL YouTube terlebih dahulu.")
