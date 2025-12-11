import streamlit as st
import pandas as pd
import google.generativeai as genai
from PIL import Image
import io
import json

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout="wide", page_title="Auto Tally Extractor")

# --- JUDUL & DESKRIPSI ---
st.title("⚓ Vessel Operation Data Extractor (Industry 4.0)")
st.markdown("""
Aplikasi ini mengekstrak data dari **Vessel Operation Report** (Gambar 1) 
dan mengubahnya menjadi format **Rekapitulasi Excel** (Gambar 2) secara otomatis.
""")

# --- SIDEBAR: KONFIGURASI API ---
with st.sidebar:
    st.header("Konfigurasi")
    
    # Cek apakah API Key ada di Secrets Streamlit Cloud
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("API Key terdeteksi dari System!")
    else:
        # Jika tidak ada di secrets, minta input manual
        api_key = st.text_input("Masukkan Google Gemini API Key", type="password")
        st.info("Dapatkan API Key gratis di: https://aistudio.google.com/")
    
    st.markdown("---")
  
# --- FUNGSI EKSTRAKSI MENGGUNAKAN AI ---
def extract_data_with_ai(image, api_key):
    """
    Mengirim gambar ke Gemini Flash untuk diekstrak menjadi JSON
    sesuai struktur tabel target.
    """
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash') # Model yang cepat dan hemat biaya

    # Prompt Engineering: Ini adalah instruksi "otak" dari aplikasi
    prompt = """
    Kamu adalah ahli data entry pelabuhan. Tugasmu adalah mengekstrak data dari gambar Laporan Operasi Kapal (Vessel Operation Report) ini.
    
    Lakukan perhitungan matematika sederhana untuk menggabungkan kategori.
    
    Aturan Mapping (PENTING):
    1. Ambil data Header: 'Vessel', 'Service' (Service Name).
    2. Untuk bagian 'EXPORT' (berasal dari kolom LOADING di gambar):
       - BOX LADEN: Jumlahkan baris 'FULL' + 'REEFER' + 'OOG' pada kolom LOADING.
       - BOX EMPTY: Ambil baris 'EMPTY' pada kolom LOADING.
    3. Untuk bagian 'TRANSHIPMENT' (berasal dari baris T/S):
       - Asumsikan T/S dari kolom LOADING atau DISCHARGE sesuai konteks (biasanya Loading T/S).
       - Jika tidak spesifik, ambil baris 'T/S FULL' sebagai Laden dan 'T/S EMPTY' sebagai Empty.
    4. Untuk bagian 'SHIFTING':
       - Ambil dari kolom SHIFTING.
    
    Keluarkan output HANYA dalam format JSON raw tanpa markdown, dengan key berikut:
    {
        "vessel_name": "nama kapal",
        "service_name": "nama service",
        "export_laden_20": integer,
        "export_laden_40": integer,
        "export_laden_45": integer,
        "export_empty_20": integer,
        "export_empty_40": integer,
        "export_empty_45": integer,
        "ts_laden_20": integer,
        "ts_laden_40": integer,
        "ts_laden_45": integer,
        "ts_empty_20": integer,
        "ts_empty_40": integer,
        "ts_empty_45": integer,
        "shift_laden_20": integer,
        "shift_laden_40": integer,
        "shift_laden_45": integer,
        "shift_empty_20": integer,
        "shift_empty_40": integer,
        "shift_empty_45": integer,
        "total_shift_box": integer,
        "total_shift_teus": float
    }
    Jika angka tidak ada, isi dengan 0.
    """

    try:
        response = model.generate_content([prompt, image])
        # Membersihkan markdown jika ada
        text_response = response.text.replace('```json', '').replace('```', '').strip()
        return json.loads(text_response)
    except Exception as e:
        st.error(f"Error extracting data: {e}")
        return None

# --- MAIN APP ---
uploaded_files = st.file_uploader("Upload Gambar Laporan (Bisa Banyak Sekaligus)", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

if uploaded_files and api_key:
    if st.button(f"Proses {len(uploaded_files)} Gambar"):
        
        progress_bar = st.progress(0)
        all_data = []

        for index, uploaded_file in enumerate(uploaded_files):
            # 1. Buka Gambar
            image = Image.open(uploaded_file)
            
            # 2. Ekstrak Data
            with st.spinner(f"Sedang memproses {uploaded_file.name}..."):
                data = extract_data_with_ai(image, api_key)
            
            if data:
                # 3. Post-Processing (Menghitung Total yang belum ada di JSON)
                # Hitung Total Box Export
                total_box_export = (data['export_laden_20'] + data['export_laden_40'] + data['export_laden_45'] + 
                                    data['export_empty_20'] + data['export_empty_40'] + data['export_empty_45'])
                
                # Hitung TEUS Export (Rumus standar: 20=1, 40=2, 45=2.25 atau sesuai kebijakan terminal)
                # Di sini saya pakai standar umum, bisa disesuaikan.
                teus_export = (data['export_laden_20'] * 1) + (data['export_laden_40'] * 2) + (data['export_laden_45'] * 2.25) + \
                              (data['export_empty_20'] * 1) + (data['export_empty_40'] * 2) + (data['export_empty_45'] * 2.25)

                # Total Box T/S
                total_box_ts = (data['ts_laden_20'] + data['ts_laden_40'] + data['ts_laden_45'] +
                                data['ts_empty_20'] + data['ts_empty_40'] + data['ts_empty_45'])
                
                # Menambahkan data file name untuk tracking
                row = {
                    "NO": index + 1,
                    "Vessel": data.get('vessel_name', '-'),
                    "Service Name": data.get('service_name', '-'),
                    "Remark": 0, # Default
                    # EXPORT SECTION
                    "EXP_LADEN_20": data['export_laden_20'],
                    "EXP_LADEN_40": data['export_laden_40'],
                    "EXP_LADEN_45": data['export_laden_45'],
                    "EXP_EMPTY_20": data['export_empty_20'],
                    "EXP_EMPTY_40": data['export_empty_40'],
                    "EXP_EMPTY_45": data['export_empty_45'],
                    "TOTAL BOX EXPORT": total_box_export,
                    "TEUS EXPORT": teus_export,
                    # T/S SECTION
                    "TS_LADEN_20": data['ts_laden_20'],
                    "TS_LADEN_40": data['ts_laden_40'],
                    "TS_LADEN_45": data['ts_laden_45'],
                    "TS_EMPTY_20": data['ts_empty_20'],
                    "TS_EMPTY_40": data['ts_empty_40'],
                    "TS_EMPTY_45": data['ts_empty_45'],
                    "TOTAL BOX T/S": total_box_ts,
                    "TEUS T/S": 0, # Placeholder, perlu rumus
                    # SHIFTING SECTION
                    "SHIFT_LADEN_20": data['shift_laden_20'],
                    "SHIFT_LADEN_40": data['shift_laden_40'],
                    "SHIFT_LADEN_45": data['shift_laden_45'],
                    "SHIFT_EMPTY_20": data['shift_empty_20'],
                    "SHIFT_EMPTY_40": data['shift_empty_40'],
                    "SHIFT_EMPTY_45": data['shift_empty_45'],
                    "TOTAL BOX SHIFTING": data['total_shift_box'],
                    "TEUS SHIFTING": data['total_shift_teus'],
                    # GRAND TOTALS
                    "Total (Boxes)": total_box_export + total_box_ts + data['total_shift_box'],
                    "Total Teus": teus_export + data['total_shift_teus'] # + teus ts
                }
                all_data.append(row)

            # Update Progress
            progress_bar.progress((index + 1) / len(uploaded_files))

        # --- TAMPILKAN HASIL ---
        if all_data:
            st.success("Proses Selesai!")
            
            # Buat DataFrame
            df = pd.DataFrame(all_data)
            
            # Tampilan Dataframe di Layar (Mirip Excel)
            st.dataframe(df, use_container_width=True)
            
            # --- DOWNLOAD BUTTON ---
            # Konversi ke Excel di memori
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
            
            st.download_button(
                label="📥 Download Excel Result",
                data=output.getvalue(),
                file_name="Rekap_Operasi_Kapal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Data tidak ditemukan atau gagal diekstrak.")

elif not api_key:
    st.warning("Silakan masukkan API Key di sidebar sebelah kiri untuk memulai.")
