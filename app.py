import streamlit as st
import pandas as pd
import google.generativeai as genai
from PIL import Image
import io
import json

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout="wide", page_title="NPCT1 Auto Tally")

# --- JUDUL ---
st.title("⚓ NPCT1 Tally Extractor (Secure Mode)")
st.markdown("""
**Mode Privasi:** Upload gambar tabel saja (crop bagian header nama kapal). 
Masukkan identitas kapal secara manual di bawah ini.
""")

# --- SIDEBAR: KUNCI & INPUT MANUAL ---
with st.sidebar:
    st.header("1. Konfigurasi API")
    # Cek Secrets untuk Streamlit Cloud
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terdeteksi")
    else:
        api_key = st.text_input("Masukkan API Key", type="password")
        st.caption("Butuh API Key jika dijalankan lokal.")

    st.divider()
    
    st.header("2. Input Identitas (Manual)")
    st.info("Karena header gambar dicrop, isi data ini agar Excel tidak kosong.")
    input_vessel = st.text_input("Nama Kapal", value="Vessel A")
    input_service = st.text_input("Service / Voyage", value="Service A")

# --- FUNGSI AI ---
def extract_table_data(image, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')

    # Prompt disederhanakan karena tidak perlu cari nama kapal lagi
    prompt = """
    Analisis gambar tabel operasi pelabuhan ini. 
    Abaikan header teks, fokus hanya pada angka di dalam grid tabel.
    
    Tugas: Ekstrak data numerik dan lakukan penjumlahan kategori.
    
    Logic Mapping:
    1. EXPORT (Loading):
       - LADEN = Penjumlahan baris FULL + REEFER + OOG di kolom LOADING.
       - EMPTY = Baris EMPTY di kolom LOADING.
    2. TRANSHIPMENT (T/S):
       - Ambil angka dari baris T/S (biasanya di kolom Loading/Discharge).
       - Jika tidak spesifik, asumsikan T/S FULL dan T/S EMPTY.
    3. SHIFTING:
       - Ambil total dari kolom SHIFTING.
    
    Output JSON (Integer only, 0 jika kosong):
    {
        "export_laden_20": int, "export_laden_40": int, "export_laden_45": int,
        "export_empty_20": int, "export_empty_40": int, "export_empty_45": int,
        "ts_laden_20": int, "ts_laden_40": int, "ts_laden_45": int,
        "ts_empty_20": int, "ts_empty_40": int, "ts_empty_45": int,
        "shift_laden_20": int, "shift_laden_40": int, "shift_laden_45": int,
        "shift_empty_20": int, "shift_empty_40": int, "shift_empty_45": int,
        "total_shift_box": int, "total_shift_teus": float
    }
    """

    try:
        response = model.generate_content([prompt, image])
        text_response = response.text.replace('```json', '').replace('```', '').strip()
        return json.loads(text_response)
    except Exception as e:
        return None

# --- MAIN AREA ---
uploaded_files = st.file_uploader("3. Upload Potongan Gambar Tabel", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

if st.button("🚀 Proses Ekstraksi") and uploaded_files and api_key:
    
    progress_bar = st.progress(0)
    all_data = []
    
    for index, uploaded_file in enumerate(uploaded_files):
        # Proses Gambar
        image = Image.open(uploaded_file)
        
        # Kirim ke AI
        data = extract_table_data(image, api_key)
        
        if data:
            # Hitung Total di Python (Double Check Logic)
            tot_exp_box = (data['export_laden_20'] + data['export_laden_40'] + data['export_laden_45'] + 
                           data['export_empty_20'] + data['export_empty_40'] + data['export_empty_45'])
            
            # Rumus TEUS (Bisa disesuaikan)
            teus_exp = (data['export_laden_20'] * 1) + (data['export_laden_40'] * 2) + (data['export_laden_45'] * 2.25) + \
                       (data['export_empty_20'] * 1) + (data['export_empty_40'] * 2) + (data['export_empty_45'] * 2.25)

            tot_ts_box = (data['ts_laden_20'] + data['ts_laden_40'] + data['ts_laden_45'] +
                          data['ts_empty_20'] + data['ts_empty_40'] + data['ts_empty_45'])

            # Susun Baris Excel
            row = {
                "NO": index + 1,
                "Vessel": input_vessel,   # Diambil dari Input Manual
                "Service Name": input_service, # Diambil dari Input Manual
                "Remark": 0,
                # --- DATA DARI AI ---
                "EXP_LADEN_20": data['export_laden_20'],
                "EXP_LADEN_40": data['export_laden_40'],
                "EXP_LADEN_45": data['export_laden_45'],
                "EXP_EMPTY_20": data['export_empty_20'],
                "EXP_EMPTY_40": data['export_empty_40'],
                "EXP_EMPTY_45": data['export_empty_45'],
                "TOTAL BOX EXPORT": tot_exp_box,
                "TEUS EXPORT": teus_exp,
                
                "TS_LADEN_20": data['ts_laden_20'],
                "TS_LADEN_40": data['ts_laden_40'],
                "TS_LADEN_45": data['ts_laden_45'],
                "TS_EMPTY_20": data['ts_empty_20'],
                "TS_EMPTY_40": data['ts_empty_40'],
                "TS_EMPTY_45": data['ts_empty_45'],
                "TOTAL BOX T/S": tot_ts_box,
                "TEUS T/S": 0, 

                "SHIFT_LADEN_20": data['shift_laden_20'],
                "SHIFT_LADEN_40": data['shift_laden_40'],
                "SHIFT_LADEN_45": data['shift_laden_45'],
                "SHIFT_EMPTY_20": data['shift_empty_20'],
                "SHIFT_EMPTY_40": data['shift_empty_40'],
                "SHIFT_EMPTY_45": data['shift_empty_45'],
                "TOTAL BOX SHIFTING": data['total_shift_box'],
                "TEUS SHIFTING": data['total_shift_teus'],
                
                "Total (Boxes)": tot_exp_box + tot_ts_box + data['total_shift_box'],
                "Total Teus": teus_exp + data['total_shift_teus']
            }
            all_data.append(row)
        
        progress_bar.progress((index + 1) / len(uploaded_files))

    # --- HASIL ---
    if all_data:
        st.success("✅ Ekstraksi Selesai")
        df = pd.DataFrame(all_data)
        st.dataframe(df, use_container_width=True)
        
        # Download Button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekap')
        
        st.download_button("📥 Download Excel", data=output.getvalue(), 
                           file_name=f"Rekap_{input_vessel}.xlsx")
