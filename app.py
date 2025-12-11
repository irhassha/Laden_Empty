import streamlit as st
import pandas as pd
import requests
import base64
import json
import io
from PIL import Image

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout="wide", page_title="NPCT1 Auto Tally")

# --- JUDUL ---
st.title("⚓ NPCT1 Tally Extractor (API Version)")
st.markdown("""
**Mode Stabil:** Menggunakan Direct API Call untuk menghindari error instalasi library.
""")

# --- SIDEBAR: KUNCI & INPUT MANUAL ---
with st.sidebar:
    st.header("1. Konfigurasi API")
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terdeteksi")
    else:
        api_key = st.text_input("Masukkan API Key", type="password")

    st.divider()
    
    st.header("2. Input Identitas (Manual)")
    st.info("Isi data ini agar Excel tidak kosong.")
    input_vessel = st.text_input("Nama Kapal", value="Vessel A")
    input_service = st.text_input("Service / Voyage", value="Service A")

# --- FUNGSI AI (VERSI REST API - NO LIB DEPENDENCY) ---
def extract_table_data(image, api_key):
    # 1. Convert Image to Base64
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    # 2. Setup Request ke Gemini 1.5 Flash
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}

    # 3. Prompt Engineering
    prompt_text = """
    Analisis gambar tabel operasi pelabuhan ini. Abaikan header teks.
    
    Logic Mapping:
    1. EXPORT (Loading):
       - LADEN = Penjumlahan baris FULL + REEFER + OOG di kolom LOADING.
       - EMPTY = Baris EMPTY di kolom LOADING.
    2. TRANSHIPMENT (T/S):
       - Ambil angka dari baris T/S. Jika tidak spesifik, asumsikan T/S FULL dan T/S EMPTY.
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

    # 4. Payload Body
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt_text},
                {
                    "inline_data": {
                        "mime_type": "image/jpeg",
                        "data": img_str
                    }
                }
            ]
        }],
        "generationConfig": {
            "response_mime_type": "application/json"
        }
    }

    try:
        # Tembak API
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        
        if response.status_code == 200:
            result = response.json()
            # Parsing Navigasi JSON Gemini
            text_response = result['candidates'][0]['content']['parts'][0]['text']
            return json.loads(text_response)
        else:
            st.error(f"API Error {response.status_code}: {response.text}")
            return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# --- MAIN AREA ---
uploaded_files = st.file_uploader("3. Upload Potongan Gambar Tabel", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

if st.button("🚀 Proses Ekstraksi") and uploaded_files and api_key:
    
    progress_bar = st.progress(0)
    all_data = []
    
    for index, uploaded_file in enumerate(uploaded_files):
        image = Image.open(uploaded_file)
        
        # Panggil fungsi REST API yang baru
        data = extract_table_data(image, api_key)
        
        if data:
            # Hitung Total (Python Logic)
            tot_exp_box = (data.get('export_laden_20',0) + data.get('export_laden_40',0) + data.get('export_laden_45',0) + 
                           data.get('export_empty_20',0) + data.get('export_empty_40',0) + data.get('export_empty_45',0))
            
            teus_exp = (data.get('export_laden_20',0) * 1) + (data.get('export_laden_40',0) * 2) + (data.get('export_laden_45',0) * 2.25) + \
                       (data.get('export_empty_20',0) * 1) + (data.get('export_empty_40',0) * 2) + (data.get('export_empty_45',0) * 2.25)

            tot_ts_box = (data.get('ts_laden_20',0) + data.get('ts_laden_40',0) + data.get('ts_laden_45',0) +
                          data.get('ts_empty_20',0) + data.get('ts_empty_40',0) + data.get('ts_empty_45',0))

            row = {
                "NO": index + 1,
                "Vessel": input_vessel,
                "Service Name": input_service,
                "Remark": 0,
                # Mapping ke Excel
                "EXP_LADEN_20": data.get('export_laden_20',0),
                "EXP_LADEN_40": data.get('export_laden_40',0),
                "EXP_LADEN_45": data.get('export_laden_45',0),
                "EXP_EMPTY_20": data.get('export_empty_20',0),
                "EXP_EMPTY_40": data.get('export_empty_40',0),
                "EXP_EMPTY_45": data.get('export_empty_45',0),
                "TOTAL BOX EXPORT": tot_exp_box,
                "TEUS EXPORT": teus_exp,
                
                "TS_LADEN_20": data.get('ts_laden_20',0),
                "TS_LADEN_40": data.get('ts_laden_40',0),
                "TS_LADEN_45": data.get('ts_laden_45',0),
                "TS_EMPTY_20": data.get('ts_empty_20',0),
                "TS_EMPTY_40": data.get('ts_empty_40',0),
                "TS_EMPTY_45": data.get('ts_empty_45',0),
                "TOTAL BOX T/S": tot_ts_box,
                "TEUS T/S": 0, 

                "SHIFT_LADEN_20": data.get('shift_laden_20',0),
                "SHIFT_LADEN_40": data.get('shift_laden_40',0),
                "SHIFT_LADEN_45": data.get('shift_laden_45',0),
                "SHIFT_EMPTY_20": data.get('shift_empty_20',0),
                "SHIFT_EMPTY_40": data.get('shift_empty_40',0),
                "SHIFT_EMPTY_45": data.get('shift_empty_45',0),
                "TOTAL BOX SHIFTING": data.get('total_shift_box',0),
                "TEUS SHIFTING": data.get('total_shift_teus',0),
                
                "Total (Boxes)": tot_exp_box + tot_ts_box + data.get('total_shift_box',0),
                "Total Teus": teus_exp + data.get('total_shift_teus',0)
            }
            all_data.append(row)
        
        progress_bar.progress((index + 1) / len(uploaded_files))

    if all_data:
        st.success("✅ Ekstraksi Selesai")
        df = pd.DataFrame(all_data)
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekap')
        
        st.download_button("📥 Download Excel", data=output.getvalue(), 
                           file_name=f"Rekap_{input_vessel}.xlsx")
