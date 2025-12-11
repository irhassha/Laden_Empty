import streamlit as st
import pandas as pd
import requests
import base64
import json
import io
import time
from PIL import Image

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout="wide", page_title="NPCT1 Auto Tally")

# --- CSS CUSTOM ---
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stAlert {
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# --- JUDUL ---
st.title("⚓ NPCT1 Tally Extractor (Smart Discovery Mode)")
st.markdown("""
**Mode Pintar:** 1. Masukkan API Key & Data Kapal.
2. Sistem akan **otomatis mencari model AI** yang tersedia di akun Google Anda.
3. Tidak perlu khawatir salah versi model lagi.
""")

# --- SIDEBAR: KUNCI & INPUT MANUAL ---
with st.sidebar:
    st.header("1. Konfigurasi API")
    
    # Cek apakah ada di Secrets Streamlit Cloud
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terdeteksi dari System")
    else:
        api_key = st.text_input("Masukkan Google Gemini API Key", type="password")
        st.caption("Dapatkan Key gratis di aistudio.google.com")
    
    # Bersihkan API Key dari spasi tidak sengaja
    if api_key:
        api_key = api_key.strip()

    st.divider()
    
    st.header("2. Input Identitas (Manual)")
    st.warning("Data ini wajib diisi agar Excel memiliki identitas.")
    input_vessel = st.text_input("Nama Kapal (Vessel Name)", value="Vessel A")
    input_service = st.text_input("Service / Voyage", value="Service A")

# --- FUNGSI BARU: AUTO-DETECT MODEL ---
def get_best_available_model(api_key):
    """
    Bertanya ke Google: Model apa yang tersedia untuk key ini?
    Menghindari error 404 karena menebak nama model.
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            # Ambil semua model yang bisa generateContent
            available_models = [
                m['name'].replace('models/', '') 
                for m in data.get('models', []) 
                if 'generateContent' in m.get('supportedGenerationMethods', [])
            ]
            
            if not available_models:
                return None, "Tidak ada model yang mendukung generateContent."

            # Prioritas Model (Cari yang Flash dulu karena cepat, lalu Pro)
            # Kita cari string partial match
            priority_order = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-1.0-pro', 'gemini-pro']
            
            for priority in priority_order:
                # Cari model yang mengandung nama prioritas (misal gemini-1.5-flash-001 cocok dgn gemini-1.5-flash)
                for model in available_models:
                    if priority in model:
                        return model, None
            
            # Jika tidak ada yang cocok dengan prioritas, ambil yang pertama saja
            return available_models[0], None
            
        else:
            return None, f"Gagal cek model. API Error: {response.status_code} - {response.text}"
    except Exception as e:
        return None, f"Koneksi error saat cek model: {e}"

# --- FUNGSI AI (VERSI REST API) ---
def extract_table_data(image, api_key):
    
    # 1. FIX IMAGE MODE
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # 2. Convert Image to Base64
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    # 3. DAPATKAN NAMA MODEL SECARA DINAMIS
    model_name, error_msg = get_best_available_model(api_key)
    
    if not model_name:
        st.error(f"Stop: {error_msg}")
        return None
    
    # Info ke user model apa yang dipakai (Optional debug)
    # st.info(f"Menggunakan Model: {model_name}") 

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}

    prompt_text = """
    Analisis gambar tabel operasi pelabuhan ini. Fokus hanya pada angka.
    
    TUGAS: Ekstrak data dan lakukan pemetaan kategori berikut:
    1. EXPORT (Loading):
       - BOX LADEN = Penjumlahan angka di baris 'FULL' + 'REEFER' + 'OOG' pada kolom LOADING.
       - BOX EMPTY = Ambil angka di baris 'EMPTY' pada kolom LOADING.
    2. TRANSHIPMENT / TS (Baris T/S):
       - Ambil angkanya. Jika tidak spesifik, asumsikan Laden.
    3. SHIFTING:
       - Ambil total dari kolom SHIFTING.
    
    OUTPUT JSON Integer (0 jika kosong):
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

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt_text},
                {"inline_data": {"mime_type": "image/jpeg", "data": img_str}}
            ]
        }],
        "generationConfig": {"response_mime_type": "application/json"}
    }

    try:
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        
        if response.status_code == 200:
            result = response.json()
            try:
                text_response = result['candidates'][0]['content']['parts'][0]['text']
                clean_json = text_response.replace('```json', '').replace('```', '').strip()
                return json.loads(clean_json)
            except (KeyError, IndexError, json.JSONDecodeError):
                st.error(f"Format JSON Salah dari model {model_name}. Coba lagi.")
                return None
        else:
            st.error(f"API Error pada {model_name} ({response.status_code}): {response.text}")
            return None

    except Exception as e:
        st.error(f"Koneksi Error: {e}")
        return None

# --- MAIN AREA ---
uploaded_files = st.file_uploader("3. Upload Potongan Gambar Tabel (Bisa Banyak)", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

if st.button("🚀 Proses Ekstraksi") and uploaded_files and api_key:
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    all_data = []
    
    for index, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Sedang memproses: {uploaded_file.name}...")
        
        image = Image.open(uploaded_file)
        data = extract_table_data(image, api_key)
        
        if data:
            # --- POST PROCESSING ---
            
            # 1. Total Box Export
            exp_laden_box = data.get('export_laden_20',0) + data.get('export_laden_40',0) + data.get('export_laden_45',0)
            exp_empty_box = data.get('export_empty_20',0) + data.get('export_empty_40',0) + data.get('export_empty_45',0)
            tot_exp_box = exp_laden_box + exp_empty_box
            
            # 2. TEUS Export (Rumus: 20=1, 40=2, 45=2.25)
            teus_exp = (data.get('export_laden_20',0) * 1) + (data.get('export_laden_40',0) * 2) + (data.get('export_laden_45',0) * 2.25) + \
                       (data.get('export_empty_20',0) * 1) + (data.get('export_empty_40',0) * 2) + (data.get('export_empty_45',0) * 2.25)

            # 3. Total Box T/S
            ts_laden_box = data.get('ts_laden_20',0) + data.get('ts_laden_40',0) + data.get('ts_laden_45',0)
            ts_empty_box = data.get('ts_empty_20',0) + data.get('ts_empty_40',0) + data.get('ts_empty_45',0)
            tot_ts_box = ts_laden_box + ts_empty_box
            
            # 4. TEUS T/S (Estimasi)
            teus_ts = (data.get('ts_laden_20',0) * 1) + (data.get('ts_laden_40',0) * 2) + (data.get('ts_laden_45',0) * 2.25) + \
                      (data.get('ts_empty_20',0) * 1) + (data.get('ts_empty_40',0) * 2) + (data.get('ts_empty_45',0) * 2.25)

            # 5. Shifting
            tot_shift_box = data.get('total_shift_box', 0)
            if tot_shift_box == 0:
                tot_shift_box = data.get('shift_laden_20',0) + data.get('shift_laden_40',0) + data.get('shift_laden_45',0) + \
                                data.get('shift_empty_20',0) + data.get('shift_empty_40',0) + data.get('shift_empty_45',0)

            # --- MAPPING DATA KE KOLOM EXCEL ---
            row = {
                "NO": index + 1,
                "Vessel": input_vessel,
                "Service Name": input_service,
                "Remark": 0,
                
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
                "TEUS T/S": teus_ts, 

                "SHIFT_LADEN_20": data.get('shift_laden_20',0),
                "SHIFT_LADEN_40": data.get('shift_laden_40',0),
                "SHIFT_LADEN_45": data.get('shift_laden_45',0),
                "SHIFT_EMPTY_20": data.get('shift_empty_20',0),
                "SHIFT_EMPTY_40": data.get('shift_empty_40',0),
                "SHIFT_EMPTY_45": data.get('shift_empty_45',0),
                "TOTAL BOX SHIFTING": tot_shift_box,
                "TEUS SHIFTING": data.get('total_shift_teus',0),
                
                "Total (Boxes)": tot_exp_box + tot_ts_box + tot_shift_box,
                "Total Teus": teus_exp + teus_ts + data.get('total_shift_teus',0)
            }
            all_data.append(row)
        
        progress_bar.progress((index + 1) / len(uploaded_files))

    status_text.text("Selesai!")

    # --- TAMPILKAN HASIL ---
    if all_data:
        st.success(f"✅ Berhasil mengekstrak {len(all_data)} data!")
        
        df = pd.DataFrame(all_data)
        
        st.dataframe(df.style.format("{:.2f}", subset=["TEUS EXPORT", "TEUS T/S", "TEUS SHIFTING", "Total Teus"]), 
                     use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
        
        st.download_button(
            label="📥 Download Excel File", 
            data=output.getvalue(), 
            file_name=f"Rekap_{input_vessel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
