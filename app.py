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
st.title("⚓ NPCT1 Tally Extractor (Super Robust Version)")
st.markdown("""
**Mode Stabil & Privat:** 1. Masukkan API Key & Data Kapal.
2. Upload potongan gambar tabel (tanpa header nama kapal).
3. Sistem akan memproses menggunakan Direct API Call dengan **Auto-Fallback Model**.
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

# --- FUNGSI AI (VERSI REST API - NO LIB DEPENDENCY) ---
def extract_table_data(image, api_key):
    """
    Mengirim gambar ke Gemini via REST API standar.
    Menggunakan mekanisme FALLBACK untuk mencoba beberapa model jika satu gagal.
    """
    
    # 1. FIX IMAGE MODE
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # 2. Convert Image to Base64
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    # Daftar Model untuk dicoba (Urutan prioritas: Cepat -> Stabil -> Kuat)
    # Menghapus alias 'latest' yang sering berubah/error 404
    models_to_try = [
        "gemini-1.5-flash",      # Alias umum (biasanya paling aman)
        "gemini-1.5-flash-001",  # Versi pinned (pasti ada)
        "gemini-1.5-flash-002",  # Versi pinned baru
        "gemini-1.5-pro",        # Alias umum Pro
        "gemini-1.5-pro-001",    # Versi pinned Pro
        "gemini-1.5-pro-002"     # Versi pinned Pro baru
    ]

    last_error = ""

    # 3. Loop Try-Catch untuk setiap model
    for model_name in models_to_try:
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
            # Kirim Request
            # st.toast(f"Mencoba model: {model_name}...", icon="🤖") # Optional: Debugging UI
            response = requests.post(url, headers=headers, data=json.dumps(payload))
            
            if response.status_code == 200:
                result = response.json()
                try:
                    text_response = result['candidates'][0]['content']['parts'][0]['text']
                    clean_json = text_response.replace('```json', '').replace('```', '').strip()
                    return json.loads(clean_json) # BERHASIL! Keluar dari loop
                except (KeyError, IndexError, json.JSONDecodeError):
                    last_error = f"Model {model_name} merespon tapi format salah."
                    continue # Coba model berikutnya
            else:
                last_error = f"Model {model_name} Error {response.status_code}: {response.text}"
                # Jika 404 (Model not found) atau 503 (Overloaded), lanjut ke model berikutnya
                if response.status_code in [404, 503, 500]:
                    time.sleep(1) # Jeda sebentar sebelum retry
                    continue
                else:
                    # Error fatal (misal API Key salah), hentikan loop
                    break

        except Exception as e:
            last_error = f"Connection Error pada {model_name}: {e}"
            continue

    # Jika loop selesai dan tidak ada return, berarti semua gagal
    st.error(f"Gagal memproses gambar. Detail Error Terakhir: {last_error}")
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
            # --- POST PROCESSING (HITUNG MATEMATIKA DI PYTHON) ---
            
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
