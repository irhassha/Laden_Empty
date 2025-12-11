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
st.title("⚓ NPCT1 Tally Extractor (Smart Failover Mode)")
st.markdown("""
**Mode Pintar:** 1. Masukkan API Key & Data Kapal.
2. Sistem akan otomatis mencari model AI yang tersedia.
3. **Anti-Limit:** Jika satu model kuotanya habis (429), sistem otomatis pindah ke model lain.
""")

# --- INITIALIZE SESSION STATE ---
# Kita butuh ini agar data tidak hilang saat kita klik tombol 'Combine' nanti
if 'extracted_data' not in st.session_state:
    st.session_state['extracted_data'] = []

# --- FUNGSI UTILITY: CACHE DATA MODEL ---
@st.cache_data(ttl=300) 
def get_prioritized_models(api_key):
    """
    Mengambil semua model yang tersedia dan mengurutkannya.
    PRIORITAS UTAMA: FLASH (Karena kuota gratis besar & cepat).
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            all_models = [
                m['name'].replace('models/', '') 
                for m in data.get('models', []) 
                if 'generateContent' in m.get('supportedGenerationMethods', [])
            ]
            
            if not all_models:
                return []
            
            sorted_models = []
            # Layer 1: Flash Models 
            flash_models = [m for m in all_models if 'flash' in m.lower()]
            # Layer 2: Pro Models
            pro_models = [m for m in all_models if 'pro' in m.lower() and m not in flash_models]
            # Layer 3: Others
            other_models = [m for m in all_models if m not in flash_models and m not in pro_models]
            
            sorted_models = flash_models + pro_models + other_models
            return sorted_models
        else:
            return []
    except Exception:
        return []

# --- SIDEBAR: KUNCI & INPUT MANUAL ---
with st.sidebar:
    st.header("1. Konfigurasi API")
    
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terdeteksi dari System")
    else:
        api_key = st.text_input("Masukkan Google Gemini API Key", type="password")
        st.caption("Dapatkan Key gratis di aistudio.google.com")
    
    if api_key:
        api_key = api_key.strip()
        
        st.divider()
        st.header("🔍 Status Koneksi")
        with st.spinner("Mengecek koneksi..."):
            active_models = get_prioritized_models(api_key)
            
        if active_models:
            st.success("✅ Terhubung & Siap")
            with st.expander(f"Lihat {len(active_models)} Model Aktif"):
                st.write(active_models)
                
            st.info("""
            **Info Batas Free Tier:**
            - 15 Request / Menit
            - 1.500 Request / Hari
            
            *Jika error 429, sistem otomatis pindah model.*
            """)
            st.markdown("**Cek Sisa Kuota Realtime:**")
            st.link_button("📊 Buka Google AI Dashboard", "https://aistudio.google.com/app/plan_information")
        else:
            st.error("❌ Koneksi Gagal / API Key Salah")

    st.divider()
    
    st.header("2. Input Identitas (Manual)")
    st.warning("Data ini wajib diisi agar Excel memiliki identitas.")
    input_vessel = st.text_input("Nama Kapal (Vessel Name)", value="Vessel A")
    input_service = st.text_input("Service / Voyage", value="Service A")


# --- FUNGSI AI (VERSI REST API + FAILOVER) ---
def extract_table_data(image, api_key):
    
    if image.mode != 'RGB':
        image = image.convert('RGB')

    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    candidate_models = get_prioritized_models(api_key)
    
    if not candidate_models:
        candidate_models = ["gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-1.5-pro"]
    
    last_error_msg = ""
    
    for model_name in candidate_models:
        if "experimental" in model_name:
            continue

        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}

        prompt_text = """
        Analisis gambar tabel operasi pelabuhan ini. Fokus hanya pada angka.
        
        TUGAS: Ekstrak data dan lakukan pemetaan kategori berikut:
        
        1. IMPORT (Ambil dari kolom DISCHARGE di gambar):
           - BOX LADEN = Penjumlahan angka di baris 'FULL' + 'REEFER' + 'OOG' pada kolom DISCHARGE.
           - BOX EMPTY = Ambil angka di baris 'EMPTY' pada kolom DISCHARGE.

        2. EXPORT (Ambil dari kolom LOADING di gambar):
           - BOX LADEN = Penjumlahan angka di baris 'FULL' + 'REEFER' + 'OOG' pada kolom LOADING.
           - BOX EMPTY = Ambil angka di baris 'EMPTY' pada kolom LOADING.
           
        3. TRANSHIPMENT / TS (Baris T/S):
           - Ambil angkanya. Jika tidak spesifik, asumsikan Laden.
           
        4. SHIFTING:
           - Ambil total dari kolom SHIFTING.
        
        OUTPUT JSON Integer (0 jika kosong):
        {
            "import_laden_20": int, "import_laden_40": int, "import_laden_45": int,
            "import_empty_20": int, "import_empty_40": int, "import_empty_45": int,
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
                except Exception:
                    last_error_msg = f"Model {model_name} output format salah."
                    continue 
            
            elif response.status_code == 429:
                last_error_msg = f"Kuota {model_name} habis (429). Mencoba model lain..."
                continue
            
            elif response.status_code in [404, 500, 503]:
                last_error_msg = f"Model {model_name} tidak ditemukan/error server ({response.status_code})."
                continue
            
            else:
                last_error_msg = f"API Error {response.status_code}: {response.text}"
                break

        except Exception as e:
            last_error_msg = f"Koneksi Error pada {model_name}: {e}"
            continue

    st.error(f"Gagal memproses. Detail: {last_error_msg}")
    return None

# --- MAIN AREA ---
uploaded_files = st.file_uploader("3. Upload Potongan Gambar Tabel (Bisa Banyak)", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

# TOMBOL PROSES EKSTRAKSI
if st.button("🚀 Proses Ekstraksi") and uploaded_files and api_key:
    
    # Reset data lama jika user upload ulang
    st.session_state['extracted_data'] = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for index, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Sedang memproses: {uploaded_file.name}...")
        
        image = Image.open(uploaded_file)
        data = extract_table_data(image, api_key)
        
        if data:
            # --- POST PROCESSING ---
            # 1. IMPORT
            imp_laden_box = data.get('import_laden_20',0) + data.get('import_laden_40',0) + data.get('import_laden_45',0)
            imp_empty_box = data.get('import_empty_20',0) + data.get('import_empty_40',0) + data.get('import_empty_45',0)
            tot_imp_box = imp_laden_box + imp_empty_box

            teus_imp = (data.get('import_laden_20',0) * 1) + (data.get('import_laden_40',0) * 2) + (data.get('import_laden_45',0) * 2.25) + \
                       (data.get('import_empty_20',0) * 1) + (data.get('import_empty_40',0) * 2) + (data.get('import_empty_45',0) * 2.25)

            # 2. EXPORT
            exp_laden_box = data.get('export_laden_20',0) + data.get('export_laden_40',0) + data.get('export_laden_45',0)
            exp_empty_box = data.get('export_empty_20',0) + data.get('export_empty_40',0) + data.get('export_empty_45',0)
            tot_exp_box = exp_laden_box + exp_empty_box
            
            teus_exp = (data.get('export_laden_20',0) * 1) + (data.get('export_laden_40',0) * 2) + (data.get('export_laden_45',0) * 2.25) + \
                       (data.get('export_empty_20',0) * 1) + (data.get('export_empty_40',0) * 2) + (data.get('export_empty_45',0) * 2.25)

            # 3. TRANSHIPMENT
            ts_laden_box = data.get('ts_laden_20',0) + data.get('ts_laden_40',0) + data.get('ts_laden_45',0)
            ts_empty_box = data.get('ts_empty_20',0) + data.get('ts_empty_40',0) + data.get('ts_empty_45',0)
            tot_ts_box = ts_laden_box + ts_empty_box
            
            teus_ts = (data.get('ts_laden_20',0) * 1) + (data.get('ts_laden_40',0) * 2) + (data.get('ts_laden_45',0) * 2.25) + \
                      (data.get('ts_empty_20',0) * 1) + (data.get('ts_empty_40',0) * 2) + (data.get('ts_empty_45',0) * 2.25)

            # 4. SHIFTING
            tot_shift_box = data.get('total_shift_box', 0)
            if tot_shift_box == 0:
                tot_shift_box = data.get('shift_laden_20',0) + data.get('shift_laden_40',0) + data.get('shift_laden_45',0) + \
                                data.get('shift_empty_20',0) + data.get('shift_empty_40',0) + data.get('shift_empty_45',0)

            # Mapping Data
            row = {
                "NO": index + 1,
                "Vessel": f"{input_vessel} ({index+1})", # Tambah index biar beda kalo upload banyak file
                "Service Name": input_service,
                "Remark": 0,
                
                "IMP_LADEN_20": data.get('import_laden_20',0),
                "IMP_LADEN_40": data.get('import_laden_40',0),
                "IMP_LADEN_45": data.get('import_laden_45',0),
                "IMP_EMPTY_20": data.get('import_empty_20',0),
                "IMP_EMPTY_40": data.get('import_empty_40',0),
                "IMP_EMPTY_45": data.get('import_empty_45',0),
                "TOTAL BOX IMPORT": tot_imp_box,
                "TEUS IMPORT": teus_imp,

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
                
                "Total (Boxes)": tot_imp_box + tot_exp_box + tot_ts_box + tot_shift_box,
                "Total Teus": teus_imp + teus_exp + teus_ts + data.get('total_shift_teus',0)
            }
            # SIMPAN KE SESSION STATE
            st.session_state['extracted_data'].append(row)
        
        progress_bar.progress((index + 1) / len(uploaded_files))
    status_text.text("Selesai!")

# --- DISPLAY LOGIC (SELALU TAMPIL JIKA ADA DATA DI STATE) ---
if st.session_state['extracted_data']:
    df = pd.DataFrame(st.session_state['extracted_data'])
    
    st.divider()
    st.header("3. Hasil Ekstraksi Data")
    
    # Kolom desimal
    teus_cols = ["TEUS IMPORT", "TEUS EXPORT", "TEUS T/S", "TEUS SHIFTING", "Total Teus"]
    st.dataframe(df.style.format("{:.2f}", subset=teus_cols), use_container_width=True)
    
    # Download Full Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
    
    st.download_button(
        label="📥 Download Excel (Semua Data)", 
        data=output.getvalue(), 
        file_name=f"Rekap_{input_vessel}_Full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- FITUR BARU: COMBINE DATA ---
    st.divider()
    st.header("4. Fitur Combine Data (Penjumlahan Kapal)")
    st.info("Pilih beberapa kapal dari daftar di bawah untuk menjumlahkan datanya.")

    # Multiselect berdasarkan NO Urut
    options = df['NO'].tolist()
    # Format label dropdown biar enak dibaca: "1 - Vessel A (1)"
    choice_labels = {row['NO']: f"{row['NO']} - {row['Vessel']}" for index, row in df.iterrows()}
    
    selected_indices = st.multiselect(
        "Pilih Data (Nomor Urut):", 
        options,
        format_func=lambda x: choice_labels.get(x)
    )

    if selected_indices:
        # Filter Data
        subset_df = df[df['NO'].isin(selected_indices)]
        
        st.write("##### Data yang dipilih:")
        st.dataframe(subset_df.style.format("{:.2f}", subset=teus_cols), use_container_width=True)
        
        # Lakukan Summing untuk kolom numerik saja
        numeric_cols = subset_df.select_dtypes(include='number').columns
        # Hapus kolom NO dan Remark dari penjumlahan biar gak aneh dan error
        cols_to_sum = [c for c in numeric_cols if c not in ['NO', 'Remark']]
        
        sum_row = subset_df[cols_to_sum].sum()
        
        # Buat Dataframe Hasil Combine
        combined_df = pd.DataFrame([sum_row])
        
        # Isi kembali kolom teks identitas
        combined_df.insert(0, "NO", "GABUNGAN")
        combined_df.insert(1, "Vessel", "MULTIPLE VESSELS")
        combined_df.insert(2, "Service Name", "COMBINED")
        combined_df.insert(3, "Remark", "-")

        st.success(f"✅ Total Gabungan dari {len(selected_indices)} Kapal:")
        st.dataframe(combined_df.style.format("{:.2f}", subset=teus_cols), use_container_width=True)
        
        # Download Combine Button
        output_combine = io.BytesIO()
        with pd.ExcelWriter(output_combine, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Combined_Data')
        
        st.download_button(
            label="📥 Download Excel (Data Gabungan Saja)", 
            data=output_combine.getvalue(), 
            file_name=f"Rekap_Gabungan_{len(selected_indices)}_Kapal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
