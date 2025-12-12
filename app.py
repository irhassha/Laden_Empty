import streamlit as st
import pandas as pd
import requests
import base64
import json
import io
import time
from PIL import Image

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    layout="wide", 
    page_title="RBM Auto Tally",
    page_icon="⚓"
)

# --- CSS CUSTOM UNTUK TAMPILAN BERSIH ---
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stAlert {
        margin-top: 1rem;
    }
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# --- INITIALIZE SESSION STATE ---
if 'extracted_data' not in st.session_state:
    st.session_state['extracted_data'] = []
if 'images' not in st.session_state:
    st.session_state['images'] = {}

# --- FUNGSI UTILITY ---
@st.cache_data(ttl=300) 
def get_prioritized_models(api_key):
    # Fixed URL format
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
            if not all_models: return []
            
            flash_models = [m for m in all_models if 'flash' in m.lower()]
            pro_models = [m for m in all_models if 'pro' in m.lower() and m not in flash_models]
            other_models = [m for m in all_models if m not in flash_models and m not in pro_models]
            return flash_models + pro_models + other_models
        else:
            return []
    except Exception:
        return []

def render_image_viewer(df, key_suffix):
    """Helper function untuk menampilkan viewer gambar yang konsisten di berbagai tab"""
    st.markdown("---")
    c_view1, c_view2 = st.columns([1, 2])
    
    with c_view1:
        st.info("🔍 **Cek Gambar Asli**")
        vessel_list = df['Vessel'].unique()
        # Gunakan key unik agar widget tidak bentrok antar tab
        selected_vessel = st.selectbox("Pilih Kapal:", vessel_list, key=f"v_sel_{key_suffix}")
        
        # Ambil ID kapal yang dipilih
        if not df.empty:
            selected_row = df[df['Vessel'] == selected_vessel].iloc[0]
            selected_id = selected_row['NO']
        else:
            selected_id = None
        
    with c_view2:
        if selected_id and selected_id in st.session_state['images']:
            st.image(st.session_state['images'][selected_id], caption=f"Dokumen Asli: {selected_vessel}", use_container_width=True)
        else:
            st.warning("Gambar asli tidak ditemukan.")
    st.markdown("---")

# --- SIDEBAR ---
with st.sidebar:
    st.title("⚙️ Pengaturan")
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terhubung")
    else:
        api_key = st.text_input("Google Gemini API Key", type="password", placeholder="Paste key disini...")
    
    if api_key:
        api_key = api_key.strip()
        st.divider()
        with st.spinner("Cek koneksi..."):
            active_models = get_prioritized_models(api_key)
        if active_models:
            st.markdown(f"**Status:** 🟢 Online ({len(active_models)} Models)")
            st.progress(100)
        else:
            st.error("API Key Invalid")
            
    st.divider()
    if st.button("🗑️ Hapus Semua Data", type="secondary", use_container_width=True):
        st.session_state['extracted_data'] = []
        st.session_state['images'] = {} 
        st.rerun()
    st.info("Versi Aplikasi: 2.4 (DG & Hatch Fix)\nFitur: Enhanced Extraction Logic")

# --- HEADER ---
st.title("⚓ RBM Auto Tally")
st.markdown("Automasi ekstraksi data operasional pelabuhan dari gambar laporan ke Excel.")
st.divider()

# --- INPUT AREA ---
st.subheader("Upload Laporan")
uploaded_files = st.file_uploader("Upload Potongan Gambar Tabel", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

# --- FUNGSI EKSTRAKSI ---
def extract_table_data(image, api_key):
    if image.mode != 'RGB': image = image.convert('RGB')
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    candidate_models = get_prioritized_models(api_key)
    if not candidate_models: candidate_models = ["gemini-1.5-flash", "gemini-1.5-flash-latest"]
    
    last_error_msg = "Unknown Error"

    for model_name in candidate_models:
        if "experimental" in model_name: continue

        # Fixed URL format
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        
        prompt_text = """
        Analisis gambar tabel operasi pelabuhan ini. Saya butuh data detail (granular) per sel matriks.
        
        STRUKTUR DATA DI GAMBAR:
        Baris (Rows): FULL, REEFER, OOG, DG (atau IMO/Dangerous), T/S FULL, T/S DG, T/S EMPTY, T/S OOG, EMPTY.
        Kolom (Cols): 
        - DISCHARGE (IMPORT): 20, 40, 45
        - LOADING (EXPORT): 20, 40, 45
        - SHIFTING (RESTOW): 20, 40, 45
        
        TUGAS PENTING:
        1. DG / IMO: Cari baris berlabel "DG", "IMO", atau "Dangerous". Ekstrak angkanya ke field '_dg'.
        2. HATCH COVER: Cari sel khusus berlabel "Hatch Cover". Nilainya biasanya KECIL (0-50). JANGAN ambil angka Total Box (ratusan/ribuan) sebagai Hatch Cover.
        3. Ekstrak angka integer. Jika sel kosong/strip, ISI 0.
        
        OUTPUT JSON FORMAT (snake_case):
        {
            "imp_20_full": int, "imp_20_dg": int, "imp_20_reefer": int, "imp_20_oog": int, "imp_20_ts_full": int, "imp_20_ts_dg": int, "imp_20_ts_empty": int, "imp_20_ts_oog": int, "imp_20_empty": int,
            "imp_40_full": int, "imp_40_dg": int, "imp_40_reefer": int, "imp_40_oog": int, "imp_40_ts_full": int, "imp_40_ts_dg": int, "imp_40_ts_empty": int, "imp_40_ts_oog": int, "imp_40_empty": int,
            "imp_45_full": int, "imp_45_dg": int, "imp_45_reefer": int, "imp_45_oog": int, "imp_45_ts_full": int, "imp_45_ts_dg": int, "imp_45_ts_empty": int, "imp_45_ts_oog": int, "imp_45_empty": int,

            "exp_20_full": int, "exp_20_dg": int, "exp_20_reefer": int, "exp_20_oog": int, "exp_20_ts_full": int, "exp_20_ts_dg": int, "exp_20_ts_empty": int, "exp_20_ts_oog": int, "exp_20_empty": int,
            "exp_40_full": int, "exp_40_dg": int, "exp_40_reefer": int, "exp_40_oog": int, "exp_40_ts_full": int, "exp_40_ts_dg": int, "exp_40_ts_empty": int, "exp_40_ts_oog": int, "exp_40_empty": int,
            "exp_45_full": int, "exp_45_dg": int, "exp_45_reefer": int, "exp_45_oog": int, "exp_45_ts_full": int, "exp_45_ts_dg": int, "exp_45_ts_empty": int, "exp_45_ts_oog": int, "exp_45_empty": int,

            "shift_20_full": int, "shift_20_reefer": int, "shift_20_empty": int,
            "shift_40_full": int, "shift_40_reefer": int, "shift_40_empty": int,
            "shift_45_full": int, "shift_45_reefer": int, "shift_45_empty": int,
            
            "hatch_cover": int
        }
        """
        payload = {"contents": [{"parts": [{"text": prompt_text}, {"inline_data": {"mime_type": "image/jpeg", "data": img_str}}]}], "generationConfig": {"response_mime_type": "application/json"}}

        try:
            response = requests.post(url, headers=headers, data=json.dumps(payload))
            if response.status_code == 200:
                try:
                    text_response = response.json()['candidates'][0]['content']['parts'][0]['text']
                    # Lebih robust dalam membersihkan JSON
                    clean_json = text_response.strip()
                    if "```json" in clean_json:
                        clean_json = clean_json.split("```json")[1].split("```")[0].strip()
                    elif "```" in clean_json:
                        clean_json = clean_json.split("```")[1].strip()
                    
                    return json.loads(clean_json)
                except Exception as e:
                    last_error_msg = f"Format JSON Salah ({model_name}): {e}"
                    continue # Coba model lain
                    
            elif response.status_code == 429: 
                last_error_msg = f"Limit Kuota ({model_name}). Mencoba model lain..."
                continue 
            elif response.status_code in [404, 500, 503]: 
                last_error_msg = f"Server Error {response.status_code} ({model_name})"
                continue 
            else: 
                last_error_msg = f"API Error {response.status_code}: {response.text}"
                break 
        except Exception as e: 
            last_error_msg = f"Koneksi Gagal: {e}"
            continue

    st.error(f"Gagal memproses gambar. Detail Error: {last_error_msg}")
    return None

# --- TOMBOL PROSES ---
if st.button("🚀 Mulai Proses Ekstraksi", type="primary", use_container_width=True):
    if not uploaded_files or not api_key:
        st.warning("Mohon lengkapi API Key dan Upload File.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for index, uploaded_file in enumerate(uploaded_files):
            status_text.caption(f"Sedang memproses: {uploaded_file.name}...")
            image = Image.open(uploaded_file)
            data = extract_table_data(image, api_key)
            
            if data:
                # --- CALCULATIONS & MAPPING (MERGE DG TO FULL) ---
                
                # Import Summary (DG is merged into Full Laden in calculation)
                # Note: Full in summary = Full Raw + DG Raw
                i_l_20 = data.get('imp_20_full',0) + data.get('imp_20_dg',0) + data.get('imp_20_reefer',0) + data.get('imp_20_oog',0)
                i_l_40 = data.get('imp_40_full',0) + data.get('imp_40_dg',0) + data.get('imp_40_reefer',0) + data.get('imp_40_oog',0)
                i_l_45 = data.get('imp_45_full',0) + data.get('imp_45_dg',0) + data.get('imp_45_reefer',0) + data.get('imp_45_oog',0)
                
                i_e_20 = data.get('imp_20_empty',0)
                i_e_40 = data.get('imp_40_empty',0)
                i_e_45 = data.get('imp_45_empty',0)
                
                # Export Summary
                e_l_20 = data.get('exp_20_full',0) + data.get('exp_20_dg',0) + data.get('exp_20_reefer',0) + data.get('exp_20_oog',0)
                e_l_40 = data.get('exp_40_full',0) + data.get('exp_40_dg',0) + data.get('exp_40_reefer',0) + data.get('exp_40_oog',0)
                e_l_45 = data.get('exp_45_full',0) + data.get('exp_45_dg',0) + data.get('exp_45_reefer',0) + data.get('exp_45_oog',0)
                
                e_e_20 = data.get('exp_20_empty',0)
                e_e_40 = data.get('exp_40_empty',0)
                e_e_45 = data.get('exp_45_empty',0)

                # TS Summary (DG kept separate or aggregated based on standard logic)
                # Note: Standard TS Summary usually just sums boxes.
                ts_l_20 = data.get('imp_20_ts_full',0) + data.get('imp_20_ts_dg',0) + data.get('imp_20_ts_oog',0) + data.get('exp_20_ts_full',0) + data.get('exp_20_ts_dg',0) + data.get('exp_20_ts_oog',0)
                ts_l_40 = data.get('imp_40_ts_full',0) + data.get('imp_40_ts_dg',0) + data.get('imp_40_ts_oog',0) + data.get('exp_40_ts_full',0) + data.get('exp_40_ts_dg',0) + data.get('exp_40_ts_oog',0)
                ts_l_45 = data.get('imp_45_ts_full',0) + data.get('imp_45_ts_dg',0) + data.get('imp_45_ts_oog',0) + data.get('exp_45_ts_full',0) + data.get('exp_45_ts_dg',0) + data.get('exp_45_ts_oog',0)
                ts_e_20 = data.get('imp_20_ts_empty',0) + data.get('exp_20_ts_empty',0)
                ts_e_40 = data.get('imp_40_ts_empty',0) + data.get('exp_40_ts_empty',0)
                ts_e_45 = data.get('imp_45_ts_empty',0) + data.get('exp_45_ts_empty',0)

                # TEUS
                teus_imp = (i_l_20*1 + i_l_40*2 + i_l_45*2.25) + (i_e_20*1 + i_e_40*2 + i_e_45*2.25)
                teus_exp = (e_l_20*1 + e_l_40*2 + e_l_45*2.25) + (e_e_20*1 + e_e_40*2 + e_e_45*2.25)
                teus_ts = (ts_l_20*1 + ts_l_40*2 + ts_l_45*2.25) + (ts_e_20*1 + ts_e_40*2 + ts_e_45*2.25)
                
                # Shifting
                s_20 = data.get('shift_20_full',0) + data.get('shift_20_reefer',0) + data.get('shift_20_empty',0)
                s_40 = data.get('shift_40_full',0) + data.get('shift_40_reefer',0) + data.get('shift_40_empty',0)
                s_45 = data.get('shift_45_full',0) + data.get('shift_45_reefer',0) + data.get('shift_45_empty',0)
                tot_shift = s_20 + s_40 + s_45
                teus_shift = (s_20*1) + (s_40*2) + (s_45*2.25)

                grand_tot_box = sum([i_l_20,i_l_40,i_l_45,i_e_20,i_e_40,i_e_45,
                                     e_l_20,e_l_40,e_l_45,e_e_20,e_e_40,e_e_45,
                                     ts_l_20,ts_l_40,ts_l_45,ts_e_20,ts_e_40,ts_e_45, tot_shift])
                grand_tot_teus = teus_imp + teus_exp + teus_ts + teus_shift

                data_id = len(st.session_state['extracted_data']) + 1
                
                # Logic Guard untuk Hatch Cover (Jika nilai tidak masuk akal > 200, set 0)
                hatch_val = data.get('hatch_cover', 0)
                if hatch_val > 200:
                    hatch_val = 0

                # Row Dict (UPDATED: AUTO NAME)
                row = {
                    "NO": data_id, "Vessel": f"Vessel ({data_id})", "Service Name": "-", "Remark": 0,
                    
                    "IMP_LADEN_20": i_l_20, "IMP_LADEN_40": i_l_40, "IMP_LADEN_45": i_l_45, "IMP_EMPTY_20": i_e_20, "IMP_EMPTY_40": i_e_40, "IMP_EMPTY_45": i_e_45,
                    "TOTAL BOX IMPORT": sum([i_l_20,i_l_40,i_l_45,i_e_20,i_e_40,i_e_45]), "TEUS IMPORT": teus_imp,

                    "EXP_LADEN_20": e_l_20, "EXP_LADEN_40": e_l_40, "EXP_LADEN_45": e_l_45, "EXP_EMPTY_20": e_e_20, "EXP_EMPTY_40": e_e_40, "EXP_EMPTY_45": e_e_45,
                    "TOTAL BOX EXPORT": sum([e_l_20,e_l_40,e_l_45,e_e_20,e_e_40,e_e_45]), "TEUS EXPORT": teus_exp,

                    "TS_LADEN_20": ts_l_20, "TS_LADEN_40": ts_l_40, "TS_LADEN_45": ts_l_45, "TS_EMPTY_20": ts_e_20, "TS_EMPTY_40": ts_e_40, "TS_EMPTY_45": ts_e_45,
                    "TOTAL BOX T/S": sum([ts_l_20,ts_l_40,ts_l_45,ts_e_20,ts_e_40,ts_e_45]), "TEUS T/S": teus_ts,

                    "TOTAL BOX SHIFTING": tot_shift, "TEUS SHIFTING": teus_shift,
                    "Total (Boxes)": grand_tot_box, "Total Teus": grand_tot_teus,

                    # RECON DATA (Updated: DG merged to Full for Import/Export, but KEPT separate for T/S)
                    
                    # IMPORT (Merge DG to Full)
                    "IMP_20_Full": data.get('imp_20_full',0) + data.get('imp_20_dg',0),
                    "IMP_20_Reefer": data.get('imp_20_reefer',0), "IMP_20_OOG": data.get('imp_20_oog',0),
                    "IMP_20_TS_Full": data.get('imp_20_ts_full',0), "IMP_20_TS_Reefer": 0, "IMP_20_TS_OOG": data.get('imp_20_ts_oog',0), "IMP_20_TS_DG": data.get('imp_20_ts_dg',0), "IMP_20_TS_Empty": data.get('imp_20_ts_empty',0), "IMP_20_Empty": data.get('imp_20_empty',0), "IMP_20_LCL": 0,
                    
                    "IMP_40_Full": data.get('imp_40_full',0) + data.get('imp_40_dg',0),
                    "IMP_40_Reefer": data.get('imp_40_reefer',0), "IMP_40_OOG": data.get('imp_40_oog',0),
                    "IMP_40_TS_Full": data.get('imp_40_ts_full',0), "IMP_40_TS_Reefer": 0, "IMP_40_TS_OOG": data.get('imp_40_ts_oog',0), "IMP_40_TS_DG": data.get('imp_40_ts_dg',0), "IMP_40_TS_Empty": data.get('imp_40_ts_empty',0), "IMP_40_Empty": data.get('imp_40_empty',0), "IMP_40_LCL": 0,
                    
                    "IMP_45_Full": data.get('imp_45_full',0) + data.get('imp_45_dg',0),
                    "IMP_45_Reefer": data.get('imp_45_reefer',0), "IMP_45_OOG": data.get('imp_45_oog',0),
                    "IMP_45_TS_Full": data.get('imp_45_ts_full',0), "IMP_45_TS_Reefer": 0, "IMP_45_TS_OOG": data.get('imp_45_ts_oog',0), "IMP_45_TS_DG": data.get('imp_45_ts_dg',0), "IMP_45_TS_Empty": data.get('imp_45_ts_empty',0), "IMP_45_Empty": data.get('imp_45_empty',0), "IMP_45_LCL": 0,

                    # EXPORT (Merge DG to Full)
                    "EXP_20_Full": data.get('exp_20_full',0) + data.get('exp_20_dg',0),
                    "EXP_20_Reefer": data.get('exp_20_reefer',0), "EXP_20_OOG": data.get('exp_20_oog',0),
                    "EXP_20_TS_Full": data.get('exp_20_ts_full',0), "EXP_20_TS_Reefer": 0, "EXP_20_TS_OOG": data.get('exp_20_ts_oog',0), "EXP_20_TS_DG": data.get('exp_20_ts_dg',0), "EXP_20_TS_Empty": data.get('exp_20_ts_empty',0), "EXP_20_Empty": data.get('exp_20_empty',0), "EXP_20_LCL": 0,
                    
                    "EXP_40_Full": data.get('exp_40_full',0) + data.get('exp_40_dg',0),
                    "EXP_40_Reefer": data.get('exp_40_reefer',0), "EXP_40_OOG": data.get('exp_40_oog',0),
                    "EXP_40_TS_Full": data.get('exp_40_ts_full',0), "EXP_40_TS_Reefer": 0, "EXP_40_TS_OOG": data.get('exp_40_ts_oog',0), "EXP_40_TS_DG": data.get('exp_40_ts_dg',0), "EXP_40_TS_Empty": data.get('exp_40_ts_empty',0), "EXP_40_Empty": data.get('exp_40_empty',0), "EXP_40_LCL": 0,
                    
                    "EXP_45_Full": data.get('exp_45_full',0) + data.get('exp_45_dg',0),
                    "EXP_45_Reefer": data.get('exp_45_reefer',0), "EXP_45_OOG": data.get('exp_45_oog',0),
                    "EXP_45_TS_Full": data.get('exp_45_ts_full',0), "EXP_45_TS_Reefer": 0, "EXP_45_TS_OOG": data.get('exp_45_ts_oog',0), "EXP_45_TS_DG": data.get('exp_45_ts_dg',0), "EXP_45_TS_Empty": data.get('exp_45_ts_empty',0), "EXP_45_Empty": data.get('exp_45_empty',0), "EXP_45_LCL": 0,

                    "SHIFT_20_Full": data.get('shift_20_full',0), "SHIFT_20_Reefer": data.get('shift_20_reefer',0), "SHIFT_20_empty": data.get('shift_20_empty',0), "SHIFT_20_LCL": 0,
                    "SHIFT_40_Full": data.get('shift_40_full',0), "SHIFT_40_Reefer": data.get('shift_40_reefer',0), "SHIFT_40_empty": data.get('shift_40_empty',0), "SHIFT_40_LCL": 0,
                    "SHIFT_45_Full": data.get('shift_45_full',0), "SHIFT_45_Reefer": data.get('shift_45_reefer',0), "SHIFT_45_empty": data.get('shift_45_empty',0), "SHIFT_45_LCL": 0,
                    
                    "Hatch Cover": hatch_val
                }
                
                st.session_state['extracted_data'].append(row)
                st.session_state['images'][data_id] = image
                
            progress_bar.progress((index + 1) / len(uploaded_files))
        
        status_text.success("Selesai!")
        time.sleep(1)
        st.rerun()

# --- TAMPILAN OUTPUT ---
if st.session_state['extracted_data']:
    df = pd.DataFrame(st.session_state['extracted_data'])
    st.divider()
    
    # Define Column Groups
    summary_cols = ["NO", "Vessel", "Service Name", "Remark", "IMP_LADEN_20", "IMP_LADEN_40", "IMP_LADEN_45", "IMP_EMPTY_20", "IMP_EMPTY_40", "IMP_EMPTY_45", "TOTAL BOX IMPORT", "TEUS IMPORT", "EXP_LADEN_20", "EXP_LADEN_40", "EXP_LADEN_45", "EXP_EMPTY_20", "EXP_EMPTY_40", "EXP_EMPTY_45", "TOTAL BOX EXPORT", "TEUS EXPORT", "TS_LADEN_20", "TS_LADEN_40", "TS_LADEN_45", "TS_EMPTY_20", "TS_EMPTY_40", "TS_EMPTY_45", "TOTAL BOX T/S", "TEUS T/S", "TOTAL BOX SHIFTING", "TEUS SHIFTING", "Total (Boxes)", "Total Teus"]
    recon_cols = ["Vessel", "Service Name"] + [c for c in df.columns if c not in summary_cols and c not in ["NO", "Vessel", "Service Name", "Remark"]]
    recon_cols.append("Hatch Cover")

    tab1, tab_recon, tab3 = st.tabs(["📋 Data Detail (Edit)", "🔬 Recon (Detail - Edit)", "➕ Gabung Data"])
    
    # --- TAB 1: SUMMARY ---
    with tab1:
        st.markdown("##### Hasil Ekstraksi (Summary)")
        render_image_viewer(df, "tab1") # Unified Viewer
        
        edited_df = st.data_editor(df[summary_cols], num_rows="dynamic", use_container_width=True, column_config={"NO": st.column_config.NumberColumn(disabled=True)})
        
        # Download Button untuk Summary Saja
        c1, c2 = st.columns([3,1])
        c1.caption("Copy data:")
        c1.code(edited_df.to_csv(index=False, sep='\t'), language='csv')
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Summary')
        c2.download_button("📥 Excel Summary", data=output.getvalue(), file_name=f"Summary_Data.xlsx", use_container_width=True)

    # --- TAB RECON: GRANULAR ---
    with tab_recon:
        st.markdown("##### Data Recon (Full Detail)")
        render_image_viewer(df, "recon") # Unified Viewer
        
        # --- FITUR SMART FILTER ---
        # FIX: Pastikan Vessel dan Service Name masuk, tapi NO dan Remark (jika ada di summary) tidak duplikat logic
        # Ambil semua kolom yg BUKAN summary, lalu tambahkan Vessel & Service di depan
        detail_cols = [c for c in df.columns if c not in summary_cols and c not in ["NO", "Vessel", "Service Name", "Remark"]]
        
        cols_for_recon = ["Vessel", "Service Name"] + detail_cols
        
        # Re-order columns to put Hatch Cover at the end if it exists
        if "Hatch Cover" in cols_for_recon:
             cols_for_recon.remove("Hatch Cover")
             cols_for_recon.append("Hatch Cover")

        df_recon_base = df[cols_for_recon]
        
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            st.caption("Mode Edit Aktif. Gunakan filter di kanan untuk menyederhanakan tampilan.")
        with col_f2:
            hide_zeros = st.checkbox("Sembunyikan kolom kosong (0)", value=True)
        
        if hide_zeros:
            # Filter kolom: Tampilkan jika Numeric & Sum != 0, ATAU jika kolom identitas (String)
            valid_cols = []
            for c in df_recon_base.columns:
                # Selalu tampilkan identitas
                if c in ["Vessel", "Service Name"]:
                    valid_cols.append(c)
                # Untuk data angka, cek apakah totalnya bukan 0
                elif pd.api.types.is_numeric_dtype(df_recon_base[c]):
                    if df_recon_base[c].sum() != 0:
                        valid_cols.append(c)
                # Kolom lain (non-numeric) tampilkan saja biar aman
                else:
                    valid_cols.append(c)
            
            df_recon_display = df_recon_base[valid_cols]
        else:
            df_recon_display = df_recon_base

        edited_df_recon = st.data_editor(df_recon_display, num_rows="dynamic", use_container_width=True)
        
        # Download Button untuk Recon Saja
        c_r1, c_r2 = st.columns([3,1])
        c_r1.caption("Copy data recon:")
        c_r1.code(edited_df_recon.to_csv(index=False, sep='\t'), language='csv')
        
        output_recon = io.BytesIO()
        with pd.ExcelWriter(output_recon, engine='openpyxl') as writer:
            # Download hasil edit (filtered/unfiltered)
            edited_df_recon.to_excel(writer, index=False, sheet_name='Recon')
        c_r2.download_button("📥 Excel Recon", data=output_recon.getvalue(), file_name=f"Recon_Data.xlsx", use_container_width=True)

    # --- TAB 3: COMBINE ---
    with tab3:
        st.markdown("##### Combine Multi-Kapal")
        options = df['NO'].tolist()
        choice_labels = {row['NO']: f"{row['NO']} - {row['Vessel']}" for index, row in df.iterrows()}
        selected_indices = st.multiselect("Pilih kapal:", options, format_func=lambda x: choice_labels.get(x))

        if selected_indices:
            subset_df = df[df['NO'].isin(selected_indices)]
            numeric_cols = subset_df.select_dtypes(include='number').columns
            cols_to_sum = [c for c in numeric_cols if c not in ['NO', 'Remark']]
            sum_row = subset_df[cols_to_sum].sum()
            
            combined_df = pd.DataFrame([sum_row])
            combined_df.insert(0, "NO", "GABUNGAN")
            combined_df.insert(1, "Vessel", "MULTIPLE VESSELS")
            combined_df.insert(2, "Service Name", "COMBINED")
            combined_df.insert(3, "Remark", "-")
            
            st.dataframe(combined_df[summary_cols], use_container_width=True)
            output_combine = io.BytesIO()
            with pd.ExcelWriter(output_combine, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Combined')
            st.download_button("📥 Download Gabungan", data=output_combine.getvalue(), file_name=f"Rekap_Gabungan.xlsx", use_container_width=True)
