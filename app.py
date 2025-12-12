import streamlit as st
import pandas as pd
import requests
import base64
import json
import io
import time
import altair as alt
from PIL import Image

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    layout="wide", 
    page_title="NPCT1 Auto Tally",
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
    /* Mempercantik Metrics/Cards */
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

# --- FUNGSI UTILITY: CACHE DATA MODEL ---
@st.cache_data(ttl=300) 
def get_prioritized_models(api_key):
    """Mengambil semua model yang tersedia dan mengurutkannya."""
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

# --- SIDEBAR: KONFIGURASI API (SET & FORGET) ---
with st.sidebar:
    st.title("⚙️ Pengaturan")
    
    # API Key Input
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ API Key Terhubung")
    else:
        api_key = st.text_input("Google Gemini API Key", type="password", placeholder="Paste key disini...")
        st.caption("[Dapatkan Key Gratis](https://aistudio.google.com/)")
    
    if api_key:
        api_key = api_key.strip()
        st.divider()
        
        # Status Koneksi Simple
        with st.spinner("Cek koneksi..."):
            active_models = get_prioritized_models(api_key)
            
        if active_models:
            st.markdown(f"**Status:** 🟢 Online ({len(active_models)} Models)")
            st.progress(100)
        else:
            st.markdown("**Status:** 🔴 Offline")
            st.error("API Key Invalid")
            
    st.divider()
    # Tombol Reset Data
    if st.button("🗑️ Hapus Semua Data", type="secondary", use_container_width=True):
        st.session_state['extracted_data'] = []
        st.rerun()

    st.info("Versi Aplikasi: 1.4 (Recon)\nMode: Smart Failover + Recon Tab")

# --- HEADER APLIKASI ---
st.title("⚓ NPCT1 Tally Extractor")
st.markdown("Automasi ekstraksi data operasional pelabuhan dari gambar laporan ke Excel.")
st.divider()

# --- INPUT AREA (STEP 1 & 2) ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Identitas Kapal")
    input_vessel = st.text_input("Nama Kapal (Vessel Name)", value="Vessel A", placeholder="Contoh: MV. SINAR SUNDA")
    input_service = st.text_input("Service / Voyage", value="Service A", placeholder="Contoh: 001N")

with col2:
    st.subheader("2. Upload Laporan")
    uploaded_files = st.file_uploader("Upload Potongan Gambar Tabel", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

# --- FUNGSI EKSTRAKSI CORE ---
def extract_table_data(image, api_key):
    if image.mode != 'RGB': image = image.convert('RGB')
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

    candidate_models = get_prioritized_models(api_key)
    if not candidate_models: candidate_models = ["gemini-1.5-flash", "gemini-1.5-flash-latest"]
    
    last_error_msg = ""
    for model_name in candidate_models:
        if "experimental" in model_name: continue

        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        
        # PROMPT DIPERBARUI UNTUK DATA RECON (GRANULAR)
        prompt_text = """
        Analisis gambar tabel operasi pelabuhan ini. Ambil SEMUA angka detail per sel.

        STRUKTUR: Baris (FULL, REEFER, OOG, T/S FULL, T/S EMPTY, T/S OOG, EMPTY) x Kolom (20', 40', 45').

        TUGAS:
        1. IMPORT (Kolom DISCHARGE): Ambil angka raw.
        2. EXPORT (Kolom LOADING): Ambil angka raw.
        3. SHIFTING: Ambil angka raw.
        4. HATCH COVER: Ambil angka total.

        OUTPUT JSON (snake_case, integer, 0 jika kosong/tidak ada):
        {
            // IMPORT (DISCHARGE)
            "imp_20_full": int, "imp_20_reefer": int, "imp_20_oog": int, "imp_20_ts_full": int, "imp_20_ts_empty": int, "imp_20_ts_oog": int, "imp_20_empty": int,
            "imp_40_full": int, "imp_40_reefer": int, "imp_40_oog": int, "imp_40_ts_full": int, "imp_40_ts_empty": int, "imp_40_ts_oog": int, "imp_40_empty": int,
            "imp_45_full": int, "imp_45_reefer": int, "imp_45_oog": int, "imp_45_ts_full": int, "imp_45_ts_empty": int, "imp_45_ts_oog": int, "imp_45_empty": int,
            "hatch_cover": int,

            // EXPORT (LOADING)
            "exp_20_full": int, "exp_20_reefer": int, "exp_20_oog": int, "exp_20_ts_full": int, "exp_20_ts_empty": int, "exp_20_ts_oog": int, "exp_20_empty": int,
            "exp_40_full": int, "exp_40_reefer": int, "exp_40_oog": int, "exp_40_ts_full": int, "exp_40_ts_empty": int, "exp_40_ts_oog": int, "exp_40_empty": int,
            "exp_45_full": int, "exp_45_reefer": int, "exp_45_oog": int, "exp_45_ts_full": int, "exp_45_ts_empty": int, "exp_45_ts_oog": int, "exp_45_empty": int,

            // SHIFTING
            "shift_total_box": int, "shift_total_teus": float
        }
        """
        payload = {"contents": [{"parts": [{"text": prompt_text}, {"inline_data": {"mime_type": "image/jpeg", "data": img_str}}]}], "generationConfig": {"response_mime_type": "application/json"}}

        try:
            response = requests.post(url, headers=headers, data=json.dumps(payload))
            if response.status_code == 200:
                clean_json = response.json()['candidates'][0]['content']['parts'][0]['text'].replace('```json', '').replace('```', '').strip()
                return json.loads(clean_json)
            elif response.status_code == 429: continue 
            elif response.status_code in [404, 500, 503]: continue 
            else: break 
        except Exception as e: continue

    st.error(f"Gagal memproses gambar. Detail: {last_error_msg}")
    return None

# --- TOMBOL AKSI ---
if st.button("🚀 Mulai Proses Ekstraksi", type="primary", use_container_width=True):
    if not uploaded_files or not api_key:
        st.warning("Mohon lengkapi API Key dan Upload File terlebih dahulu.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for index, uploaded_file in enumerate(uploaded_files):
            status_text.caption(f"Sedang memproses: {uploaded_file.name}...")
            image = Image.open(uploaded_file)
            data = extract_table_data(image, api_key)
            
            if data:
                # --- CALCULATION LOGIC (MAPPING RAW TO SUMMARY) ---
                
                # Helper untuk hitung Laden/Empty 20/40/45 dari raw data
                # Import Summary
                i_l_20 = data.get('imp_20_full',0) + data.get('imp_20_reefer',0) + data.get('imp_20_oog',0)
                i_l_40 = data.get('imp_40_full',0) + data.get('imp_40_reefer',0) + data.get('imp_40_oog',0)
                i_l_45 = data.get('imp_45_full',0) + data.get('imp_45_reefer',0) + data.get('imp_45_oog',0)
                
                i_e_20 = data.get('imp_20_empty',0)
                i_e_40 = data.get('imp_40_empty',0)
                i_e_45 = data.get('imp_45_empty',0)
                
                # Export Summary
                e_l_20 = data.get('exp_20_full',0) + data.get('exp_20_reefer',0) + data.get('exp_20_oog',0)
                e_l_40 = data.get('exp_40_full',0) + data.get('exp_40_reefer',0) + data.get('exp_40_oog',0)
                e_l_45 = data.get('exp_45_full',0) + data.get('exp_45_reefer',0) + data.get('exp_45_oog',0)
                
                e_e_20 = data.get('exp_20_empty',0)
                e_e_40 = data.get('exp_40_empty',0)
                e_e_45 = data.get('exp_45_empty',0)

                # TS Summary (Dari baris TS di Import + TS di Export jika format tabel terpisah, 
                # tapi biasanya TS adalah baris sendiri. Di script ini kita mapping T/S import/export rows ke TS Summary)
                ts_l_20 = data.get('imp_20_ts_full',0) + data.get('imp_20_ts_oog',0) + data.get('exp_20_ts_full',0) + data.get('exp_20_ts_oog',0)
                ts_l_40 = data.get('imp_40_ts_full',0) + data.get('imp_40_ts_oog',0) + data.get('exp_40_ts_full',0) + data.get('exp_40_ts_oog',0)
                ts_l_45 = data.get('imp_45_ts_full',0) + data.get('imp_45_ts_oog',0) + data.get('exp_45_ts_full',0) + data.get('exp_45_ts_oog',0)
                
                ts_e_20 = data.get('imp_20_ts_empty',0) + data.get('exp_20_ts_empty',0)
                ts_e_40 = data.get('imp_40_ts_empty',0) + data.get('exp_40_ts_empty',0)
                ts_e_45 = data.get('imp_45_ts_empty',0) + data.get('exp_45_ts_empty',0)

                # TEUS CALC
                teus_imp = (i_l_20*1 + i_l_40*2 + i_l_45*2.25) + (i_e_20*1 + i_e_40*2 + i_e_45*2.25)
                teus_exp = (e_l_20*1 + e_l_40*2 + e_l_45*2.25) + (e_e_20*1 + e_e_40*2 + e_e_45*2.25)
                teus_ts = (ts_l_20*1 + ts_l_40*2 + ts_l_45*2.25) + (ts_e_20*1 + ts_e_40*2 + ts_e_45*2.25)
                
                # TOTALS
                tot_shift = data.get('shift_total_box', 0)
                grand_tot_box = (sum([i_l_20,i_l_40,i_l_45,i_e_20,i_e_40,i_e_45]) + 
                                 sum([e_l_20,e_l_40,e_l_45,e_e_20,e_e_40,e_e_45]) + 
                                 sum([ts_l_20,ts_l_40,ts_l_45,ts_e_20,ts_e_40,ts_e_45]) + tot_shift)
                
                grand_tot_teus = teus_imp + teus_exp + teus_ts + data.get('shift_total_teus', 0)

                # Row Construction (Gabungan Data Summary & Raw Data untuk Recon)
                row = {
                    "NO": len(st.session_state['extracted_data']) + 1,
                    "Vessel": f"{input_vessel} ({len(st.session_state['extracted_data']) + 1})",
                    "Service Name": input_service,
                    "Remark": 0,
                    
                    # --- SUMMARY COLUMNS (Untuk Tab 1) ---
                    "IMP_LADEN_20": i_l_20, "IMP_LADEN_40": i_l_40, "IMP_LADEN_45": i_l_45,
                    "IMP_EMPTY_20": i_e_20, "IMP_EMPTY_40": i_e_40, "IMP_EMPTY_45": i_e_45,
                    "TOTAL BOX IMPORT": sum([i_l_20,i_l_40,i_l_45,i_e_20,i_e_40,i_e_45]), "TEUS IMPORT": teus_imp,

                    "EXP_LADEN_20": e_l_20, "EXP_LADEN_40": e_l_40, "EXP_LADEN_45": e_l_45,
                    "EXP_EMPTY_20": e_e_20, "EXP_EMPTY_40": e_e_40, "EXP_EMPTY_45": e_e_45,
                    "TOTAL BOX EXPORT": sum([e_l_20,e_l_40,e_l_45,e_e_20,e_e_40,e_e_45]), "TEUS EXPORT": teus_exp,

                    "TS_LADEN_20": ts_l_20, "TS_LADEN_40": ts_l_40, "TS_LADEN_45": ts_l_45,
                    "TS_EMPTY_20": ts_e_20, "TS_EMPTY_40": ts_e_40, "TS_EMPTY_45": ts_e_45,
                    "TOTAL BOX T/S": sum([ts_l_20,ts_l_40,ts_l_45,ts_e_20,ts_e_40,ts_e_45]), "TEUS T/S": teus_ts,

                    "TOTAL BOX SHIFTING": tot_shift, "TEUS SHIFTING": data.get('shift_total_teus',0),
                    "Total (Boxes)": grand_tot_box, "Total Teus": grand_tot_teus,

                    # --- RAW DATA COLUMNS (Untuk Tab Recon) ---
                    # Import Detailed
                    "I_20_Full": data.get('imp_20_full',0), "I_20_Reefer": data.get('imp_20_reefer',0), "I_20_OOG": data.get('imp_20_oog',0),
                    "I_20_TS_Full": data.get('imp_20_ts_full',0), "I_20_TS_Reefer": 0, "I_20_TS_OOG": data.get('imp_20_ts_oog',0), "I_20_TS_Empty": data.get('imp_20_ts_empty',0), "I_20_Empty": data.get('imp_20_empty',0),
                    
                    "I_40_Full": data.get('imp_40_full',0), "I_40_Reefer": data.get('imp_40_reefer',0), "I_40_OOG": data.get('imp_40_oog',0),
                    "I_40_TS_Full": data.get('imp_40_ts_full',0), "I_40_TS_Reefer": 0, "I_40_TS_OOG": data.get('imp_40_ts_oog',0), "I_40_TS_Empty": data.get('imp_40_ts_empty',0), "I_40_Empty": data.get('imp_40_empty',0),
                    
                    "I_45_Full": data.get('imp_45_full',0), "I_45_Reefer": data.get('imp_45_reefer',0), "I_45_OOG": data.get('imp_45_oog',0),
                    "I_45_TS_Full": data.get('imp_45_ts_full',0), "I_45_TS_Reefer": 0, "I_45_TS_OOG": data.get('imp_45_ts_oog',0), "I_45_TS_Empty": data.get('imp_45_ts_empty',0), "I_45_Empty": data.get('imp_45_empty',0),
                    
                    "Hatch Cover": data.get('hatch_cover', 0),

                    # Export Detailed
                    "E_20_Full": data.get('exp_20_full',0), "E_20_Reefer": data.get('exp_20_reefer',0), "E_20_OOG": data.get('exp_20_oog',0),
                    "E_20_TS_Full": data.get('exp_20_ts_full',0), "E_20_TS_Reefer": 0, "E_20_TS_OOG": data.get('exp_20_ts_oog',0), "E_20_TS_Empty": data.get('exp_20_ts_empty',0), "E_20_Empty": data.get('exp_20_empty',0),
                    
                    "E_40_Full": data.get('exp_40_full',0), "E_40_Reefer": data.get('exp_40_reefer',0), "E_40_OOG": data.get('exp_40_oog',0),
                    "E_40_TS_Full": data.get('exp_40_ts_full',0), "E_40_TS_Reefer": 0, "E_40_TS_OOG": data.get('exp_40_ts_oog',0), "E_40_TS_Empty": data.get('exp_40_ts_empty',0), "E_40_Empty": data.get('exp_40_empty',0),
                    
                    "E_45_Full": data.get('exp_45_full',0), "E_45_Reefer": data.get('exp_45_reefer',0), "E_45_OOG": data.get('exp_45_oog',0),
                    "E_45_TS_Full": data.get('exp_45_ts_full',0), "E_45_TS_Reefer": 0, "E_45_TS_OOG": data.get('exp_45_ts_oog',0), "E_45_TS_Empty": data.get('exp_45_ts_empty',0), "E_45_Empty": data.get('exp_45_empty',0),
                }
                st.session_state['extracted_data'].append(row)
            progress_bar.progress((index + 1) / len(uploaded_files))
        
        status_text.success("Selesai!")
        time.sleep(1)
        st.rerun()

# --- DISPLAY HASIL (TABS VIEW) ---
if st.session_state['extracted_data']:
    
    df_initial = pd.DataFrame(st.session_state['extracted_data'])
    
    st.divider()
    
    # DEFINISI KOLOM UNTUK TAB 1 (SUMMARY)
    summary_cols = [
        "NO", "Vessel", "Service Name", "Remark",
        "IMP_LADEN_20", "IMP_LADEN_40", "IMP_LADEN_45", "IMP_EMPTY_20", "IMP_EMPTY_40", "IMP_EMPTY_45", "TOTAL BOX IMPORT", "TEUS IMPORT",
        "EXP_LADEN_20", "EXP_LADEN_40", "EXP_LADEN_45", "EXP_EMPTY_20", "EXP_EMPTY_40", "EXP_EMPTY_45", "TOTAL BOX EXPORT", "TEUS EXPORT",
        "TS_LADEN_20", "TS_LADEN_40", "TS_LADEN_45", "TS_EMPTY_20", "TS_EMPTY_40", "TS_EMPTY_45", "TOTAL BOX T/S", "TEUS T/S",
        "TOTAL BOX SHIFTING", "TEUS SHIFTING", "Total (Boxes)", "Total Teus"
    ]

    # DEFINISI KOLOM UNTUK TAB RECON (DETAILED)
    recon_cols = [
        "Vessel", "Service Name",
        # IMPORT
        "I_20_Full", "I_20_Reefer", "I_20_OOG", "I_20_TS_Full", "I_20_TS_Reefer", "I_20_TS_OOG", "I_20_TS_Empty", "I_20_Empty",
        "I_40_Full", "I_40_Reefer", "I_40_OOG", "I_40_TS_Full", "I_40_TS_Reefer", "I_40_TS_OOG", "I_40_TS_Empty", "I_40_Empty",
        "I_45_Full", "I_45_Reefer", "I_45_OOG", "I_45_TS_Full", "I_45_TS_Reefer", "I_45_TS_OOG", "I_45_TS_Empty", "I_45_Empty",
        "Hatch Cover", "TOTAL BOX IMPORT", "TEUS IMPORT",
        # EXPORT
        "E_20_Full", "E_20_Reefer", "E_20_OOG", "E_20_TS_Full", "E_20_TS_Reefer", "E_20_TS_OOG", "E_20_TS_Empty", "E_20_Empty",
        "E_40_Full", "E_40_Reefer", "E_40_OOG", "E_40_TS_Full", "E_40_TS_Reefer", "E_40_TS_OOG", "E_40_TS_Empty", "E_40_Empty",
        "E_45_Full", "E_45_Reefer", "E_45_OOG", "E_45_TS_Full", "E_45_TS_Reefer", "E_45_TS_OOG", "E_45_TS_Empty", "E_45_Empty",
        "TOTAL BOX EXPORT", "TEUS EXPORT"
    ]
    
    tab1, tab_recon, tab2, tab3 = st.tabs(["📋 Data Detail (Edit)", "🔬 Recon (Detail)", "📊 Dashboard", "➕ Gabung Data"])
    
    # --- TAB 1: DATA EDITOR (SUMMARY) ---
    with tab1:
        st.markdown("##### Hasil Ekstraksi (Summary)")
        
        # Filter hanya kolom summary untuk ditampilkan disini agar tidak pusing
        df_summary = df_initial[summary_cols]
        
        edited_df = st.data_editor(
            df_summary, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={"NO": st.column_config.NumberColumn(disabled=True)}
        )
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.caption("Copy data summary:")
            st.code(edited_df.to_csv(index=False, sep='\t'), language='csv')
        with c2:
            st.write("") 
            st.write("") 
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Summary')
            st.download_button("📥 Excel Summary", data=output.getvalue(), file_name=f"Rekap_{input_vessel}.xlsx", use_container_width=True)

    # --- TAB RECON: DATA GRANULAR ---
    with tab_recon:
        st.markdown("##### Data Recon (Full Detail)")
        st.caption("Tabel ini berisi rincian raw per jenis container (Full, Reefer, OOG, dll) untuk keperluan rekonsiliasi.")
        
        # Filter kolom recon
        df_recon = df_initial[recon_cols]
        st.dataframe(df_recon, use_container_width=True)
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.caption("Copy data recon:")
            st.code(df_recon.to_csv(index=False, sep='\t'), language='csv')
        with c2:
            st.write("") 
            st.write("") 
            output_recon = io.BytesIO()
            with pd.ExcelWriter(output_recon, engine='openpyxl') as writer:
                df_recon.to_excel(writer, index=False, sheet_name='Recon')
            st.download_button("📥 Excel Recon", data=output_recon.getvalue(), file_name=f"Recon_{input_vessel}.xlsx", use_container_width=True)

    # --- TAB 2: DASHBOARD ---
    with tab2:
        st.markdown("##### Ringkasan Volume (Total TEUS)")
        if not df_initial.empty:
            summary_data = {
                "Activity": ["Import", "Export", "Transhipment", "Shifting"],
                "Total TEUs": [
                    df_initial['TEUS IMPORT'].sum(),
                    df_initial['TEUS EXPORT'].sum(),
                    df_initial['TEUS T/S'].sum(),
                    df_initial['TEUS SHIFTING'].sum()
                ]
            }
            chart_df = pd.DataFrame(summary_data)
            chart = alt.Chart(chart_df).mark_bar().encode(
                x=alt.X('Total TEUs', title='Total TEUs'),
                y=alt.Y('Activity', sort='-x', title='Aktivitas'),
                color=alt.Color('Activity', legend=None), 
                tooltip=['Activity', 'Total TEUs']
            ).properties(height=300).interactive()
            st.altair_chart(chart, use_container_width=True)

    # --- TAB 3: COMBINE DATA ---
    with tab3:
        st.markdown("##### Fitur Penjumlahan Multi-Kapal")
        options = df_initial['NO'].tolist()
        choice_labels = {row['NO']: f"{row['NO']} - {row['Vessel']}" for index, row in df_initial.iterrows()}
        selected_indices = st.multiselect("Pilih kapal:", options, format_func=lambda x: choice_labels.get(x))

        if selected_indices:
            subset_df = df_initial[df_initial['NO'].isin(selected_indices)]
            numeric_cols = subset_df.select_dtypes(include='number').columns
            cols_to_sum = [c for c in numeric_cols if c not in ['NO', 'Remark']]
            sum_row = subset_df[cols_to_sum].sum()
            
            combined_df = pd.DataFrame([sum_row])
            combined_df.insert(0, "NO", "GABUNGAN")
            combined_df.insert(1, "Vessel", "MULTIPLE VESSELS")
            combined_df.insert(2, "Service Name", "COMBINED")
            combined_df.insert(3, "Remark", "-")
            
            # Tampilkan Summary Combine
            st.dataframe(combined_df[summary_cols], use_container_width=True)
            
            output_combine = io.BytesIO()
            with pd.ExcelWriter(output_combine, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Combined')
            st.download_button("📥 Download Gabungan", data=output_combine.getvalue(), file_name=f"Rekap_Gabungan.xlsx", use_container_width=True)
