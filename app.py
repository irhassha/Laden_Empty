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

    st.info("Versi Aplikasi: 1.3 (Enhanced)\nMode: Smart Failover + Editable")

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
        
        prompt_text = """
        Analisis gambar tabel operasi pelabuhan ini. Fokus hanya pada angka.
        TUGAS: Ekstrak data dan lakukan pemetaan kategori berikut:
        1. IMPORT (Kolom DISCHARGE): BOX LADEN (FULL+REEFER+OOG), BOX EMPTY.
        2. EXPORT (Kolom LOADING): BOX LADEN (FULL+REEFER+OOG), BOX EMPTY.
        3. TRANSHIPMENT / TS (Baris T/S): Ambil angkanya.
        4. SHIFTING: Ambil total kolom SHIFTING.
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
        # Jangan reset st.session_state['extracted_data'] di sini agar bisa menambah data (append)
        # Atau jika ingin reset setiap kali klik tombol, uncomment baris bawah:
        # st.session_state['extracted_data'] = [] 
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for index, uploaded_file in enumerate(uploaded_files):
            status_text.caption(f"Sedang memproses: {uploaded_file.name}...")
            image = Image.open(uploaded_file)
            data = extract_table_data(image, api_key)
            
            if data:
                # --- CALCULATION LOGIC ---
                # Import
                imp_l = [data.get(f'import_laden_{x}',0) for x in [20,40,45]]
                imp_e = [data.get(f'import_empty_{x}',0) for x in [20,40,45]]
                teus_imp = (imp_l[0]*1 + imp_l[1]*2 + imp_l[2]*2.25) + (imp_e[0]*1 + imp_e[1]*2 + imp_e[2]*2.25)
                
                # Export
                exp_l = [data.get(f'export_laden_{x}',0) for x in [20,40,45]]
                exp_e = [data.get(f'export_empty_{x}',0) for x in [20,40,45]]
                teus_exp = (exp_l[0]*1 + exp_l[1]*2 + exp_l[2]*2.25) + (exp_e[0]*1 + exp_e[1]*2 + exp_e[2]*2.25)

                # Transhipment
                ts_l = [data.get(f'ts_laden_{x}',0) for x in [20,40,45]]
                ts_e = [data.get(f'ts_empty_{x}',0) for x in [20,40,45]]
                teus_ts = (ts_l[0]*1 + ts_l[1]*2 + ts_l[2]*2.25) + (ts_e[0]*1 + ts_e[1]*2 + ts_e[2]*2.25)

                # Shifting
                sh_l = [data.get(f'shift_laden_{x}',0) for x in [20,40,45]]
                sh_e = [data.get(f'shift_empty_{x}',0) for x in [20,40,45]]
                tot_shift_box = data.get('total_shift_box', sum(sh_l)+sum(sh_e))

                # Row Construction
                row = {
                    "NO": len(st.session_state['extracted_data']) + 1, # Increment NO based on existing data
                    "Vessel": f"{input_vessel} ({len(st.session_state['extracted_data']) + 1})",
                    "Service Name": input_service,
                    "Remark": 0,
                    # Import
                    "IMP_LADEN_20": imp_l[0], "IMP_LADEN_40": imp_l[1], "IMP_LADEN_45": imp_l[2],
                    "IMP_EMPTY_20": imp_e[0], "IMP_EMPTY_40": imp_e[1], "IMP_EMPTY_45": imp_e[2],
                    "TOTAL BOX IMPORT": sum(imp_l)+sum(imp_e), "TEUS IMPORT": teus_imp,
                    # Export
                    "EXP_LADEN_20": exp_l[0], "EXP_LADEN_40": exp_l[1], "EXP_LADEN_45": exp_l[2],
                    "EXP_EMPTY_20": exp_e[0], "EXP_EMPTY_40": exp_e[1], "EXP_EMPTY_45": exp_e[2],
                    "TOTAL BOX EXPORT": sum(exp_l)+sum(exp_e), "TEUS EXPORT": teus_exp,
                    # TS
                    "TS_LADEN_20": ts_l[0], "TS_LADEN_40": ts_l[1], "TS_LADEN_45": ts_l[2],
                    "TS_EMPTY_20": ts_e[0], "TS_EMPTY_40": ts_e[1], "TS_EMPTY_45": ts_e[2],
                    "TOTAL BOX T/S": sum(ts_l)+sum(ts_e), "TEUS T/S": teus_ts,
                    # Shifting
                    "SHIFT_LADEN_20": sh_l[0], "SHIFT_LADEN_40": sh_l[1], "SHIFT_LADEN_45": sh_l[2],
                    "SHIFT_EMPTY_20": sh_e[0], "SHIFT_EMPTY_40": sh_e[1], "SHIFT_EMPTY_45": sh_e[2],
                    "TOTAL BOX SHIFTING": tot_shift_box, "TEUS SHIFTING": data.get('total_shift_teus',0),
                    # Grand Total
                    "Total (Boxes)": (sum(imp_l)+sum(imp_e)) + (sum(exp_l)+sum(exp_e)) + (sum(ts_l)+sum(ts_e)) + tot_shift_box,
                    "Total Teus": teus_imp + teus_exp + teus_ts + data.get('total_shift_teus',0)
                }
                st.session_state['extracted_data'].append(row)
            progress_bar.progress((index + 1) / len(uploaded_files))
        
        status_text.success("Selesai!")
        time.sleep(1)
        st.rerun()

# --- DISPLAY HASIL (TABS VIEW) ---
if st.session_state['extracted_data']:
    
    # Konversi state ke DataFrame awal
    df_initial = pd.DataFrame(st.session_state['extracted_data'])
    teus_cols = ["TEUS IMPORT", "TEUS EXPORT", "TEUS T/S", "TEUS SHIFTING", "Total Teus"]
    
    st.divider()
    
    # TABS: Detail, Dashboard, Combine
    tab1, tab2, tab3 = st.tabs(["📋 Data Detail (Edit)", "📊 Dashboard", "➕ Gabung Data (Combine)"])
    
    # --- TAB 1: DATA EDITOR ---
    with tab1:
        st.markdown("##### Hasil Ekstraksi (Bisa Diedit)")
        st.caption("Klik dua kali pada sel untuk mengoreksi angka jika ada kesalahan OCR.")
        
        # FITUR UTAMA: DATA EDITOR
        edited_df = st.data_editor(
            df_initial, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "NO": st.column_config.NumberColumn(disabled=True),
                "Total Teus": st.column_config.NumberColumn(format="%.2f"),
            }
        )
        
        # Download & Copy
        c1, c2 = st.columns([3, 1])
        with c1:
            st.caption("Copy data di bawah ini untuk Paste ke Excel:")
            st.code(edited_df.to_csv(index=False, sep='\t'), language='csv')
        with c2:
            st.write("") 
            st.write("") 
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
            st.download_button(
                label="📥 Download Excel", 
                data=output.getvalue(), 
                file_name=f"Rekap_{input_vessel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # --- TAB 2: DASHBOARD (SIMPLE VISUALIZATION) ---
    with tab2:
        st.markdown("##### Ringkasan Volume (Total TEUS)")
        
        # Siapkan data untuk grafik
        if not edited_df.empty:
            chart_data = edited_df[["Vessel", "TEUS IMPORT", "TEUS EXPORT", "TEUS T/S", "TEUS SHIFTING"]].set_index("Vessel")
            st.bar_chart(chart_data)
            
            # Kartu Metrik Total
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Import (TEUs)", f"{edited_df['TEUS IMPORT'].sum():,.2f}")
            m2.metric("Total Export (TEUs)", f"{edited_df['TEUS EXPORT'].sum():,.2f}")
            m3.metric("Grand Total (TEUs)", f"{edited_df['Total Teus'].sum():,.2f}")

    # --- TAB 3: COMBINE DATA ---
    with tab3:
        st.markdown("##### Fitur Penjumlahan Multi-Kapal")
        
        # Gunakan edited_df agar koreksi user terbawa ke sini
        options = edited_df['NO'].tolist()
        choice_labels = {row['NO']: f"{row['NO']} - {row['Vessel']}" for index, row in edited_df.iterrows()}
        selected_indices = st.multiselect("Pilih kapal yang ingin dijumlahkan:", options, format_func=lambda x: choice_labels.get(x))

        if selected_indices:
            subset_df = edited_df[edited_df['NO'].isin(selected_indices)]
            st.info(f"Menjumlahkan {len(selected_indices)} data terpilih...")
            
            numeric_cols = subset_df.select_dtypes(include='number').columns
            cols_to_sum = [c for c in numeric_cols if c not in ['NO', 'Remark']]
            sum_row = subset_df[cols_to_sum].sum()
            
            combined_df = pd.DataFrame([sum_row])
            combined_df.insert(0, "NO", "GABUNGAN")
            combined_df.insert(1, "Vessel", "MULTIPLE VESSELS")
            combined_df.insert(2, "Service Name", "COMBINED")
            combined_df.insert(3, "Remark", "-")
            
            st.dataframe(combined_df.style.format("{:.2f}", subset=teus_cols), use_container_width=True)
            
            c1, c2 = st.columns([3, 1])
            with c1:
                st.code(combined_df.to_csv(index=False, sep='\t'), language='csv')
            with c2:
                st.write("")
                st.write("")
                output_combine = io.BytesIO()
                with pd.ExcelWriter(output_combine, engine='openpyxl') as writer:
                    combined_df.to_excel(writer, index=False, sheet_name='Combined')
                st.download_button(
                    label="📥 Download Gabungan", 
                    data=output_combine.getvalue(), 
                    file_name=f"Rekap_Gabungan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
