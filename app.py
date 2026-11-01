import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(layout="wide", page_title="PKP & PPKP Workforce Norm Optimizer", page_icon="⚖️")

st.title("⚖️ PKP & PPKP Staffing Norm & Distribution Matrix")
st.caption("Workload-Driven Relative Value Unit (RVU) & Span-of-Control Allocation Engine | KKM")
st.markdown("---")

# =============================================================================
# 1. DATA INGESTION & ROBUST COLUMN CLEANING
# =============================================================================
EXCEL_FILENAME = "Data Perjawatan dan Beban Tugas PKP PPKP di PKD 16102026.xlsx"

@st.cache_data
def load_and_prep_data(filepath):
    df_raw = pd.read_excel(filepath, header=[0, 1])
    df = df_raw.copy()

    # Flatten multi-index column headers
    flat_cols = []
    for c1, c2 in df.columns:
        c1_s = str(c1).strip() if "Unnamed" not in str(c1) else ""
        c2_s = str(c2).strip() if "Unnamed" not in str(c2) else ""
        if c1_s and c2_s:
            name = f"{c1_s}_{c2_s}"
        elif c1_s:
            name = c1_s
        else:
            name = c2_s
        
        # Clean special characters
        name = (name.replace("\n", "")
                    .replace(" ", "_")
                    .replace("/", "_")
                    .replace("-", "_")
                    .replace("(", "")
                    .replace(")", "")
                    .replace(",", "")
                    .replace("%", "pct"))
        flat_cols.append(name)

    df.columns = flat_cols

    # Exclude total row at the bottom
    df = df[df['Daerah'].notna() & (df['Daerah'] != 'Total') & (df['Kategori'] != 'Total')].copy()

    # Directly and safely map the staffing columns
    df['Bil_PKP_Total'] = (
        pd.to_numeric(df.get('Bil_PKP', 0), errors='coerce').fillna(0) +
        pd.to_numeric(df.get('Bil_PKP_Kontrak_2025', 0), errors='coerce').fillna(0)
    )
    df['Bil_PPKP_Total'] = (
        pd.to_numeric(df.get('Bil_PPKP', 0), errors='coerce').fillna(0) +
        pd.to_numeric(df.get('Bil_PPKP_Kontrak_2025', 0), errors='coerce').fillna(0)
    )

    # Convert all burden task columns to float
    task_cols = [c for c in df.columns if c.startswith('Beban_Tugas_')]
    for col in task_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, task_cols

try:
    df, task_cols = load_and_prep_data(EXCEL_FILENAME)
except Exception as e:
    st.error(f"⚠️ Error loading file: {e}")
    st.info("Ensure the file exists in the directory and run using: `streamlit run app.py`")
    st.stop()

# =============================================================================
# 2. SIDEBAR: CADRE SELECTOR & RVU COMPLEXITY WEIGHTS
# =============================================================================
st.sidebar.header("🎯 Target Cadre & Policy Model")
selected_cadre = st.sidebar.radio("Select Cadre to Optimize:", ["PPKP (Ground Operational)", "PKP (Supervisory & Governance)"])

st.sidebar.markdown("---")
st.sidebar.header("⚖️ Task Complexity Multipliers (RVU)")
st.sidebar.caption("Assign relative weight (1 = Routine/Short, 10 = Highly Complex/Time-Intensive):")

with st.sidebar.expander("1. Kawalan Penyakit & Wabak", expanded=True):
    w_outbreak = st.slider("Kawalan Penyakit Berjangkit", 1, 10, 3)
    w_sampling = st.slider("Persampelan Penyakit", 1, 10, 4)
    w_tracing = st.slider("Siasatan Kes & Kontak", 1, 10, 4)
    w_bilik_gerakan = st.slider("Bilik Gerakan Dibuka", 1, 10, 7)
    w_survelans = st.slider("Laporan Survelans", 1, 10, 3)

with st.sidebar.expander("2. Perundangan & Pendakwaan", expanded=False):
    w_risikan = st.slider("Risikan Undang-Undang", 1, 10, 5)
    w_notis = st.slider("Notis Kompaun / Kompaun", 1, 10, 3)
    w_kertas_siasatan = st.slider("Kertas Siasatan (IP Prosecution)", 1, 10, 9)
    w_audit_pewasapan = st.slider("Audit Pewasapan", 1, 10, 4)

with st.sidebar.expander("3. Kesihatan Persekitaran & Pekerjaan", expanded=False):
    w_premis = st.slider("Pemeriksaan Kesihatan Premis", 1, 10, 3)
    w_makanan = st.slider("Penilaian Risiko Premis Makanan", 1, 10, 4)
    w_osh_audit = st.slider("Audit KKP (OSH) di Tempat Kerja", 1, 10, 6)
    w_osh_siasatan = st.slider("Siasatan Kemalangan/Penyakit OSH", 1, 10, 6)
    w_air = st.slider("Pemantauan Kualiti Air Minum", 1, 10, 3)

with st.sidebar.expander("4. Kawalan Pintu Masuk Antarabangsa (PMA)", expanded=False):
    w_kapal = st.slider("Pemeriksaan Kapal / Vessel / Penumpang", 1, 10, 2)
    w_mayat_patogen = st.slider("Import/Eksport Mayat / Bahan Patogenik", 1, 10, 6)
    w_makanan_import = st.slider("Pemeriksaan Makanan Import", 1, 10, 4)

with st.sidebar.expander("5. Tadbir Urus & Latihan", expanded=False):
    w_latihan = st.slider("Latihan Teknikal PKP / PPKP", 1, 10, 4)
    w_advokasi = st.slider("Advokasi & Pendidikan Kesihatan", 1, 10, 2)
    w_penilaian_risiko = st.slider("Penilaian Risiko Persekitaran & Bencana", 1, 10, 5)

# Supervisory Span-of-Control parameter for PKP
if selected_cadre.startswith("PKP"):
    st.sidebar.markdown("---")
    st.sidebar.subheader("👔 PKP Span of Control")
    span_ratio = st.sidebar.slider("PPKP per PKP Ratio (Span-of-Control)", 10, 35, 20, 
                                   help="Target number of PPKP supervised per 1 PKP.")

# Minimum baseline staffing constraint
st.sidebar.markdown("---")
min_floor = st.sidebar.number_input(f"Minimum Baseline Floor per District ({selected_cadre[:4]})", min_value=1, max_value=5, value=1 if selected_cadre.startswith("PKP") else 5)

# =============================================================================
# 3. WORKLOAD CALCULATION ENGINE
# =============================================================================
def get_c(pattern):
    matches = [c for c in df.columns if pattern.lower() in c.lower()]
    return df[matches[0]] if matches else pd.Series(0, index=df.index)

# Operational RVU (PPKP Burden)
df['Operational_RVU'] = (
    get_c('Kawalan_penyakit_berjangkit') * w_outbreak +
    get_c('Persampelan_penyakit') * w_sampling +
    get_c('Siasatan_kes_dan_kontak') * w_tracing +
    get_c('Bilik_gerakan') * w_bilik_gerakan +
    get_c('Laporan_aktiviti_survelans') * w_survelans +
    get_c('Risikan_Undang') * w_risikan +
    get_c('Notis_kompaun') * w_notis +
    get_c('Kertas_siasatan') * w_kertas_siasatan +
    get_c('Audit_pewasapan') * w_audit_pewasapan +
    get_c('Pemeriksaan_kesihatan_persekitaran_di_premis') * w_premis +
    get_c('Penilaian_risiko_premis_makanan') * w_makanan +
    get_c('Audit_keselamatan_dan_kesihatan') * w_osh_audit +
    get_c('Notifikasi_dan_siasatan_kemalangan') * w_osh_siasatan +
    get_c('Pemantauan_kualiti_air') * w_air +
    get_c('kapal_vessel') * w_kapal +
    get_c('mayat') * w_mayat_patogen +
    get_c('produk_makanan_import') * w_makanan_import +
    get_c('Latihan_teknikal') * w_latihan +
    get_c('Advokasi_dan_pendidikan') * w_advokasi +
    get_c('Penilaian_risiko_kesihatan_persekitaran_penyakit_berjangkit_dan_bencana') * w_penilaian_risiko
)

# Managerial & Governance Burden (PKP Burden)
df['Supervisory_RVU'] = (
    get_c('Kertas_siasatan') * (w_kertas_siasatan * 1.5) +
    get_c('Bilik_gerakan') * (w_bilik_gerakan * 2.0) +
    get_c('Latihan_teknikal') * (w_latihan * 1.5) +
    get_c('Penilaian_risiko_kesihatan_persekitaran_penyakit_berjangkit_dan_bencana') * (w_penilaian_risiko * 1.5) +
    (df['Bil_PPKP_Total'] * 50)
)

# Allocation logic
if selected_cadre.startswith("PPKP"):
    pool = df['Bil_PPKP_Total'].sum()
    df['Current_Staff'] = df['Bil_PPKP_Total']
    workload_metric = df['Operational_RVU']
else:
    pool = df['Bil_PKP_Total'].sum()
    df['Current_Staff'] = df['Bil_PKP_Total']
    workload_metric = (df['Supervisory_RVU'] * 0.6) + ((df['Bil_PPKP_Total'] / span_ratio) * 0.4)

total_workload = workload_metric.sum()
n_districts = len(df)
baseline_reserved = n_districts * min_floor
distributable_pool = max(0, pool - baseline_reserved)

if total_workload > 0 and distributable_pool > 0:
    df['Norm_Allocation'] = (min_floor + (workload_metric / total_workload) * distributable_pool).round(1)
else:
    df['Norm_Allocation'] = float(min_floor)

df['Variance'] = (df['Current_Staff'] - df['Norm_Allocation']).round(1)

# =============================================================================
# 4. DASHBOARD INTERFACE & METRICS
# =============================================================================
col_a, col_b, col_c, col_d = st.columns(4)
col_a.metric("Total National Cadre Pool", int(pool))
col_b.metric("Total Operational Districts", n_districts)

critical_shortage = df.loc[df['Variance'].idxmin()]
col_c.metric("Critical Deficit District", f"{critical_shortage['Daerah']}", f"{critical_shortage['Variance']} posts")

highest_surplus = df.loc[df['Variance'].idxmax()]
col_d.metric("Highest Surplus Capacity", f"{highest_surplus['Daerah']}", f"+{highest_surplus['Variance']} posts")

st.markdown("---")

filter_col1, filter_col2 = st.columns(2)
with filter_col1:
    selected_state = st.multiselect("Filter by State (Negeri):", options=sorted(df['Negeri'].dropna().unique()), default=[])
with filter_col2:
    status_filter = st.radio("Display Scope:", ["All Districts", "Deficit Only (Needs Staff)", "Surplus Only (Potential Donors)"], horizontal=True)

df_filtered = df.copy()
if selected_state:
    df_filtered = df_filtered[df_filtered['Negeri'].isin(selected_state)]

if status_filter == "Deficit Only (Needs Staff)":
    df_filtered = df_filtered[df_filtered['Variance'] < 0]
elif status_filter == "Surplus Only (Potential Donors)":
    df_filtered = df_filtered[df_filtered['Variance'] > 0]

tab_ledger, tab_chart = st.tabs(["📋 District Allocation Norms", "📊 Disparity Visualizer"])

with tab_ledger:
    st.write(f"Showing **{len(df_filtered)}** districts. Negative variance indicates an understaffed district relative to its actual workload.")
    cols_to_show = ['Negeri', 'Daerah', 'Kategori', 'Current_Staff', 'Norm_Allocation', 'Variance']
    
    st.dataframe(
        df_filtered[cols_to_show].sort_values(by='Variance', ascending=True)
        .style.background_gradient(subset=['Variance'], cmap='RdYlGn', vmin=-10, vmax=10)
        .format({'Norm_Allocation': '{:.1f}', 'Variance': '{:+.1f}'}),
        use_container_width=True
    )
    
    csv_data = df_filtered[cols_to_show].to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download Norms Summary (CSV)", data=csv_data, file_name=f"{selected_cadre[:4]}_Norms_Distribution.csv", mime="text/csv")

with tab_chart:
    st.subheader(f"Top 20 Districts with Largest Staff Deficit ({selected_cadre[:4]})")
    deficit_chart = (df.sort_values(by='Variance', ascending=True)
                     .head(20)
                     .set_index('Daerah')[['Current_Staff', 'Norm_Allocation']])
    st.bar_chart(deficit_chart)