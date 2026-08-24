from __future__ import annotations

from pathlib import Path
import re
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# KONFIGURASI
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

FILE_MAHASISWA = DATA_DIR / "v_mahasiswa.xls"
FILE_DIFABEL = DATA_DIR / "difabel.xls"
FILE_YUDISIUM = DATA_DIR / "vs_yudisium_mahasiswa.xls"
FILE_PMB = DATA_DIR / "vs_rekap_pmb.xls"
FILE_AKREDITASI = DATA_DIR / "akreditasiprodi.xls"

TARGET_IKU = {
    "IKU-01-01": 1.0,
    "IKU-01-05": 15.0,
    "IKU-01-07": 2.0,
    "IKU-02-01": 1.0,
    "IKU-04-1": 35.0,
    "IKU-04-3": 52.5,
    "IKU-35-54": 80.0,
}

DAERAH_TERTINGGAL_2025_2029 = {
    "NIAS UTARA", "SUMBA TENGAH", "SUMBA BARAT DAYA", "SABU RAIJUA", 
    "KEPULAUAN SULA", "DONGGALA", "TAMBRAUW", "MAYBRAT", "TELUK WONDAMA", 
    "PEGUNUNGAN ARFAK", "PANIAI", "PUNCAK JAYA", "PUNCAK", "DOGIYAI", 
    "INTAN JAYA", "DEIYAI", "JAYAWIJAYA", "YAHUKIMO", "PEGUNUNGAN BINTANG", 
    "TOLIKARA", "NDUGA", "LANNY JAYA", "MAMBERAMO TENGAH", "YALIMO", 
    "WAROPEN", "MAMBERAMO RAYA", "KEEROM", "BOVEN DIGOEL", "MAPPI", "ASMAT",
}

C = {
    "primary": "#00695C",
    "secondary": "#00897B",
    "light": "#E0F2F1",
    "yellow": "#F9A825",
    "orange": "#F97316",
    "blue": "#1E88E5",
    "purple": "#7B1FA2",
    "gray": "#64748B",
    "bg": "#F4F7FA",
    "text": "#172033",
    "border": "#E2E8F0",
}

st.set_page_config(
    page_title="Dashboard E-Semesta UIN Jakarta",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# TEMA & OVERRIDE CSS
# ============================================================

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Poppins:wght@500;600;700&display=swap');

    html, body, [class*="css"] {{ font-family: "Inter", sans-serif; }}
    .stApp {{ background: {C["bg"]}; color: {C["text"]}; }}
    .block-container {{ max-width: 1720px; padding-top: 1rem; padding-bottom: 3rem; }}

    /* ======= SIDEBAR & FILTER PANEL ======= */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #063F39 0%, #075E54 42%, #0A7A6B 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.08);
    }}
    [data-testid="stSidebar"] * {{ color: white; }}
    [data-testid="stSidebar"] hr {{ border-color: rgba(255, 255, 255, 0.14); margin: 0.8rem 0; }}
    [data-testid="stSidebar"] label {{ color: rgba(255, 255, 255, 0.94) !important; font-size: 11px !important; font-weight: 650 !important; }}
    
    .filter-heading {{
        margin: 2px 0 8px; color: rgba(255, 255, 255, 0.68); font-size: 10px; 
        font-weight: 700; letter-spacing: 0.10em; text-transform: uppercase;
    }}

    /* KOTAK INPUT FILTER & TEKS MURNI HITAM AGAR TERBACA */
    [data-testid="stSidebar"] [data-baseweb="select"] {{
        background-color: white !important;
        border-radius: 8px;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"] span,
    [data-testid="stSidebar"] [data-baseweb="select"] input,
    [data-testid="stSidebar"] [data-baseweb="select"] div {{
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 12px !important;
    }}
    [data-testid="stSidebar"] input::placeholder {{
        color: #64748B !important;
        -webkit-text-fill-color: #64748B !important;
    }}
    
    /* MULTISELECT TAGS WARNA GELAP */
    [data-testid="stSidebar"] [data-baseweb="tag"] {{
        background-color: #E2E8F0 !important;
        border-radius: 6px !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="tag"] span,
    [data-testid="stSidebar"] [data-baseweb="tag"] svg {{
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
    }}

    /* TOMBOL RESET FILTER */
    [data-testid="stSidebar"] .stButton > button {{
        width: 100%; min-height: 40px; border-radius: 8px; font-size: 11px; font-weight: 700;
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }}
    [data-testid="stSidebar"] .stButton > button:hover {{
        background-color: #E2E8F0 !important;
        border-color: #94A3B8 !important;
    }}
    
    /* ======= EXPANDER DI SIDEBAR ======= */
    [data-testid="stSidebar"] [data-testid="stExpander"] {{
        background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 10px;
    }}
    [data-testid="stSidebar"] [data-testid="stExpander"] summary p {{
        font-weight: 600; color: white !important; -webkit-text-fill-color: white !important;
    }}

    /* ======= TAB STYLING UTAMA ======= */
    [data-testid="stTabs"] button[role="tab"] {{
        font-family: "Poppins", sans-serif;
        font-size: 15px;
        font-weight: 600;
        color: #64748B;
    }}
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
        color: #00695C;
    }}

    /* ======= KARTU INFO ======= */
    .hero {{ padding: 25px 29px; margin-bottom: 18px; color: white; border-radius: 20px; background: linear-gradient(120deg, {C["primary"]}, {C["secondary"]}); box-shadow: 0 12px 35px rgba(0, 105, 92, 0.18); }}
    .hero h1 {{ margin: 0; font-family: "Poppins", sans-serif; font-size: 29px; }}
    .hero p {{ margin: 6px 0 0; font-size: 13px; opacity: 0.88; }}
    .section-title {{ margin-top: 12px; margin-bottom: 2px; font-family: "Poppins", sans-serif; font-size: 18px; font-weight: 700; }}
    .section-subtitle {{ margin-bottom: 12px; color: {C["gray"]}; font-size: 12px; }}
    
    .kpi-card, .iku-card, .empty-card {{ padding: 17px 18px; background: white; border: 1px solid {C["border"]}; border-radius: 17px; box-shadow: 0 5px 18px rgba(15, 23, 42, 0.045); }}
    .kpi-card {{ min-height: 143px; }}
    .iku-card {{ min-height: 235px; }}
    .empty-card {{ padding: 38px; text-align: center; }}
    .kpi-icon {{ display: flex; align-items: center; justify-content: center; width: 38px; height: 38px; border-radius: 12px; background: {C["light"]}; font-size: 19px; }}
    .kpi-value {{ margin-top: 10px; font-family: "Poppins", sans-serif; font-size: 26px; font-weight: 700; }}
    .kpi-label {{ color: {C["gray"]}; font-size: 12px; font-weight: 600; }}
    .kpi-note {{ margin-top: 3px; color: {C["gray"]}; font-size: 10px; }}
    
    .iku-code {{ display: inline-block; padding: 4px 8px; color: {C["primary"]}; background: {C["light"]}; border-radius: 8px; font-size: 11px; font-weight: 700; }}
    .iku-title {{ min-height: 51px; margin-top: 9px; font-size: 12px; font-weight: 700; }}
    .iku-value {{ margin-top: 7px; font-family: "Poppins", sans-serif; font-size: 25px; font-weight: 700; }}
    .iku-detail {{ margin-top: 7px; color: {C["gray"]}; font-size: 10px; line-height: 1.5; }}
    
    .badge-success {{ color: #166534; background: #DCFCE7; display: inline-block; padding: 3px 8px; border-radius: 9px; font-size: 10px; font-weight: 700;}}
    .badge-warning {{ color: #92400E; background: #FEF3C7; display: inline-block; padding: 3px 8px; border-radius: 9px; font-size: 10px; font-weight: 700;}}
    .badge-danger {{ color: #991B1B; background: #FEE2E2; display: inline-block; padding: 3px 8px; border-radius: 9px; font-size: 10px; font-weight: 700;}}
    
    div[data-testid="stPlotlyChart"] {{ padding: 8px; background: white; border: 1px solid {C["border"]}; border-radius: 17px; }}
    div[data-testid="stDataFrame"] {{ overflow: hidden; border: 1px solid {C["border"]}; border-radius: 15px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FUNGSI BANTU
# ============================================================

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = result.columns.astype(str).str.strip().str.lower().str.replace(r"\s+", "_", regex=True)
    return result

def format_number(value: Any) -> str:
    if value is None or pd.isna(value): return "–"
    return f"{float(value):,.0f}".replace(",", ".")

def format_percent(value: float | None) -> str:
    if value is None or pd.isna(value): return "Data belum tersedia"
    return f"{value:.2f}%".replace(".", ",")

def percentage(numerator: float, denominator: float) -> float | None:
    if denominator is None or denominator == 0 or pd.isna(denominator): return None
    return float(numerator) / float(denominator) * 100

def growth(current: float, previous: float) -> float | None:
    if previous is None or previous == 0 or pd.isna(previous): return None
    return (float(current) - float(previous)) / float(previous) * 100

def normalize_status(value: Any) -> str:
    return str(value).strip().lower()

def status_category(value: Any) -> str:
    text = normalize_status(value)
    if text in {"aktif", "active"}: return "Aktif"
    if "lulus" in text or "alumni" in text: return "Lulus"
    if "cuti" in text: return "Cuti"
    if "tidak aktif" in text or "non aktif" in text or "nonaktif" in text or text in {"na", "n/a"}: return "Tidak Aktif"
    if "drop out" in text or "dropout" in text or "putus studi" in text or text == "do" or text.startswith("do "): return "DO/Putus Studi"
    if "mengundurkan diri" in text or "undur diri" in text or "resign" in text: return "Mengundurkan Diri"
    if "meninggal" in text or "wafat" in text: return "Meninggal"
    if "skors" in text or "suspend" in text: return "Skorsing"
    if "transfer" in text or "pindah" in text: return "Pindah/Transfer"
    if text in {"", "nan", "none", "tidak diketahui"}: return "Tidak Diketahui"
    return str(value).strip().title()

def normalize_gender(value: Any) -> str:
    text = str(value).strip().lower()
    if text in {"l", "laki-laki", "laki laki", "pria"}: return "Laki-laki"
    if text in {"p", "perempuan", "wanita"}: return "Perempuan"
    return "Tidak diketahui"

def normalize_boolean(value: Any) -> bool:
    if pd.isna(value): return False
    return str(value).strip().lower() in {"t", "true", "ya", "y", "1", "tepat waktu"}

def normalize_region(value: Any) -> str:
    if pd.isna(value): return ""
    text = str(value).upper().strip()
    text = re.sub(r"^(KABUPATEN|KAB\.?|KOTA ADMINISTRASI|KOTA)\s+", "", text)
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def normalize_level(value: Any) -> str:
    text = str(value).strip().upper().replace("STRATA ", "S").replace("(S-", "S").replace(")", "").replace("-", "")
    aliases = {"SARJANA": "S1", "MAGISTER": "S2", "DOKTOR": "S3"}
    return aliases.get(text, text)

def hero(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="hero"><h1>{title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)

def section(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="section-title">{title}</div><div class="section-subtitle">{subtitle}</div>', unsafe_allow_html=True)

def subheading(text: str) -> None:
    st.markdown(
        f'''
        <div style="background-color: #D1FAE5; color: #065F46; padding: 12px 18px; 
        border-radius: 8px; font-family: 'Poppins', sans-serif; font-weight: 600; 
        font-size: 16px; margin-top: 35px; margin-bottom: 20px; 
        border-left: 6px solid #059669; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            {text}
        </div>
        ''',
        unsafe_allow_html=True
    )

def kpi(icon: str, label: str, value: str, note: str = "") -> None:
    st.markdown(f'<div class="kpi-card"><div class="kpi-icon">{icon}</div><div class="kpi-value">{value}</div><div class="kpi-label">{label}</div><div class="kpi-note">{note}</div></div>', unsafe_allow_html=True)

def iku_card(code: str, title: str, value: float | None, formula: str, note: str, status: str = "Tersedia") -> None:
    target = TARGET_IKU.get(code)
    if status == "Tersedia":
        if value is not None and (target is None or value >= target): badge, label = "badge-success", "Target tercapai"
        else: badge, label = "badge-warning", "Belum mencapai target"
    elif status in {"Sebagian", "Proxy"}: badge, label = "badge-warning", status
    else: badge, label = "badge-danger", "Data belum tersedia"

    st.markdown(f'<div class="iku-card"><span class="iku-code">{code}</span><div class="iku-title">{title}</div><div class="iku-value">{format_percent(value)}</div><span class="{badge}">{label}</span><div class="iku-detail"><b>Target:</b> {format_percent(target)}<br><b>Rumus:</b> {formula}<br><b>Catatan:</b> {note}</div></div>', unsafe_allow_html=True)

def empty_dashboard(title: str, year: int, icon: str) -> None:
    st.markdown(f'<div class="empty-card"><div style="font-size:42px">{icon}</div><h3>Konten sedang disiapkan</h3><p>Halaman ini sudah memiliki filter waktu sinkron. Data dan KPI akan dimuat saat tersedia.</p></div>', unsafe_allow_html=True)


# ============================================================
# MEMBACA DATA
# ============================================================

@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    required_files = [FILE_MAHASISWA, FILE_DIFABEL, FILE_YUDISIUM, FILE_PMB, FILE_AKREDITASI]
    for file_path in required_files:
        if not file_path.exists(): raise FileNotFoundError(f"File tidak ditemukan: {file_path}")

    mahasiswa = normalize_columns(pd.read_excel(FILE_MAHASISWA, engine="xlrd"))
    difabel = normalize_columns(pd.read_excel(FILE_DIFABEL, engine="xlrd"))
    yudisium = normalize_columns(pd.read_excel(FILE_YUDISIUM, engine="xlrd"))
    pmb = normalize_columns(pd.read_excel(FILE_PMB, engine="xlrd"))
    akreditasi = normalize_columns(pd.read_excel(FILE_AKREDITASI, engine="xlrd"))

    for frame in [mahasiswa, difabel, yudisium]:
        if "nim" in frame.columns: frame["nim"] = frame["nim"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

    mahasiswa["tahun_angkatan"] = pd.to_numeric(mahasiswa["tahun_angkatan"], errors="coerce").astype("Int64")
    mahasiswa["status_normal"] = mahasiswa["status"].fillna("").apply(normalize_status)
    mahasiswa["status_kategori"] = mahasiswa["status"].fillna("").apply(status_category)
    mahasiswa["gender_normal"] = mahasiswa["kelamin"].fillna("").apply(normalize_gender)
    mahasiswa["jenjang_normal"] = mahasiswa["jenjang"].fillna("").apply(normalize_level)
    mahasiswa["kota_normal"] = mahasiswa["kota"].apply(normalize_region)
    mahasiswa["asal_daerah_tertinggal"] = mahasiswa["kota_normal"].isin(DAERAH_TERTINGGAL_2025_2029)
    
    # Inisialisasi kolom jenis sekolah jika tidak ada di sumber aslinya
    if "jenis_sekolah" not in mahasiswa.columns:
        mahasiswa["jenis_sekolah"] = "Tidak diketahui"

    for column in ["fakultas", "jurusan", "jenis_seleksi", "propinsi", "kota", "jenis_sekolah"]:
        if column in mahasiswa.columns:
            mahasiswa[column] = mahasiswa[column].fillna("Tidak diketahui").astype(str).str.strip()

    difabel["tahun_angkatan"] = pd.to_numeric(difabel["tahun_angkatan"], errors="coerce").astype("Int64")
    difabel["status_normal"] = difabel["status"].fillna("").apply(normalize_status)
    difabel["gender_normal"] = difabel["kelamin"].fillna("").apply(normalize_gender)
    difabel = difabel.drop_duplicates(subset=["nim"], keep="first")

    yudisium["tanggal_yudisium_mahasiswa"] = pd.to_datetime(yudisium["tanggal_yudisium_mahasiswa"], errors="coerce")
    yudisium["tanggal_lulus"] = yudisium["tanggal_yudisium_mahasiswa"]
    yudisium["tahun_lulus"] = yudisium["tanggal_lulus"].dt.year.astype("Int64")
    yudisium["ipk"] = pd.to_numeric(yudisium["ipk"], errors="coerce")
    yudisium["semester"] = pd.to_numeric(yudisium["semester"], errors="coerce")
    yudisium["tepat_waktu_bool"] = yudisium["tepat_waktu"].apply(normalize_boolean)
    yudisium["tepat_waktu_ipk_325"] = yudisium["tepat_waktu_bool"] & yudisium["ipk"].ge(3.25)
    yudisium["jenjang_normal"] = yudisium["jenjang"].fillna("").apply(normalize_level)
    yudisium = yudisium.sort_values("tanggal_lulus").drop_duplicates(subset=["nim"], keep="last")

    pmb["tahun"] = pd.to_numeric(pmb["tahun"], errors="coerce").astype("Int64")
    for column in ["peminat", "lulus_seleksi", "daftar_ulang"]: pmb[column] = pd.to_numeric(pmb[column], errors="coerce").fillna(0)
    pmb["jenjang_normal"] = pmb["jenjang"].fillna("").apply(normalize_level)
    for column in ["fakultas", "jurusan", "jenis_seleksi"]: pmb[column] = pmb[column].fillna("Tidak diketahui").astype(str).str.strip()

    for column in ["tanggal_sk_akreditasi", "masa_mulai_sk_akreditasi", "masa_sk_akhir_akreditasi"]: akreditasi[column] = pd.to_datetime(akreditasi[column], errors="coerce")
    for column in ["jenjang", "fakultas", "prodi", "no_sk_akreditasi", "peringkat"]: akreditasi[column] = akreditasi[column].fillna("Tidak diketahui").astype(str).str.strip()
    akreditasi["jenjang_normal"] = akreditasi["jenjang"].apply(normalize_level)

    hari_ini = pd.Timestamp.today().normalize()
    akreditasi["status_akreditasi"] = "Aktif"
    akreditasi.loc[akreditasi["masa_sk_akhir_akreditasi"].isna(), "status_akreditasi"] = "Tanggal akhir belum tersedia"
    akreditasi.loc[akreditasi["masa_sk_akhir_akreditasi"].notna() & (akreditasi["masa_sk_akhir_akreditasi"] < hari_ini), "status_akreditasi"] = "Kedaluwarsa"
    akreditasi["sisa_hari"] = (akreditasi["masa_sk_akhir_akreditasi"] - hari_ini).dt.days
    akreditasi["masa_berakhir"] = "Normal"
    akreditasi.loc[akreditasi["sisa_hari"].between(0, 365, inclusive="both"), "masa_berakhir"] = "Berakhir ≤ 1 tahun"
    akreditasi.loc[akreditasi["sisa_hari"] < 0, "masa_berakhir"] = "Sudah berakhir"
    akreditasi.loc[akreditasi["sisa_hari"].isna(), "masa_berakhir"] = "Belum diketahui"

    return mahasiswa, difabel, yudisium, pmb, akreditasi

try:
    df, df_difabel, df_yudisium, df_pmb, df_akreditasi = load_data()
except Exception as error:
    st.error(f"Gagal membaca data: {error}")
    st.stop()


# ============================================================
# PERSIAPAN OPSI FILTER
# ============================================================

MENU_TABS = [
    "Eksekutif", "Akademik", "SDM", "Keuangan", "Aset", 
    "Riset", "Alumni", "Prestasi", "Paperless (PLO)", "IKU"
]

landing_years = sorted(set(df["tahun_angkatan"].dropna().astype(int).tolist()) | set(df_yudisium["tahun_lulus"].dropna().astype(int).tolist()) | set(df_pmb["tahun"].dropna().astype(int).tolist()), reverse=True)
academic_year_options = ["2026/2027 Genap", "2026/2027 Ganjil", "2025/2026 Genap", "2025/2026 Ganjil", "2024/2025 Genap", "2024/2025 Ganjil"]
cohort_options = sorted(df["tahun_angkatan"].dropna().astype(int).unique().tolist(), reverse=True)
level_options = sorted(df["jenjang_normal"].dropna().unique().tolist())
faculty_options = sorted(df["fakultas"].dropna().unique().tolist())
path_options = sorted(set(df["jenis_seleksi"].dropna().astype(str).tolist()) | set(df_pmb["jenis_seleksi"].dropna().astype(str).tolist()))
province_options = sorted(df["propinsi"].dropna().unique().tolist())
status_options = sorted(df["status_kategori"].dropna().astype(str).unique().tolist())

# Ambil opsi jenis_sekolah dari dataframe, jika tersedia
school_options = sorted(df["jenis_sekolah"].dropna().unique().tolist()) if "jenis_sekolah" in df.columns else []

available_iku_years = sorted(set(df["tahun_angkatan"].dropna().astype(int).tolist()) | set(df_yudisium["tahun_lulus"].dropna().astype(int).tolist()), reverse=True)


# ============================================================
# HEADER & SIDEBAR FILTER
# ============================================================

st.markdown(
    """
    <div style="padding-top: 5px; padding-bottom: 20px;">
        <h2 style="margin: 0; color: #00695C; font-family: 'Poppins', sans-serif;">🎓 E-Semesta UIN Jakarta</h2>
        <p style="margin: 0; color: #64748B; font-size: 14px;">Dashboard Eksekutif Terintegrasi</p>
    </div>
    """,
    unsafe_allow_html=True
)

with st.sidebar:
    st.markdown(
        """
        <div style="padding-bottom: 0.5rem;">
            <div style="font-family: 'Poppins', sans-serif; font-size: 20px; font-weight: bold; margin-bottom: 2px;">⚙️ Filter Panel</div>
            <div style="font-size: 12px; color: rgba(255,255,255,0.7);">Sesuaikan parameter visualisasi</div>
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.markdown("<hr style='margin-top: 0;'>", unsafe_allow_html=True)

    st.markdown('<div class="filter-heading">FILTER GLOBAL (TAHUN)</div>', unsafe_allow_html=True)
    global_year = st.selectbox("Pilih Tahun Referensi", landing_years, index=0, key="global_year")
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="filter-heading">FILTER SPESIFIK AKADEMIK</div>', unsafe_allow_html=True)
    with st.expander("Buka Filter Akademik", expanded=True):
        academic_period = st.selectbox("Tahun Akademik", academic_year_options, index=3, key="ac_period")
        selected_cohorts = st.multiselect("Tahun Angkatan", cohort_options, key="ac_cohorts")
        selected_levels = st.multiselect("Jenjang", level_options, key="ac_levels")
        selected_faculties = st.multiselect("Fakultas", faculty_options, key="ac_faculties")
        
        program_source = df[df["fakultas"].isin(selected_faculties)] if selected_faculties else df
        program_options = sorted(program_source["jurusan"].dropna().unique().tolist())
        selected_programs = st.multiselect("Program Studi", program_options, key="ac_programs")
        
        selected_paths = st.multiselect("Jalur Masuk", path_options, key="ac_paths")
        selected_provinces = st.multiselect("Propinsi Asal", province_options, key="ac_provinces")
        selected_schools = st.multiselect("Jenis Sekolah", school_options, key="ac_schools")
        selected_statuses = st.multiselect("Status Mahasiswa", status_options, key="ac_statuses")

        if st.button("↻ Reset Filter Akademik", key="reset_ac_filter"):
            for k in ["ac_period", "ac_cohorts", "ac_levels", "ac_faculties", "ac_programs", "ac_paths", "ac_provinces", "ac_schools", "ac_statuses"]:
                st.session_state.pop(k, None)
            st.rerun()


# ============================================================
# TAB NAVIGASI UTAMA
# ============================================================

main_tabs = st.tabs(MENU_TABS)

# --- 1. DASHBOARD EKSEKUTIF ---
with main_tabs[0]:
    hero("Dashboard Eksekutif UIN Jakarta", "Ringkasan strategis mahasiswa, PMB, kelulusan, dan indikator institusi.")

    landing_students = df[df["tahun_angkatan"].eq(global_year)]
    landing_active = df[df["status_normal"].eq("aktif")]
    landing_graduates = df_yudisium[df_yudisium["tahun_lulus"].eq(global_year)]

    pmb_years = sorted(df_pmb["tahun"].dropna().astype(int).unique().tolist())
    landing_pmb_year = global_year if global_year in pmb_years else max(pmb_years)
    landing_pmb = df_pmb[df_pmb["tahun"].eq(landing_pmb_year)]

    total_peminat = float(landing_pmb["peminat"].sum())
    total_lulus = float(landing_pmb["lulus_seleksi"].sum())
    total_daftar = float(landing_pmb["daftar_ulang"].sum())
    landing_yield = percentage(total_daftar, total_lulus)

    if landing_pmb_year != global_year:
        st.warning(f"Data PMB tahun {global_year} belum tersedia. Ringkasan PMB memakai tahun terbaru {landing_pmb_year}.")

    executive_cols = st.columns(6)
    executive_values = [
        ("👥", "Mahasiswa aktif", format_number(len(landing_active)), "Data aktual semua angkatan"),
        ("🆕", f"Angkatan {global_year}", format_number(len(landing_students)), "Data aktual"),
        ("📝", "Peminat PMB", format_number(total_peminat), f"Data aktual PMB {landing_pmb_year}"),
        ("🎓", "Lulusan", format_number(landing_graduates["nim"].nunique()), f"Data aktual yudisium {global_year}"),
        ("🎯", "Yield rate", format_percent(landing_yield), f"Data aktual PMB {landing_pmb_year}"),
        ("🏫", "Daya tampung", "12.500", "Sampel sementara"),
    ]
    for column, item in zip(executive_cols, executive_values):
        with column: kpi(*item)

    chart_a, chart_b = st.columns([1.4, 1])
    with chart_a:
        faculty_summary = landing_active.groupby("fakultas").size().reset_index(name="Mahasiswa").sort_values("Mahasiswa")
        figure = px.bar(faculty_summary, x="Mahasiswa", y="fakultas", orientation="h", text="Mahasiswa", title="Mahasiswa Aktif per Fakultas", color_discrete_sequence=[C["primary"]])
        figure.update_layout(height=470, plot_bgcolor="white", paper_bgcolor="white", showlegend=False)
        st.plotly_chart(figure, use_container_width=True)

    with chart_b:
        level_summary = landing_active["jenjang_normal"].value_counts().rename_axis("Jenjang").reset_index(name="Mahasiswa")
        figure = px.pie(level_summary, names="Jenjang", values="Mahasiswa", hole=0.58, title="Komposisi Jenjang Mahasiswa Aktif", color_discrete_sequence=[C["primary"], C["yellow"], C["blue"], C["orange"], C["purple"]])
        figure.update_layout(height=470)
        st.plotly_chart(figure, use_container_width=True)

    sample_cols = st.columns(4)
    sample_values = [
        ("📚", "Rasio dosen-mahasiswa", "1 : 27", "Sampel sementara"),
        ("🔬", "Publikasi terindeks", "1.248", "Sampel sementara"),
        ("💰", "Realisasi anggaran", "78,4%", "Sampel sementara"),
        ("🤝", "Mitra aktif", "186", "Sampel sementara"),
    ]
    for column, item in zip(sample_cols, sample_values):
        with column: kpi(*item)


# --- 2. DASHBOARD AKADEMIK ---
with main_tabs[1]:
    hero("Dashboard Akademik", "Profil akademik terintegrasi. Seluruh seksi di bawah ini merespons filter di panel kiri secara serentak.")

    academic_df = df.copy()
    if selected_cohorts: academic_df = academic_df[academic_df["tahun_angkatan"].isin(selected_cohorts)]
    if selected_levels: academic_df = academic_df[academic_df["jenjang_normal"].isin(selected_levels)]
    if selected_faculties: academic_df = academic_df[academic_df["fakultas"].isin(selected_faculties)]
    if selected_programs: academic_df = academic_df[academic_df["jurusan"].isin(selected_programs)]
    if selected_paths: academic_df = academic_df[academic_df["jenis_seleksi"].isin(selected_paths)]
    if selected_provinces: academic_df = academic_df[academic_df["propinsi"].isin(selected_provinces)]
    if selected_schools: academic_df = academic_df[academic_df["jenis_sekolah"].isin(selected_schools)]
    if selected_statuses: academic_df = academic_df[academic_df["status_kategori"].isin(selected_statuses)]

    academic_active = academic_df[academic_df["status_normal"].eq("aktif")]
    academic_start_year = int(academic_period.split("/")[0])


    # =========================================================
    # A. DATA PMB
    # =========================================================
    subheading("A. Data Penerimaan Mahasiswa Baru (PMB)")

    pmb_years = sorted(df_pmb["tahun"].dropna().astype(int).unique().tolist())
    academic_pmb_year = academic_start_year if academic_start_year in pmb_years else max(pmb_years)
    academic_pmb = df_pmb[df_pmb["tahun"].eq(academic_pmb_year)].copy()

    if selected_levels: academic_pmb = academic_pmb[academic_pmb["jenjang_normal"].isin(selected_levels)]
    if selected_faculties: academic_pmb = academic_pmb[academic_pmb["fakultas"].isin(selected_faculties)]
    if selected_programs: academic_pmb = academic_pmb[academic_pmb["jurusan"].isin(selected_programs)]
    if selected_paths: academic_pmb = academic_pmb[academic_pmb["jenis_seleksi"].isin(selected_paths)]

    if academic_pmb_year != academic_start_year:
        st.warning(f"Data PMB tahun {academic_start_year} belum tersedia. Data ini memakai tahun terbaru {academic_pmb_year}.")

    total_peminat = float(academic_pmb["peminat"].sum())
    total_lulus = float(academic_pmb["lulus_seleksi"].sum())
    total_daftar = float(academic_pmb["daftar_ulang"].sum())

    cols = st.columns(4)
    pmb_values = [
        ("📝", "Peminat", format_number(total_peminat)),
        ("✅", "Lulus seleksi", format_number(total_lulus)),
        ("📋", "Daftar ulang", format_number(total_daftar)),
        ("🎯", "Yield rate", format_percent(percentage(total_daftar, total_lulus))),
    ]
    for column, item in zip(cols, pmb_values):
        with column: kpi(item[0], item[1], item[2], f"PMB {academic_pmb_year}")

    funnel = go.Figure(go.Funnel(y=["Peminat", "Lulus Seleksi", "Daftar Ulang"], x=[total_peminat, total_lulus, total_daftar], textinfo="value+percent initial"))
    funnel.update_layout(title=f"Funnel PMB {academic_pmb_year}")
    st.plotly_chart(funnel, use_container_width=True)

    pmb_summary_export = academic_pmb.groupby(["fakultas", "jurusan", "jenjang_normal", "jenis_seleksi"], dropna=False).agg(Peminat=("peminat", "sum"), Lulus_Seleksi=("lulus_seleksi", "sum"), Daftar_Ulang=("daftar_ulang", "sum")).reset_index()
    pmb_daftar_ulang_numeric = pd.to_numeric(pmb_summary_export["Daftar_Ulang"], errors="coerce")
    pmb_lulus_numeric = pd.to_numeric(pmb_summary_export["Lulus_Seleksi"], errors="coerce")
    pmb_summary_export["Yield_Rate"] = pmb_daftar_ulang_numeric.div(pmb_lulus_numeric.where(pmb_lulus_numeric.ne(0))).mul(100).round(2)

    st.markdown("**Statistik PMB Berdasarkan Filter** (Tampilan dibatasi 10 baris pertama)")
    st.dataframe(pmb_summary_export.head(10), use_container_width=True, hide_index=True)


    # =========================================================
    # B. DATA MAHASISWA
    # =========================================================
    subheading("B. Data Mahasiswa")
    
    status_counts = academic_df["status_kategori"].fillna("Tidak Diketahui").value_counts().to_dict()
    preferred_status_order = ["Aktif", "Lulus", "Cuti", "Tidak Aktif", "DO/Putus Studi", "Mengundurkan Diri", "Meninggal", "Skorsing", "Pindah/Transfer", "Tidak Diketahui"]
    status_icon_map = {"Aktif": "✅", "Lulus": "🎓", "Cuti": "⏸️", "Tidak Aktif": "🚫", "DO/Putus Studi": "⚠️", "Mengundurkan Diri": "↩️", "Meninggal": "🕊️", "Skorsing": "⛔", "Pindah/Transfer": "🔄", "Tidak Diketahui": "❓"}

    visible_statuses = [s for s in preferred_status_order if status_counts.get(s, 0) > 0]
    other_statuses = sorted(s for s in status_counts if s not in preferred_status_order)
    visible_statuses.extend(other_statuses)

    student_cards = [("👥", "Total mahasiswa", format_number(len(academic_df)), "Sesuai filter akademik")]
    for status_name in visible_statuses:
        student_cards.append((status_icon_map.get(status_name, "📌"), status_name, format_number(status_counts.get(status_name, 0)), "Status mahasiswa"))

    for start_index in range(0, len(student_cards), 5):
        card_columns = st.columns(min(5, len(student_cards) - start_index))
        for column, item in zip(card_columns, student_cards[start_index:start_index + 5]):
            with column: kpi(*item)

    st.markdown("<br>", unsafe_allow_html=True)
    mhs_col1, mhs_col2 = st.columns([1.45, 0.75])
    
    with mhs_col1:
        prodi_base = academic_df.groupby(["fakultas", "jurusan", "jenjang_normal"], dropna=False).agg(Total_Mahasiswa=("nim", "nunique")).reset_index()
        chart_df = prodi_base.nlargest(20, "Total_Mahasiswa").sort_values("Total_Mahasiswa")
        figure = px.bar(chart_df, x="Total_Mahasiswa", y="jurusan", orientation="h", text="Total_Mahasiswa", title="20 Program Studi dengan Mahasiswa Terbanyak", color_discrete_sequence=[C["primary"]], hover_data={"fakultas": True, "jenjang_normal": True, "Total_Mahasiswa": True})
        figure.update_layout(height=620, plot_bgcolor="white", paper_bgcolor="white", showlegend=False, xaxis_title="Total Mahasiswa", yaxis_title=None, margin=dict(l=10, r=35, t=55, b=25))
        figure.update_traces(textposition="outside")
        st.plotly_chart(figure, use_container_width=True)

    with mhs_col2:
        gender_df = academic_df["gender_normal"].value_counts().rename_axis("Jenis Kelamin").reset_index(name="Mahasiswa")
        figure = px.pie(gender_df, names="Jenis Kelamin", values="Mahasiswa", hole=0.58, title="Komposisi Jenis Kelamin", color_discrete_sequence=[C["primary"], C["orange"], C["gray"]])
        figure.update_layout(height=500, paper_bgcolor="white", legend=dict(orientation="h", y=-0.08))
        figure.update_traces(textinfo="label+value+percent")
        st.plotly_chart(figure, use_container_width=True)

    province_coordinates = {
        "ACEH": (4.6951, 96.7494), "SUMATERA UTARA": (2.1154, 99.5451), "SUMATERA BARAT": (-0.7399, 100.8000), "RIAU": (0.2933, 101.7068),
        "KEPULAUAN RIAU": (3.9457, 108.1429), "JAMBI": (-1.4852, 102.4381), "SUMATERA SELATAN": (-3.3194, 103.9144), "KEPULAUAN BANGKA BELITUNG": (-2.7411, 106.4406),
        "BENGKULU": (-3.5778, 102.3464), "LAMPUNG": (-4.5586, 105.4068), "DKI JAKARTA": (-6.2088, 106.8456), "JAWA BARAT": (-6.9175, 107.6191),
        "BANTEN": (-6.4058, 106.0640), "JAWA TENGAH": (-7.1510, 110.1403), "DI YOGYAKARTA": (-7.8754, 110.4262), "JAWA TIMUR": (-7.5361, 112.2384),
        "BALI": (-8.3405, 115.0920), "NUSA TENGGARA BARAT": (-8.6529, 117.3616), "NUSA TENGGARA TIMUR": (-8.6574, 121.0794), "KALIMANTAN BARAT": (-0.2788, 111.4753),
        "KALIMANTAN TENGAH": (-1.6815, 113.3824), "KALIMANTAN SELATAN": (-3.0926, 115.2838), "KALIMANTAN TIMUR": (0.5387, 116.4194), "KALIMANTAN UTARA": (3.0731, 116.0414),
        "SULAWESI UTARA": (0.6247, 123.9750), "GORONTALO": (0.6999, 122.4467), "SULAWESI TENGAH": (-1.4300, 121.4456), "SULAWESI BARAT": (-2.8441, 119.2321),
        "SULAWESI SELATAN": (-3.6688, 119.9741), "SULAWESI TENGGARA": (-4.1449, 122.1746), "MALUKU": (-3.2385, 130.1453), "MALUKU UTARA": (1.5700, 127.8088),
        "PAPUA BARAT": (-1.3361, 133.1747), "PAPUA BARAT DAYA": (-1.1307, 131.2416), "PAPUA": (-4.2699, 138.0804), "PAPUA TENGAH": (-3.7048, 136.6798),
        "PAPUA PEGUNUNGAN": (-4.2699, 138.6667), "PAPUA SELATAN": (-7.1327, 139.2310),
    }
    province_aliases = {"NANGGROE ACEH DARUSSALAM": "ACEH", "NAD": "ACEH", "SUMUT": "SUMATERA UTARA", "SUMBAR": "SUMATERA BARAT", "SUMSEL": "SUMATERA SELATAN", "KEPRI": "KEPULAUAN RIAU", "BANGKA BELITUNG": "KEPULAUAN BANGKA BELITUNG", "DKI": "DKI JAKARTA", "JAKARTA": "DKI JAKARTA", "JABAR": "JAWA BARAT", "JATENG": "JAWA TENGAH", "DAERAH ISTIMEWA YOGYAKARTA": "DI YOGYAKARTA", "D.I. YOGYAKARTA": "DI YOGYAKARTA", "DIY": "DI YOGYAKARTA", "JATIM": "JAWA TIMUR", "NTB": "NUSA TENGGARA BARAT", "NTT": "NUSA TENGGARA TIMUR", "KALBAR": "KALIMANTAN BARAT", "KALTENG": "KALIMANTAN TENGAH", "KALSEL": "KALIMANTAN SELATAN", "KALTIM": "KALIMANTAN TIMUR", "KALTARA": "KALIMANTAN UTARA", "SULUT": "SULAWESI UTARA", "SULTENG": "SULAWESI TENGAH", "SULBAR": "SULAWESI BARAT", "SULSEL": "SULAWESI SELATAN", "SULTRA": "SULAWESI TENGGARA", "IRIAN JAYA": "PAPUA"}

    province_map_df = academic_df["propinsi"].fillna("Tidak diketahui").astype(str).str.strip().str.upper().replace(province_aliases).value_counts().rename_axis("Propinsi").reset_index(name="Mahasiswa")
    province_map_df["Latitude"] = province_map_df["Propinsi"].map(lambda v: province_coordinates.get(v, (None, None))[0])
    province_map_df["Longitude"] = province_map_df["Propinsi"].map(lambda v: province_coordinates.get(v, (None, None))[1])
    mapped_provinces = province_map_df.dropna(subset=["Latitude", "Longitude"]).copy()

    if not mapped_provinces.empty:
        figure = px.scatter_geo(mapped_provinces, lat="Latitude", lon="Longitude", size="Mahasiswa", color="Mahasiswa", hover_name="Propinsi", hover_data={"Mahasiswa": ":,", "Latitude": False, "Longitude": False}, size_max=42, title="Peta Sebaran Asal Propinsi Mahasiswa", color_continuous_scale="Teal")
        figure.update_geos(projection_type="mercator", showland=True, landcolor="#EEF4F2", showocean=True, oceancolor="#EAF4FB", showcountries=True, countrycolor="#CBD5E1", coastlinecolor="#94A3B8", lataxis_range=[-12, 7], lonaxis_range=[94, 142])
        figure.update_layout(height=500, paper_bgcolor="white", margin=dict(l=5, r=5, t=55, b=5), coloraxis_colorbar_title="Mahasiswa")
        st.plotly_chart(figure, use_container_width=True)

    st.markdown("**Detail Data Mahasiswa** (Tampilan dibatasi 10 baris pertama)")
    selected_mhs_columns = [col for col in ["nim", "nama", "fakultas", "jurusan", "jenjang_normal", "tahun_angkatan", "status_kategori", "gender_normal", "propinsi", "jenis_seleksi", "jenis_sekolah"] if col in academic_df.columns]
    mhs_display_df = academic_df[selected_mhs_columns]
    st.dataframe(
        mhs_display_df.head(10), use_container_width=True, hide_index=True,
        column_config={"nim": "NIM", "nama": "Nama", "fakultas": "Fakultas", "jurusan": "Program Studi", "jenjang_normal": "Jenjang", "tahun_angkatan": "Angkatan", "status_kategori": "Status", "gender_normal": "L/P", "propinsi": "Propinsi", "jenis_seleksi": "Jalur Seleksi", "jenis_sekolah": "Jenis Sekolah"}
    )


    # =========================================================
    # C. DATA ASAL / JENIS SEKOLAH
    # =========================================================
    subheading("C. Data Asal / Jenis Sekolah")

    if "jenis_sekolah" in academic_df.columns and not academic_df["jenis_sekolah"].eq("Tidak diketahui").all():
        school_data = academic_df[academic_df["jenis_sekolah"] != "Tidak diketahui"]
        
        cols = st.columns(3)
        with cols[0]: kpi("🏫", "Data Sekolah Terdata", format_number(len(school_data)), "Berdasarkan filter")
        with cols[1]: kpi("📊", "Jenis Sekolah Terbanyak", school_data["jenis_sekolah"].mode()[0] if not school_data.empty else "-", "Kategori Dominan")
        with cols[2]: kpi("📈", "Jumlah Kategori Sekolah", format_number(school_data["jenis_sekolah"].nunique()), "Variasi Jenis")

        if not school_data.empty:
            school_summary = school_data["jenis_sekolah"].value_counts().rename_axis("Jenis Sekolah").reset_index(name="Jumlah Mahasiswa")
            
            sch_col1, sch_col2 = st.columns([1, 1])
            with sch_col1:
                figure = px.pie(school_summary, names="Jenis Sekolah", values="Jumlah Mahasiswa", hole=0.58, title="Komposisi Jenis Sekolah", color_discrete_sequence=[C["primary"], C["orange"], C["yellow"], C["blue"], C["gray"]])
                figure.update_layout(height=430, paper_bgcolor="white", legend=dict(orientation="h", y=-0.15))
                figure.update_traces(textinfo="label+value+percent")
                st.plotly_chart(figure, use_container_width=True)
            
            with sch_col2:
                st.markdown("<br>**Detail Data Jenis Sekolah** (Tampilan dibatasi 10 baris pertama)", unsafe_allow_html=True)
                st.dataframe(school_summary.head(10), use_container_width=True, hide_index=True)
    else:
        empty_dashboard("Data Asal / Jenis Sekolah", global_year, "🏫")
        st.info("Catatan: Kolom 'jenis_sekolah' belum ditemukan di sumber data Anda atau isinya kosong. Grafik akan otomatis muncul jika datanya tersedia.")


    # =========================================================
    # D. DATA KEBUTUHAN KHUSUS
    # =========================================================
    subheading("D. Data Mahasiswa Berkebutuhan Khusus")

    filtered_difabel = df_difabel.copy()
    if selected_cohorts: filtered_difabel = filtered_difabel[filtered_difabel["tahun_angkatan"].isin(selected_cohorts)]
    if selected_faculties: filtered_difabel = filtered_difabel[filtered_difabel["fakultas"].isin(selected_faculties)]
    if selected_programs: filtered_difabel = filtered_difabel[filtered_difabel["jurusan"].isin(selected_programs)]
    if selected_provinces: filtered_difabel = filtered_difabel[filtered_difabel["propinsi"].isin(selected_provinces)]

    cols = st.columns(3)
    with cols[0]: kpi("♿", "Total mahasiswa", format_number(filtered_difabel["nim"].nunique()), "NIM unik")
    with cols[1]: kpi("✅", "Status aktif", format_number(filtered_difabel.loc[filtered_difabel["status_normal"].eq("aktif"), "nim"].nunique()), "")
    with cols[2]: kpi("📚", "Program Studi", format_number(filtered_difabel["jurusan"].nunique()), "")

    if not filtered_difabel.empty:
        need_df = filtered_difabel["kebutuhan_khusus"].value_counts().rename_axis("Kebutuhan Khusus").reset_index(name="Mahasiswa").sort_values("Mahasiswa")
        figure = px.bar(need_df, x="Mahasiswa", y="Kebutuhan Khusus", orientation="h", title="Jenis Kebutuhan Khusus", color_discrete_sequence=[C["primary"]])
        st.plotly_chart(figure, use_container_width=True)

    st.markdown("**Detail Mahasiswa Kebutuhan Khusus** (Tampilan dibatasi 10 baris pertama)")
    difabel_export_cols = [c for c in ["nim", "nama", "fakultas", "jurusan", "tahun_angkatan", "kebutuhan_khusus", "status_normal", "gender_normal", "propinsi"] if c in filtered_difabel.columns]
    st.dataframe(filtered_difabel[difabel_export_cols].head(10), use_container_width=True, hide_index=True)


    # =========================================================
    # E. DATA DAERAH TERTINGGAL
    # =========================================================
    subheading("E. Data Mahasiswa Daerah Tertinggal")

    disadvantaged = academic_df[academic_df["asal_daerah_tertinggal"]]

    cols = st.columns(3)
    with cols[0]: kpi("🗺️", "Mahasiswa daerah tertinggal", format_number(len(disadvantaged)), "")
    with cols[1]: kpi("🏘️", "Kabupaten teridentifikasi", format_number(disadvantaged["kota_normal"].nunique()), "")
    with cols[2]: kpi("🌍", "Propinsi", format_number(disadvantaged["propinsi"].nunique()), "")

    st.markdown("**Detail Mahasiswa Daerah Tertinggal** (Tampilan dibatasi 10 baris pertama)")
    dt_export_cols = [c for c in ["nim", "nama", "fakultas", "jurusan", "tahun_angkatan", "status_kategori", "gender_normal", "kota_normal", "propinsi"] if c in disadvantaged.columns]
    st.dataframe(disadvantaged[dt_export_cols].head(10), use_container_width=True, hide_index=True)


    # =========================================================
    # F. DATA KELULUSAN
    # =========================================================
    subheading("F. Data Kelulusan")

    academic_yud = df_yudisium[df_yudisium["tahun_lulus"].eq(academic_start_year)].copy()

    if selected_levels: academic_yud = academic_yud[academic_yud["jenjang_normal"].isin(selected_levels)]
    if selected_faculties: academic_yud = academic_yud[academic_yud["fakultas"].isin(selected_faculties)]
    if selected_programs: academic_yud = academic_yud[academic_yud["prodi"].isin(selected_programs)]

    graduate_count = academic_yud["nim"].nunique()
    on_time_count = academic_yud.loc[academic_yud["tepat_waktu_bool"], "nim"].nunique()
    on_time_ipk_count = academic_yud.loc[academic_yud["tepat_waktu_ipk_325"], "nim"].nunique()
    on_time_percentage = percentage(on_time_count, graduate_count)

    graduation_cols = st.columns(4)
    graduation_values = [
        ("🎓", "Jumlah lulusan", format_number(graduate_count), f"Yudisium {academic_start_year}"),
        ("⏱️", "Lulus tepat waktu", format_number(on_time_count), f"Yudisium {academic_start_year}"),
        ("🏅", "Tepat waktu dan IPK ≥ 3,25", format_number(on_time_ipk_count), f"Yudisium {academic_start_year}"),
        ("📈", "Kelulusan tepat waktu", format_percent(on_time_percentage), f"Yudisium {academic_start_year}"),
    ]
    for column, item in zip(graduation_cols, graduation_values):
        with column: kpi(*item)

    if academic_yud.empty:
        st.info(f"Data yudisium tahun {academic_start_year} tidak tersedia untuk filter yang dipilih.")
    else:
        graduation_status = pd.DataFrame({"Kategori": ["Lulus tepat waktu", "Tidak tepat waktu"], "Lulusan": [on_time_count, max(graduate_count - on_time_count, 0)]})
        graduation_ipk = pd.DataFrame({"Kategori": ["Tepat waktu dan IPK ≥ 3,25", "Tepat waktu dan IPK < 3,25", "Tidak tepat waktu"], "Lulusan": [on_time_ipk_count, max(on_time_count - on_time_ipk_count, 0), max(graduate_count - on_time_count, 0)]})

        graduation_col, ipk_col = st.columns(2)
        with graduation_col:
            figure = px.pie(graduation_status, names="Kategori", values="Lulusan", hole=0.58, title=f"Komposisi Ketepatan Waktu Yudisium {academic_start_year}", color_discrete_sequence=[C["primary"], C["orange"]])
            figure.update_layout(height=430, paper_bgcolor="white")
            figure.update_traces(textinfo="label+value+percent")
            st.plotly_chart(figure, use_container_width=True)

        with ipk_col:
            figure = px.bar(graduation_ipk, x="Kategori", y="Lulusan", text="Lulusan", title="Ketepatan Waktu dan IPK Lulusan", color="Kategori", color_discrete_sequence=[C["primary"], C["yellow"], C["orange"]])
            figure.update_layout(height=430, paper_bgcolor="white", plot_bgcolor="white", showlegend=False, xaxis_title=None, yaxis_title="Lulusan", margin=dict(l=10, r=10, t=55, b=90))
            figure.update_xaxes(tickangle=-20)
            st.plotly_chart(figure, use_container_width=True)

        graduation_faculty = academic_yud.groupby("fakultas", dropna=False).agg(Jumlah_Lulusan=("nim", "nunique"), Lulus_Tepat_Waktu=("tepat_waktu_bool", "sum"), Tepat_Waktu_IPK_325=("tepat_waktu_ipk_325", "sum"), Rata_Rata_IPK=("ipk", "mean")).reset_index()
        jumlah_lulusan_numeric = pd.to_numeric(graduation_faculty["Jumlah_Lulusan"], errors="coerce")
        lulus_tepat_waktu_numeric = pd.to_numeric(graduation_faculty["Lulus_Tepat_Waktu"], errors="coerce")
        graduation_faculty["Kelulusan_Tepat_Waktu"] = lulus_tepat_waktu_numeric.div(jumlah_lulusan_numeric.where(jumlah_lulusan_numeric.ne(0))).mul(100).round(2)
        tepat_waktu_ipk_numeric = pd.to_numeric(graduation_faculty["Tepat_Waktu_IPK_325"], errors="coerce")
        graduation_faculty["Tepat_Waktu_IPK_325_Persen"] = tepat_waktu_ipk_numeric.div(jumlah_lulusan_numeric.where(jumlah_lulusan_numeric.ne(0))).mul(100).round(2)
        graduation_faculty["Rata_Rata_IPK"] = graduation_faculty["Rata_Rata_IPK"].round(2)

        figure = px.bar(graduation_faculty.sort_values("Kelulusan_Tepat_Waktu"), x="Kelulusan_Tepat_Waktu", y="fakultas", orientation="h", text="Kelulusan_Tepat_Waktu", title="Persentase Kelulusan Tepat Waktu per Fakultas", color_discrete_sequence=[C["primary"]])
        figure.update_layout(height=520, paper_bgcolor="white", plot_bgcolor="white", showlegend=False, xaxis_title="Persentase", yaxis_title=None)
        figure.update_xaxes(range=[0, 100])
        figure.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        st.plotly_chart(figure, use_container_width=True)

        yudisium_columns = [c for c in ["nim", "nama", "tanggal_lulus", "tahun_lulus", "jenjang_normal", "fakultas", "prodi", "semester", "ipk", "tepat_waktu"] if c in academic_yud.columns]
        graduation_export_data = academic_yud[yudisium_columns].copy()

        st.markdown("**Detail Data Kelulusan** (Tampilan dibatasi 10 baris pertama)")
        st.dataframe(
            graduation_export_data.head(10), use_container_width=True, hide_index=True, 
            column_config={"nim": "NIM", "nama": "Nama", "tanggal_lulus": st.column_config.DateColumn("Tanggal Yudisium/Lulus", format="DD/MM/YYYY"), "tahun_lulus": "Tahun Lulus", "jenjang_normal": "Jenjang", "fakultas": "Fakultas", "prodi": "Program Studi", "semester": "Semester", "ipk": st.column_config.NumberColumn("IPK", format="%.2f"), "tepat_waktu": "Tepat Waktu"}
        )


    # =========================================================
    # G. DATA AKREDITASI PROGRAM STUDI
    # =========================================================
    subheading("G. Data Akreditasi Program Studi")
    
    akreditasi_filter = df_akreditasi.copy()
    if selected_levels: akreditasi_filter = akreditasi_filter[akreditasi_filter["jenjang_normal"].isin(selected_levels)]
    if selected_faculties: akreditasi_filter = akreditasi_filter[akreditasi_filter["fakultas"].isin(selected_faculties)]
    if selected_programs: akreditasi_filter = akreditasi_filter[akreditasi_filter["prodi"].isin(selected_programs)]

    total_prodi_akreditasi = len(akreditasi_filter)
    akreditasi_aktif = int(akreditasi_filter["status_akreditasi"].eq("Aktif").sum())
    akreditasi_kedaluwarsa = int(akreditasi_filter["status_akreditasi"].eq("Kedaluwarsa").sum())
    segera_berakhir = int(akreditasi_filter["masa_berakhir"].eq("Berakhir ≤ 1 tahun").sum())
    jumlah_unggul = int(akreditasi_filter["peringkat"].str.upper().eq("UNGGUL").sum())

    accreditation_kpis = st.columns(5)
    accreditation_values = [
        ("📚", "Program studi", total_prodi_akreditasi, "Sesuai filter akademik"),
        ("✅", "Akreditasi aktif", akreditasi_aktif, "Masa berlaku belum berakhir"),
        ("🏅", "Peringkat Unggul", jumlah_unggul, "Peringkat akreditasi Unggul"),
        ("⏳", "Berakhir ≤ 1 tahun", segera_berakhir, "Perlu perhatian"),
        ("⚠️", "Kedaluwarsa", akreditasi_kedaluwarsa, "Masa berlaku berakhir"),
    ]
    for column, item in zip(accreditation_kpis, accreditation_values):
        with column: kpi(item[0], item[1], format_number(item[2]), item[3])

    accreditation_left, accreditation_right = st.columns([1, 1.25])
    with accreditation_left:
        rank_summary = akreditasi_filter["peringkat"].fillna("Tidak diketahui").value_counts().rename_axis("Peringkat").reset_index(name="Program Studi")
        figure = px.pie(rank_summary, names="Peringkat", values="Program Studi", hole=0.56, title="Komposisi Peringkat Akreditasi", color_discrete_sequence=[C["primary"], C["yellow"], C["blue"], C["orange"], C["purple"], C["gray"]])
        figure.update_layout(height=440, paper_bgcolor="white", legend=dict(orientation="h", y=-0.12))
        figure.update_traces(textinfo="label+value+percent")
        st.plotly_chart(figure, use_container_width=True)

    with accreditation_right:
        faculty_accreditation = akreditasi_filter.groupby(["fakultas", "peringkat"], dropna=False).size().reset_index(name="Program Studi")
        figure = px.bar(faculty_accreditation, x="fakultas", y="Program Studi", color="peringkat", barmode="stack", title="Peringkat Akreditasi per Fakultas", color_discrete_sequence=[C["primary"], C["yellow"], C["blue"], C["orange"], C["purple"], C["gray"]])
        figure.update_layout(height=440, paper_bgcolor="white", plot_bgcolor="white", xaxis_title=None, yaxis_title="Program Studi", legend_title_text="Peringkat", margin=dict(l=10, r=10, t=55, b=100))
        figure.update_xaxes(tickangle=-35)
        st.plotly_chart(figure, use_container_width=True)

    accreditation_table = (
        akreditasi_filter[["jenjang_normal", "fakultas", "prodi", "peringkat", "no_sk_akreditasi", "tanggal_sk_akreditasi", "masa_mulai_sk_akreditasi", "masa_sk_akhir_akreditasi", "status_akreditasi", "masa_berakhir", "sisa_hari"]]
        .assign(prioritas_urutan=lambda frame: frame["masa_berakhir"].map({"Sudah berakhir": 1, "Berakhir ≤ 1 tahun": 2, "Normal": 3, "Belum diketahui": 4}).fillna(5))
        .sort_values(["prioritas_urutan", "masa_sk_akhir_akreditasi", "fakultas", "prodi"], ascending=[True, True, True, True])
        .drop(columns=["prioritas_urutan"])
    )

    st.markdown("**Daftar dan Masa Berlaku Akreditasi Program Studi** (Tampilan dibatasi 10 baris pertama)")
    st.dataframe(
        accreditation_table.head(10), use_container_width=True, hide_index=True,
        column_config={
            "jenjang_normal": "Jenjang", "fakultas": "Fakultas", "prodi": "Program Studi", "peringkat": "Peringkat", "no_sk_akreditasi": "Nomor SK",
            "tanggal_sk_akreditasi": st.column_config.DateColumn("Tanggal SK", format="DD/MM/YYYY"),
            "masa_mulai_sk_akreditasi": st.column_config.DateColumn("Mulai Berlaku", format="DD/MM/YYYY"),
            "masa_sk_akhir_akreditasi": st.column_config.DateColumn("Akhir Berlaku", format="DD/MM/YYYY"),
            "status_akreditasi": "Status Akreditasi", "masa_berakhir": "Status Masa Berlaku",
            "sisa_hari": st.column_config.NumberColumn("Sisa Hari", format="%d"),
        },
    )


# --- 3 S/D 8. DASHBOARD KOSONG ---
with main_tabs[2]: hero("Dashboard SDM", "Kerangka dashboard sumber daya manusia sedang disiapkan."); empty_dashboard("Dashboard SDM", global_year, "👩‍🏫")
with main_tabs[3]: hero("Dashboard Keuangan", "Kerangka dashboard keuangan sedang disiapkan."); empty_dashboard("Dashboard Keuangan", global_year, "💰")
with main_tabs[4]: hero("Dashboard Aset", "Kerangka dashboard aset sedang disiapkan."); empty_dashboard("Dashboard Aset", global_year, "🏢")
with main_tabs[5]: hero("Dashboard Riset", "Kerangka dashboard riset dan publikasi sedang disiapkan."); empty_dashboard("Dashboard Riset", global_year, "🔬")
with main_tabs[6]: hero("Dashboard Alumni", "Kerangka dashboard alumni dan tracer study sedang disiapkan."); empty_dashboard("Dashboard Alumni", global_year, "🎓")
with main_tabs[7]: hero("Dashboard Prestasi Mahasiswa", "Kerangka dashboard prestasi akademik dan nonakademik mahasiswa sedang disiapkan."); empty_dashboard("Dashboard Prestasi Mahasiswa", global_year, "🏆")
with main_tabs[8]: hero("Dashboard Paperless (PLO)", "Kerangka dashboard layanan dan proses administrasi paperless sedang disiapkan."); empty_dashboard("Dashboard Paperless (PLO)", global_year, "📄")

# --- 9. DASHBOARD IKU ---
with main_tabs[9]:
    hero("Dashboard Capaian IKU", "Capaian indikator berdasarkan sumber data yang tersedia (Berbasis Filter Tahun Global).")

    iku_sub_section = st.radio(
        "Sub-Menu IKU",
        ["IKU Akademik", "IKU SDM", "IKU Keuangan", "IKU Aset", "IKU Riset", "IKU Alumni"],
        horizontal=True,
        label_visibility="collapsed"
    )
    st.markdown("<br>", unsafe_allow_html=True)

    if iku_sub_section == "IKU Akademik":
        new_t = int(df["tahun_angkatan"].eq(global_year).sum())
        new_previous = int(df["tahun_angkatan"].eq(global_year - 1).sum())
        iku_01_01 = growth(new_t, new_previous)

        new_students = df[df["tahun_angkatan"].eq(global_year)]
        disadvantaged_count = int(new_students["asal_daerah_tertinggal"].sum())
        iku_01_07 = percentage(disadvantaged_count, len(new_students))

        dif_t = int(df_difabel["tahun_angkatan"].eq(global_year).sum())
        dif_previous = int(df_difabel["tahun_angkatan"].eq(global_year - 1).sum())
        iku_02_01 = growth(dif_t, dif_previous)

        yud_year = df_yudisium[df_yudisium["tahun_lulus"].eq(global_year)]
        graduate_count = yud_year["nim"].nunique()
        on_time_count = yud_year.loc[yud_year["tepat_waktu_bool"], "nim"].nunique()
        on_time_ipk_count = yud_year.loc[yud_year["tepat_waktu_ipk_325"], "nim"].nunique()

        iku_04_1 = percentage(on_time_ipk_count, graduate_count)
        iku_04_3 = percentage(on_time_count, graduate_count)

        pmb_years = sorted(df_pmb["tahun"].dropna().astype(int).unique().tolist())
        iku_pmb_year = global_year if global_year in pmb_years else max(pmb_years)
        pmb_year_df = df_pmb[df_pmb["tahun"].eq(iku_pmb_year)]

        total_lulus = float(pmb_year_df["lulus_seleksi"].sum())
        total_daftar = float(pmb_year_df["daftar_ulang"].sum())
        iku_35_54 = percentage(total_daftar, total_lulus)

        if iku_pmb_year != global_year:
            st.warning(f"Data PMB tahun {global_year} belum tersedia. IKU-35-54 memakai data PMB terbaru {iku_pmb_year}.")

        iku_rows = [
            ("IKU-01-01", "Persentase peningkatan mahasiswa pada PTK", iku_01_01, "(Mahasiswa baru t − mahasiswa baru t−1) ÷ mahasiswa baru t−1 × 100%", f"{new_t} dibanding {new_previous}.", ("Tersedia" if iku_01_01 is not None else "Sebagian")),
            ("IKU-01-05", "Persentase lulusan pesantren yang ditampung", None, "Mahasiswa baru lulusan pesantren ÷ total mahasiswa baru × 100%", "Belum tersedia data asal pesantren.", "Belum tersedia"),
            ("IKU-01-07", "Persentase mahasiswa baru dari daerah tertinggal", iku_01_07, "Mahasiswa baru daerah tertinggal ÷ total mahasiswa baru × 100%", f"{disadvantaged_count} dari {len(new_students)} mahasiswa baru.", "Tersedia"),
            ("IKU-02-01", "Persentase peningkatan mahasiswa berkebutuhan khusus", iku_02_01, "(Mahasiswa kebutuhan khusus t − t−1) ÷ mahasiswa kebutuhan khusus t−1 × 100%", f"{dif_t} dibanding {dif_previous}.", ("Tersedia" if iku_02_01 is not None else "Sebagian")),
            ("IKU-04-1", "Persentase mahasiswa lulus tepat waktu dengan IPK ≥ 3,25", iku_04_1, "Lulusan tepat waktu dengan IPK ≥ 3,25 ÷ seluruh lulusan × 100%", f"{on_time_ipk_count} dari {graduate_count} lulusan.", ("Tersedia" if iku_04_1 is not None else "Belum tersedia")),
            ("IKU-04-3", "Persentase kelulusan tepat waktu", iku_04_3, "Lulusan tepat waktu ÷ seluruh lulusan × 100%", f"{on_time_count} dari {graduate_count} lulusan.", ("Tersedia" if iku_04_3 is not None else "Belum tersedia")),
            ("IKU-35-54", "Yield Rate Mahasiswa Baru", iku_35_54, "Daftar ulang ÷ lulus seleksi × 100%", f"PMB {iku_pmb_year}: {format_number(total_daftar)} dari {format_number(total_lulus)}.", ("Tersedia" if iku_35_54 is not None else "Belum tersedia")),
        ]

        for start in range(0, len(iku_rows), 3):
            columns = st.columns(3)
            for column, row in zip(columns, iku_rows[start:start + 3]):
                with column: iku_card(*row)

        section("Ketersediaan Dashboard Profil", "Komponen yang mendukung IKU-48-06.")
        profile_readiness = pd.DataFrame([
            ["Total mahasiswa", "Tersedia"], ["Jenis kelamin per universitas", "Tersedia"], ["Mahasiswa per fakultas", "Tersedia"],
            ["Mahasiswa per program studi", "Tersedia"], ["Mahasiswa per jenjang", "Tersedia"], ["Mahasiswa per propinsi", "Tersedia"],
            ["Profil alumni", "Belum tersedia"],
        ], columns=["Komponen", "Status"])
        st.dataframe(profile_readiness, use_container_width=True, hide_index=True)

    elif iku_sub_section == "IKU SDM": empty_dashboard("IKU SDM", global_year, "👩‍🏫")
    elif iku_sub_section == "IKU Keuangan": empty_dashboard("IKU Keuangan", global_year, "💰")
    elif iku_sub_section == "IKU Aset": empty_dashboard("IKU Aset", global_year, "🏢")
    elif iku_sub_section == "IKU Riset": empty_dashboard("IKU Riset", global_year, "🔬")
    elif iku_sub_section == "IKU Alumni": empty_dashboard("IKU Alumni", global_year, "🎓")

st.markdown("<br>", unsafe_allow_html=True)
st.caption("Sumber data: v_mahasiswa.xls, difabel.xls, vs_yudisium_mahasiswa.xls, vs_rekap_pmb.xls, dan akreditasiprodi.xls.")
