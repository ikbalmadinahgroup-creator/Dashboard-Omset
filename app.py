"""
Dashboard Omset MFlash
=======================
Streamlit dashboard menampilkan:
  1. Omset All
  2. Omset Service
  3. Omset Gadget & Aksesoris
  4. Omset Marketing Corporate

Filter: Tahun, Bulan, Cabang
Data diisi lewat tombol upload file Excel (tidak perlu edit source code / repo GitHub).

Sumber data:
  - Sheet "Faktur Penjualan" pada file utama -> Omset All, Omset Service, Omset Gadget & Aksesoris
    (dihitung dari kolom CABANG, TGL FAKTUR, KATEGORI PENJUALAN, TOTAL HARGA)
  - File terpisah "Data Marketing Corporate" (template disediakan di sidebar) -> Omset Marketing Corporate
    (kolom: Tahun, Bulan, Cabang, Omset)

Jalankan dengan:
    streamlit run app.py
"""

import io
import os
import re
from datetime import datetime

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------------------------------------------------------------------------
# Konfigurasi umum
# --------------------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard Omset MFlash", layout="wide", page_icon="📊")

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
os.makedirs(MAIN_DATA_DIR, exist_ok=True)
CORPORATE_DATA_PATH = os.path.join(DATA_DIR, "corporate_data.xlsx")

MAIN_SHEET_NAME = "Faktur Penjualan"
CORE_COLUMNS = ["TGL FAKTUR", "KATEGORI PENJUALAN", "TOTAL HARGA"]
REQUIRED_COLUMNS = ["CABANG"] + CORE_COLUMNS
MAX_MAIN_FILES = 50

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
BULAN_MAP = {name.lower(): i + 1 for i, name in enumerate(BULAN_ID)}
# beberapa alias singkatan umum
BULAN_ALIAS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
    "jul": 7, "agu": 8, "aug": 8, "sep": 9, "okt": 10, "oct": 10,
    "nov": 11, "des": 12, "dec": 12,
}


def format_rupiah(value: float) -> str:
    if value is None or pd.isna(value):
        value = 0
    return "Rp " + f"{value:,.0f}".replace(",", ".")


def sanitize_filename(name: str) -> str:
    base = os.path.basename(name)
    base = re.sub(r"[^A-Za-z0-9._\- ]+", "_", base)
    return base.strip() or "file.xlsx"


_FILENAME_NOISE_WORDS = {
    w.upper() for w in
    ["FP", "FAKTUR", "PENJUALAN", "DATA", "CABANG", "REPORT", "LAPORAN",
     "EXPORT", "OMSET", "BULAN", "PERIODE", "MINGGU", "TRIWULAN"] + BULAN_ID
}


def branch_from_filename(filename: str) -> str:
    """Tebak nama cabang dari nama file, dipakai jika file tidak punya kolom CABANG.
    Strategi: buang kata-kata umum (FP, Data, nama bulan, dst) & angka, lalu ambil
    kata pertama yang tersisa sebagai nama cabang."""
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"[_\-–]+", " ", name)
    name = re.sub(r"([A-Za-z])(\d)", r"\1 \2", name)
    name = re.sub(r"(\d)([A-Za-z])", r"\1 \2", name)
    tokens = [t for t in re.split(r"\s+", name) if t]
    cleaned = [t for t in tokens if t.upper() not in _FILENAME_NOISE_WORDS and not t.isdigit()]
    return cleaned[0].upper() if cleaned else "TIDAK DIKETAHUI"


def classify_kategori(kategori: str) -> str:
    """Kelompokkan KATEGORI PENJUALAN menjadi Service / Gadget & Aksesoris / Lainnya."""
    if not kategori:
        return "Lainnya"
    k = str(kategori).strip().upper()
    if k.startswith("SERVICE"):
        return "Service"
    if k.startswith("PENJUALAN"):
        return "Gadget & Aksesoris"
    return "Lainnya"


def parse_bulan(value) -> int | None:
    """Terima nama bulan (Indonesia), singkatan, atau angka 1-12."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        v = int(value)
        return v if 1 <= v <= 12 else None
    s = str(value).strip().lower()
    if s in BULAN_MAP:
        return BULAN_MAP[s]
    if s in BULAN_ALIAS:
        return BULAN_ALIAS[s]
    try:
        v = int(float(s))
        return v if 1 <= v <= 12 else None
    except ValueError:
        return None


# --------------------------------------------------------------------------------------
# Loader data utama (Faktur Penjualan) - streaming read_only agar cepat & hemat memori
# --------------------------------------------------------------------------------------

def _find_data_sheet(wb):
    """Cari sheet yang punya kolom TGL FAKTUR, KATEGORI PENJUALAN, TOTAL HARGA.
    Diprioritaskan sheet bernama 'Faktur Penjualan', lalu sheet lain yang cocok
    (berguna untuk file per-cabang yang mungkin punya nama sheet berbeda)."""
    candidates = sorted(
        wb.sheetnames,
        key=lambda n: 0 if n.strip().lower() == MAIN_SHEET_NAME.lower() else 1,
    )
    for name in candidates:
        ws = wb[name]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row is None:
            continue
        col_idx = {}
        for i, h in enumerate(header_row):
            if h is None:
                continue
            key = str(h).strip().upper()
            if key in REQUIRED_COLUMNS:
                col_idx[key] = i
        if all(c in col_idx for c in CORE_COLUMNS):
            return ws, col_idx
    return None, None


@st.cache_data(show_spinner=False)
def load_main_data(file_bytes: bytes, filename_hint: str = "") -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    ws, col_idx = _find_data_sheet(wb)
    if ws is None:
        raise ValueError(
            f"Tidak ada sheet dengan kolom {', '.join(CORE_COLUMNS)}. "
            f"Sheet yang tersedia: {', '.join(wb.sheetnames)}"
        )

    has_cabang_col = "CABANG" in col_idx
    fallback_branch = branch_from_filename(filename_hint) if filename_hint else "TIDAK DIKETAHUI"

    max_scan_row = min(ws.max_row, 300_000)
    rows = []
    empty_streak = 0
    for row in ws.iter_rows(min_row=2, max_row=max_scan_row, values_only=True):
        cabang = row[col_idx["CABANG"]] if has_cabang_col else None
        if not cabang or str(cabang).strip() == "":
            cabang = fallback_branch

        tgl = row[col_idx["TGL FAKTUR"]]
        kategori = row[col_idx["KATEGORI PENJUALAN"]]
        total = row[col_idx["TOTAL HARGA"]]

        if tgl is None and kategori is None and total is None:
            empty_streak += 1
            if empty_streak > 3000:
                break
            continue
        empty_streak = 0

        if not isinstance(tgl, datetime):
            continue
        if total is None:
            total = 0
        try:
            total = float(total)
        except (TypeError, ValueError):
            continue

        rows.append(
            {
                "Cabang": str(cabang).strip().upper(),
                "Tanggal": tgl,
                "Tahun": tgl.year,
                "Bulan": tgl.month,
                "Kategori": str(kategori).strip().upper() if kategori else "LAINNYA",
                "Omset": total,
                "SumberFile": filename_hint,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Kelompok"] = df["Kategori"].apply(classify_kategori)
    return df


def load_all_main_data(main_dir: str) -> tuple[pd.DataFrame, list[str]]:
    """Baca semua file .xlsx di folder main_dir (satu file = biasanya satu cabang) dan gabungkan."""
    files = sorted(f for f in os.listdir(main_dir) if f.lower().endswith(".xlsx"))
    frames = []
    errors = []
    for fname in files:
        fpath = os.path.join(main_dir, fname)
        try:
            with open(fpath, "rb") as f:
                file_bytes = f.read()
            df = load_main_data(file_bytes, filename_hint=fname)
            if not df.empty:
                frames.append(df)
        except ValueError as e:
            errors.append(f"{fname}: {e}")
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, errors


@st.cache_data(show_spinner=False)
def load_corporate_data(file_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    for name in xls.sheet_names:
        if "corporate" in name.lower() or "marketing" in name.lower():
            sheet_name = name
            break

    raw = pd.read_excel(xls, sheet_name=sheet_name)
    raw.columns = [str(c).strip().upper() for c in raw.columns]

    required = ["TAHUN", "BULAN", "CABANG", "OMSET"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Kolom berikut tidak ditemukan di file Marketing Corporate: {', '.join(missing)}. "
            f"Gunakan template yang tersedia di sidebar."
        )

    df = raw[required].copy()
    df = df.dropna(subset=["TAHUN", "BULAN", "CABANG"])
    df["Tahun"] = df["TAHUN"].astype(int)
    df["Bulan"] = df["BULAN"].apply(parse_bulan)
    df["Cabang"] = df["CABANG"].astype(str).str.strip().str.upper()
    df["Omset"] = pd.to_numeric(df["OMSET"], errors="coerce").fillna(0)
    df = df.dropna(subset=["Bulan"])
    df["Bulan"] = df["Bulan"].astype(int)
    return df[["Tahun", "Bulan", "Cabang", "Omset"]]


def make_corporate_template() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Marketing Corporate"
    ws.append(["Tahun", "Bulan", "Cabang", "Omset"])
    contoh = [
        [2026, "Januari", "KLENDER", 50000000],
        [2026, "Januari", "RADJIMAN", 12000000],
        [2026, "Februari", "KLENDER", 43000000],
    ]
    for r in contoh:
        ws.append(r)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16

    info = wb.create_sheet("Petunjuk")
    info.append(["Petunjuk pengisian Data Marketing Corporate"])
    info.append([""])
    info.append(["- Tahun: angka 4 digit, contoh 2026"])
    info.append(["- Bulan: nama bulan dalam Bahasa Indonesia, contoh Januari, Februari, ... Desember"])
    info.append(["- Cabang: nama cabang, harus konsisten dengan nama cabang di data utama (Faktur Penjualan)"])
    info.append(["- Omset: angka total omset marketing corporate untuk cabang & bulan tersebut"])
    info.append(["- Satu baris = satu kombinasi Tahun + Bulan + Cabang"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Sidebar: upload data
# --------------------------------------------------------------------------------------

st.sidebar.title("📥 Data")

st.sidebar.caption(
    "Upload file data per cabang (masing-masing file = data 1 cabang, format sama seperti "
    "sheet 'Faktur Penjualan'). Bisa pilih banyak file sekaligus, sampai 50 file."
)
main_uploads = st.sidebar.file_uploader(
    "Upload file data cabang",
    type=["xlsx"],
    accept_multiple_files=True,
    key="main_uploads",
)
if main_uploads:
    existing = len(os.listdir(MAIN_DATA_DIR))
    if existing + len(main_uploads) > MAX_MAIN_FILES:
        st.sidebar.error(
            f"Maksimal {MAX_MAIN_FILES} file. Saat ini sudah ada {existing} file tersimpan, "
            f"hapus beberapa dulu di bawah sebelum upload {len(main_uploads)} file baru."
        )
    else:
        for uf in main_uploads:
            safe_name = sanitize_filename(uf.name)
            with open(os.path.join(MAIN_DATA_DIR, safe_name), "wb") as f:
                f.write(uf.getbuffer())
        st.sidebar.success(f"{len(main_uploads)} file tersimpan.")

stored_main_files = sorted(os.listdir(MAIN_DATA_DIR))
if stored_main_files:
    with st.sidebar.expander(f"📁 File cabang tersimpan ({len(stored_main_files)}/{MAX_MAIN_FILES})"):
        for fname in stored_main_files:
            fc1, fc2 = st.columns([4, 1])
            fc1.write(fname)
            if fc2.button("🗑️", key=f"del_{fname}"):
                os.remove(os.path.join(MAIN_DATA_DIR, fname))
                st.rerun()
        if st.button("Hapus semua file cabang", key="del_all_main"):
            for fname in stored_main_files:
                os.remove(os.path.join(MAIN_DATA_DIR, fname))
            st.rerun()

st.sidebar.markdown("---")

corp_upload = st.sidebar.file_uploader(
    "Upload file Data Marketing Corporate",
    type=["xlsx"],
    key="corp_upload",
)
if corp_upload is not None:
    with open(CORPORATE_DATA_PATH, "wb") as f:
        f.write(corp_upload.getbuffer())
    st.sidebar.success("File Marketing Corporate tersimpan.")

st.sidebar.download_button(
    "⬇️ Download template Marketing Corporate",
    data=make_corporate_template(),
    file_name="template_marketing_corporate.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.sidebar.caption(
    "File yang sudah diupload akan otomatis dipakai lagi setiap kali dashboard dibuka, "
    "sampai diganti dengan file baru."
)

# --------------------------------------------------------------------------------------
# Load data
# --------------------------------------------------------------------------------------

st.title("📊 Dashboard Omset MFlash")

if not os.listdir(MAIN_DATA_DIR):
    st.info("Silakan upload file data cabang (bisa lebih dari satu) lewat sidebar untuk mulai.")
    st.stop()

with st.spinner(f"Memproses {len(os.listdir(MAIN_DATA_DIR))} file data cabang..."):
    df_main, load_errors = load_all_main_data(MAIN_DATA_DIR)

if load_errors:
    st.warning(
        "Sebagian file dilewati karena formatnya tidak cocok:\n\n"
        + "\n".join(f"- {e}" for e in load_errors)
    )

if df_main.empty:
    st.warning("Tidak ada data transaksi yang bisa dibaca dari file yang diupload.")
    st.stop()

df_corp = pd.DataFrame(columns=["Tahun", "Bulan", "Cabang", "Omset"])
if os.path.exists(CORPORATE_DATA_PATH):
    with open(CORPORATE_DATA_PATH, "rb") as f:
        corp_bytes = f.read()
    try:
        df_corp = load_corporate_data(corp_bytes)
    except ValueError as e:
        st.warning(f"Data Marketing Corporate tidak dipakai: {e}")
else:
    st.sidebar.info(
        "Belum ada data Marketing Corporate. Download template di atas, isi, lalu upload."
    )

# --------------------------------------------------------------------------------------
# Filter
# --------------------------------------------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.title("🔎 Filter")

tahun_options = sorted(set(df_main["Tahun"]).union(set(df_corp["Tahun"])), reverse=True)
bulan_options = list(range(1, 13))
cabang_options = sorted(set(df_main["Cabang"]).union(set(df_corp["Cabang"])))

sel_tahun = st.sidebar.multiselect("Tahun", tahun_options, default=tahun_options)
sel_bulan = st.sidebar.multiselect(
    "Bulan", bulan_options, default=bulan_options, format_func=lambda m: BULAN_ID[m - 1]
)
sel_cabang = st.sidebar.multiselect("Cabang", cabang_options, default=cabang_options)

if not sel_tahun or not sel_bulan or not sel_cabang:
    st.warning("Pilih minimal satu Tahun, Bulan, dan Cabang di sidebar.")
    st.stop()

f_main = df_main[
    df_main["Tahun"].isin(sel_tahun)
    & df_main["Bulan"].isin(sel_bulan)
    & df_main["Cabang"].isin(sel_cabang)
]
f_corp = df_corp[
    df_corp["Tahun"].isin(sel_tahun)
    & df_corp["Bulan"].isin(sel_bulan)
    & df_corp["Cabang"].isin(sel_cabang)
]

# --------------------------------------------------------------------------------------
# KPI
# --------------------------------------------------------------------------------------

omset_all = f_main["Omset"].sum()
omset_service = f_main.loc[f_main["Kelompok"] == "Service", "Omset"].sum()
omset_gadget = f_main.loc[f_main["Kelompok"] == "Gadget & Aksesoris", "Omset"].sum()
omset_corporate = f_corp["Omset"].sum()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Omset All", format_rupiah(omset_all))
c2.metric("Omset Service", format_rupiah(omset_service))
c3.metric("Omset Gadget & Aksesoris", format_rupiah(omset_gadget))
c4.metric("Omset Marketing Corporate", format_rupiah(omset_corporate))

st.markdown("---")

# --------------------------------------------------------------------------------------
# Tren bulanan
# --------------------------------------------------------------------------------------

st.subheader("Tren Omset per Bulan")

trend_all = (
    f_main.groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index().assign(Kategori="Omset All")
)
trend_service = (
    f_main[f_main["Kelompok"] == "Service"]
    .groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index().assign(Kategori="Omset Service")
)
trend_gadget = (
    f_main[f_main["Kelompok"] == "Gadget & Aksesoris"]
    .groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index().assign(Kategori="Omset Gadget & Aksesoris")
)
trend_corp = (
    f_corp.groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index().assign(Kategori="Omset Marketing Corporate")
    if not f_corp.empty
    else pd.DataFrame({"Tahun": pd.Series(dtype="int64"), "Bulan": pd.Series(dtype="int64"),
                        "Omset": pd.Series(dtype="float64"), "Kategori": pd.Series(dtype="object")})
)

trend_parts = [d for d in [trend_all, trend_service, trend_gadget, trend_corp] if not d.empty]
trend = pd.concat(trend_parts, ignore_index=True) if trend_parts else trend_corp
if not trend.empty:
    trend["Periode"] = trend.apply(lambda r: f"{BULAN_ID[int(r['Bulan']) - 1][:3]} {int(r['Tahun'])}", axis=1)
    trend["urut"] = trend["Tahun"] * 100 + trend["Bulan"]
    trend = trend.sort_values("urut")
    fig_trend = px.line(
        trend, x="Periode", y="Omset", color="Kategori", markers=True,
        category_orders={"Periode": trend["Periode"].unique().tolist()},
    )
    fig_trend.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Omset (Rp)")
    st.plotly_chart(fig_trend, width="stretch")
else:
    st.info("Tidak ada data untuk kombinasi filter ini.")

# --------------------------------------------------------------------------------------
# Breakdown per cabang
# --------------------------------------------------------------------------------------

st.subheader("Omset per Cabang")

by_branch_all = f_main.groupby("Cabang")["Omset"].sum().rename("Omset All")
by_branch_service = f_main[f_main["Kelompok"] == "Service"].groupby("Cabang")["Omset"].sum().rename("Omset Service")
by_branch_gadget = f_main[f_main["Kelompok"] == "Gadget & Aksesoris"].groupby("Cabang")["Omset"].sum().rename("Omset Gadget & Aksesoris")
by_branch_corp = f_corp.groupby("Cabang")["Omset"].sum().rename("Omset Marketing Corporate")

by_branch = pd.concat(
    [by_branch_all, by_branch_service, by_branch_gadget, by_branch_corp], axis=1
).astype(float).fillna(0.0).reset_index()

bcol1, bcol2 = st.columns([2, 1])
with bcol1:
    branch_long = by_branch.melt(id_vars="Cabang", var_name="Kategori", value_name="Omset")
    fig_branch = px.bar(
        branch_long, x="Cabang", y="Omset", color="Kategori", barmode="group"
    )
    fig_branch.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Omset (Rp)")
    st.plotly_chart(fig_branch, width="stretch")

with bcol2:
    display_df = by_branch.copy()
    for col in display_df.columns[1:]:
        display_df[col] = display_df[col].apply(format_rupiah)
    st.dataframe(display_df, width="stretch", hide_index=True)

st.caption(
    "Omset All & Omset Service & Omset Gadget dihitung dari sheet 'Faktur Penjualan'. "
    "Omset Marketing Corporate dihitung dari file terpisah yang diupload (lihat sidebar)."
)
