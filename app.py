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
os.makedirs(DATA_DIR, exist_ok=True)
MAIN_DATA_PATH = os.path.join(DATA_DIR, "main_data.xlsx")
CORPORATE_DATA_PATH = os.path.join(DATA_DIR, "corporate_data.xlsx")

MAIN_SHEET_NAME = "Faktur Penjualan"
REQUIRED_COLUMNS = ["CABANG", "TGL FAKTUR", "KATEGORI PENJUALAN", "TOTAL HARGA"]

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

@st.cache_data(show_spinner=False)
def load_main_data(file_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    sheet_name = None
    for name in wb.sheetnames:
        if name.strip().lower() == MAIN_SHEET_NAME.lower():
            sheet_name = name
            break
    if sheet_name is None:
        raise ValueError(
            f"Sheet '{MAIN_SHEET_NAME}' tidak ditemukan di file. "
            f"Sheet yang tersedia: {', '.join(wb.sheetnames)}"
        )

    ws = wb[sheet_name]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    col_idx = {}
    for i, h in enumerate(header_row):
        if h is None:
            continue
        key = str(h).strip().upper()
        if key in REQUIRED_COLUMNS:
            col_idx[key] = i

    missing = [c for c in REQUIRED_COLUMNS if c not in col_idx]
    if missing:
        raise ValueError(
            f"Kolom berikut tidak ditemukan di sheet '{sheet_name}': {', '.join(missing)}"
        )

    max_scan_row = min(ws.max_row, 300_000)
    rows = []
    empty_streak = 0
    for row in ws.iter_rows(min_row=2, max_row=max_scan_row, values_only=True):
        cabang = row[col_idx["CABANG"]]
        if cabang is None or str(cabang).strip() == "":
            empty_streak += 1
            if empty_streak > 3000:
                break
            continue
        empty_streak = 0

        tgl = row[col_idx["TGL FAKTUR"]]
        kategori = row[col_idx["KATEGORI PENJUALAN"]]
        total = row[col_idx["TOTAL HARGA"]]

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
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Kelompok"] = df["Kategori"].apply(classify_kategori)
    return df


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

main_upload = st.sidebar.file_uploader(
    "Upload file data utama (mengandung sheet 'Faktur Penjualan')",
    type=["xlsx"],
    key="main_upload",
)
if main_upload is not None:
    with open(MAIN_DATA_PATH, "wb") as f:
        f.write(main_upload.getbuffer())
    st.sidebar.success("File data utama tersimpan.")

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

if not os.path.exists(MAIN_DATA_PATH):
    st.info("Silakan upload file data utama (sheet 'Faktur Penjualan') lewat sidebar untuk mulai.")
    st.stop()

with open(MAIN_DATA_PATH, "rb") as f:
    main_bytes = f.read()

try:
    with st.spinner("Memproses data utama..."):
        df_main = load_main_data(main_bytes)
except ValueError as e:
    st.error(str(e))
    st.stop()

if df_main.empty:
    st.warning("Tidak ada data transaksi yang bisa dibaca dari sheet 'Faktur Penjualan'.")
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
