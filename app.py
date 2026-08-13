"""
Dashboard Omset MFlash
=======================
Menampilkan omset untuk 3 kategori utama:
  1. Omset All
  2. Omset Service
  3. Omset Gadget & Aksesoris
(+ Omset Marketing Corporate)

Ada 2 tampilan:
  - Tab "Ringkasan": kartu KPI berwarna + progress ring % pencapaian + tren bulanan,
    bisa difilter Tahun / Bulan / Cabang.
  - Tab "Scoreboard": tabel per cabang gaya scoreboard (Target / Expected Value /
    Pencapaian / Gap / Kejar Target Per Hari / rata-rata omset bulan lalu vs bulan ini)
    + grafik progress harian (aktual kumulatif vs target pace).

Semua data diisi lewat tombol upload (tidak perlu edit source code / repo GitHub).
Kalau file yang diupload punya sheet "Scoreboard" (file master yang biasa dipakai),
Target & Expected Value (untuk Omset All/Service/Gadget & Aksesoris) dan Scoreboard
Marketing Corporate (per nama sales) otomatis kebaca dari situ.

Jalankan dengan:
    streamlit run app.py
"""

import calendar
import io
import os
import re
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------------------
# Konfigurasi umum
# --------------------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard Omset MFlash", layout="wide", page_icon="📊")

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
os.makedirs(MAIN_DATA_DIR, exist_ok=True)
ADS_DATA_DIR = os.path.join(DATA_DIR, "ads")
os.makedirs(ADS_DATA_DIR, exist_ok=True)
CORPORATE_DATA_PATH = os.path.join(DATA_DIR, "corporate_data.xlsx")
TARGET_DATA_PATH = os.path.join(DATA_DIR, "target_data.xlsx")

MAIN_SHEET_NAME = "Faktur Penjualan"
# Kolom wajib untuk hitung Omset Service vs Gadget & Aksesoris:
#   - Omset Service = baris dengan KATEGORI BARANG "Jasa" atau "Sparepart"
#   - Omset Gadget & Aksesoris = Omset All - Omset Service (semua baris lainnya)
CORE_COLUMNS = ["TGL FAKTUR", "KATEGORI BARANG", "TOTAL HARGA"]
REQUIRED_COLUMNS = ["CABANG"] + CORE_COLUMNS
MAX_MAIN_FILES = 50
MAX_ADS_FILES = 50

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
BULAN_MAP = {name.lower(): i + 1 for i, name in enumerate(BULAN_ID)}
BULAN_ALIAS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
    "jul": 7, "agu": 8, "aug": 8, "sep": 9, "okt": 10, "oct": 10,
    "nov": 11, "des": 12, "dec": 12,
}

# Urutan cabang default, mengikuti urutan di dashboard/scoreboard existing (bukan abjad).
# Cabang yang tidak ada di daftar ini otomatis ditambahkan di akhir (urut abjad).
BRANCH_ORDER = [
    "KLENDER", "CEGER", "BINTARA", "RADJIMAN", "JATIMULYA", "DRAMAGA", "CONDET",
    "JATIBENING", "SAWANGAN", "WARBONG", "CINERE", "CIBINONG", "KARAWANG",
    "JATIWARINGIN", "CIKAMPEK", "CILANGKAP", "PEJATEN", "CIBUBUR",
]
_BRANCH_RANK = {b: i for i, b in enumerate(BRANCH_ORDER)}


def order_branches(branches) -> list:
    """Urutkan daftar cabang sesuai BRANCH_ORDER; cabang baru/tak dikenal ditaruh di akhir (abjad)."""
    branches = list(branches)
    return sorted(branches, key=lambda b: (_BRANCH_RANK.get(b, len(BRANCH_ORDER)), b))


# --------------------------------------------------------------------------------------
# Helper umum
# --------------------------------------------------------------------------------------

def format_rupiah(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        value = 0
    return "Rp " + f"{value:,.0f}".replace(",", ".")


def format_number(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"
    return f"{value:,.0f}".replace(",", ".")


def format_percent(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"
    s = f"{value * 100:,.2f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return s + "%"


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


_ADS_BRANCH_OVERRIDES = {
    "WARUNG BONGKOK": "WARBONG",
    "WARBONG": "WARBONG",
}


def branch_from_campaign_name(campaign_name: str) -> str:
    """Tebak cabang dari nama campaign Meta Ads, contoh: '12 Cinere - MATOT !!!' -> CINERE,
    '10 Warung Bongkok - L - MATOT !!!' -> WARBONG. Campaign tanpa prefix nomor cabang
    (mis. 'SV - LCD 78K - CTWA') dikelompokkan sebagai 'LAINNYA'."""
    if not campaign_name:
        return "LAINNYA"
    name = str(campaign_name).strip()
    m = re.match(r"^\s*\d{1,2}\s+(.+?)\s*-", name)
    if not m:
        return "LAINNYA"
    chunk = re.sub(r"\s+", " ", m.group(1)).strip().upper()
    if chunk in _ADS_BRANCH_OVERRIDES:
        return _ADS_BRANCH_OVERRIDES[chunk]
    for b in BRANCH_ORDER:
        if chunk == b or chunk.startswith(b) or b.startswith(chunk):
            return b
    if "WARUNG" in chunk:
        return "WARBONG"
    return chunk


SERVICE_BARANG_VALUES = {"JASA", "SPAREPART"}


def classify_kategori(kategori_barang: str) -> str:
    """Kelompokkan baris transaksi berdasarkan KATEGORI BARANG:
    - Omset Service = baris dengan KATEGORI BARANG "Jasa" atau "Sparepart"
    - Omset Gadget & Aksesoris = semua baris lainnya (Omset All dikurangi Omset Service)"""
    if kategori_barang and str(kategori_barang).strip().upper() in SERVICE_BARANG_VALUES:
        return "Service"
    return "Gadget & Aksesoris"


def parse_bulan(value):
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


def to_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return pd.to_datetime(value).date()
    except (ValueError, TypeError):
        return None


# --------------------------------------------------------------------------------------
# Loader data utama (transaksi per cabang) - streaming read_only agar cepat & hemat memori
# --------------------------------------------------------------------------------------

def _find_data_sheet(wb):
    """Cari sheet yang punya kolom TGL FAKTUR, KATEGORI BARANG, TOTAL HARGA.
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
        barang = row[col_idx["KATEGORI BARANG"]]
        total = row[col_idx["TOTAL HARGA"]]

        if tgl is None and barang is None and total is None:
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
                "Tanggal": pd.Timestamp(tgl.date()),
                "Tahun": tgl.year,
                "Bulan": tgl.month,
                "KategoriBarang": str(barang).strip().upper() if barang else "",
                "Omset": total,
                "SumberFile": filename_hint,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Kelompok"] = df["KategoriBarang"].apply(classify_kategori)
    return df


def load_all_main_data(main_dir: str):
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
        except Exception as e:  # noqa: BLE001 - tampilkan apa adanya, jangan hentikan file lain
            errors.append(f"{fname}: {e}")
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, errors


# --------------------------------------------------------------------------------------
# Loader data Iklan Meta Ads (export "Campaigns" dari Ads Manager)
# Fokus: Messaging Conversations Started & Cost per Messaging Conversation Started
# --------------------------------------------------------------------------------------

_ADS_REQUIRED_COLS = [
    "CAMPAIGN NAME",
    "AMOUNT SPENT (IDR)",
    "MESSAGING CONVERSATIONS STARTED",
    "COST PER MESSAGING CONVERSATION STARTED (IDR)",
]
_ADS_COLUMNS = [
    "Cabang", "Campaign", "Status", "PeriodeMulai", "PeriodeSelesai", "Spend", "MsgConv",
    "CostPerMsg", "Results", "Impressions", "LinkClicks", "CTR", "CPM", "SumberFile",
]


def _num_or(v, default=0.0):
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _num_or_none(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


@st.cache_data(show_spinner=False)
def load_ads_data(file_bytes: bytes, filename_hint: str = "") -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header_row is None:
        raise ValueError("File kosong / tidak ada header.")

    col_idx = {}
    for i, h in enumerate(header_row):
        if h is None:
            continue
        col_idx[str(h).strip().upper()] = i

    missing = [c for c in _ADS_REQUIRED_COLS if c not in col_idx]
    if missing:
        raise ValueError(
            f"File ini sepertinya bukan export 'Campaigns' dari Meta Ads Manager (kolom {', '.join(missing)} "
            f"tidak ditemukan). Export dari Ads Manager > pilih kolom Messaging Conversation, lalu Export > Excel."
        )

    def g(row, name):
        idx = col_idx.get(name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        campaign = g(row, "CAMPAIGN NAME")
        if not campaign or str(campaign).strip() == "":
            continue
        rows.append(
            {
                "Cabang": branch_from_campaign_name(campaign),
                "Campaign": str(campaign).strip(),
                "Status": str(g(row, "CAMPAIGN DELIVERY") or "-").strip().title(),
                "PeriodeMulai": to_date(g(row, "REPORTING STARTS")),
                "PeriodeSelesai": to_date(g(row, "REPORTING ENDS")),
                "Spend": _num_or(g(row, "AMOUNT SPENT (IDR)")),
                "MsgConv": _num_or(g(row, "MESSAGING CONVERSATIONS STARTED")),
                "CostPerMsg": _num_or_none(g(row, "COST PER MESSAGING CONVERSATION STARTED (IDR)")),
                "Results": _num_or_none(g(row, "RESULTS")),
                "Impressions": _num_or(g(row, "IMPRESSIONS")),
                "LinkClicks": _num_or(g(row, "LINK CLICKS")),
                "CTR": _num_or_none(g(row, "CTR (LINK CLICK-THROUGH RATE)")),
                "CPM": _num_or_none(g(row, "CPM (COST PER 1,000 IMPRESSIONS) (IDR)")),
                "SumberFile": filename_hint,
            }
        )

    return pd.DataFrame(rows, columns=_ADS_COLUMNS)


def load_all_ads_data(ads_dir: str):
    files = sorted(f for f in os.listdir(ads_dir) if f.lower().endswith(".xlsx"))
    frames, errors = [], []
    for fname in files:
        fpath = os.path.join(ads_dir, fname)
        try:
            with open(fpath, "rb") as f:
                file_bytes = f.read()
            df = load_ads_data(file_bytes, filename_hint=fname)
            if not df.empty:
                frames.append(df)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{fname}: {e}")
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_ADS_COLUMNS)
    return combined, errors


def aggregate_ads_by_branch(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Cabang", "Spend", "MsgConv", "CostPerMsg", "Impressions", "LinkClicks", "CTR", "CPM", "JumlahCampaign"])
    g = df.groupby("Cabang").agg(
        Spend=("Spend", "sum"),
        MsgConv=("MsgConv", "sum"),
        Impressions=("Impressions", "sum"),
        LinkClicks=("LinkClicks", "sum"),
        JumlahCampaign=("Campaign", "nunique"),
    ).reset_index()
    g["CostPerMsg"] = g.apply(lambda r: (r["Spend"] / r["MsgConv"]) if r["MsgConv"] else None, axis=1)
    g["CTR"] = g.apply(lambda r: (r["LinkClicks"] / r["Impressions"]) if r["Impressions"] else None, axis=1)
    g["CPM"] = g.apply(lambda r: (r["Spend"] / r["Impressions"] * 1000) if r["Impressions"] else None, axis=1)
    return g.sort_values(by="CostPerMsg", na_position="last")


def generate_ads_insights(df: pd.DataFrame):
    """Rekomendasi otomatis berbasis data: bandingkan tiap campaign terhadap rata-rata
    tertimbang Cost per Messaging Conversation, dan tandai campaign yang boros tanpa hasil."""
    insights = []
    if df.empty:
        return insights, 0.0, 0.0, None

    total_spend = df["Spend"].sum()
    total_msg = df["MsgConv"].sum()
    avg_cost = (total_spend / total_msg) if total_msg else None

    # 1) Spend signifikan tapi 0 Messaging Conversation -> boros, evaluasi/pause
    zero_conv = df[(df["Spend"] >= 5000) & (df["MsgConv"] == 0)].sort_values("Spend", ascending=False)
    for _, r in zero_conv.iterrows():
        insights.append({
            "level": "bad", "title": f"{r['Campaign']} ({r['Cabang']})",
            "text": f"Sudah menghabiskan {format_rupiah(r['Spend'])} tapi belum menghasilkan Messaging "
                    f"Conversation sama sekali. Evaluasi ulang audience/creative, atau pause campaign ini "
                    f"supaya budget tidak terus terbuang.",
        })

    if avg_cost:
        # 2) Cost per messaging jauh di atas rata-rata -> mahal
        high = df[(df["MsgConv"] > 0) & (df["CostPerMsg"] > avg_cost * 1.5)].sort_values("CostPerMsg", ascending=False)
        for _, r in high.iterrows():
            ratio = r["CostPerMsg"] / avg_cost
            insights.append({
                "level": "warn", "title": f"{r['Campaign']} ({r['Cabang']})",
                "text": f"Cost per Messaging Conversation {format_rupiah(r['CostPerMsg'])} — {ratio:.1f}x lebih "
                        f"mahal dari rata-rata semua campaign ({format_rupiah(avg_cost)}). Coba ganti creative/"
                        f"copy, sempitkan targeting, atau turunkan budget harian.",
            })

        # 3) Cost per messaging jauh di bawah rata-rata dengan volume cukup -> efisien, scale up
        low = df[(df["MsgConv"] >= 3) & (df["CostPerMsg"] < avg_cost * 0.7)].sort_values("CostPerMsg")
        for _, r in low.iterrows():
            insights.append({
                "level": "good", "title": f"{r['Campaign']} ({r['Cabang']})",
                "text": f"Paling efisien: Cost per Messaging Conversation hanya {format_rupiah(r['CostPerMsg'])} "
                        f"dari {int(r['MsgConv'])} conversation. Kandidat kuat untuk dinaikkan budgetnya (scale up).",
            })

    # 4) Campaign berstatus Inactive tapi masih tercatat ada spend kecil -> pastikan benar-benar berhenti
    ghost = df[(df["Status"].str.lower() == "inactive") & (df["Spend"] > 0) & (df["Spend"] < 1000)]
    for _, r in ghost.iterrows():
        insights.append({
            "level": "warn", "title": f"{r['Campaign']} ({r['Cabang']})",
            "text": f"Status sudah Inactive tapi masih tercatat sisa spend {format_rupiah(r['Spend'])}. "
                    f"Cek ulang di Ads Manager untuk memastikan campaign benar-benar berhenti menarik budget.",
        })

    return insights, total_spend, total_msg, avg_cost


def render_insight_card(title: str, text: str, level: str) -> str:
    styles = {
        "bad": ("#fee2e2", "#991b1b", "🚨"),
        "warn": ("#fef9c3", "#854d0e", "⚠️"),
        "good": ("#dcfce7", "#166534", "✅"),
    }
    bg, fg, icon = styles.get(level, ("#f3f4f6", "#111827", "ℹ️"))
    return (
        f'<div style="background:{bg};color:{fg};border-radius:10px;padding:10px 14px;margin-bottom:8px;">'
        f'<b>{icon} {title}</b><br><span style="font-size:13px;">{text}</span></div>'
    )


_SALES_TOTAL_LABELS = ("SMM", "TOTAL", "GRAND TOTAL", "HEAD OF CORPORATE")


def generate_sales_insights(board_df: pd.DataFrame, category_label: str, name_label: str = "CABANG"):
    """Rekomendasi otomatis untuk setiap cabang/nama di sebuah scoreboard: tandai penurunan
    rata-rata omset harian (bulan ini vs bulan lalu) dan pencapaian yang masih di bawah 85%."""
    insights = []
    if board_df is None or board_df.empty:
        return insights

    rows = board_df[~board_df[name_label].astype(str).str.upper().isin(_SALES_TOTAL_LABELS)]
    for _, r in rows.iterrows():
        nama = r.get(name_label)
        if nama is None or (isinstance(nama, float) and pd.isna(nama)):
            continue

        gap = r.get("GAP")
        bulan_lalu = r.get("PERIODE BULAN LALU")
        bulan_ini = r.get("PERIODE BULAN INI")
        if gap is not None and not pd.isna(gap) and bulan_lalu is not None and not pd.isna(bulan_lalu) and bulan_lalu > 0:
            pct_turun = gap / bulan_lalu
            if pct_turun <= -0.20:
                insights.append({
                    "level": "bad", "title": f"{category_label} — {nama}",
                    "text": f"Rata-rata omset harian turun {format_percent(abs(pct_turun))} dibanding bulan lalu "
                            f"({format_rupiah(bulan_lalu)} → {format_rupiah(bulan_ini)} per hari). Perlu evaluasi "
                            f"segera: cek ketersediaan stok/teknisi, promo yang sedang berjalan, dan aktivitas "
                            f"kompetitor di sekitar cabang ini.",
                })
            elif pct_turun <= -0.05:
                insights.append({
                    "level": "warn", "title": f"{category_label} — {nama}",
                    "text": f"Rata-rata omset harian turun {format_percent(abs(pct_turun))} dibanding bulan lalu "
                            f"({format_rupiah(bulan_lalu)} → {format_rupiah(bulan_ini)} per hari). Pantau terus, "
                            f"pertimbangkan tambahan promo/aktivasi lokal supaya tidak berlanjut turun.",
                })

        pct = r.get("% PENCAPAIAN")
        if pct is not None and not pd.isna(pct) and pct < 0.85:
            kejar = r.get("KEJAR TARGET PERHARI")
            saran = (
                f" Kejar tambahan {format_rupiah(kejar)}/hari sampai akhir periode untuk bisa capai target."
                if kejar is not None and not pd.isna(kejar) and kejar > 0 else ""
            )
            insights.append({
                "level": "warn" if pct >= 0.7 else "bad",
                "title": f"{category_label} — {nama}",
                "text": f"Pencapaian baru {format_percent(pct)} dari expected value (di bawah target 85%).{saran}",
            })

    return insights


def render_kpi_card_text(label: str, value_text: str, color1: str, color2: str, icon: str) -> str:
    return f"""
    <div style="background:linear-gradient(135deg,{color1},{color2});border-radius:14px;
                padding:16px 18px;color:white;box-shadow:0 2px 8px rgba(0,0,0,.12);height:100%;">
      <div style="font-size:26px;line-height:1;">{icon}</div>
      <div style="font-size:13px;opacity:.9;margin-top:8px;">{label}</div>
      <div style="font-size:22px;font-weight:700;margin-top:2px;">{value_text}</div>
    </div>
    """


# --------------------------------------------------------------------------------------
# Loader data Marketing Corporate manual (Tahun, Bulan, Cabang, Omset) - fallback
# kalau sheet "Scoreboard" tidak ditemukan di file yang diupload
# --------------------------------------------------------------------------------------

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
    for r in [
        [2026, "Januari", "KLENDER", 50000000],
        [2026, "Januari", "RADJIMAN", 12000000],
        [2026, "Februari", "KLENDER", 43000000],
    ]:
        ws.append(r)
    for col, w in zip("ABCD", [10, 14, 16, 16]):
        ws.column_dimensions[col].width = w

    info = wb.create_sheet("Petunjuk")
    info.append(["Petunjuk pengisian Data Marketing Corporate"])
    info.append([""])
    info.append(["- Dipakai HANYA kalau file yang diupload tidak punya sheet 'Scoreboard'."])
    info.append(["- Tahun: angka 4 digit, contoh 2026"])
    info.append(["- Bulan: nama bulan dalam Bahasa Indonesia, contoh Januari, Februari, ... Desember"])
    info.append(["- Cabang: nama cabang, harus konsisten dengan nama cabang di data utama"])
    info.append(["- Omset: angka total omset marketing corporate untuk cabang & bulan tersebut"])
    info.append(["- Satu baris = satu kombinasi Tahun + Bulan + Cabang"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Auto-ekstrak Target & Scoreboard Marketing Corporate dari sheet "Scoreboard"
# (kalau file master diupload apa adanya, tidak perlu input ulang manual)
# --------------------------------------------------------------------------------------

_SECTION_PATTERN = re.compile(
    r"SCOREBOARD\s+OMSET\s+(ALL|SERVICE|GADGET\s*&?\s*AKSESORIS|MARKETING\s+CORPORATE)", re.IGNORECASE
)
_TARGET_DF_COLUMNS = ["Cabang", "PeriodeMulai", "PeriodeSelesai", "TargetService", "TargetGadget", "TargetAll", "TargetCorporate"]
_CORP_COLS = [
    "NAMA", "OMSET SAMURAI", "OMSET HARIAN (DR TARGET)", "EXPECTED VALUE", "HARI INI", "S/D HARI INI",
    "% PENCAPAIAN", "GAP VS EXPECTED", "TOTAL GAP SAMURAI", "KEJAR TARGET PERHARI",
    "PERIODE BULAN LALU", "PERIODE BULAN INI", "GAP",
]


def _empty_target_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_TARGET_DF_COLUMNS)


def _read_scoreboard_sections(rows):
    """Parse mentah semua section 'SCOREBOARD OMSET ...' di sheet Scoreboard.
    Return dict kategori -> {"tanggal", "sisa_hari", "col_idx", "data": {nama: {header: value}}}."""
    sections = {}
    i = 0
    while i < len(rows):
        row = rows[i]
        title_cell = next((str(c) for c in row if isinstance(c, str) and _SECTION_PATTERN.search(c)), None)
        header_row_idx = None
        if title_cell:
            m = _SECTION_PATTERN.search(title_cell)
            kategori = re.sub(r"\s+", " ", m.group(1).upper())
            if "GADGET" in kategori:
                kategori = "GADGET & AKSESORIS"
            elif "MARKETING" in kategori:
                kategori = "MARKETING CORPORATE"

            tanggal = None
            if any(isinstance(c, str) and c.strip().upper() == "TANGGAL" for c in row):
                for c in row:
                    if isinstance(c, (datetime, date)):
                        tanggal = c.date() if isinstance(c, datetime) else c
                        break

            sisa_hari = None
            if i + 1 < len(rows):
                next_row = rows[i + 1]
                if any(isinstance(c, str) and c.strip().upper() == "SISA HARI" for c in next_row):
                    for c in next_row:
                        if isinstance(c, (int, float)) and not isinstance(c, bool):
                            sisa_hari = int(c)
                            break

            col_idx = {}
            for j in range(i + 1, min(i + 8, len(rows))):
                hrow = rows[j]
                if any(isinstance(c, str) and c.strip().upper() == "CABANG" for c in hrow):
                    for k, c in enumerate(hrow):
                        if isinstance(c, str) and c.strip():
                            col_idx.setdefault(c.strip().upper(), k)
                    header_row_idx = j
                    break

            data = {}
            if header_row_idx is not None and "CABANG" in col_idx:
                name_col = col_idx["CABANG"]
                for j in range(header_row_idx + 1, len(rows)):
                    brow = rows[j]
                    if name_col >= len(brow):
                        break
                    nval = brow[name_col]
                    if nval is None or str(nval).strip() == "":
                        break
                    nname = str(nval).strip().upper()
                    row_values = {}
                    for hname, hidx in col_idx.items():
                        if hidx < len(brow):
                            v = brow[hidx]
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                row_values[hname] = float(v)
                    data[nname] = row_values
                    if nname in ("SMM", "TOTAL", "GRAND TOTAL", "HEAD OF CORPORATE"):
                        pass  # baris total tetap disimpan, ditandai lewat is_total saat render

            if tanggal is not None and sisa_hari is not None and data:
                sections[kategori] = {
                    "tanggal": tanggal, "sisa_hari": sisa_hari, "col_idx": col_idx, "data": data,
                }
        i = header_row_idx if header_row_idx is not None else i
        i += 1
    return sections


@st.cache_data(show_spinner=False)
def extract_scoreboard_target(file_bytes: bytes) -> pd.DataFrame:
    """Ambil OMSET SAMURAI (target) & OMSET HARIAN per cabang dari section SERVICE &
    GADGET & AKSESORIS, lalu turunkan periode target dari TANGGAL + SISA HARI."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return _empty_target_df()

    sheet_name = next((n for n in wb.sheetnames if "scoreboard" in n.strip().lower()), None)
    if sheet_name is None:
        return _empty_target_df()

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 500), values_only=True))
    sections = _read_scoreboard_sections(rows)

    svc = sections.get("SERVICE")
    gdg = sections.get("GADGET & AKSESORIS")
    if not svc and not gdg:
        return _empty_target_df()

    def get_vals(section, name):
        row = section["data"].get(name, {}) if section else {}
        return row.get("OMSET SAMURAI"), row.get("OMSET HARIAN (DR TARGET SAMURAI)")

    branch_names = set()
    if svc:
        branch_names |= {n for n in svc["data"] if n not in ("SMM", "TOTAL", "GRAND TOTAL")}
    if gdg:
        branch_names |= {n for n in gdg["data"] if n not in ("SMM", "TOTAL", "GRAND TOTAL")}

    rows_out = []
    for cabang in branch_names:
        svc_target, svc_harian = get_vals(svc, cabang)
        gdg_target, gdg_harian = get_vals(gdg, cabang)

        ref_section = svc if (svc and cabang in svc["data"]) else gdg
        tanggal, sisa = ref_section["tanggal"], ref_section["sisa_hari"]

        total_hari = None
        if svc_target and svc_harian:
            total_hari = round(svc_target / svc_harian)
        elif gdg_target and gdg_harian:
            total_hari = round(gdg_target / gdg_harian)
        if not total_hari:
            continue

        periode_selesai = tanggal + timedelta(days=sisa)
        periode_mulai = periode_selesai - timedelta(days=total_hari - 1)

        rows_out.append(
            {
                "Cabang": cabang,
                "PeriodeMulai": periode_mulai,
                "PeriodeSelesai": periode_selesai,
                "TargetService": svc_target or 0.0,
                "TargetGadget": gdg_target or 0.0,
                "TargetAll": (svc_target or 0.0) + (gdg_target or 0.0),
                "TargetCorporate": 0.0,
            }
        )

    return pd.DataFrame(rows_out, columns=_TARGET_DF_COLUMNS) if rows_out else _empty_target_df()


def extract_scoreboard_target_all(main_dir: str) -> pd.DataFrame:
    frames = []
    for fname in sorted(os.listdir(main_dir)):
        if not fname.lower().endswith(".xlsx"):
            continue
        try:
            with open(os.path.join(main_dir, fname), "rb") as f:
                file_bytes = f.read()
            df = extract_scoreboard_target(file_bytes)
            if not df.empty:
                frames.append(df)
        except Exception:  # noqa: BLE001
            continue
    if not frames:
        return _empty_target_df()
    combined = pd.concat(frames, ignore_index=True)
    return combined.drop_duplicates(subset=["Cabang", "PeriodeMulai", "PeriodeSelesai"], keep="last")


@st.cache_data(show_spinner=False)
def extract_scoreboard_corporate(file_bytes: bytes):
    """Ambil section 'SCOREBOARD OMSET MARKETING CORPORATE' (per nama sales) apa adanya
    (snapshot statis, karena tidak ada data transaksi harian corporate yang diupload)."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return None, None

    sheet_name = next((n for n in wb.sheetnames if "scoreboard" in n.strip().lower()), None)
    if sheet_name is None:
        return None, None

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 500), values_only=True))
    sections = _read_scoreboard_sections(rows)
    corp = sections.get("MARKETING CORPORATE")
    if not corp:
        return None, None

    out_rows = []
    for name, vals in corp["data"].items():
        out_rows.append(
            {
                "CABANG": name,
                "OMSET SAMURAI": vals.get("OMSET SAMURAI"),
                "OMSET HARIAN (DR TARGET)": vals.get("OMSET HARIAN (DR TARGET SAMURAI)"),
                "EXPECTED VALUE": vals.get("EXPECTED VALUE"),
                "HARI INI": None,
                "S/D HARI INI": vals.get("S/D HARI INI"),
                "% PENCAPAIAN": vals.get("% PENCAPAIAN"),
                "GAP VS EXPECTED": vals.get("S/D HARI NI"),
                "TOTAL GAP SAMURAI": vals.get("TOTAL GAP SAMURAI"),
                "KEJAR TARGET PERHARI": vals.get("KEJAR TARGET PERHARI"),
                "PERIODE BULAN LALU": vals.get("PERIODE BULAN LALU"),
                "PERIODE BULAN INI": vals.get("PERIODE BULAN INI"),
                "GAP": vals.get("GAP"),
                "SISA HARI": corp["sisa_hari"],
            }
        )
    if not out_rows:
        return None, None

    df = pd.DataFrame(out_rows)
    # baris total (HEAD OF CORPORATE / SMM / TOTAL) ditaruh paling bawah
    is_total_mask = df["CABANG"].isin(["SMM", "TOTAL", "GRAND TOTAL", "HEAD OF CORPORATE"])
    df = pd.concat([df[~is_total_mask], df[is_total_mask]], ignore_index=True)
    return df, corp["tanggal"]


def extract_scoreboard_corporate_all(main_dir: str):
    for fname in sorted(os.listdir(main_dir)):
        if not fname.lower().endswith(".xlsx"):
            continue
        try:
            with open(os.path.join(main_dir, fname), "rb") as f:
                file_bytes = f.read()
            df, tanggal = extract_scoreboard_corporate(file_bytes)
            if df is not None:
                return df, tanggal
        except Exception:  # noqa: BLE001
            continue
    return None, None


# --------------------------------------------------------------------------------------
# Loader data Target manual (untuk tab Scoreboard) - fallback kalau sheet Scoreboard
# tidak ditemukan
# --------------------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_target_data(file_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    for name in xls.sheet_names:
        if "target" in name.lower():
            sheet_name = name
            break

    raw = pd.read_excel(xls, sheet_name=sheet_name)
    raw.columns = [str(c).strip().upper() for c in raw.columns]

    required = ["CABANG", "PERIODE MULAI", "PERIODE SELESAI", "TARGET SERVICE", "TARGET GADGET & AKSESORIS"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Kolom berikut tidak ditemukan di file Target: {', '.join(missing)}. "
            f"Gunakan template yang tersedia di sidebar."
        )

    df = raw[required].copy()
    df = df.dropna(subset=["CABANG", "PERIODE MULAI", "PERIODE SELESAI"])
    df["Cabang"] = df["CABANG"].astype(str).str.strip().str.upper()
    df["PeriodeMulai"] = df["PERIODE MULAI"].apply(to_date)
    df["PeriodeSelesai"] = df["PERIODE SELESAI"].apply(to_date)
    df["TargetService"] = pd.to_numeric(df["TARGET SERVICE"], errors="coerce").fillna(0)
    df["TargetGadget"] = pd.to_numeric(df["TARGET GADGET & AKSESORIS"], errors="coerce").fillna(0)
    df["TargetAll"] = df["TargetService"] + df["TargetGadget"]
    if "TARGET MARKETING CORPORATE" in raw.columns:
        df["TargetCorporate"] = pd.to_numeric(raw["TARGET MARKETING CORPORATE"], errors="coerce").fillna(0)
    else:
        df["TargetCorporate"] = 0.0
    df = df.dropna(subset=["PeriodeMulai", "PeriodeSelesai"])
    return df[_TARGET_DF_COLUMNS]


def make_target_template() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Target"
    ws.append([
        "Cabang", "Periode Mulai", "Periode Selesai",
        "Target Service", "Target Gadget & Aksesoris", "Target Marketing Corporate",
    ])
    today = date.today()
    for r in [
        ["KLENDER", today.replace(day=1), today.replace(day=1) + timedelta(days=91), 988358000, 515642000, 100000000],
        ["RADJIMAN", today.replace(day=1), today.replace(day=1) + timedelta(days=91), 303000000, 655642350, 30000000],
    ]:
        ws.append(r)
    for cell_row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=3):
        for cell in cell_row:
            cell.number_format = "DD/MM/YYYY"
    for col, w in zip("ABCDEF", [16, 16, 16, 16, 22, 24]):
        ws.column_dimensions[col].width = w

    info = wb.create_sheet("Petunjuk")
    info.append(["Petunjuk pengisian Data Target"])
    info.append([""])
    info.append(["- Dipakai HANYA kalau file yang diupload tidak punya sheet 'Scoreboard'."])
    info.append(["- Cabang: nama cabang, harus konsisten dengan nama cabang di data utama"])
    info.append(["- Periode Mulai / Periode Selesai: tanggal mulai & selesai periode target berjalan"])
    info.append(["  (mis. target 3 bulan berjalan -> Periode Mulai = awal periode, Periode Selesai = akhir periode)"])
    info.append(["- Target Service: target omset Service untuk cabang tsb selama periode itu"])
    info.append(["- Target Gadget & Aksesoris: target omset Gadget & Aksesoris untuk cabang tsb selama periode itu"])
    info.append(["- Target Marketing Corporate: target omset Marketing Corporate untuk cabang tsb selama periode itu (opsional, boleh 0)"])
    info.append(["- Target Omset All dihitung otomatis = Target Service + Target Gadget & Aksesoris"])
    info.append(["- Kalau target berganti tiap periode, tambahkan baris baru untuk periode berikutnya"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Perhitungan Scoreboard (live, dari data transaksi + target)
# --------------------------------------------------------------------------------------

def build_scoreboard(df_kategori: pd.DataFrame, df_target: pd.DataFrame, target_col: str,
                      branches: list, tanggal_acuan: date) -> pd.DataFrame:
    """df_kategori: subset transaksi (Cabang, Tanggal, Omset) untuk 1 kategori.
    target_col: 'TargetAll' / 'TargetService' / 'TargetGadget'."""
    ts_acuan = pd.Timestamp(tanggal_acuan)
    bulan_ini_awal = pd.Timestamp(tanggal_acuan.replace(day=1))
    bulan_lalu_akhir = bulan_ini_awal - timedelta(days=1)
    bulan_lalu_awal = bulan_lalu_akhir.replace(day=1)

    rows = []
    for cabang in branches:
        sub = df_kategori[df_kategori["Cabang"] == cabang]

        trow = None
        if not df_target.empty:
            cand = df_target[
                (df_target["Cabang"] == cabang)
                & (df_target["PeriodeMulai"] <= tanggal_acuan)
                & (df_target["PeriodeSelesai"] >= tanggal_acuan)
            ]
            if not cand.empty:
                trow = cand.iloc[0]

        target = total_hari = hari_berjalan = sisa_hari = omset_harian = expected_value = None
        periode_mulai = periode_selesai = None
        if trow is not None:
            target = float(trow[target_col])
            periode_mulai = trow["PeriodeMulai"]
            periode_selesai = trow["PeriodeSelesai"]
            total_hari = (periode_selesai - periode_mulai).days + 1
            hari_berjalan = max(0, min(total_hari, (tanggal_acuan - periode_mulai).days + 1))
            sisa_hari = total_hari - hari_berjalan
            omset_harian = target / total_hari if total_hari else None
            expected_value = omset_harian * hari_berjalan if omset_harian is not None else None

        hari_ini = sub.loc[sub["Tanggal"] == ts_acuan, "Omset"].sum()
        if periode_mulai is not None:
            sd_hari_ini = sub.loc[
                (sub["Tanggal"] >= pd.Timestamp(periode_mulai)) & (sub["Tanggal"] <= ts_acuan), "Omset"
            ].sum()
        else:
            sd_hari_ini = sub.loc[sub["Tanggal"] <= ts_acuan, "Omset"].sum()

        pct = (sd_hari_ini / expected_value) if expected_value else None
        gap_vs_expected = (sd_hari_ini - expected_value) if expected_value is not None else None
        total_gap_samurai = (target - sd_hari_ini) if target is not None else None
        kejar_perhari = (total_gap_samurai / sisa_hari) if sisa_hari else None

        hari_lalu_terpakai = (bulan_lalu_akhir - bulan_lalu_awal).days + 1
        omset_bulan_lalu = sub.loc[
            (sub["Tanggal"] >= bulan_lalu_awal) & (sub["Tanggal"] <= bulan_lalu_akhir), "Omset"
        ].sum()
        periode_bulan_lalu = omset_bulan_lalu / hari_lalu_terpakai if hari_lalu_terpakai else 0

        hari_ini_terpakai = (ts_acuan - bulan_ini_awal).days + 1
        omset_bulan_ini = sub.loc[
            (sub["Tanggal"] >= bulan_ini_awal) & (sub["Tanggal"] <= ts_acuan), "Omset"
        ].sum()
        periode_bulan_ini = omset_bulan_ini / hari_ini_terpakai if hari_ini_terpakai > 0 else 0

        gap_rata2 = periode_bulan_ini - periode_bulan_lalu

        rows.append(
            {
                "CABANG": cabang,
                "OMSET SAMURAI": target,
                "OMSET HARIAN (DR TARGET)": omset_harian,
                "EXPECTED VALUE": expected_value,
                "HARI INI": hari_ini,
                "S/D HARI INI": sd_hari_ini,
                "% PENCAPAIAN": pct,
                "GAP VS EXPECTED": gap_vs_expected,
                "TOTAL GAP SAMURAI": total_gap_samurai,
                "KEJAR TARGET PERHARI": kejar_perhari,
                "PERIODE BULAN LALU": periode_bulan_lalu,
                "PERIODE BULAN INI": periode_bulan_ini,
                "GAP": gap_rata2,
                "SISA HARI": sisa_hari,
                "PERIODE MULAI": periode_mulai,
                "PERIODE SELESAI": periode_selesai,
            }
        )

    result = pd.DataFrame(rows)
    return _finalize_scoreboard(result)


def _finalize_scoreboard(result: pd.DataFrame) -> pd.DataFrame:
    """Tambahkan baris total (SMM) dengan agregasi yang konsisten untuk semua tabel scoreboard."""
    total = {"CABANG": "SMM"}
    for col in ["OMSET SAMURAI", "EXPECTED VALUE", "HARI INI", "S/D HARI INI",
                "TOTAL GAP SAMURAI", "PERIODE BULAN LALU", "PERIODE BULAN INI"]:
        total[col] = result[col].sum(skipna=True) if result[col].notna().any() else None
    total["OMSET HARIAN (DR TARGET)"] = (
        result["OMSET HARIAN (DR TARGET)"].sum(skipna=True) if result["OMSET HARIAN (DR TARGET)"].notna().any() else None
    )
    total["% PENCAPAIAN"] = (
        total["S/D HARI INI"] / total["EXPECTED VALUE"]
        if total.get("EXPECTED VALUE") else None
    )
    total["GAP VS EXPECTED"] = (
        total["S/D HARI INI"] - total["EXPECTED VALUE"]
        if total.get("EXPECTED VALUE") is not None else None
    )
    sisa_hari_total = result["SISA HARI"].dropna()
    sisa_hari_repr = sisa_hari_total.iloc[0] if not sisa_hari_total.empty else None
    total["KEJAR TARGET PERHARI"] = (
        total["TOTAL GAP SAMURAI"] / sisa_hari_repr
        if total.get("TOTAL GAP SAMURAI") is not None and sisa_hari_repr
        else None
    )
    total["GAP"] = (
        (total["PERIODE BULAN INI"] - total["PERIODE BULAN LALU"])
        if total.get("PERIODE BULAN INI") is not None
        else None
    )
    total["SISA HARI"] = sisa_hari_repr
    if "PERIODE MULAI" in result.columns:
        pm = result["PERIODE MULAI"].dropna()
        ps = result["PERIODE SELESAI"].dropna()
        total["PERIODE MULAI"] = pm.iloc[0] if not pm.empty else None
        total["PERIODE SELESAI"] = ps.iloc[0] if not ps.empty else None

    return pd.concat([result, pd.DataFrame([total])], ignore_index=True)


def build_scoreboard_corporate_manual(df_corp: pd.DataFrame, df_target: pd.DataFrame,
                                       branches: list, tanggal_acuan: date) -> pd.DataFrame:
    """Fallback: Scoreboard Marketing Corporate dari data bulanan manual (Tahun, Bulan, Cabang, Omset),
    dipakai kalau sheet 'Scoreboard' (per nama sales) tidak ditemukan."""
    bulan_ini_awal = pd.Timestamp(tanggal_acuan.replace(day=1))
    bulan_lalu_akhir = bulan_ini_awal - timedelta(days=1)
    ym_acuan = tanggal_acuan.year * 100 + tanggal_acuan.month

    rows = []
    for cabang in branches:
        sub = df_corp[df_corp["Cabang"] == cabang] if not df_corp.empty else df_corp
        sub_ym = (sub["Tahun"] * 100 + sub["Bulan"]) if not sub.empty else pd.Series(dtype="int64")

        trow = None
        if not df_target.empty:
            cand = df_target[
                (df_target["Cabang"] == cabang)
                & (df_target["PeriodeMulai"] <= tanggal_acuan)
                & (df_target["PeriodeSelesai"] >= tanggal_acuan)
            ]
            if not cand.empty:
                trow = cand.iloc[0]

        target = total_hari = hari_berjalan = sisa_hari = omset_harian = expected_value = None
        periode_mulai = None
        if trow is not None:
            target = float(trow["TargetCorporate"])
            periode_mulai = trow["PeriodeMulai"]
            periode_selesai = trow["PeriodeSelesai"]
            total_hari = (periode_selesai - periode_mulai).days + 1
            hari_berjalan = max(0, min(total_hari, (tanggal_acuan - periode_mulai).days + 1))
            sisa_hari = total_hari - hari_berjalan
            omset_harian = target / total_hari if total_hari else None
            expected_value = omset_harian * hari_berjalan if omset_harian is not None else None

        if periode_mulai is not None:
            ym_mulai = periode_mulai.year * 100 + periode_mulai.month
            sd_hari_ini = sub.loc[(sub_ym >= ym_mulai) & (sub_ym <= ym_acuan), "Omset"].sum() if not sub.empty else 0
        else:
            sd_hari_ini = sub.loc[sub_ym <= ym_acuan, "Omset"].sum() if not sub.empty else 0

        pct = (sd_hari_ini / expected_value) if expected_value else None
        gap_vs_expected = (sd_hari_ini - expected_value) if expected_value is not None else None
        total_gap_samurai = (target - sd_hari_ini) if target is not None else None
        kejar_perhari = (total_gap_samurai / sisa_hari) if sisa_hari else None

        hari_lalu_terpakai = calendar.monthrange(bulan_lalu_akhir.year, bulan_lalu_akhir.month)[1]
        omset_bulan_lalu = (
            sub.loc[(sub["Tahun"] == bulan_lalu_akhir.year) & (sub["Bulan"] == bulan_lalu_akhir.month), "Omset"].sum()
            if not sub.empty else 0
        )
        periode_bulan_lalu = omset_bulan_lalu / hari_lalu_terpakai if hari_lalu_terpakai else 0

        hari_ini_terpakai = (pd.Timestamp(tanggal_acuan) - bulan_ini_awal).days + 1
        omset_bulan_ini = (
            sub.loc[(sub["Tahun"] == tanggal_acuan.year) & (sub["Bulan"] == tanggal_acuan.month), "Omset"].sum()
            if not sub.empty else 0
        )
        periode_bulan_ini = omset_bulan_ini / hari_ini_terpakai if hari_ini_terpakai > 0 else 0

        gap_rata2 = periode_bulan_ini - periode_bulan_lalu

        rows.append(
            {
                "CABANG": cabang,
                "OMSET SAMURAI": target,
                "OMSET HARIAN (DR TARGET)": omset_harian,
                "EXPECTED VALUE": expected_value,
                "HARI INI": None,
                "S/D HARI INI": sd_hari_ini,
                "% PENCAPAIAN": pct,
                "GAP VS EXPECTED": gap_vs_expected,
                "TOTAL GAP SAMURAI": total_gap_samurai,
                "KEJAR TARGET PERHARI": kejar_perhari,
                "PERIODE BULAN LALU": periode_bulan_lalu,
                "PERIODE BULAN INI": periode_bulan_ini,
                "GAP": gap_rata2,
                "SISA HARI": sisa_hari,
            }
        )

    return _finalize_scoreboard(pd.DataFrame(rows))


# --------------------------------------------------------------------------------------
# Progress harian (aktual kumulatif vs target pace lurus)
# --------------------------------------------------------------------------------------

def build_daily_progress(df_kategori: pd.DataFrame, df_target: pd.DataFrame, target_col: str,
                          branches: list, tanggal_acuan: date):
    if df_target.empty:
        return None
    cand = df_target[
        df_target["Cabang"].isin(branches)
        & (df_target["PeriodeMulai"] <= tanggal_acuan)
        & (df_target["PeriodeSelesai"] >= tanggal_acuan)
    ]
    if cand.empty:
        return None

    periode_mulai = cand["PeriodeMulai"].iloc[0]
    periode_selesai = cand["PeriodeSelesai"].iloc[0]
    total_target = cand[target_col].sum()
    total_hari = (periode_selesai - periode_mulai).days + 1
    if total_hari <= 0:
        return None
    omset_harian = total_target / total_hari

    sub = df_kategori[df_kategori["Cabang"].isin(branches)]
    full_range = pd.date_range(periode_mulai, periode_selesai)
    actual_range = pd.date_range(periode_mulai, min(tanggal_acuan, periode_selesai))

    daily = sub[(sub["Tanggal"] >= pd.Timestamp(periode_mulai)) & (sub["Tanggal"] <= pd.Timestamp(tanggal_acuan))]
    daily = daily.groupby("Tanggal")["Omset"].sum()
    daily = daily.reindex(actual_range, fill_value=0)
    cum_actual = daily.cumsum()

    target_line = pd.Series([omset_harian * (i + 1) for i in range(len(full_range))], index=full_range)

    df_target_line = pd.DataFrame({"Tanggal": full_range, "Nilai": target_line.values, "Seri": "Target (pace lurus)"})
    df_actual_line = pd.DataFrame({"Tanggal": cum_actual.index, "Nilai": cum_actual.values, "Seri": "Aktual (kumulatif)"})
    chart_df = pd.concat([df_target_line, df_actual_line], ignore_index=True)
    return chart_df


def render_daily_progress_chart(df_kategori, df_target, target_col, branches, tanggal_acuan, color_actual):
    chart_df = build_daily_progress(df_kategori, df_target, target_col, branches, tanggal_acuan)
    if chart_df is None:
        return None
    fig = px.line(
        chart_df, x="Tanggal", y="Nilai", color="Seri",
        color_discrete_map={"Target (pace lurus)": "#9ca3af", "Aktual (kumulatif)": color_actual},
    )
    fig.update_traces(selector=dict(name="Target (pace lurus)"), line=dict(dash="dash"))
    fig.add_vline(x=pd.Timestamp(tanggal_acuan), line_dash="dot", line_color="#ef4444")
    fig.update_layout(
        legend_title_text="", xaxis_title="", yaxis_title="Omset kumulatif (Rp)",
        margin=dict(t=10, b=10, l=10, r=10), height=280,
    )
    return fig


# --------------------------------------------------------------------------------------
# Riwayat / History pencapaian harian - tidak hilang walau tanggal sudah lewat, difilter
# bebas lewat Tanggal/Bulan/Tahun/Cabang (independen dari Tanggal Acuan)
# --------------------------------------------------------------------------------------

def build_daily_history(df_kategori: pd.DataFrame, df_target: pd.DataFrame, target_col: str,
                         branches: list, date_start: date, date_end: date) -> pd.DataFrame:
    """Omset aktual & target per HARI (bukan kumulatif) untuk rentang tanggal bebas,
    supaya histori pencapaian hari-hari sebelumnya tetap bisa dilihat kapan saja."""
    if date_end < date_start:
        date_start, date_end = date_end, date_start
    dates = pd.date_range(date_start, date_end)

    sub = df_kategori[df_kategori["Cabang"].isin(branches)]
    daily_actual = sub[(sub["Tanggal"] >= pd.Timestamp(date_start)) & (sub["Tanggal"] <= pd.Timestamp(date_end))]
    daily_actual = daily_actual.groupby("Tanggal")["Omset"].sum().reindex(dates, fill_value=0.0)

    daily_target = pd.Series(0.0, index=dates)
    has_target = pd.Series(False, index=dates)
    if not df_target.empty:
        cand = df_target[df_target["Cabang"].isin(branches)]
        for _, trow in cand.iterrows():
            pm, ps = trow["PeriodeMulai"], trow["PeriodeSelesai"]
            total_hari = (ps - pm).days + 1
            if total_hari <= 0 or pd.isna(trow[target_col]):
                continue
            daily_rate = trow[target_col] / total_hari
            mask = [(d.date() >= pm) and (d.date() <= ps) for d in dates]
            mask = pd.Series(mask, index=dates)
            daily_target.loc[mask] += daily_rate
            has_target.loc[mask] = True

    result = pd.DataFrame({
        "Tanggal": dates,
        "OmsetAktual": daily_actual.values,
        "OmsetTarget": daily_target.values,
        "AdaTarget": has_target.values,
    })
    result["PctPencapaian"] = result.apply(
        lambda r: (r["OmsetAktual"] / r["OmsetTarget"]) if r["AdaTarget"] and r["OmsetTarget"] > 0 else None, axis=1
    )
    return result


def render_daily_history_chart(hist_df: pd.DataFrame):
    if hist_df is None or hist_df.empty:
        return None
    colors = [pencapaian_color(p) for p in hist_df["PctPencapaian"]]
    labels = [format_rupiah(v) for v in hist_df["OmsetAktual"]]
    hover = [
        f"{d.strftime('%d/%m/%Y')}<br>Omset: {format_rupiah(a)}<br>Target harian: {format_rupiah(t) if ht else '-'}"
        f"<br>% Pencapaian: {format_percent(p) if p is not None else '-'}"
        for d, a, t, ht, p in zip(hist_df["Tanggal"], hist_df["OmsetAktual"], hist_df["OmsetTarget"], hist_df["AdaTarget"], hist_df["PctPencapaian"])
    ]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=hist_df["Tanggal"], y=hist_df["OmsetAktual"], marker_color=colors,
        text=labels, textposition="outside", textfont=dict(size=9),
        hovertext=hover, hoverinfo="text", name="Omset Aktual per Hari",
    ))
    if hist_df["AdaTarget"].any():
        fig.add_trace(go.Scatter(
            x=hist_df["Tanggal"], y=hist_df["OmsetTarget"].where(hist_df["AdaTarget"]),
            mode="lines", line=dict(color="#6b7280", dash="dash"), name="Target Harian",
        ))
    fig.update_layout(
        legend_title_text="", xaxis_title="", yaxis_title="Omset per hari (Rp)",
        margin=dict(t=20, b=10, l=10, r=10), height=340, showlegend=True,
    )
    return fig


# --------------------------------------------------------------------------------------
# Progress ring (donut) untuk % pencapaian
# --------------------------------------------------------------------------------------

def pencapaian_color(pct) -> str:
    """Standar warna pencapaian dipakai di ring, badge KPI, dan tabel scoreboard:
    >=100% hijau, 85%-99,9% kuning, <85% merah."""
    if pct is None or (isinstance(pct, float) and pd.isna(pct)):
        return "#9ca3af"
    if pct >= 1.0:
        return "#22c55e"
    if pct >= 0.85:
        return "#facc15"
    return "#ef4444"


def render_progress_ring(pct, track_color="#e5e7eb"):
    p = 0 if pct is None else max(0.0, min(pct, 1.0))
    color = pencapaian_color(pct)
    fig = go.Figure(
        data=[
            go.Pie(
                values=[p, 1 - p],
                hole=0.78,
                marker=dict(colors=[color, track_color], line=dict(width=0)),
                textinfo="none",
                sort=False,
                direction="clockwise",
                rotation=0,
                showlegend=False,
                hoverinfo="skip",
            )
        ]
    )
    label = format_percent(pct) if pct is not None else "-"
    fig.update_layout(
        annotations=[dict(text=f"<b>{label}</b>", x=0.5, y=0.5, font=dict(size=26, color="#111827"), showarrow=False)],
        margin=dict(t=6, b=6, l=6, r=6),
        height=200,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


MONEY_COLS = [
    "OMSET SAMURAI", "OMSET HARIAN (DR TARGET)", "EXPECTED VALUE", "HARI INI", "S/D HARI INI",
    "GAP VS EXPECTED", "TOTAL GAP SAMURAI", "KEJAR TARGET PERHARI",
    "PERIODE BULAN LALU", "PERIODE BULAN INI", "GAP",
]

# Grup kolom & warna banner, meniru gaya scoreboard existing (abu/hijau/oranye/biru)
_SCOREBOARD_GROUPS = [
    ("DETAIL TARGET", ["OMSET SAMURAI", "OMSET HARIAN (DR TARGET)", "EXPECTED VALUE"], "#6b7280"),
    ("DETAIL PENCAPAIAN", ["HARI INI", "S/D HARI INI", "% PENCAPAIAN"], "#2e7d32"),
    ("DETAIL GAP", ["GAP VS EXPECTED", "TOTAL GAP SAMURAI", "KEJAR TARGET PERHARI"], "#e65100"),
    ("MONITORING PROGRESS RATA-RATA / HARI", ["PERIODE BULAN LALU", "PERIODE BULAN INI", "GAP"], "#1565c0"),
]


def _cell_color(col: str, v) -> str:
    if pd.isna(v):
        return ""
    if col == "% PENCAPAIAN":
        if v >= 1:
            return "background-color:#c6efce;color:#006100"
        if v >= 0.85:
            return "background-color:#ffeb9c;color:#9c6500"
        return "background-color:#ffc7ce;color:#9c0006"
    if col in ("GAP VS EXPECTED", "GAP"):
        return "background-color:#c6efce;color:#006100" if v >= 0 else "background-color:#ffc7ce;color:#9c0006"
    if col == "TOTAL GAP SAMURAI":
        return "background-color:#ffc7ce;color:#9c0006" if v > 0 else "background-color:#c6efce;color:#006100"
    return ""


def render_scoreboard_html(title: str, df: pd.DataFrame, accent: str = "#1f2937", name_label: str = "CABANG") -> str:
    cols = [name_label] + [c for _, group_cols, _ in _SCOREBOARD_GROUPS for c in group_cols]

    header_group_html = f'<th rowspan="2" style="background:{accent};color:white;padding:8px;text-align:left;">{name_label}</th>'
    for gname, gcols, gcolor in _SCOREBOARD_GROUPS:
        header_group_html += (
            f'<th colspan="{len(gcols)}" style="background:{gcolor};color:white;padding:6px;'
            f'text-align:center;font-size:12px;">{gname}</th>'
        )

    header_col_html = ""
    for _, gcols, gcolor in _SCOREBOARD_GROUPS:
        for c in gcols:
            header_col_html += (
                f'<th style="background:{gcolor}22;color:#111;padding:6px 8px;text-align:right;'
                f'font-size:11px;border:1px solid #e5e7eb;white-space:nowrap;">{c}</th>'
            )

    body_html = ""
    n = len(df)
    for idx, (_, row) in enumerate(df.iterrows()):
        is_total = idx == n - 1
        row_bg = "#eef2f7" if is_total else "#ffffff"
        fw = "700" if is_total else "400"
        body_html += f'<tr style="background:{row_bg};font-weight:{fw};">'
        body_html += f'<td style="padding:5px 8px;border:1px solid #e5e7eb;">{row["CABANG"]}</td>'
        for _, gcols, _ in _SCOREBOARD_GROUPS:
            for c in gcols:
                v = row[c]
                txt = format_percent(v) if c == "% PENCAPAIAN" else format_number(v)
                color = _cell_color(c, v)
                body_html += (
                    f'<td style="padding:5px 8px;border:1px solid #e5e7eb;text-align:right;{color}">{txt}</td>'
                )
        body_html += "</tr>"

    return f"""
    <div style="margin-bottom:28px;overflow-x:auto;">
      <table style="border-collapse:collapse;width:100%;font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-size:13px;">
        <tr>
          <td colspan="{len(cols)}" style="background:{accent};color:white;padding:10px 14px;
              font-size:17px;font-weight:700;letter-spacing:.3px;">{title}</td>
        </tr>
        <tr>{header_group_html}</tr>
        <tr>{header_col_html}</tr>
        {body_html}
      </table>
    </div>
    """


def render_kpi_card(label: str, value: float, color1: str, color2: str, icon: str, pct=None) -> str:
    badge = ""
    if pct is not None:
        badge_bg = "rgba(198,239,206,.9)" if pct >= 1 else ("rgba(255,235,156,.9)" if pct >= 0.85 else "rgba(255,199,206,.9)")
        badge_fg = "#006100" if pct >= 1 else ("#9c6500" if pct >= 0.85 else "#9c0006")
        badge = (
            f'<div style="display:inline-block;background:{badge_bg};color:{badge_fg};'
            f'border-radius:999px;padding:2px 10px;font-size:12px;font-weight:700;margin-top:6px;">'
            f'{format_percent(pct)} dari target</div>'
        )
    return f"""
    <div style="background:linear-gradient(135deg,{color1},{color2});border-radius:14px;
                padding:16px 18px;color:white;box-shadow:0 2px 8px rgba(0,0,0,.12);height:100%;">
      <div style="font-size:26px;line-height:1;">{icon}</div>
      <div style="font-size:13px;opacity:.9;margin-top:8px;">{label}</div>
      <div style="font-size:22px;font-weight:700;margin-top:2px;">{format_rupiah(value)}</div>
      {badge}
    </div>
    """


# --------------------------------------------------------------------------------------
# Sidebar: upload data
# --------------------------------------------------------------------------------------

st.sidebar.title("📥 Data")

st.sidebar.caption(
    "Upload file data cabang (bisa file per-cabang, atau langsung file master yang punya sheet "
    "'Faktur Penjualan' + 'Scoreboard' — target & scoreboard corporate otomatis kebaca dari situ). "
    "Bisa pilih banyak file sekaligus, sampai 50 file."
)
main_uploads = st.sidebar.file_uploader(
    "Upload file data cabang", type=["xlsx"], accept_multiple_files=True, key="main_uploads"
)
if main_uploads:
    existing = len(os.listdir(MAIN_DATA_DIR))
    if existing + len(main_uploads) > MAX_MAIN_FILES:
        st.sidebar.error(
            f"Maksimal {MAX_MAIN_FILES} file. Saat ini sudah ada {existing} file tersimpan, "
            f"hapus beberapa dulu di bawah sebelum upload {len(main_uploads)} file baru."
        )
    else:
        saved, failed = 0, []
        for uf in main_uploads:
            safe_name = sanitize_filename(uf.name)
            try:
                with open(os.path.join(MAIN_DATA_DIR, safe_name), "wb") as f:
                    f.write(uf.getbuffer())
                saved += 1
            except Exception as e:  # noqa: BLE001
                failed.append(f"{uf.name}: {e}")
        if saved:
            st.sidebar.success(f"{saved} file tersimpan.")
        if failed:
            st.sidebar.error("Gagal menyimpan:\n" + "\n".join(failed))

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
st.sidebar.caption("Opsional — hanya perlu diisi kalau file di atas TIDAK punya sheet 'Scoreboard':")

corp_upload = st.sidebar.file_uploader("Upload file Data Marketing Corporate (manual)", type=["xlsx"], key="corp_upload")
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

target_upload = st.sidebar.file_uploader("Upload file Data Target (manual)", type=["xlsx"], key="target_upload")
if target_upload is not None:
    with open(TARGET_DATA_PATH, "wb") as f:
        f.write(target_upload.getbuffer())
    st.sidebar.success("File Target tersimpan.")

st.sidebar.download_button(
    "⬇️ Download template Target",
    data=make_target_template(),
    file_name="template_target.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.sidebar.markdown("---")
st.sidebar.title("📣 Data Iklan Meta Ads")
st.sidebar.caption(
    "Upload hasil export 'Campaigns' dari Meta Ads Manager (Export > Excel). Bisa upload beberapa "
    "file sekaligus (mis. per periode), sampai 50 file. Dipakai di tab 'Iklan (Meta Ads)'."
)
ads_uploads = st.sidebar.file_uploader(
    "Upload file export Meta Ads", type=["xlsx"], accept_multiple_files=True, key="ads_uploads"
)
if ads_uploads:
    existing_ads = len(os.listdir(ADS_DATA_DIR))
    if existing_ads + len(ads_uploads) > MAX_ADS_FILES:
        st.sidebar.error(
            f"Maksimal {MAX_ADS_FILES} file. Saat ini sudah ada {existing_ads} file tersimpan, "
            f"hapus beberapa dulu di bawah sebelum upload {len(ads_uploads)} file baru."
        )
    else:
        saved, failed = 0, []
        for uf in ads_uploads:
            safe_name = sanitize_filename(uf.name)
            try:
                with open(os.path.join(ADS_DATA_DIR, safe_name), "wb") as f:
                    f.write(uf.getbuffer())
                saved += 1
            except Exception as e:  # noqa: BLE001
                failed.append(f"{uf.name}: {e}")
        if saved:
            st.sidebar.success(f"{saved} file iklan tersimpan.")
        if failed:
            st.sidebar.error("Gagal menyimpan:\n" + "\n".join(failed))

stored_ads_files = sorted(os.listdir(ADS_DATA_DIR))
if stored_ads_files:
    with st.sidebar.expander(f"📁 File iklan tersimpan ({len(stored_ads_files)}/{MAX_ADS_FILES})"):
        for fname in stored_ads_files:
            fc1, fc2 = st.columns([4, 1])
            fc1.write(fname)
            if fc2.button("🗑️", key=f"del_ads_{fname}"):
                os.remove(os.path.join(ADS_DATA_DIR, fname))
                st.rerun()
        if st.button("Hapus semua file iklan", key="del_all_ads"):
            for fname in stored_ads_files:
                os.remove(os.path.join(ADS_DATA_DIR, fname))
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(
    "File yang sudah diupload akan otomatis dipakai lagi setiap kali dashboard dibuka, "
    "sampai diganti/dihapus. Upload ulang dengan nama file sama = update (bukan duplikat)."
)

# --------------------------------------------------------------------------------------
# Load data
# --------------------------------------------------------------------------------------

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
logo_col, title_col = st.columns([1, 6])
with logo_col:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=110)
with title_col:
    st.title("📊 Dashboard Omset MFlash")
    st.caption("Apapun Gadgetnya, MFlash Solusinya")

if not os.listdir(MAIN_DATA_DIR):
    st.info("Silakan upload file data cabang (bisa lebih dari satu) lewat sidebar untuk mulai.")
    st.stop()

with st.spinner(f"Memproses {len(os.listdir(MAIN_DATA_DIR))} file data cabang..."):
    df_main, load_errors = load_all_main_data(MAIN_DATA_DIR)

if load_errors:
    st.warning("Sebagian file dilewati karena formatnya tidak cocok:\n\n" + "\n".join(f"- {e}" for e in load_errors))

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

df_target = _empty_target_df()
target_source = None
if os.path.exists(TARGET_DATA_PATH):
    with open(TARGET_DATA_PATH, "rb") as f:
        target_bytes = f.read()
    try:
        df_target = load_target_data(target_bytes)
        if not df_target.empty:
            target_source = "manual"
    except ValueError as e:
        st.warning(f"Data Target tidak dipakai: {e}")

if df_target.empty:
    df_target_auto = extract_scoreboard_target_all(MAIN_DATA_DIR)
    if not df_target_auto.empty:
        df_target = df_target_auto
        target_source = "auto"

corp_scoreboard_df, corp_scoreboard_tanggal = extract_scoreboard_corporate_all(MAIN_DATA_DIR)

# --------------------------------------------------------------------------------------
# Filter (dipakai di tab Ringkasan)
# --------------------------------------------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.title("🔎 Filter Ringkasan")

tahun_options = sorted(set(df_main["Tahun"]).union(set(df_corp["Tahun"])), reverse=True)
bulan_options = list(range(1, 13))
cabang_options = order_branches(set(df_main["Cabang"]).union(set(df_corp["Cabang"])).union(set(df_target["Cabang"])))

sel_tahun = st.sidebar.multiselect("Tahun", tahun_options, default=tahun_options)
sel_bulan = st.sidebar.multiselect("Bulan", bulan_options, default=bulan_options, format_func=lambda m: BULAN_ID[m - 1])
sel_cabang = st.sidebar.multiselect("Cabang", cabang_options, default=cabang_options)

st.sidebar.markdown("---")
st.sidebar.title("🎯 Tanggal Acuan (Scoreboard)")
default_tanggal_acuan = df_main["Tanggal"].max().date() if not df_main.empty else date.today()
tanggal_acuan = st.sidebar.date_input(
    "Dianggap sebagai 'Hari Ini'", value=default_tanggal_acuan,
    help="Default = tanggal transaksi terakhir di data yang diupload, supaya kolom 'Hari Ini' tidak kosong.",
)

if not sel_tahun or not sel_bulan or not sel_cabang:
    st.warning("Pilih minimal satu Tahun, Bulan, dan Cabang di sidebar.")
    st.stop()

f_main = df_main[df_main["Tahun"].isin(sel_tahun) & df_main["Bulan"].isin(sel_bulan) & df_main["Cabang"].isin(sel_cabang)]
f_corp = df_corp[df_corp["Tahun"].isin(sel_tahun) & df_corp["Bulan"].isin(sel_bulan) & df_corp["Cabang"].isin(sel_cabang)]

# --------------------------------------------------------------------------------------
# Hitung scoreboard sekali (dipakai di tab Ringkasan buat ring/badge dan di tab Scoreboard)
# --------------------------------------------------------------------------------------

branches_sb = sel_cabang if sel_cabang else cabang_options

kategori_map = [
    ("SCOREBOARD OMSET ALL", df_main, "TargetAll", "#111827"),
    ("SCOREBOARD OMSET SERVICE", df_main[df_main["Kelompok"] == "Service"], "TargetService", "#0f766e"),
    ("SCOREBOARD OMSET GADGET & AKSESORIS", df_main[df_main["Kelompok"] == "Gadget & Aksesoris"], "TargetGadget", "#6d28d9"),
]
boards = {}
for title, subset, target_col, accent in kategori_map:
    boards[title] = (build_scoreboard(subset, df_target, target_col, branches_sb, tanggal_acuan), accent, subset, target_col)

if corp_scoreboard_df is not None:
    board_corp = corp_scoreboard_df
    corp_is_auto = True
else:
    board_corp = build_scoreboard_corporate_manual(df_corp, df_target, branches_sb, tanggal_acuan)
    corp_is_auto = False


def _smm_row(board: pd.DataFrame):
    smm = board[board["CABANG"].isin(["SMM", "TOTAL", "GRAND TOTAL", "HEAD OF CORPORATE"])]
    if smm.empty and len(board) > 0:
        smm = board.iloc[[-1]]
    return smm.iloc[0] if not smm.empty else None


def _pct_and_target_ratio(board: pd.DataFrame):
    row = _smm_row(board)
    if row is None:
        return None, None
    pct = row.get("% PENCAPAIAN")
    pct = None if pd.isna(pct) else pct
    sd = row.get("S/D HARI INI")
    target = row.get("OMSET SAMURAI")
    ratio_target = None
    if target and not pd.isna(target) and sd is not None and not pd.isna(sd):
        ratio_target = sd / target
    return pct, ratio_target


pct_all, ratio_all = _pct_and_target_ratio(boards["SCOREBOARD OMSET ALL"][0])
pct_service, ratio_service = _pct_and_target_ratio(boards["SCOREBOARD OMSET SERVICE"][0])
pct_gadget, ratio_gadget = _pct_and_target_ratio(boards["SCOREBOARD OMSET GADGET & AKSESORIS"][0])
pct_corp, ratio_corp = _pct_and_target_ratio(board_corp)

tab_ringkasan, tab_scoreboard, tab_ads = st.tabs(["📈 Ringkasan", "🏆 Scoreboard", "📣 Iklan (Meta Ads)"])

# --------------------------------------------------------------------------------------
# TAB 1: Ringkasan
# --------------------------------------------------------------------------------------

with tab_ringkasan:
    omset_all = f_main["Omset"].sum()
    omset_service = f_main.loc[f_main["Kelompok"] == "Service", "Omset"].sum()
    omset_gadget = f_main.loc[f_main["Kelompok"] == "Gadget & Aksesoris", "Omset"].sum()
    omset_corporate = f_corp["Omset"].sum() if not f_corp.empty else 0.0
    if corp_is_auto:
        crow = _smm_row(board_corp)
        if crow is not None and crow.get("S/D HARI INI") is not None and not pd.isna(crow.get("S/D HARI INI")):
            omset_corporate = crow["S/D HARI INI"]

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_kpi_card("Omset All", omset_all, "#111827", "#374151", "📊", pct_all), unsafe_allow_html=True)
    c2.markdown(render_kpi_card("Omset Service", omset_service, "#0f766e", "#14b8a6", "🛠️", pct_service), unsafe_allow_html=True)
    c3.markdown(render_kpi_card("Omset Gadget & Aksesoris", omset_gadget, "#6d28d9", "#a78bfa", "📱", pct_gadget), unsafe_allow_html=True)
    c4.markdown(render_kpi_card("Omset Marketing Corporate", omset_corporate, "#b45309", "#f59e0b", "🤝", pct_corp), unsafe_allow_html=True)

    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    st.subheader("🎯 Progress Pencapaian")
    st.caption(
        f"Posisi per tanggal acuan **{tanggal_acuan.strftime('%d/%m/%Y')}** (ubah di sidebar). "
        "Ring = % pencapaian terhadap ekspektasi pace saat ini (🟢 ≥100% · 🟡 85-99,9% · 🔴 <85%). "
        "Teks di bawah = % dari total target periode."
    )
    r1, r2, r3 = st.columns(3)
    for col, title, pct, ratio in [
        (r1, "Omset All", pct_all, ratio_all),
        (r2, "Omset Service", pct_service, ratio_service),
        (r3, "Omset Gadget & Aksesoris", pct_gadget, ratio_gadget),
    ]:
        with col:
            st.markdown(f"<div style='text-align:center;font-weight:700;'>{title}</div>", unsafe_allow_html=True)
            st.plotly_chart(render_progress_ring(pct), width="stretch", config={"displayModeBar": False})
            ratio_txt = format_percent(ratio) if ratio is not None else "-"
            st.markdown(
                f"<div style='text-align:center;color:#374151;'>Anda telah mencapai "
                f"<b>{ratio_txt}</b> dari keseluruhan target periode ini</div>",
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.subheader("Tren Omset per Bulan")

    trend_all = f_main.groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index().assign(Kategori="Omset All")
    trend_service = (
        f_main[f_main["Kelompok"] == "Service"].groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index()
        .assign(Kategori="Omset Service")
    )
    trend_gadget = (
        f_main[f_main["Kelompok"] == "Gadget & Aksesoris"].groupby(["Tahun", "Bulan"])["Omset"].sum().reset_index()
        .assign(Kategori="Omset Gadget & Aksesoris")
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

    st.subheader("Omset per Cabang")

    by_branch_all = f_main.groupby("Cabang")["Omset"].sum().rename("Omset All")
    by_branch_service = f_main[f_main["Kelompok"] == "Service"].groupby("Cabang")["Omset"].sum().rename("Omset Service")
    by_branch_gadget = f_main[f_main["Kelompok"] == "Gadget & Aksesoris"].groupby("Cabang")["Omset"].sum().rename("Omset Gadget & Aksesoris")
    by_branch_corp = f_corp.groupby("Cabang")["Omset"].sum().rename("Omset Marketing Corporate")

    by_branch = pd.concat(
        [by_branch_all, by_branch_service, by_branch_gadget, by_branch_corp], axis=1
    ).astype(float).fillna(0.0).reset_index()
    branch_order_local = order_branches(by_branch["Cabang"])
    by_branch["Cabang"] = pd.Categorical(by_branch["Cabang"], categories=branch_order_local, ordered=True)
    by_branch = by_branch.sort_values("Cabang").reset_index(drop=True)

    bcol1, bcol2 = st.columns([2, 1])
    with bcol1:
        branch_long = by_branch.melt(id_vars="Cabang", var_name="Kategori", value_name="Omset")
        branch_long["Label"] = branch_long["Omset"].apply(lambda v: format_rupiah(v) if v else "")
        fig_branch = px.bar(
            branch_long, x="Cabang", y="Omset", color="Kategori", barmode="group",
            category_orders={"Cabang": branch_order_local}, text="Label",
        )
        fig_branch.update_traces(textposition="outside", textfont_size=9, textangle=-90, cliponaxis=False)
        fig_branch.update_layout(
            legend_title_text="", xaxis_title="", yaxis_title="Omset (Rp)",
            margin=dict(t=40), height=520,
        )
        st.plotly_chart(fig_branch, width="stretch")
    with bcol2:
        num_cols = list(by_branch.columns[1:])
        try:
            sty = (
                by_branch.style
                .background_gradient(cmap="RdYlGn", subset=num_cols, axis=0)
                .format({c: format_rupiah for c in num_cols})
            )
            st.dataframe(sty, width="stretch", hide_index=True)
        except ImportError:
            display_df = by_branch.copy()
            for col in num_cols:
                display_df[col] = display_df[col].apply(format_rupiah)
            st.dataframe(display_df, width="stretch", hide_index=True)

    st.caption(
        "Omset All & Omset Service & Omset Gadget dihitung dari data cabang yang diupload. "
        "Omset Marketing Corporate dihitung dari sheet Scoreboard (kalau ada) atau file manual (lihat sidebar)."
    )

# --------------------------------------------------------------------------------------
# TAB 2: Scoreboard
# --------------------------------------------------------------------------------------

with tab_scoreboard:
    sisa_hari_caption = ""
    if not df_target.empty:
        cand = df_target[(df_target["PeriodeMulai"] <= tanggal_acuan) & (df_target["PeriodeSelesai"] >= tanggal_acuan)]
        if not cand.empty:
            sisa = (cand.iloc[0]["PeriodeSelesai"] - tanggal_acuan).days
            sisa_hari_caption = f" · Sisa hari periode: **{sisa}**"
    st.caption(
        f"📅 Tanggal acuan: **{tanggal_acuan.strftime('%d/%m/%Y')}**{sisa_hari_caption}. "
        "Ubah di sidebar kalau mau lihat posisi di tanggal lain."
    )

    st.markdown("### 📅 Riwayat Pencapaian Harian")
    st.caption(
        "History ini tidak hilang walau harinya sudah lewat — pilih sendiri rentang tanggal / bulan / "
        "tahun / cabang untuk lihat pencapaian di hari mana pun. Warna bar: 🟢 ≥100% · 🟡 85-99,9% · 🔴 <85%."
    )
    hist_kategori_options = ["Omset All", "Omset Service", "Omset Gadget & Aksesoris"]
    _hist_map = {
        "Omset All": boards["SCOREBOARD OMSET ALL"],
        "Omset Service": boards["SCOREBOARD OMSET SERVICE"],
        "Omset Gadget & Aksesoris": boards["SCOREBOARD OMSET GADGET & AKSESORIS"],
    }
    h1, h2, h3, h4 = st.columns([1.3, 1, 1, 1.7])
    with h1:
        hist_kategori = st.selectbox("Kategori", hist_kategori_options, key="hist_kategori")
    tahun_hist_options = sorted(df_main["Tahun"].unique(), reverse=True) or [date.today().year]
    with h2:
        hist_tahun = st.selectbox("Tahun", tahun_hist_options, key="hist_tahun")
    with h3:
        hist_bulan = st.selectbox("Bulan", list(range(1, 13)), index=date.today().month - 1, format_func=lambda m: BULAN_ID[m - 1], key="hist_bulan")
    with h4:
        hist_cabang = st.multiselect("Cabang", cabang_options, default=branches_sb, key="hist_cabang")

    _last_day = calendar.monthrange(hist_tahun, hist_bulan)[1]
    _default_start = date(hist_tahun, hist_bulan, 1)
    _default_end = date(hist_tahun, hist_bulan, _last_day)
    hist_range = st.date_input(
        "Rentang tanggal (bisa disesuaikan manual)", value=(_default_start, _default_end),
        key="hist_range",
    )
    if isinstance(hist_range, tuple) and len(hist_range) == 2:
        hist_start, hist_end = hist_range
    else:
        hist_start, hist_end = _default_start, _default_end

    if not hist_cabang:
        st.info("Pilih minimal satu cabang untuk lihat riwayat pencapaian harian.")
    else:
        _, _, hist_subset, hist_target_col = _hist_map[hist_kategori]
        hist_df = build_daily_history(hist_subset, df_target, hist_target_col, hist_cabang, hist_start, hist_end)
        fig_hist = render_daily_history_chart(hist_df)
        if fig_hist is not None:
            st.plotly_chart(fig_hist, width="stretch")
            if not hist_df["AdaTarget"].any():
                st.caption("Belum ada data Target untuk periode ini, jadi bar tetap ditampilkan abu-abu (tidak bisa hitung % pencapaian).")
        else:
            st.info("Tidak ada data untuk kombinasi filter ini.")

    st.markdown("---")

    if df_target.empty:
        st.info(
            "Belum ada data Target. Kalau file yang diupload punya sheet **Scoreboard** (seperti file "
            "master yang biasa dipakai), targetnya otomatis kebaca — coba upload ulang file itu lewat "
            "'Upload file data cabang'. Atau, download template Target di sidebar, isi manual, lalu upload. "
            "Tanpa data ini, kolom Target / Expected Value / % Pencapaian / Gap / Kejar Target tidak bisa "
            "dihitung — tabel di bawah tetap tampil dengan kolom lain saja."
        )
    elif target_source == "auto":
        st.success("🎯 Target & Expected Value otomatis terbaca dari sheet **Scoreboard** di file yang diupload.")
    else:
        st.success("🎯 Target & Expected Value dihitung dari file Data Target yang diupload.")

    for title, (board, accent, subset, target_col) in boards.items():
        st.markdown(render_scoreboard_html(title, board, accent), unsafe_allow_html=True)
        with st.expander(f"📈 Progress harian — {title}"):
            fig = render_daily_progress_chart(subset, df_target, target_col, branches_sb, tanggal_acuan, accent)
            if fig is not None:
                st.plotly_chart(fig, width="stretch")
            else:
                st.info("Belum ada data Target untuk periode ini, grafik progress harian tidak bisa ditampilkan.")

    if corp_is_auto:
        tgl_txt = corp_scoreboard_tanggal.strftime("%d/%m/%Y") if corp_scoreboard_tanggal else "-"
        st.caption(
            f"ℹ️ Tabel di bawah adalah snapshot **per nama sales**, apa adanya dari sheet Scoreboard "
            f"(per tanggal **{tgl_txt}** di file yang diupload) — tidak berubah walau Tanggal Acuan diganti, "
            "karena tidak ada data transaksi harian corporate yang diupload."
        )
        st.markdown(
            render_scoreboard_html("SCOREBOARD OMSET MARKETING CORPORATE (per Sales)", board_corp, "#b45309", name_label="NAMA SALES"),
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            render_scoreboard_html("SCOREBOARD OMSET MARKETING CORPORATE", board_corp, "#b45309"),
            unsafe_allow_html=True,
        )
        st.caption(
            "Data Marketing Corporate bersifat bulanan (bukan harian), jadi kolom **HARI INI** tidak tersedia "
            "(ditampilkan \"-\"). Upload file dengan sheet **Scoreboard** untuk otomatis dapat scoreboard "
            "per nama sales."
        )

    st.markdown(
        "**Cara baca warna:** hijau = sudah di atas ekspektasi / target tercapai (≥100%), "
        "kuning = mendekati (85-99,9% dari ekspektasi), merah = di bawah ekspektasi / masih ada gap (<85%)."
    )

    st.markdown("---")
    st.subheader("💡 Insight & Rekomendasi Perbaikan")
    st.caption(
        "Dihitung otomatis dari data Service, Gadget & Aksesoris (Penjualan), dan Marketing Corporate: "
        "membandingkan rata-rata omset harian bulan ini vs bulan lalu, serta % Pencapaian terhadap Expected Value."
    )
    sales_insights = (
        generate_sales_insights(boards["SCOREBOARD OMSET SERVICE"][0], "Omset Service")
        + generate_sales_insights(boards["SCOREBOARD OMSET GADGET & AKSESORIS"][0], "Penjualan Gadget & Aksesoris")
        + generate_sales_insights(board_corp, "Marketing Corporate")
    )
    if not sales_insights:
        st.success("Tidak ada penurunan atau gap pencapaian signifikan yang terdeteksi di Service, Gadget & Aksesoris, maupun Marketing Corporate. Pertahankan performa saat ini.")
    else:
        _order = {"bad": 0, "warn": 1, "good": 2}
        sales_insights_sorted = sorted(sales_insights, key=lambda i: _order.get(i["level"], 9))[:20]
        st.markdown("".join(render_insight_card(i["title"], i["text"], i["level"]) for i in sales_insights_sorted), unsafe_allow_html=True)
        if len(sales_insights) > 20:
            st.caption(f"Menampilkan 20 insight paling kritis dari total {len(sales_insights)} yang terdeteksi.")

# --------------------------------------------------------------------------------------
# TAB 3: Iklan (Meta Ads)
# --------------------------------------------------------------------------------------

with tab_ads:
    st.caption(
        "Upload hasil export **Campaigns** dari Meta Ads Manager (Export > Excel) lewat sidebar. "
        "Fokus di sini: **Messaging Conversations Started** & **Cost per Messaging Conversation Started**, "
        "plus rekomendasi otomatis berdasarkan performa tiap campaign."
    )

    if not os.listdir(ADS_DATA_DIR):
        st.info(
            "Belum ada file iklan yang diupload. Upload file export Campaigns dari Meta Ads Manager "
            "lewat sidebar ('📣 Data Iklan Meta Ads') untuk melihat performa Messaging Conversation "
            "per cabang beserta rekomendasi improvement-nya."
        )
    else:
        with st.spinner(f"Memproses {len(os.listdir(ADS_DATA_DIR))} file iklan..."):
            df_ads, ads_errors = load_all_ads_data(ADS_DATA_DIR)

        if ads_errors:
            st.warning("Sebagian file dilewati karena formatnya tidak cocok:\n\n" + "\n".join(f"- {e}" for e in ads_errors))

        if df_ads.empty:
            st.warning("Tidak ada data campaign yang bisa dibaca dari file yang diupload.")
        else:
            ads_branches = order_branches([b for b in df_ads["Cabang"].unique() if b != "LAINNYA"])
            if "LAINNYA" in df_ads["Cabang"].unique():
                ads_branches = ads_branches + ["LAINNYA"]
            ads_status_options = sorted(df_ads["Status"].unique())

            fa1, fa2 = st.columns([2, 1])
            with fa1:
                sel_ads_branch = st.multiselect("Cabang", ads_branches, default=ads_branches, key="ads_branch_filter")
            with fa2:
                sel_ads_status = st.multiselect("Status Campaign", ads_status_options, default=ads_status_options, key="ads_status_filter")

            f_ads = df_ads[df_ads["Cabang"].isin(sel_ads_branch) & df_ads["Status"].isin(sel_ads_status)]

            if f_ads.empty:
                st.warning("Tidak ada campaign untuk kombinasi filter ini.")
            else:
                total_spend = f_ads["Spend"].sum()
                total_msg = f_ads["MsgConv"].sum()
                avg_cost = (total_spend / total_msg) if total_msg else None
                active_n = f_ads.loc[f_ads["Status"].str.lower() == "active", "Campaign"].nunique()

                a1, a2, a3, a4 = st.columns(4)
                a1.markdown(render_kpi_card_text("Total Spend", format_rupiah(total_spend), "#1e3a8a", "#3b82f6", "💰"), unsafe_allow_html=True)
                a2.markdown(render_kpi_card_text("Messaging Conversations", format_number(total_msg), "#0f766e", "#14b8a6", "💬"), unsafe_allow_html=True)
                a3.markdown(render_kpi_card_text("Rata-rata Cost / Messaging", format_rupiah(avg_cost) if avg_cost else "-", "#b45309", "#f59e0b", "📉"), unsafe_allow_html=True)
                a4.markdown(render_kpi_card_text("Campaign Aktif", format_number(active_n), "#6d28d9", "#a78bfa", "🚀"), unsafe_allow_html=True)

                st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
                st.subheader("Performa per Cabang")

                branch_agg = aggregate_ads_by_branch(f_ads)
                branch_agg_display = branch_agg.copy()
                branch_agg_display["Cabang"] = pd.Categorical(
                    branch_agg_display["Cabang"], categories=order_branches(branch_agg_display["Cabang"]) if "LAINNYA" not in branch_agg_display["Cabang"].values else order_branches([b for b in branch_agg_display["Cabang"] if b != "LAINNYA"]) + ["LAINNYA"],
                    ordered=True,
                )
                branch_agg_display = branch_agg_display.sort_values("Cabang")

                gcol1, gcol2 = st.columns(2)
                with gcol1:
                    cost_data = branch_agg.sort_values("CostPerMsg").copy()
                    cost_data["Label"] = cost_data["CostPerMsg"].apply(lambda v: format_rupiah(v) if pd.notna(v) else "")
                    fig_cost = px.bar(
                        cost_data, x="Cabang", y="CostPerMsg",
                        color="CostPerMsg", color_continuous_scale="RdYlGn_r", text="Label",
                        title="Cost per Messaging Conversation per Cabang (makin rendah makin baik)",
                    )
                    fig_cost.update_traces(textposition="outside", textfont_size=9, cliponaxis=False)
                    fig_cost.update_layout(xaxis_title="", yaxis_title="Cost per Messaging (Rp)", coloraxis_showscale=False, margin=dict(t=60))
                    st.plotly_chart(fig_cost, width="stretch")
                with gcol2:
                    msg_data = branch_agg.sort_values("MsgConv", ascending=False).copy()
                    msg_data["Label"] = msg_data["MsgConv"].apply(lambda v: format_number(v) if pd.notna(v) else "")
                    fig_msg = px.bar(
                        msg_data, x="Cabang", y="MsgConv",
                        color="MsgConv", color_continuous_scale="Greens", text="Label",
                        title="Jumlah Messaging Conversation per Cabang",
                    )
                    fig_msg.update_traces(textposition="outside", textfont_size=9, cliponaxis=False)
                    fig_msg.update_layout(xaxis_title="", yaxis_title="Messaging Conversations", coloraxis_showscale=False, margin=dict(t=60))
                    st.plotly_chart(fig_msg, width="stretch")

                branch_table = branch_agg_display.rename(columns={
                    "Spend": "Spend (Rp)", "MsgConv": "Messaging Conversation", "CostPerMsg": "Cost per Messaging (Rp)",
                    "Impressions": "Impressions", "LinkClicks": "Link Clicks", "CTR": "CTR", "CPM": "CPM (Rp)",
                    "JumlahCampaign": "Jumlah Campaign",
                })[["Cabang", "Jumlah Campaign", "Spend (Rp)", "Messaging Conversation", "Cost per Messaging (Rp)", "CTR", "CPM (Rp)", "Impressions", "Link Clicks"]]
                try:
                    sty = (
                        branch_table.style
                        .background_gradient(cmap="RdYlGn_r", subset=["Cost per Messaging (Rp)"])
                        .background_gradient(cmap="Greens", subset=["Messaging Conversation"])
                        .format({
                            "Spend (Rp)": format_rupiah, "Cost per Messaging (Rp)": lambda v: format_rupiah(v) if pd.notna(v) else "-",
                            "CPM (Rp)": lambda v: format_rupiah(v) if pd.notna(v) else "-",
                            "CTR": lambda v: format_percent(v) if pd.notna(v) else "-",
                            "Impressions": format_number, "Link Clicks": format_number, "Messaging Conversation": format_number,
                        })
                    )
                    st.dataframe(sty, width="stretch", hide_index=True)
                except ImportError:
                    disp = branch_table.copy()
                    disp["Spend (Rp)"] = disp["Spend (Rp)"].apply(format_rupiah)
                    disp["Cost per Messaging (Rp)"] = disp["Cost per Messaging (Rp)"].apply(lambda v: format_rupiah(v) if pd.notna(v) else "-")
                    st.dataframe(disp, width="stretch", hide_index=True)

                st.subheader("Detail per Campaign")
                detail = f_ads.sort_values("CostPerMsg", na_position="last")[
                    ["Cabang", "Campaign", "Status", "Spend", "MsgConv", "CostPerMsg", "CTR", "CPM", "Impressions", "LinkClicks"]
                ].rename(columns={
                    "Spend": "Spend (Rp)", "MsgConv": "Messaging Conversation", "CostPerMsg": "Cost per Messaging (Rp)",
                    "CPM": "CPM (Rp)", "LinkClicks": "Link Clicks",
                })
                disp_detail = detail.copy()
                disp_detail["Spend (Rp)"] = disp_detail["Spend (Rp)"].apply(format_rupiah)
                disp_detail["Cost per Messaging (Rp)"] = disp_detail["Cost per Messaging (Rp)"].apply(lambda v: format_rupiah(v) if pd.notna(v) else "-")
                disp_detail["CPM (Rp)"] = disp_detail["CPM (Rp)"].apply(lambda v: format_rupiah(v) if pd.notna(v) else "-")
                disp_detail["CTR"] = disp_detail["CTR"].apply(lambda v: format_percent(v) if pd.notna(v) else "-")
                disp_detail["Impressions"] = disp_detail["Impressions"].apply(format_number)
                disp_detail["Link Clicks"] = disp_detail["Link Clicks"].apply(format_number)
                disp_detail["Messaging Conversation"] = disp_detail["Messaging Conversation"].apply(format_number)
                st.dataframe(disp_detail, width="stretch", hide_index=True)

                st.markdown("---")
                st.subheader("💡 Insight & Rekomendasi Improvement")
                st.caption(
                    "Dihitung otomatis dari data yang diupload: dibandingkan terhadap rata-rata Cost per "
                    "Messaging Conversation tertimbang semua campaign yang tampil di filter ini."
                )
                insights, _, _, _ = generate_ads_insights(f_ads)
                if not insights:
                    st.success("Performa semua campaign relatif merata, tidak ada anomali signifikan yang terdeteksi.")
                else:
                    order = {"bad": 0, "warn": 1, "good": 2}
                    insights_sorted = sorted(insights, key=lambda i: order.get(i["level"], 9))
                    html = "".join(render_insight_card(i["title"], i["text"], i["level"]) for i in insights_sorted)
                    st.markdown(html, unsafe_allow_html=True)

                st.caption(
                    "Cabang 'LAINNYA' = campaign yang namanya tidak diawali kode/nama cabang "
                    "(mis. campaign umum/brand). Cabang tidak dikenali lainnya akan tetap tampil apa adanya."
                )
