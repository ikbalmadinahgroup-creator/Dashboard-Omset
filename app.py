"""Dashboard Omset MFlash - Streamlit dashboard untuk monitoring Omset, Iklan, Walk-in,
dan 6 Pilar MFlash, 18 cabang MFlash.

Fitur utama:
- Tab Ringkasan, Scoreboard, Iklan, Walk-in, 6 Pilar
- Auto-deteksi 2 format file Omset utama: file master (sheet Faktur Penjualan +
  Scoreboard) ATAU file per-cabang "Rincian Faktur Penjualan" (bisa banyak file sekaligus)
- 6 Pilar MFlash: klasifikasi omset dari kolom KATEGORI PILAR yang HANYA ada di file
  per-cabang (Rincian Faktur Penjualan), tidak ada di file master lama. PENTING: nama
  kolom ini TIDAK seragam antar file - kadang "KATEGORI PILAR Sales Invoice", kadang
  "KATEGORI PILAR Faktur Penjualan" - jadi dideteksi dengan cara mencari kolom mana pun
  yang diawali "KATEGORI PILAR", bukan exact match ke satu nama saja (lihat
  _find_pilar_column_index()). Setiap Pilar juga dilengkapi Gross Profit (TOTAL HARGA -
  HARGA BELI, atau kolom GROSS PROFIT yang sudah dihitung di file master kalau ada) dan
  Qty khusus untuk Pilar Penyewaan Corporate & Maintenance Corporate.
- Auto-ekstrak Target & Scoreboard Marketing Corporate dari sheet Scoreboard (kalau ada),
  dengan fallback ke file Target manual yang diupload - file manual DIPRIORITASKAN kalau
  auto-ekstrak gagal menemukan angka target yang valid (lihat _has_target_signal())
- Target MFlash berlaku per KUARTAL (Jan-Mar/Apr-Jun/Jul-Sep/Okt-Des), % Pencapaian
  dihitung kumulatif sejak awal kuartal s/d tanggal acuan (bukan per bulan)
- Insight & Rekomendasi Perbaikan yang dikelompokkan rapi per kategori, dengan rencana
  aksi Online/Offline terpisah
- Export laporan presentasi CEO (PPTX & PDF): Penyajian Data, Evaluasi, Perbaikan
- Ledger permanen (histori upload) supaya grafik riwayat pencapaian harian tidak hilang
  meski file lama diganti/dihapus
- Auto-backup semua file yang diupload ke repo GitHub (opsional, lihat README) supaya
  data tidak hilang saat app di Streamlit Community Cloud sleep/restart/redeploy
- Dashboard TIDAK berhenti render total kalau data Omset belum diupload - tab Iklan &
  Walk-in tetap bisa dipakai independen
"""

import base64
import calendar
import io
import os
import re
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    ListFlowable, ListItem, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)


# --------------------------------------------------------------------------------------
# Logo / favicon (base64, disematkan langsung supaya tidak tergantung file eksternal)
# --------------------------------------------------------------------------------------

LOGO_BASE64 = "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"


def _page_icon():
    if not LOGO_BASE64 or "PLACEHOLDER" in LOGO_BASE64:
        return "📊"
    try:
        from PIL import Image
        return Image.open(io.BytesIO(base64.b64decode(LOGO_BASE64)))
    except Exception:  # noqa: BLE001
        return "📊"


st.set_page_config(
    page_title="Dashboard Omset MFlash",
    layout="wide",
    page_icon=_page_icon(),
    initial_sidebar_state="expanded",
)

# PENTING: hanya sembunyikan #MainMenu (hamburger menu) dan footer bawaan Streamlit -
# JANGAN sentuh header/toolbar/decoration sama sekali, karena di versi Streamlit
# terbaru tombol panah buka/tutup sidebar ("»") ikut berada di strip toolbar itu.
# Kalau toolbar disembunyikan (mis. lewat [data-testid="stToolbar"] {display:none}),
# tombol buka sidebar ikut hilang dan sidebar "Kelola Data" jadi tidak bisa dibuka sama
# sekali - ini bug yang pernah terjadi di versi sebelumnya, jangan diulangi.
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------------------
# GitHub Auto-Backup - supaya file yang diupload tidak hilang saat app di Streamlit
# Community Cloud sleep / restart / redeploy (disk container bersifat sementara)
# --------------------------------------------------------------------------------------

_GH_API_BASE = "https://api.github.com"
_GH_MAX_BYTES = 90 * 1024 * 1024


def _gh_config():
    try:
        gh = st.secrets.get("github", {})
    except Exception:  # noqa: BLE001
        gh = {}
    return gh.get("token"), gh.get("repo"), gh.get("branch", "main")


_GH_TOKEN, _GH_REPO, _GH_BRANCH = _gh_config()
_GH_ENABLED = bool(_GH_TOKEN and _GH_REPO)


def _gh_headers():
    return {
        "Authorization": f"Bearer {_GH_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _gh_repo_path(local_path: str) -> str:
    return os.path.relpath(local_path, start=os.path.dirname(os.path.abspath(__file__)))


def github_get_file_sha(repo_path: str):
    if not _GH_ENABLED:
        return None
    url = f"{_GH_API_BASE}/repos/{_GH_REPO}/contents/{repo_path}"
    try:
        resp = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=15)
        if resp.status_code == 200:
            return resp.json().get("sha")
    except Exception:  # noqa: BLE001
        pass
    return None


def github_upload_file(local_path: str, repo_path: str, message: str) -> str:
    """Return 'ok' / 'too_large' / 'error' / 'disabled'."""
    if not _GH_ENABLED:
        return "disabled"
    try:
        size = os.path.getsize(local_path)
        if size > _GH_MAX_BYTES:
            return "too_large"
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")
        sha = github_get_file_sha(repo_path)
        url = f"{_GH_API_BASE}/repos/{_GH_REPO}/contents/{repo_path}"
        payload = {"message": message, "content": content_b64, "branch": _GH_BRANCH}
        if sha:
            payload["sha"] = sha
        resp = requests.put(url, headers=_gh_headers(), json=payload, timeout=30)
        return "ok" if resp.status_code in (200, 201) else "error"
    except Exception:  # noqa: BLE001
        return "error"


def github_delete_file(repo_path: str, message: str) -> bool:
    if not _GH_ENABLED:
        return False
    sha = github_get_file_sha(repo_path)
    if not sha:
        return False
    try:
        url = f"{_GH_API_BASE}/repos/{_GH_REPO}/contents/{repo_path}"
        resp = requests.delete(
            url, headers=_gh_headers(),
            json={"message": message, "sha": sha, "branch": _GH_BRANCH}, timeout=15,
        )
        return resp.status_code == 200
    except Exception:  # noqa: BLE001
        return False


def github_download_file(repo_path: str, local_path: str) -> bool:
    if not _GH_ENABLED:
        return False
    try:
        url = f"{_GH_API_BASE}/repos/{_GH_REPO}/contents/{repo_path}"
        resp = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=20)
        if resp.status_code != 200:
            return False
        data = resp.json()
        if data.get("encoding") != "base64":
            return False
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(base64.b64decode(data["content"]))
        return True
    except Exception:  # noqa: BLE001
        return False


def github_list_dir(repo_path: str):
    if not _GH_ENABLED:
        return []
    try:
        url = f"{_GH_API_BASE}/repos/{_GH_REPO}/contents/{repo_path}"
        resp = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=15)
        if resp.status_code != 200:
            return []
        return [item["name"] for item in resp.json() if item.get("type") == "file"]
    except Exception:  # noqa: BLE001
        return []


def sync_data_from_github():
    """Tarik kembali semua file (data omset/iklan/walkin + target/corporate manual +
    ledger histori) dari GitHub ke disk lokal - dipanggil sekali di awal tiap sesi baru."""
    if not _GH_ENABLED:
        return
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for sub in ("data/main", "data/ads", "data/walkin"):
        local_dir = os.path.join(base_dir, sub)
        os.makedirs(local_dir, exist_ok=True)
        for fname in github_list_dir(sub):
            local_path = os.path.join(local_dir, fname)
            if not os.path.exists(local_path):
                github_download_file(f"{sub}/{fname}", local_path)
    for fname in ("data/target.xlsx", "data/corporate.xlsx", "data/history_log.csv", "data/corp_history_log.csv"):
        local_path = os.path.join(base_dir, fname)
        if not os.path.exists(local_path):
            github_download_file(fname, local_path)


# --------------------------------------------------------------------------------------
# Konstanta & lokasi data
# --------------------------------------------------------------------------------------

MAIN_SHEET_NAME = "Faktur Penjualan"
_FAKTUR_SHEET_NAME = "Rincian Faktur Penjualan"

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
ADS_DATA_DIR = os.path.join(DATA_DIR, "ads")
WALKIN_DATA_DIR = os.path.join(DATA_DIR, "walkin")
CORPORATE_DATA_PATH = os.path.join(DATA_DIR, "corporate.xlsx")
TARGET_DATA_PATH = os.path.join(DATA_DIR, "target.xlsx")
HISTORY_LOG_PATH = os.path.join(DATA_DIR, "history_log.csv")
CORP_HISTORY_LOG_PATH = os.path.join(DATA_DIR, "corp_history_log.csv")

CORE_COLUMNS = ["TGL FAKTUR", "KATEGORI BARANG", "TOTAL HARGA"]
REQUIRED_COLUMNS = ["CABANG"] + CORE_COLUMNS

MAX_MAIN_FILES = 50
MAX_ADS_FILES = 50
MAX_WALKIN_FILES = 60

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
BULAN_MAP = {name.upper(): i + 1 for i, name in enumerate(BULAN_ID)}
BULAN_ALIAS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MEI": 5, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGU": 8, "AGT": 8, "AUG": 8, "SEP": 9, "OKT": 10, "OCT": 10,
    "NOV": 11, "DES": 12, "DEC": 12,
}

BRANCH_ORDER = [
    "KLENDER", "RADJIMAN", "CEGER", "BINTARA", "JATIMULYA", "DRAMAGA", "CONDET",
    "JATIBENING", "SAWANGAN", "WARBONG", "CINERE", "CIBINONG", "KARAWANG",
    "JATIWARINGIN", "CIKAMPEK", "CILANGKAP", "PEJATEN", "CIBUBUR",
]
_BRANCH_RANK = {b: i for i, b in enumerate(BRANCH_ORDER)}


def order_branches(branches) -> list:
    return sorted(branches, key=lambda b: (_BRANCH_RANK.get(str(b).upper(), 999), str(b)))


# --------------------------------------------------------------------------------------
# Helper format
# --------------------------------------------------------------------------------------

def format_rupiah(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    return f"Rp {v:,.0f}".replace(",", ".")


def format_number(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    return f"{v:,.0f}".replace(",", ".")


def format_decimal(v, digits: int = 1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    return f"{v:,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_percent(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    return f"{v * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def sanitize_filename(name: str) -> str:
    base = os.path.basename(name)
    return re.sub(r"[^A-Za-z0-9._\-]+", "_", base)


def branch_from_filename(filename: str) -> str:
    name = os.path.splitext(os.path.basename(filename))[0]
    chunk = re.sub(r"[_\-]+", " ", name).strip().upper()
    for b in BRANCH_ORDER:
        if chunk == b or b in chunk:
            return b
    return chunk or "TIDAK DIKETAHUI"


_ADS_BRANCH_OVERRIDES = {
    "RADJIMAN": "RADJIMAN", "JVA": "TIDAK DIKETAHUI",
}


def branch_from_campaign_name(campaign_name: str) -> str:
    chunk = str(campaign_name).upper()
    for b in BRANCH_ORDER:
        if b in chunk:
            return b
    for key, val in _ADS_BRANCH_OVERRIDES.items():
        if key in chunk:
            return val
    return "TIDAK DIKETAHUI"


def _branch_from_rincian_filename(filename: str, prefix_pattern: str) -> str:
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(prefix_pattern, "", name, flags=re.IGNORECASE)
    name = re.sub(r"[_\-]?\d{6,}$", "", name)
    name = re.sub(r"^\d+", "", name)
    name = re.sub(r"mflash", "", name, flags=re.IGNORECASE)
    chunk = re.sub(r"[_\-]+", " ", name).strip().upper()
    if not chunk:
        return "TIDAK DIKETAHUI"
    for b in BRANCH_ORDER:
        if chunk == b or chunk.startswith(b) or b.startswith(chunk) or b in chunk:
            return b
    return chunk


def branch_from_walkin_filename(filename: str) -> str:
    return _branch_from_rincian_filename(filename, r"^rincian[_\-]?pengiriman[_\-]?pesanan[_\-]?")


def branch_from_faktur_filename(filename: str) -> str:
    return _branch_from_rincian_filename(filename, r"^rincian[_\-]?faktur[_\-]?penjualan[_\-]?")


SERVICE_BARANG_VALUES = {"JASA", "SPAREPART"}


def classify_kategori(kategori_barang: str) -> str:
    if str(kategori_barang).strip().upper() in SERVICE_BARANG_VALUES:
        return "Service"
    return "Gadget & Aksesoris"


# --------------------------------------------------------------------------------------
# 6 Pilar MFlash - klasifikasi dari kolom KATEGORI PILAR yang HANYA ada di file per-cabang
# (Rincian Faktur Penjualan). File master lama tidak punya kolom ini, jadi omset dari file
# master akan otomatis masuk "Belum Dikategorikan".
#
# PENTING: nama kolom KATEGORI PILAR TIDAK seragam antar export/cabang - ada yang bernama
# "KATEGORI PILAR Sales Invoice", ada yang "KATEGORI PILAR Faktur Penjualan" (keduanya
# sudah ditemukan langsung di file nyata yang diupload user). Jadi jangan exact-match ke
# satu nama saja - cari kolom mana pun yang header-nya DIAWALI "KATEGORI PILAR".
# --------------------------------------------------------------------------------------

PILAR_ORDER = [
    "1. Service", "2. Penjualan Ritel", "3. Pengadaan Corporate",
    "4. Penyewaan Corporate", "5. Maintenance Corporate", "6. Internet & Connectivity",
    "Cicilan Syariah", "Belum Dikategorikan",
]
PILAR_ICONS = {
    "1. Service": "🛠️", "2. Penjualan Ritel": "📱", "3. Pengadaan Corporate": "🏢",
    "4. Penyewaan Corporate": "📦", "5. Maintenance Corporate": "🔧",
    "6. Internet & Connectivity": "📶", "Cicilan Syariah": "🕌", "Belum Dikategorikan": "❓",
}
PILAR_COLORS = {
    "1. Service": "#0f766e", "2. Penjualan Ritel": "#6d28d9", "3. Pengadaan Corporate": "#1d4ed8",
    "4. Penyewaan Corporate": "#b45309", "5. Maintenance Corporate": "#0891b2",
    "6. Internet & Connectivity": "#65a30d", "Cicilan Syariah": "#be185d", "Belum Dikategorikan": "#9ca3af",
}

# Pilar yang perlu ditampilkan metrik Qty tambahan di KPI card (selain Omset & Gross
# Profit yang berlaku untuk semua pilar) - sesuai permintaan eksplisit user.
_PILAR_SHOW_QTY = {"4. Penyewaan Corporate", "5. Maintenance Corporate"}


def _pilar_label(p: str) -> str:
    return p.split(". ", 1)[-1] if ". " in p else p


def _find_pilar_column_index(col_idx: dict):
    """Cari index kolom KATEGORI PILAR di dict {header_upper: index}. Nama kolom ini
    TIDAK seragam antar file export (contoh nyata yang sudah ditemukan: 'KATEGORI PILAR
    SALES INVOICE' dan 'KATEGORI PILAR FAKTUR PENJUALAN'), jadi dicari dengan prefix
    match ke 'KATEGORI PILAR', bukan exact match ke satu variasi nama saja."""
    for header, idx in col_idx.items():
        if header.startswith("KATEGORI PILAR"):
            return idx
    return None


def classify_pilar(raw_value) -> str:
    """Klasifikasi nilai kolom KATEGORI PILAR ke salah satu dari 6 Pilar resmi MFlash,
    plus 2 kategori tambahan: 'Cicilan Syariah' (kadang tercatat di kolom ini sebagai
    metode pembayaran) dan 'Belum Dikategorikan' (kosong/tidak dikenali/file tidak punya
    kolom ini sama sekali)."""
    if raw_value is None:
        return "Belum Dikategorikan"
    v = str(raw_value).strip().upper()
    if not v:
        return "Belum Dikategorikan"
    if "CICILAN" in v:
        return "Cicilan Syariah"
    if "SERVICE" in v:
        return "1. Service"
    if "RITEL" in v or "RETAIL" in v:
        return "2. Penjualan Ritel"
    if "PENGADAAN" in v:
        return "3. Pengadaan Corporate"
    if "PENYEWAAN" in v or "SEWA" in v:
        return "4. Penyewaan Corporate"
    if "MAINTENANCE" in v or "PERAWATAN" in v:
        return "5. Maintenance Corporate"
    if "INTERNET" in v or "WIFI" in v or "CONNECTIVITY" in v:
        return "6. Internet & Connectivity"
    return "Belum Dikategorikan"


def parse_bulan(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        iv = int(v)
        return iv if 1 <= iv <= 12 else None
    s = str(v).strip().upper()
    if s in BULAN_MAP:
        return BULAN_MAP[s]
    if s in BULAN_ALIAS:
        return BULAN_ALIAS[s]
    try:
        return int(s)
    except ValueError:
        return None


def to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return pd.to_datetime(v).date()
    except Exception:  # noqa: BLE001
        return None


def _to_float_or_none(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------------------
# Loader data Omset utama - mendukung DUA format:
# 1) Format lama (file master): sheet "Faktur Penjualan" + kolom CABANG, dibaca streaming
#    (read_only=True) untuk performa karena filenya besar (puluhan MB, semua cabang).
#    Format ini TIDAK punya kolom KATEGORI PILAR -> semua barisnya "Belum Dikategorikan".
#    Tapi format ini SUDAH punya kolom "GROSS PROFIT" yang dihitung sendiri, jadi dipakai
#    langsung kalau ada (lebih akurat daripada dihitung ulang).
# 2) Format baru (per cabang): sheet "Rincian Faktur Penjualan", TIDAK ada kolom CABANG
#    (cabang diambil dari nama file), dan filenya punya bug tag <dimension> yang salah
#    (cuma mendeklarasikan 1 sel) sehingga read_only=True gagal membaca semua baris -
#    harus di-parse penuh (bukan streaming). Format ini PUNYA kolom KATEGORI PILAR
#    (nama persisnya bervariasi, lihat _find_pilar_column_index()), tapi TIDAK ada kolom
#    GROSS PROFIT siap pakai, jadi dihitung sendiri: GROSS PROFIT = TOTAL HARGA - HARGA
#    BELI (sudah diverifikasi silang terhadap kolom GROSS PROFIT bawaan file master -
#    HARGA BELI di kedua format adalah harga beli TOTAL per baris, BUKAN per unit, jadi
#    tidak perlu dikalikan QTY lagi).
# --------------------------------------------------------------------------------------

_FAKTUR_REQUIRED_COLS = ["TGL FAKTUR", "KATEGORI BARANG", "TOTAL HARGA"]


def _find_data_sheet(wb):
    """Cari sheet format lama: harus punya kolom CABANG + CORE_COLUMNS di baris header."""
    candidates = [n for n in wb.sheetnames if n.strip().lower() == MAIN_SHEET_NAME.lower()]
    candidates += [n for n in wb.sheetnames if n not in candidates]
    for name in candidates:
        ws = wb[name]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row is None:
            continue
        col_idx = {}
        for i, h in enumerate(header_row):
            if h is not None:
                col_idx[str(h).strip().upper()] = i
        if all(c in col_idx for c in REQUIRED_COLUMNS):
            return ws, col_idx
    return None, None


def _extract_qty_gp(row, col_idx: dict, total: float):
    """Ambil Qty & Gross Profit dari satu baris. Prioritas Gross Profit: pakai kolom
    'GROSS PROFIT' bawaan file kalau ada (file master), kalau tidak dihitung dari
    TOTAL HARGA - HARGA BELI (HARGA BELI di file ini adalah total per baris, bukan
    per unit, jadi tidak dikalikan Qty lagi)."""
    qty_idx = col_idx.get("QTY")
    hargabeli_idx = col_idx.get("HARGA BELI")
    gp_idx = col_idx.get("GROSS PROFIT")

    qty_val = _to_float_or_none(row[qty_idx]) if (qty_idx is not None and qty_idx < len(row)) else None
    hargabeli_val = _to_float_or_none(row[hargabeli_idx]) if (hargabeli_idx is not None and hargabeli_idx < len(row)) else None
    gp_val = _to_float_or_none(row[gp_idx]) if (gp_idx is not None and gp_idx < len(row)) else None

    if gp_val is None:
        gp_val = (total - hargabeli_val) if hargabeli_val is not None else 0.0

    return (qty_val if qty_val is not None else 0.0), gp_val


def _load_faktur_sheet(ws, filename_hint: str) -> pd.DataFrame:
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header_row is None:
        raise ValueError(f"Sheet '{_FAKTUR_SHEET_NAME}' kosong / tidak ada header.")
    col_idx = {}
    for i, h in enumerate(header_row):
        if h is not None:
            col_idx[str(h).strip().upper()] = i
    missing = [c for c in _FAKTUR_REQUIRED_COLS if c not in col_idx]
    if missing:
        raise ValueError(f"Kolom berikut tidak ditemukan di sheet '{_FAKTUR_SHEET_NAME}': {', '.join(missing)}.")
    branch = branch_from_faktur_filename(filename_hint) if filename_hint else "TIDAK DIKETAHUI"
    tgl_idx = col_idx["TGL FAKTUR"]
    barang_idx = col_idx["KATEGORI BARANG"]
    total_idx = col_idx["TOTAL HARGA"]
    pilar_idx = _find_pilar_column_index(col_idx)
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if tgl_idx >= len(row) or barang_idx >= len(row) or total_idx >= len(row):
            continue
        tgl = row[tgl_idx]
        barang = row[barang_idx]
        total = row[total_idx]
        if not isinstance(tgl, datetime):
            continue
        if total is None:
            total = 0
        try:
            total = float(total)
        except (TypeError, ValueError):
            continue
        pilar_raw = row[pilar_idx] if (pilar_idx is not None and pilar_idx < len(row)) else None
        qty_val, gp_val = _extract_qty_gp(row, col_idx, total)
        rows.append({
            "Cabang": branch, "Tanggal": pd.Timestamp(tgl.date()), "Tahun": tgl.year, "Bulan": tgl.month,
            "KategoriBarang": str(barang).strip().upper() if barang else "", "Omset": total,
            "Pilar": classify_pilar(pilar_raw), "Qty": qty_val, "GrossProfit": gp_val, "SumberFile": filename_hint,
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Kelompok"] = df["KategoriBarang"].apply(classify_kategori)
    return df


def _load_master_sheet(ws, col_idx, filename_hint: str) -> pd.DataFrame:
    """Baca sheet format lama secara streaming (ws sudah dibuka read_only=True). Format ini
    tidak pernah punya kolom KATEGORI PILAR, tapi tetap dicek (prefix match) untuk berjaga-
    jaga kalau suatu saat kolom itu ditambahkan ke file master juga."""
    cabang_idx = col_idx["CABANG"]
    tgl_idx = col_idx["TGL FAKTUR"]
    barang_idx = col_idx["KATEGORI BARANG"]
    total_idx = col_idx["TOTAL HARGA"]
    pilar_idx = _find_pilar_column_index(col_idx)
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        max_idx = max(cabang_idx, tgl_idx, barang_idx, total_idx)
        if max_idx >= len(row):
            continue
        cabang = row[cabang_idx]
        tgl = row[tgl_idx]
        barang = row[barang_idx]
        total = row[total_idx]
        if not cabang or not isinstance(tgl, datetime):
            continue
        if total is None:
            total = 0
        try:
            total = float(total)
        except (TypeError, ValueError):
            continue
        cabang_norm = str(cabang).strip().upper()
        pilar_raw = row[pilar_idx] if (pilar_idx is not None and pilar_idx < len(row)) else None
        qty_val, gp_val = _extract_qty_gp(row, col_idx, total)
        rows.append({
            "Cabang": cabang_norm, "Tanggal": pd.Timestamp(tgl.date()), "Tahun": tgl.year, "Bulan": tgl.month,
            "KategoriBarang": str(barang).strip().upper() if barang else "", "Omset": total,
            "Pilar": classify_pilar(pilar_raw), "Qty": qty_val, "GrossProfit": gp_val, "SumberFile": filename_hint,
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Kelompok"] = df["KategoriBarang"].apply(classify_kategori)
    return df


def _looks_like_ads_export(file_bytes: bytes) -> bool:
    """Heuristik: file ini kemungkinan besar export Meta Ads Manager yang ke-upload di
    slot yang salah (bukan slot Omset), supaya errornya tidak ditampilkan berisik."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        for name in wb.sheetnames:
            ws = wb[name]
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if header_row is None:
                continue
            headers = {str(h).strip().upper() for h in header_row if h is not None}
            if "CAMPAIGN NAME" in headers or "AMOUNT SPENT (IDR)" in headers:
                return True
    except Exception:  # noqa: BLE001
        pass
    return False


@st.cache_data(show_spinner=False)
def load_main_data(file_bytes: bytes, filename_hint: str = "") -> pd.DataFrame:
    """Auto-deteksi format: sheet 'Rincian Faktur Penjualan' (format baru, per cabang) ATAU
    sheet format lama (Faktur Penjualan + kolom CABANG). Peek nama sheet dulu pakai
    read_only=True (murah, tidak scan baris) untuk menentukan strategi parsing."""
    try:
        wb_peek = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheetnames = wb_peek.sheetnames
    except Exception as e:  # noqa: BLE001
        raise ValueError(f"Gagal membuka file Excel: {e}") from e

    faktur_sheet_name = next((n for n in sheetnames if n.strip().lower() == _FAKTUR_SHEET_NAME.lower()), None)
    if faktur_sheet_name is not None:
        wb_full = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        return _load_faktur_sheet(wb_full[faktur_sheet_name], filename_hint)

    ws, col_idx = _find_data_sheet(wb_peek)
    if ws is None:
        raise ValueError(
            f"Tidak ditemukan sheet dengan kolom {', '.join(REQUIRED_COLUMNS)} "
            f"maupun sheet '{_FAKTUR_SHEET_NAME}'. Pastikan file sesuai format yang didukung."
        )
    return _load_master_sheet(ws, col_idx, filename_hint)


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
            if _looks_like_ads_export(file_bytes):
                continue  # kemungkinan besar file export Meta Ads, ke-upload di slot yang salah - lewati diam-diam
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
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header_row is None:
        raise ValueError("File Iklan kosong / tidak ada header.")
    col_idx = {}
    for i, h in enumerate(header_row):
        if h is not None:
            col_idx[str(h).strip().upper()] = i
    missing = [c for c in _ADS_REQUIRED_COLS if c not in col_idx]
    if missing:
        raise ValueError(f"Kolom berikut tidak ditemukan di file Iklan: {', '.join(missing)}.")

    def _get(row, key):
        idx = col_idx.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        campaign = _get(row, "CAMPAIGN NAME")
        if not campaign:
            continue
        spend = _num_or(_get(row, "AMOUNT SPENT (IDR)"))
        msg_conv = _num_or(_get(row, "MESSAGING CONVERSATIONS STARTED"))
        cost_per_msg = _num_or_none(_get(row, "COST PER MESSAGING CONVERSATION STARTED (IDR)"))
        if cost_per_msg is None and msg_conv:
            cost_per_msg = spend / msg_conv
        status = _get(row, "DELIVERY STATUS") or _get(row, "AD SET DELIVERY") or _get(row, "STATUS") or ""
        impressions = _num_or(_get(row, "IMPRESSIONS"))
        link_clicks = _num_or(_get(row, "LINK CLICKS"))
        ctr = _num_or_none(_get(row, "CTR (LINK CLICK-THROUGH RATE)")) or _num_or_none(_get(row, "CTR (ALL)"))
        if ctr is not None and ctr > 1:
            ctr = ctr / 100.0
        cpm = _num_or_none(_get(row, "CPM (COST PER 1,000 IMPRESSIONS) (IDR)"))
        results = _num_or_none(_get(row, "RESULTS"))
        rows.append({
            "Cabang": branch_from_campaign_name(str(campaign)),
            "Campaign": str(campaign), "Status": str(status).strip(),
            "PeriodeMulai": to_date(_get(row, "REPORTING STARTS")),
            "PeriodeSelesai": to_date(_get(row, "REPORTING ENDS")),
            "Spend": spend, "MsgConv": msg_conv, "CostPerMsg": cost_per_msg or 0.0,
            "Results": results, "Impressions": impressions, "LinkClicks": link_clicks,
            "CTR": ctr, "CPM": cpm, "SumberFile": filename_hint,
        })
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
        Spend=("Spend", "sum"), MsgConv=("MsgConv", "sum"),
        Impressions=("Impressions", "sum"), LinkClicks=("LinkClicks", "sum"),
        JumlahCampaign=("Campaign", "nunique"),
    ).reset_index()
    g["CostPerMsg"] = g.apply(lambda r: (r["Spend"] / r["MsgConv"]) if r["MsgConv"] else None, axis=1)
    g["CTR"] = g.apply(lambda r: (r["LinkClicks"] / r["Impressions"]) if r["Impressions"] else None, axis=1)
    g["CPM"] = g.apply(lambda r: (r["Spend"] / r["Impressions"] * 1000) if r["Impressions"] else None, axis=1)
    g = g[["Cabang", "Spend", "MsgConv", "CostPerMsg", "Impressions", "LinkClicks", "CTR", "CPM", "JumlahCampaign"]]
    return g.sort_values("Spend", ascending=False).reset_index(drop=True)


def generate_ads_insights(df: pd.DataFrame):
    """Rekomendasi otomatis berbasis data: bandingkan tiap campaign terhadap rata-rata
    tertimbang Cost per Messaging Conversation, dan tandai campaign yang boros tanpa hasil.
    Juga mendiagnosa dari sisi KONTEN: dibandingkan dengan CTR median, supaya rekomendasi
    lebih spesifik (ganti creative vs benahi funnel/kecepatan respon admin)."""
    insights = []
    if df.empty:
        return insights, 0.0, 0.0, None

    total_spend = df["Spend"].sum()
    total_msg = df["MsgConv"].sum()
    avg_cost = (total_spend / total_msg) if total_msg else None
    ctr_values = df.loc[df["CTR"].notna() & (df["CTR"] > 0), "CTR"]
    median_ctr = ctr_values.median() if not ctr_values.empty else None

    def _content_diagnosis(row) -> str:
        ctr = row.get("CTR")
        if median_ctr is None or ctr is None or pd.isna(ctr):
            return (
                " Evaluasi dari sisi konten: pastikan gambar/video utama langsung menunjukkan produk & "
                "harga promo dengan jelas dalam 3 detik pertama, judul singkat & menarik, serta ada CTA "
                "yang jelas (mis. 'Chat Sekarang'). Cek juga apakah audience yang ditarget relevan "
                "(radius lokasi cabang, minat gadget/HP)."
            )
        if ctr < median_ctr * 0.7:
            return (
                f" CTR campaign ini ({format_percent(ctr)}) di bawah rata-rata ({format_percent(median_ctr)}) — "
                f"kemungkinan besar masalah ada di KONTEN iklan: ganti hook/thumbnail 3 detik pertama, "
                f"perjelas judul & harga promo di gambar/video utama, dan pastikan CTA (mis. 'Chat Sekarang') "
                f"terlihat jelas supaya lebih banyak yang klik."
            )
        return (
            f" CTR campaign ini ({format_percent(ctr)}) sudah cukup baik (≥ rata-rata {format_percent(median_ctr)}), "
            f"jadi kemungkinan bukan di kontennya — cek funnel SETELAH orang klik: pastikan auto-reply/greeting "
            f"WhatsApp atau Messenger aktif, dan admin membalas cepat supaya calon customer tidak kabur "
            f"sebelum sempat jadi conversation."
        )

    zero_conv = df[(df["Spend"] >= 5000) & (df["MsgConv"] == 0)].sort_values("Spend", ascending=False)
    for _, r in zero_conv.iterrows():
        insights.append({
            "level": "bad", "title": f"{r['Campaign']} ({r['Cabang']})",
            "text": f"Sudah menghabiskan {format_rupiah(r['Spend'])} tapi belum menghasilkan Messaging "
                    f"Conversation sama sekali. Evaluasi ulang audience/creative, atau pause campaign ini "
                    f"supaya budget tidak terus terbuang." + _content_diagnosis(r),
        })

    if avg_cost:
        high = df[(df["MsgConv"] > 0) & (df["CostPerMsg"] > avg_cost * 1.5)].sort_values("CostPerMsg", ascending=False)
        for _, r in high.iterrows():
            ratio = r["CostPerMsg"] / avg_cost
            insights.append({
                "level": "warn", "title": f"{r['Campaign']} ({r['Cabang']})",
                "text": f"Cost per Messaging Conversation {format_rupiah(r['CostPerMsg'])} — {ratio:.1f}x lebih "
                        f"mahal dari rata-rata semua campaign ({format_rupiah(avg_cost)})." + _content_diagnosis(r),
            })

        low = df[(df["MsgConv"] >= 3) & (df["CostPerMsg"] < avg_cost * 0.7)].sort_values("CostPerMsg")
        for _, r in low.iterrows():
            insights.append({
                "level": "good", "title": f"{r['Campaign']} ({r['Cabang']})",
                "text": f"Paling efisien: Cost per Messaging Conversation hanya {format_rupiah(r['CostPerMsg'])} "
                        f"dari {int(r['MsgConv'])} conversation. Kandidat kuat untuk dinaikkan budgetnya (scale up) "
                        f"— creative & audience-nya bisa dijadikan referensi untuk campaign cabang lain yang masih mahal.",
            })

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

_CATEGORY_ACTION_PLANS = {
    "Omset Service": {
        "online": [
            "refresh & post ulang promo harga servis (LCD, baterai, konektor cas, dll) di Instagram/Facebook Story "
            "& Feed cabang minimal 3-4x/minggu, dengan harga & before-after jelas",
            "aktifkan atau naikkan budget iklan Meta Ads Click-to-WhatsApp khusus radius cabang ini (cek Cost per "
            "Messaging Conversation-nya di tab Iklan, prioritaskan creative yang CTR-nya sudah terbukti bagus)",
            "broadcast WhatsApp ke database customer lama: reminder servis berkala, promo garansi habis, upgrade sparepart",
        ],
        "offline": [
            "cek ketersediaan sparepart & jadwal teknisi supaya antrean servis tidak menumpuk/lambat (kecepatan "
            "servis mendorong repeat customer & referral)",
            "briefing tim frontliner untuk aktif menawarkan cek/servis tambahan (upsell) ke customer yang datang",
            "pasang banner/signage promo servis yang jelas terlihat dari jalan",
        ],
    },
    "Penjualan Gadget & Aksesoris": {
        "online": [
            "posting katalog & harga terbaru HP/aksesoris di Instagram/Facebook/TikTok minimal 3-4x/minggu dengan "
            "foto produk jelas & harga langsung terlihat",
            "aktifkan/naikkan budget iklan Meta Ads untuk produk yang stoknya menumpuk atau margin bagus",
            "broadcast WhatsApp ke database customer untuk promo bundling HP+aksesoris atau trade-in",
        ],
        "offline": [
            "cek display & etalase toko supaya produk best-seller terlihat jelas dari depan",
            "briefing tim sales untuk aktif menawarkan aksesoris pelengkap (upsell) saat customer beli HP",
            "adakan promo/bundling berkala untuk produk yang perputarannya lambat",
        ],
    },
    "Marketing Corporate": {
        "online": [
            "follow-up ulang leads/kontak corporate yang sempat masuk tapi belum closing lewat WhatsApp/email",
            "posting studi kasus/testimoni kerja sama corporate sebelumnya di LinkedIn/Instagram untuk membangun kredibilitas",
            "kirim proposal/penawaran khusus ke database perusahaan/instansi di sekitar radius cabang",
        ],
        "offline": [
            "jadwalkan kunjungan langsung (sales visit) ke perusahaan/instansi target di sekitar cabang",
            "ikuti/adakan event networking B2B lokal untuk memperluas jaringan corporate",
            "evaluasi ulang skema harga/kontrak corporate supaya lebih kompetitif dibanding kompetitor",
        ],
    },
}


def generate_sales_insights(board_df: pd.DataFrame, category_label: str, name_label: str = "CABANG"):
    """Rekomendasi otomatis untuk setiap cabang/nama di sebuah scoreboard: tandai penurunan
    rata-rata omset harian (bulan ini vs bulan lalu) dan pencapaian yang masih di bawah 85%.
    Setiap insight menyimpan 'problem' (ringkasan masalah) terpisah dari 'online'/'offline'
    (daftar rencana aksi per kategori) supaya bisa dirender rapi per bagian, bukan satu
    paragraf panjang."""
    insights = []
    if board_df is None or board_df.empty or name_label not in board_df.columns:
        return insights

    plan = _CATEGORY_ACTION_PLANS.get(category_label, {"online": [], "offline": []})
    rows = board_df[~board_df[name_label].astype(str).str.upper().isin(_SALES_TOTAL_LABELS)]
    for _, r in rows.iterrows():
        nama = r.get(name_label)
        if nama is None or (isinstance(nama, float) and pd.isna(nama)):
            continue

        periode_lalu = r.get("PERIODE BULAN LALU")
        periode_ini = r.get("PERIODE BULAN INI")
        if periode_lalu is not None and periode_ini is not None and not pd.isna(periode_lalu) and periode_lalu > 0:
            pct_turun = (periode_ini - periode_lalu) / periode_lalu
            saran = (
                f" Rekomendasi: fokuskan rencana aksi di bawah untuk mendorong kembali omset {category_label.lower()} "
                f"cabang ini."
            )
            if pct_turun <= -0.30:
                insights.append({
                    "level": "bad", "category": category_label, "title": str(nama),
                    "problem": f"Rata-rata omset harian turun {format_percent(abs(pct_turun))} dibanding bulan lalu "
                               f"({format_rupiah(periode_lalu)} → {format_rupiah(periode_ini)} per hari).{saran}",
                    "online": plan["online"], "offline": plan["offline"],
                })
                continue
            if pct_turun <= -0.15:
                insights.append({
                    "level": "warn", "category": category_label, "title": str(nama),
                    "problem": f"Rata-rata omset harian turun {format_percent(abs(pct_turun))} dibanding bulan lalu "
                               f"({format_rupiah(periode_lalu)} → {format_rupiah(periode_ini)} per hari).{saran}",
                    "online": plan["online"], "offline": plan["offline"],
                })
                continue

        pct = r.get("% PENCAPAIAN")
        if pct is not None and not pd.isna(pct) and pct < 0.85:
            saran = (
                f" Rekomendasi: fokuskan rencana aksi di bawah untuk mengejar target {category_label.lower()} "
                f"cabang ini."
            )
            insights.append({
                "level": "warn" if pct >= 0.7 else "bad",
                "category": category_label, "title": str(nama),
                "problem": f"Pencapaian baru {format_percent(pct)} dari expected value (di bawah target 85%).{saran}",
                "online": plan["online"], "offline": plan["offline"],
            })

    return insights


def generate_all_sales_insights(sb_service, sb_gadget, sb_all, sb_corp, selected_branches=None):
    """Gabungkan insight dari scoreboard Service, Gadget & Aksesoris, dan Marketing Corporate
    menjadi satu list (dipakai di tab Scoreboard & di export laporan)."""
    combined = []
    combined.extend(generate_sales_insights(sb_service, "Omset Service"))
    combined.extend(generate_sales_insights(sb_gadget, "Penjualan Gadget & Aksesoris"))
    if sb_corp is not None:
        combined.extend(generate_sales_insights(sb_corp, "Marketing Corporate"))
    return combined


def _render_online_offline_html(online: list, offline: list) -> str:
    def _bullets(items):
        return "".join(f"<li style='margin-bottom:3px;'>{x}</li>" for x in items)

    cols_html = ""
    if online:
        cols_html += (
            '<div style="flex:1;min-width:220px;">'
            '<b style="font-size:12.5px;">🌐 Online</b>'
            f'<ul style="margin:4px 0 0 18px;padding:0;font-size:12.5px;">{_bullets(online)}</ul>'
            '</div>'
        )
    if offline:
        cols_html += (
            '<div style="flex:1;min-width:220px;">'
            '<b style="font-size:12.5px;">🏬 Offline</b>'
            f'<ul style="margin:4px 0 0 18px;padding:0;font-size:12.5px;">{_bullets(offline)}</ul>'
            '</div>'
        )
    if not cols_html:
        return ""
    return f'<div style="display:flex;gap:16px;flex-wrap:wrap;margin-top:8px;">{cols_html}</div>'


def render_structured_insight_card(item: dict) -> str:
    """Kartu insight yang rapi: pernyataan masalah di atas, lalu rencana aksi Online &
    Offline ditampilkan berdampingan sebagai bullet list terpisah (bukan satu paragraf)."""
    styles = {
        "bad": ("#fee2e2", "#991b1b", "#fca5a5", "🚨"),
        "warn": ("#fef9c3", "#854d0e", "#fde68a", "⚠️"),
        "good": ("#dcfce7", "#166534", "#86efac", "✅"),
    }
    bg, fg, border, icon = styles.get(item["level"], ("#f3f4f6", "#111827", "#e5e7eb", "ℹ️"))
    plan_html = _render_online_offline_html(item.get("online") or [], item.get("offline") or [])
    return (
        f'<div style="background:{bg};color:{fg};border-radius:10px;padding:12px 16px;'
        f'margin-bottom:10px;border-left:4px solid {border};">'
        f'<b>{icon} {item["title"]}</b>'
        f'<div style="font-size:13px;margin-top:4px;">{item["problem"]}</div>'
        f'{plan_html}'
        f'</div>'
    )


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
# Loader data Walk-in (Rincian Pengiriman Pesanan) - satu file biasanya = satu cabang.
# Walk-in dihitung dari jumlah NOMOR PENGIRIMAN PESANAN yang UNIK (bukan jumlah baris),
# karena satu order/pengiriman bisa berisi beberapa baris/item.
# --------------------------------------------------------------------------------------

_WALKIN_REQUIRED_COLS = ["NOMOR PENGIRIMAN PESANAN", "TGL PENGIRIMAN"]
_WALKIN_COLUMNS = ["Cabang", "NomorPengiriman", "Tanggal", "Tahun", "Bulan", "SumberFile"]


def _find_walkin_sheet(wb):
    for name in wb.sheetnames:
        ws = wb[name]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row is None:
            continue
        col_idx = {}
        for i, h in enumerate(header_row):
            if h is not None:
                col_idx[str(h).strip().upper()] = i
        if all(c in col_idx for c in _WALKIN_REQUIRED_COLS):
            return ws, col_idx
    return None, None


@st.cache_data(show_spinner=False)
def load_walkin_data(file_bytes: bytes, filename_hint: str = "") -> pd.DataFrame:
    """File walk-in punya bug tag <dimension> yang salah (hanya deklarasikan 1 sel) di
    beberapa export, jadi HARUS di-parse penuh (bukan read_only) supaya semua baris terbaca."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws, col_idx = _find_walkin_sheet(wb)
    if ws is None:
        raise ValueError(f"Kolom berikut tidak ditemukan di file Walk-in: {', '.join(_WALKIN_REQUIRED_COLS)}.")

    branch = branch_from_walkin_filename(filename_hint) if filename_hint else "TIDAK DIKETAHUI"
    nomor_idx = col_idx["NOMOR PENGIRIMAN PESANAN"]
    tgl_idx = col_idx["TGL PENGIRIMAN"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if nomor_idx >= len(row) or tgl_idx >= len(row):
            continue
        nomor = row[nomor_idx]
        tgl = row[tgl_idx]
        if not nomor or not isinstance(tgl, datetime):
            continue
        rows.append({
            "Cabang": branch, "NomorPengiriman": str(nomor).strip(),
            "Tanggal": pd.Timestamp(tgl.date()), "Tahun": tgl.year, "Bulan": tgl.month,
            "SumberFile": filename_hint,
        })
    return pd.DataFrame(rows, columns=_WALKIN_COLUMNS)


def load_all_walkin_data(walkin_dir: str):
    files = sorted(f for f in os.listdir(walkin_dir) if f.lower().endswith(".xlsx"))
    frames, errors = [], []
    for fname in files:
        fpath = os.path.join(walkin_dir, fname)
        try:
            with open(fpath, "rb") as f:
                file_bytes = f.read()
            df = load_walkin_data(file_bytes, filename_hint=fname)
            if not df.empty:
                frames.append(df)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{fname}: {e}")
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_WALKIN_COLUMNS)
    if not combined.empty:
        combined = combined.drop_duplicates(subset=["Cabang", "NomorPengiriman"], keep="first")
    return combined, errors


def aggregate_walkin_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """Total walk-in (jumlah NOMOR PENGIRIMAN PESANAN unik) per cabang per bulan, plus
    rata-rata walk-in per hari. Bulan yang paling baru di seluruh data dianggap 'bulan
    berjalan' (dibagi hari yang sudah lewat sampai tanggal terakhir tercatat di bulan itu),
    bulan-bulan sebelumnya dianggap sudah penuh sebulan (dibagi jumlah hari kalender)."""
    if df.empty:
        return pd.DataFrame(columns=["Cabang", "Tahun", "Bulan", "TotalWalkin", "HariEfektif", "RataRataPerHari"])

    overall_max = df["Tanggal"].max()
    latest_ym = (overall_max.year, overall_max.month)

    rows = []
    for (cabang, tahun, bulan), g in df.groupby(["Cabang", "Tahun", "Bulan"]):
        total = g["NomorPengiriman"].nunique()
        if (int(tahun), int(bulan)) == latest_ym:
            hari_efektif = g["Tanggal"].max().day
        else:
            hari_efektif = calendar.monthrange(int(tahun), int(bulan))[1]
        rata2 = (total / hari_efektif) if hari_efektif else 0.0
        rows.append({
            "Cabang": cabang, "Tahun": int(tahun), "Bulan": int(bulan),
            "TotalWalkin": total, "HariEfektif": hari_efektif, "RataRataPerHari": rata2,
        })
    return pd.DataFrame(rows).sort_values(["Tahun", "Bulan", "Cabang"]).reset_index(drop=True)


_WALKIN_ACTION_PLAN = {
    "online": [
        "aktifkan promo/booking servis online (WhatsApp/Instagram) supaya calon customer bisa reservasi "
        "duluan sebelum datang ke toko",
        "posting testimoni & before-after hasil servis di media sosial untuk membangun kepercayaan calon "
        "walk-in baru",
        "jalankan iklan Meta Ads Click-to-WhatsApp radius cabang khusus promo servis cepat/cek gratis",
    ],
    "offline": [
        "evaluasi kecepatan & keramahan pelayanan CS/teknisi — pengalaman walk-in yang baik mendorong "
        "repeat visit & referral dari mulut ke mulut",
        "pasang signage/banner yang jelas terlihat dari jalan supaya orang lewat tertarik mampir",
        "adakan promo/event kecil berkala (mis. diskon cek gratis di akhir pekan) untuk menarik walk-in baru",
    ],
}


def generate_walkin_insights(agg_df: pd.DataFrame):
    """Bandingkan rata-rata walk-in per hari bulan terakhir vs bulan sebelumnya, per cabang,
    untuk mendeteksi penurunan trafik servis. Dipakai untuk laporan/insight, disusun dengan
    struktur yang sama (problem + online/offline) seperti insight omset."""
    insights = []
    if agg_df.empty:
        return insights
    for cabang, g in agg_df.groupby("Cabang"):
        g = g.sort_values(["Tahun", "Bulan"])
        if len(g) < 2:
            continue
        prev, latest = g.iloc[-2], g.iloc[-1]
        if prev["RataRataPerHari"] <= 0:
            continue
        pct_change = (latest["RataRataPerHari"] - prev["RataRataPerHari"]) / prev["RataRataPerHari"]
        if pct_change <= -0.15:
            insights.append({
                "level": "bad" if pct_change <= -0.30 else "warn",
                "category": "Walk-in Cabang", "title": str(cabang),
                "problem": f"Rata-rata walk-in/hari turun {format_percent(abs(pct_change))} dibanding bulan "
                           f"sebelumnya ({format_decimal(prev['RataRataPerHari'])} → "
                           f"{format_decimal(latest['RataRataPerHari'])} per hari).",
                "online": _WALKIN_ACTION_PLAN["online"], "offline": _WALKIN_ACTION_PLAN["offline"],
            })
    return insights


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

    required = ["CABANG", "TAHUN", "BULAN", "OMSET"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Kolom berikut tidak ditemukan di file Corporate: {', '.join(missing)}. "
            f"Gunakan template yang tersedia di sidebar."
        )

    df = raw[required].copy()
    df = df.dropna(subset=["CABANG", "TAHUN", "BULAN"])
    df["Cabang"] = df["CABANG"].astype(str).str.strip().str.upper()
    df["Tahun"] = pd.to_numeric(df["TAHUN"], errors="coerce")
    df["Bulan"] = df["BULAN"].apply(parse_bulan)
    df["Omset"] = pd.to_numeric(df["OMSET"], errors="coerce").fillna(0)
    df = df.dropna(subset=["Tahun", "Bulan"])
    df["Tahun"] = df["Tahun"].astype(int)
    df["Bulan"] = df["Bulan"].astype(int)
    return df[["Cabang", "Tahun", "Bulan", "Omset"]]


def make_corporate_template() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Corporate"
    ws.append(["Cabang", "Tahun", "Bulan", "Omset"])
    today = date.today()
    for r in [
        ["KLENDER", today.year, BULAN_ID[today.month - 1], 25000000],
        ["RADJIMAN", today.year, BULAN_ID[today.month - 1], 18500000],
    ]:
        ws.append(r)
    for col, w in zip("ABCD", [16, 10, 14, 18]):
        ws.column_dimensions[col].width = w

    info = wb.create_sheet("Petunjuk")
    info.append(["Petunjuk pengisian Data Corporate"])
    info.append([""])
    info.append(["- Dipakai HANYA kalau file yang diupload tidak punya section 'SCOREBOARD OMSET MARKETING CORPORATE'."])
    info.append(["- Cabang: nama cabang, harus konsisten dengan nama cabang di data utama"])
    info.append(["- Tahun & Bulan: periode omset corporate tsb (Bulan boleh nama, mis. 'Agustus', atau angka 1-12)"])
    info.append(["- Omset: total omset Marketing Corporate cabang tsb pada bulan itu"])
    info.append(["- Setiap kali ada omset baru, tambahkan baris baru (jangan timpa baris lama)"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Export laporan presentasi CEO (PPTX & PDF): Penyajian Data, Evaluasi, Perbaikan
# --------------------------------------------------------------------------------------

_BRAND_DARK = RGBColor(0x1E, 0x3A, 0x8A)
_BRAND_ACCENT = RGBColor(0x1D, 0x4E, 0xD8)
_BRAND_LIGHT = RGBColor(0xEF, 0xF6, 0xFF)


def _pptx_blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _pptx_bg(slide, prs, color: RGBColor):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    rect.shadow.inherit = False
    slide.shapes._spTree.remove(rect._element)
    slide.shapes._spTree.insert(2, rect._element)
    return rect


def _pptx_title(slide, prs, text: str, subtitle: str = ""):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), prs.slide_width - Inches(1.2), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = _BRAND_DARK
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(13)
        p2.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    return tb


def _pptx_section_slide(prs, title: str, subtitle: str = ""):
    slide = _pptx_blank_slide(prs)
    _pptx_bg(slide, prs, _BRAND_DARK)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), prs.slide_width - Inches(1.6), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(16)
        p2.font.color.rgb = RGBColor(0xBF, 0xDB, 0xFE)
    return slide


def _pptx_bullet_slide(prs, title: str, bullets: list, subtitle: str = ""):
    slide = _pptx_blank_slide(prs)
    _pptx_bg(slide, prs, RGBColor(0xFF, 0xFF, 0xFF))
    _pptx_title(slide, prs, title, subtitle)
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), prs.slide_width - Inches(1.4), prs.slide_height - Inches(2.0))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = f"•  {b}"
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
        p.space_after = Pt(8)
    return slide


def _pptx_table_slide(prs, title: str, headers: list, rows: list, subtitle: str = ""):
    slide = _pptx_blank_slide(prs)
    _pptx_bg(slide, prs, RGBColor(0xFF, 0xFF, 0xFF))
    _pptx_title(slide, prs, title, subtitle)
    n_rows = len(rows) + 1
    n_cols = len(headers)
    left, top = Inches(0.5), Inches(1.5)
    width, height = prs.slide_width - Inches(1.0), Inches(0.35) * min(n_rows, 14)
    table_shape = slide.shapes.add_table(min(n_rows, 15), n_cols, left, top, width, height)
    table = table_shape.table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = str(h)
        cell.text_frame.paragraphs[0].font.size = Pt(11)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _BRAND_ACCENT
    for r, row in enumerate(rows[:14], start=1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            cell.text_frame.paragraphs[0].font.size = Pt(10.5)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xF9, 0xFA, 0xFB) if r % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
    return slide


def _pptx_chart_slide(prs, title: str, categories: list, series_name: str, values: list, subtitle: str = ""):
    slide = _pptx_blank_slide(prs)
    _pptx_bg(slide, prs, RGBColor(0xFF, 0xFF, 0xFF))
    _pptx_title(slide, prs, title, subtitle)
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    x, y, cx, cy = Inches(0.6), Inches(1.5), prs.slide_width - Inches(1.2), prs.slide_height - Inches(2.0)
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
    return slide


def _build_report_sections(
    sales_insights, ads_insights, walkin_insights,
    omset_all, omset_service, omset_gadget,
    ads_spend, ads_leads, walkin_total, walkin_konversi,
):
    """Bangun struktur section laporan (dipakai bersama oleh generator PPTX & PDF):
    setiap section punya 'heading' dan 'body' (list baris teks)."""
    penyajian = [
        f"Omset All: {format_rupiah(omset_all)}",
        f"Omset Service: {format_rupiah(omset_service)}",
        f"Penjualan Gadget & Aksesoris: {format_rupiah(omset_gadget)}",
        f"Total Spend Iklan: {format_rupiah(ads_spend)}",
        f"Total Messaging Conversation dari Iklan: {format_number(ads_leads)}",
        f"Total Walk-in: {format_number(walkin_total)}",
        f"Rata-rata Walk-in per Hari (semua cabang): {format_decimal(walkin_konversi)}",
    ]

    evaluasi = []
    for item in sales_insights:
        evaluasi.append(f"[{item['category']}] {item['title']}: {item['problem']}")
    for ins in ads_insights:
        evaluasi.append(f"[Iklan] {ins['title']}: {ins['text']}")
    for item in walkin_insights:
        evaluasi.append(f"[{item['category']}] {item['title']}: {item['problem']}")
    if not evaluasi:
        evaluasi = ["Tidak ada catatan evaluasi khusus untuk periode ini — semua indikator dalam kondisi baik."]

    perbaikan = []
    seen = set()
    for item in sales_insights:
        for b in item.get("online", []) + item.get("offline", []):
            if b not in seen:
                seen.add(b)
                perbaikan.append(b)
    for item in walkin_insights:
        for b in item.get("online", []) + item.get("offline", []):
            if b not in seen:
                seen.add(b)
                perbaikan.append(b)
    if not perbaikan:
        perbaikan = ["Pertahankan strategi yang berjalan saat ini — tidak ada rekomendasi perbaikan mendesak."]

    return {"penyajian": penyajian, "evaluasi": evaluasi, "perbaikan": perbaikan}


def generate_pptx_report(sections: dict, periode_label: str, chart_data_list: list) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    _pptx_section_slide(prs, "Laporan Omset MFlash", f"Periode: {periode_label}")

    _pptx_section_slide(prs, "1. Penyajian Data")
    _pptx_bullet_slide(prs, "Ringkasan Angka Utama", sections["penyajian"], periode_label)
    for chart in chart_data_list:
        _pptx_chart_slide(prs, chart["title"], chart["categories"], chart["series_name"], chart["values"], periode_label)

    _pptx_section_slide(prs, "2. Evaluasi")
    chunk = 8
    ev = sections["evaluasi"]
    for i in range(0, len(ev), chunk):
        _pptx_bullet_slide(prs, "Evaluasi & Temuan", ev[i:i + chunk], periode_label)

    _pptx_section_slide(prs, "3. Rencana Perbaikan")
    pb = sections["perbaikan"]
    for i in range(0, len(pb), chunk):
        _pptx_bullet_slide(prs, "Rencana Aksi", pb[i:i + chunk], periode_label)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pdf_report(sections: dict, periode_label: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleX", parent=styles["Title"], textColor=rl_colors.HexColor("#1e3a8a"))
    h_style = ParagraphStyle("HeadingX", parent=styles["Heading2"], textColor=rl_colors.HexColor("#1d4ed8"), spaceBefore=14, spaceAfter=6)
    body_style = ParagraphStyle("BodyX", parent=styles["BodyText"], fontSize=10.5, leading=15)

    story = [
        Paragraph("Laporan Omset MFlash", title_style),
        Paragraph(f"Periode: {periode_label}", styles["Normal"]),
        Spacer(1, 0.6 * cm),
    ]

    story.append(Paragraph("1. Penyajian Data", h_style))
    story.append(ListFlowable([ListItem(Paragraph(x, body_style)) for x in sections["penyajian"]], bulletType="bullet"))

    story.append(Paragraph("2. Evaluasi", h_style))
    story.append(ListFlowable([ListItem(Paragraph(x, body_style)) for x in sections["evaluasi"]], bulletType="bullet"))

    story.append(Paragraph("3. Rencana Perbaikan", h_style))
    story.append(ListFlowable([ListItem(Paragraph(x, body_style)) for x in sections["perbaikan"]], bulletType="bullet"))

    doc.build(story)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Scoreboard: Omset per kelompok (Service / Gadget & Aksesoris / All), dengan
# perbandingan periode bulan lalu (s/d tanggal yang sama) vs bulan ini, dan pencapaian
# terhadap target (SMM = Sesuai Masih Mungkin / expected value harian x hari berjalan).
#
# PENTING: target MFlash ditetapkan per KUARTAL (3 bulan sekali: Jan-Mar, Apr-Jun,
# Jul-Sep, Okt-Des) - BUKAN per bulan. Jadi EXPECTED VALUE & OMSET S/D HARI INI (yang
# jadi basis % PENCAPAIAN) dihitung kumulatif sejak AWAL KUARTAL s/d tanggal acuan.
# Kolom PERIODE BULAN LALU/PERIODE BULAN INI tetap dihitung per bulan - itu cuma
# indikator tren rata-rata omset harian bulan-ke-bulan, tidak dipakai untuk % Pencapaian.
# --------------------------------------------------------------------------------------

def _quarter_bounds(d: date):
    """Kembalikan (tanggal_awal, tanggal_akhir) kuartal kalender yang memuat tanggal d."""
    q_start_month = ((d.month - 1) // 3) * 3 + 1
    start = date(d.year, q_start_month, 1)
    end_month = q_start_month + 2
    end = date(d.year, end_month, calendar.monthrange(d.year, end_month)[1])
    return start, end


def build_scoreboard(df: pd.DataFrame, kelompok: str, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    d = df[df["Kelompok"] == kelompok].copy() if kelompok != "All" else df.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame()

    tahun_ini, bulan_ini = tanggal_acuan.year, tanggal_acuan.month
    if bulan_ini == 1:
        tahun_lalu, bulan_lalu = tahun_ini - 1, 12
    else:
        tahun_lalu, bulan_lalu = tahun_ini, bulan_ini - 1

    hari_acuan = tanggal_acuan.day
    days_in_bulan_lalu = calendar.monthrange(tahun_lalu, bulan_lalu)[1]
    hari_lalu_cutoff = min(hari_acuan, days_in_bulan_lalu)

    # Bulan berjalan - HANYA dipakai untuk indikator tren rata-rata omset harian bulan-ke-
    # bulan ("PERIODE BULAN LALU"/"PERIODE BULAN INI"), BUKAN untuk % Pencapaian terhadap
    # target (target MFlash berlaku per kuartal, lihat _quarter_bounds()).
    ini_bulan = d[(d["Tahun"] == tahun_ini) & (d["Bulan"] == bulan_ini) & (d["Tanggal"].dt.day <= hari_acuan)]
    lalu = d[(d["Tahun"] == tahun_lalu) & (d["Bulan"] == bulan_lalu) & (d["Tanggal"].dt.day <= hari_lalu_cutoff)]
    lalu_full = d[(d["Tahun"] == tahun_lalu) & (d["Bulan"] == bulan_lalu)]

    sum_ini_bulan = ini_bulan.groupby("Cabang")["Omset"].sum()
    sum_lalu = lalu.groupby("Cabang")["Omset"].sum()
    sum_lalu_full = lalu_full.groupby("Cabang")["Omset"].sum()

    # Kuartal berjalan (Jan-Mar/Apr-Jun/Jul-Sep/Okt-Des) - dasar untuk kolom "OMSET S/D
    # HARI INI" yang dibandingkan dengan EXPECTED VALUE untuk hitung % Pencapaian.
    quarter_start, _ = _quarter_bounds(tanggal_acuan)
    qtd = d[(d["Tanggal"].dt.date >= quarter_start) & (d["Tanggal"].dt.date <= tanggal_acuan)]
    sum_qtd = qtd.groupby("Cabang")["Omset"].sum()

    branches = order_branches(set(sum_ini_bulan.index) | set(sum_lalu.index) | set(sum_lalu_full.index) | set(sum_qtd.index))
    rows = []
    for b in branches:
        oi_bulan = float(sum_ini_bulan.get(b, 0.0))
        oi_kuartal = float(sum_qtd.get(b, 0.0))
        ol = float(sum_lalu.get(b, 0.0))
        olf = float(sum_lalu_full.get(b, 0.0))
        avg_ini = oi_bulan / hari_acuan if hari_acuan else 0.0
        avg_lalu = ol / hari_lalu_cutoff if hari_lalu_cutoff else 0.0
        rows.append({
            "CABANG": b, "OMSET BULAN LALU (FULL)": olf, "OMSET S/D HARI INI": oi_kuartal,
            "OMSET BULAN LALU S/D HARI SAMA": ol, "PERIODE BULAN LALU": avg_lalu, "PERIODE BULAN INI": avg_ini,
        })
    return pd.DataFrame(rows)


def _finalize_scoreboard(sb: pd.DataFrame, df_target: pd.DataFrame, target_col: str, tanggal_acuan: date) -> pd.DataFrame:
    """Tambahkan kolom TARGET & % PENCAPAIAN (Sesuai Masih Mungkin) dan baris total SMM.
    TARGET diasumsikan sudah berupa total 1 KUARTAL PENUH (3 bulan) - EXPECTED VALUE
    dihitung dari rata-rata harian target kuartal dikali jumlah hari yang sudah berjalan
    sejak awal kuartal s/d tanggal acuan (bukan dari awal bulan)."""
    if sb.empty:
        return sb
    quarter_start, quarter_end = _quarter_bounds(tanggal_acuan)
    days_in_quarter = (quarter_end - quarter_start).days + 1
    hari_berjalan_kuartal = (tanggal_acuan - quarter_start).days + 1

    target_map = {}
    if df_target is not None and not df_target.empty and target_col in df_target.columns:
        for _, r in df_target.iterrows():
            target_map[str(r.get("CABANG", "")).strip().upper()] = float(r.get(target_col) or 0)

    sb = sb.copy()
    sb["TARGET"] = sb["CABANG"].map(lambda b: target_map.get(str(b).upper(), 0.0))
    sb["EXPECTED VALUE"] = sb["TARGET"].apply(lambda t: (t / days_in_quarter) * hari_berjalan_kuartal if days_in_quarter else 0.0)
    sb["% PENCAPAIAN"] = sb.apply(lambda r: (r["OMSET S/D HARI INI"] / r["EXPECTED VALUE"]) if r["EXPECTED VALUE"] else None, axis=1)

    total_row = {
        "CABANG": "SMM",
        "OMSET BULAN LALU (FULL)": sb["OMSET BULAN LALU (FULL)"].sum(),
        "OMSET S/D HARI INI": sb["OMSET S/D HARI INI"].sum(),
        "OMSET BULAN LALU S/D HARI SAMA": sb["OMSET BULAN LALU S/D HARI SAMA"].sum(),
        "PERIODE BULAN LALU": sb["PERIODE BULAN LALU"].sum(),
        "PERIODE BULAN INI": sb["PERIODE BULAN INI"].sum(),
        "TARGET": sb["TARGET"].sum(),
        "EXPECTED VALUE": sb["EXPECTED VALUE"].sum(),
    }
    total_row["% PENCAPAIAN"] = (total_row["OMSET S/D HARI INI"] / total_row["EXPECTED VALUE"]) if total_row["EXPECTED VALUE"] else None
    return pd.concat([sb, pd.DataFrame([total_row])], ignore_index=True)


def build_scoreboard_corporate_manual(df_corp: pd.DataFrame, df_target: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    if df_corp is None or df_corp.empty:
        return pd.DataFrame()
    d = df_corp.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame()

    tahun_ini, bulan_ini = tanggal_acuan.year, tanggal_acuan.month
    if bulan_ini == 1:
        tahun_lalu, bulan_lalu = tahun_ini - 1, 12
    else:
        tahun_lalu, bulan_lalu = tahun_ini, bulan_ini - 1

    ini = d[(d["Tahun"] == tahun_ini) & (d["Bulan"] == bulan_ini)]
    lalu = d[(d["Tahun"] == tahun_lalu) & (d["Bulan"] == bulan_lalu)]

    sum_ini = ini.groupby("Cabang")["Omset"].sum()
    sum_lalu = lalu.groupby("Cabang")["Omset"].sum()

    # Data Corporate hanya beresolusi bulanan (Tahun/Bulan, bukan tanggal harian), jadi
    # kumulatif kuartal-berjalan dijumlah dari semua bulan sejak awal kuartal s/d bulan
    # berjalan (tahun yang sama) - dipakai sebagai basis % Pencapaian terhadap target kuartal.
    quarter_start, _ = _quarter_bounds(tanggal_acuan)
    qtd = d[(d["Tahun"] == quarter_start.year) & (d["Bulan"] >= quarter_start.month) & (d["Bulan"] <= bulan_ini)]
    sum_qtd = qtd.groupby("Cabang")["Omset"].sum()

    branches = order_branches(set(sum_ini.index) | set(sum_lalu.index) | set(sum_qtd.index))
    hari_acuan = tanggal_acuan.day
    days_in_bulan_lalu = calendar.monthrange(tahun_lalu, bulan_lalu)[1]
    hari_lalu_cutoff = min(hari_acuan, days_in_bulan_lalu)

    rows = []
    for b in branches:
        oi = float(sum_ini.get(b, 0.0))
        oi_kuartal = float(sum_qtd.get(b, 0.0))
        ol = float(sum_lalu.get(b, 0.0))
        rows.append({
            "CABANG": b, "OMSET BULAN LALU (FULL)": ol, "OMSET S/D HARI INI": oi_kuartal,
            "OMSET BULAN LALU S/D HARI SAMA": ol,
            "PERIODE BULAN LALU": (ol / hari_lalu_cutoff) if hari_lalu_cutoff else 0.0,
            "PERIODE BULAN INI": (oi / hari_acuan) if hari_acuan else 0.0,
        })
    sb = pd.DataFrame(rows)
    return _finalize_scoreboard(sb, df_target, "TargetCorp", tanggal_acuan)


def build_daily_progress(df: pd.DataFrame, kelompok: str, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Hari", "Kumulatif"])
    d = df[df["Kelompok"] == kelompok].copy() if kelompok != "All" else df.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    d = d[(d["Tahun"] == tanggal_acuan.year) & (d["Bulan"] == tanggal_acuan.month) & (d["Tanggal"].dt.day <= tanggal_acuan.day)]
    if d.empty:
        return pd.DataFrame(columns=["Hari", "Kumulatif"])
    daily = d.groupby(d["Tanggal"].dt.day)["Omset"].sum().reindex(range(1, tanggal_acuan.day + 1), fill_value=0)
    return pd.DataFrame({"Hari": daily.index, "Kumulatif": daily.cumsum().values})


def render_daily_progress_chart(df: pd.DataFrame, df_target: pd.DataFrame, target_col: str, selected_branches, tanggal_acuan: date, color: str):
    prog = build_daily_progress(df, "All", tanggal_acuan, selected_branches)
    if prog.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=prog["Hari"], y=prog["Kumulatif"], mode="lines+markers", name="Aktual", line=dict(color=color, width=3)))

    target_total = 0.0
    if df_target is not None and not df_target.empty and target_col in df_target.columns:
        dft = df_target
        if selected_branches:
            dft = dft[dft["CABANG"].isin(selected_branches)]
        target_total = float(dft[target_col].sum())
    if target_total:
        days_in_month = calendar.monthrange(tanggal_acuan.year, tanggal_acuan.month)[1]
        pace = [target_total / days_in_month * h for h in prog["Hari"]]
        fig.add_trace(go.Scatter(x=prog["Hari"], y=pace, mode="lines", name="Target Pace Lurus", line=dict(color="#9ca3af", width=2, dash="dash")))

    fig.update_layout(height=320, margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02))
    fig.update_xaxes(title="Tanggal")
    fig.update_yaxes(title="Omset Kumulatif (Rp)")
    return fig


def build_daily_history(df: pd.DataFrame, kelompok: str, selected_branches=None, n_months: int = 6) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Bulan", "Omset"])
    d = df[df["Kelompok"] == kelompok].copy() if kelompok != "All" else df.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame(columns=["Bulan", "Omset"])
    d["Periode"] = d["Tahun"].astype(str) + "-" + d["Bulan"].astype(str).str.zfill(2)
    g = d.groupby("Periode")["Omset"].sum().reset_index().sort_values("Periode")
    g = g.tail(n_months)
    g["Bulan"] = g["Periode"].apply(lambda p: f"{BULAN_ID[int(p.split('-')[1]) - 1][:3]} {p.split('-')[0]}")
    return g.rename(columns={"Omset": "Omset"})[["Bulan", "Omset"]]


def render_daily_history_chart(hist_df: pd.DataFrame):
    if hist_df.empty:
        return None
    fig = px.bar(hist_df, x="Bulan", y="Omset", text_auto=".2s", color_discrete_sequence=["#1d4ed8"])
    fig.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10))
    fig.update_yaxes(title="Omset (Rp)")
    return fig


def pencapaian_color(pct):
    if pct is None or pd.isna(pct):
        return "#9ca3af"
    if pct >= 1.0:
        return "#16a34a"
    if pct >= 0.85:
        return "#f59e0b"
    return "#dc2626"


def render_progress_ring(pct, track_color: str = "#e5e7eb"):
    """Donut ring persentase pencapaian (0-100%+), dipakai untuk ring KPI di tab Ringkasan."""
    if pct is None or pd.isna(pct):
        pct_display, pct_frac = 0.0, 0.0
    else:
        pct_display, pct_frac = pct * 100, pct
    color = pencapaian_color(pct)
    frac_capped = min(max(pct_frac, 0.0), 1.0)
    fig = go.Figure(data=[go.Pie(
        values=[frac_capped, 1 - frac_capped], hole=0.72, sort=False, direction="clockwise",
        marker=dict(colors=[color, track_color]), textinfo="none", hoverinfo="skip",
    )])
    fig.update_layout(
        showlegend=False, margin=dict(l=10, r=10, t=10, b=10), height=180,
        annotations=[dict(text=f"<b>{format_decimal(pct_display)}%</b>", x=0.5, y=0.5, font_size=20, showarrow=False, font_color="#111827")],
    )
    return fig


def render_contribution_pie(omset_service_val, omset_gadget_val, omset_corp_val):
    labels, values = [], []
    for label, val in [("Service", omset_service_val), ("Gadget & Aksesoris", omset_gadget_val), ("Marketing Corporate", omset_corp_val)]:
        if val and val > 0:
            labels.append(label)
            values.append(val)
    if not values:
        return None
    fig = px.pie(names=labels, values=values, hole=0.45, color_discrete_sequence=["#0f766e", "#6d28d9", "#f59e0b"])
    fig.update_traces(textinfo="label+percent")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=20, b=10), showlegend=True)
    return fig


MONEY_COLS = [
    "OMSET BULAN LALU (FULL)", "OMSET S/D HARI INI", "OMSET BULAN LALU S/D HARI SAMA",
    "PERIODE BULAN LALU", "PERIODE BULAN INI", "TARGET", "EXPECTED VALUE",
]

_SCOREBOARD_GROUPS = ["Service", "Gadget & Aksesoris", "All"]


def _cell_color(col: str, val, row: dict) -> str:
    if col == "% PENCAPAIAN":
        return pencapaian_color(val)
    if col == "PERIODE BULAN INI":
        lalu = row.get("PERIODE BULAN LALU")
        if lalu is not None and not pd.isna(lalu) and lalu > 0 and val is not None and not pd.isna(val):
            return "#16a34a" if val >= lalu else "#dc2626"
    return "#111827"


def render_scoreboard_html(sb: pd.DataFrame, name_label: str = "CABANG") -> str:
    if sb.empty:
        return "<p>Tidak ada data untuk periode ini.</p>"
    cols = [name_label] + [c for c in MONEY_COLS if c in sb.columns] + (["% PENCAPAIAN"] if "% PENCAPAIAN" in sb.columns else [])
    header_html = "".join(f'<th style="padding:8px 10px;text-align:right;background:#1e3a8a;color:white;font-size:12px;position:sticky;top:0;">{c}</th>' if c != name_label else f'<th style="padding:8px 10px;text-align:left;background:#1e3a8a;color:white;font-size:12px;position:sticky;top:0;left:0;">{c}</th>' for c in cols)

    rows_html = ""
    for _, r in sb.iterrows():
        is_total = str(r.get(name_label, "")).upper() in _SALES_TOTAL_LABELS
        row_bg = "#eff6ff" if is_total else "#ffffff"
        fw = "700" if is_total else "400"
        cells = f'<td style="padding:7px 10px;font-weight:{fw};background:{row_bg};position:sticky;left:0;">{r.get(name_label, "")}</td>'
        for c in cols[1:]:
            val = r.get(c)
            if c == "% PENCAPAIAN":
                text = format_percent(val)
            else:
                text = format_rupiah(val)
            color = _cell_color(c, val, r)
            cells += f'<td style="padding:7px 10px;text-align:right;font-weight:{fw};background:{row_bg};color:{color};">{text}</td>'
        rows_html += f"<tr>{cells}</tr>"

    return (
        '<div style="overflow-x:auto;max-height:420px;overflow-y:auto;border:1px solid #e5e7eb;border-radius:8px;">'
        f'<table style="border-collapse:collapse;width:100%;font-size:12.5px;">'
        f"<thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table></div>"
    )


def render_kpi_card(label: str, value: float, color1: str, color2: str, icon: str, pct=None) -> str:
    pct_html = ""
    if pct is not None and not pd.isna(pct):
        pct_html = f'<div style="font-size:12px;opacity:.85;margin-top:4px;">Pencapaian: {format_percent(pct)}</div>'
    return f"""
    <div style="background:linear-gradient(135deg,{color1},{color2});border-radius:14px;
                padding:16px 18px;color:white;box-shadow:0 2px 8px rgba(0,0,0,.12);height:100%;">
      <div style="font-size:26px;line-height:1;">{icon}</div>
      <div style="font-size:13px;opacity:.9;margin-top:8px;">{label}</div>
      <div style="font-size:20px;font-weight:700;margin-top:2px;">{format_rupiah(value)}</div>
      {pct_html}
    </div>
    """


# --------------------------------------------------------------------------------------
# Agregasi & rendering untuk tab 6 Pilar MFlash
# --------------------------------------------------------------------------------------

def build_pilar_summary(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    """Omset, Gross Profit & Qty per Pilar bulan berjalan (s/d tanggal acuan) vs bulan lalu
    (s/d tanggal yang sama), dipakai untuk KPI card & indikator tren naik/turun per pilar."""
    empty = pd.DataFrame(columns=["Pilar", "OmsetBulanIni", "OmsetBulanLalu", "PctChange", "GrossProfitBulanIni", "QtyBulanIni"])
    if df is None or df.empty or "Pilar" not in df.columns:
        return empty
    d = df.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return empty

    tahun_ini, bulan_ini = tanggal_acuan.year, tanggal_acuan.month
    if bulan_ini == 1:
        tahun_lalu, bulan_lalu = tahun_ini - 1, 12
    else:
        tahun_lalu, bulan_lalu = tahun_ini, bulan_ini - 1
    hari_acuan = tanggal_acuan.day
    days_in_bulan_lalu = calendar.monthrange(tahun_lalu, bulan_lalu)[1]
    hari_lalu_cutoff = min(hari_acuan, days_in_bulan_lalu)

    ini = d[(d["Tahun"] == tahun_ini) & (d["Bulan"] == bulan_ini) & (d["Tanggal"].dt.day <= hari_acuan)]
    lalu = d[(d["Tahun"] == tahun_lalu) & (d["Bulan"] == bulan_lalu) & (d["Tanggal"].dt.day <= hari_lalu_cutoff)]

    sum_ini = ini.groupby("Pilar")["Omset"].sum()
    sum_lalu = lalu.groupby("Pilar")["Omset"].sum()
    gp_ini = ini.groupby("Pilar")["GrossProfit"].sum() if "GrossProfit" in ini.columns else pd.Series(dtype=float)
    qty_ini = ini.groupby("Pilar")["Qty"].sum() if "Qty" in ini.columns else pd.Series(dtype=float)

    rows = []
    for p in PILAR_ORDER:
        oi = float(sum_ini.get(p, 0.0))
        ol = float(sum_lalu.get(p, 0.0))
        pct = ((oi - ol) / ol) if ol else None
        rows.append({
            "Pilar": p, "OmsetBulanIni": oi, "OmsetBulanLalu": ol, "PctChange": pct,
            "GrossProfitBulanIni": float(gp_ini.get(p, 0.0)), "QtyBulanIni": float(qty_ini.get(p, 0.0)),
        })
    return pd.DataFrame(rows)


def build_pilar_by_branch(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    """Cross-tab Cabang x Pilar untuk omset bulan berjalan (s/d tanggal acuan), plus baris
    & kolom TOTAL."""
    if df is None or df.empty or "Pilar" not in df.columns:
        return pd.DataFrame()
    d = df.copy()
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    tahun_ini, bulan_ini, hari_acuan = tanggal_acuan.year, tanggal_acuan.month, tanggal_acuan.day
    d = d[(d["Tahun"] == tahun_ini) & (d["Bulan"] == bulan_ini) & (d["Tanggal"].dt.day <= hari_acuan)]
    if d.empty:
        return pd.DataFrame()

    pivot = d.pivot_table(index="Cabang", columns="Pilar", values="Omset", aggfunc="sum", fill_value=0.0)
    for p in PILAR_ORDER:
        if p not in pivot.columns:
            pivot[p] = 0.0
    pivot = pivot[PILAR_ORDER]
    pivot["TOTAL"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()

    branches_present = order_branches(pivot["Cabang"].tolist())
    pivot["Cabang"] = pd.Categorical(pivot["Cabang"], categories=branches_present, ordered=True)
    pivot = pivot.sort_values("Cabang").reset_index(drop=True)
    pivot["Cabang"] = pivot["Cabang"].astype(str)

    total_row = {"Cabang": "TOTAL"}
    for p in PILAR_ORDER + ["TOTAL"]:
        total_row[p] = pivot[p].sum()
    pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)
    return pivot


def generate_pilar_insights(pilar_by_branch: pd.DataFrame) -> list:
    """Insight kualitas data: tandai cabang dengan porsi omset 'Belum Dikategorikan' yang
    tinggi (>= 40% dari total omset cabang tsb) supaya diingatkan mengisi kolom KATEGORI
    PILAR dengan lebih rapi di faktur penjualan."""
    insights = []
    if pilar_by_branch.empty or "Belum Dikategorikan" not in pilar_by_branch.columns:
        return insights
    rows = pilar_by_branch[pilar_by_branch["Cabang"] != "TOTAL"]
    for _, r in rows.iterrows():
        total = r.get("TOTAL", 0.0)
        if not total:
            continue
        belum = r.get("Belum Dikategorikan", 0.0)
        pct = belum / total
        if pct >= 0.4:
            insights.append({
                "level": "warn" if pct < 0.7 else "bad",
                "title": str(r["Cabang"]),
                "text": f"{format_percent(pct)} dari omset cabang ini ({format_rupiah(belum)} dari "
                        f"{format_rupiah(total)}) belum diberi kategori di kolom KATEGORI PILAR. "
                        f"Ingatkan admin cabang untuk mengisi kolom ini di setiap transaksi supaya "
                        f"data 6 Pilar akurat.",
            })
    return insights


def render_pilar_kpi_card(pilar: str, value: float, pct_change, gross_profit=None, qty=None) -> str:
    color = PILAR_COLORS.get(pilar, "#374151")
    icon = PILAR_ICONS.get(pilar, "📊")
    trend_html = ""
    if pct_change is not None and not pd.isna(pct_change):
        arrow = "▲" if pct_change >= 0 else "▼"
        trend_html = (
            f'<div style="font-size:11.5px;margin-top:4px;opacity:.9;">{arrow} '
            f'{format_percent(abs(pct_change))} vs bulan lalu</div>'
        )
    gp_html = ""
    if gross_profit is not None:
        margin_html = ""
        if value:
            margin_html = f" ({format_percent(gross_profit / value)})"
        gp_html = (
            f'<div style="font-size:11.5px;margin-top:6px;padding-top:6px;'
            f'border-top:1px solid rgba(255,255,255,.25);opacity:.95;">'
            f'💹 Gross Profit: {format_rupiah(gross_profit)}{margin_html}</div>'
        )
    qty_html = ""
    if qty is not None:
        qty_html = f'<div style="font-size:11.5px;margin-top:3px;opacity:.95;">📦 Qty: {format_number(qty)} unit</div>'
    label = _pilar_label(pilar)
    return f"""
    <div style="background:{color};border-radius:14px;padding:14px 16px;color:white;
                box-shadow:0 2px 8px rgba(0,0,0,.12);height:100%;">
      <div style="font-size:22px;line-height:1;">{icon}</div>
      <div style="font-size:12px;opacity:.9;margin-top:6px;">{label}</div>
      <div style="font-size:17px;font-weight:700;margin-top:2px;">{format_rupiah(value)}</div>
      {trend_html}
      {gp_html}
      {qty_html}
    </div>
    """


def render_pilar_table_html(pilar_by_branch: pd.DataFrame) -> str:
    if pilar_by_branch.empty:
        return "<p>Tidak ada data untuk periode ini.</p>"
    display_cols = ["Cabang"] + [p for p in PILAR_ORDER if p in pilar_by_branch.columns] + ["TOTAL"]
    header_html = "".join(
        f'<th style="padding:8px 10px;text-align:{"left" if c == "Cabang" else "right"};'
        f'background:#1e3a8a;color:white;font-size:11px;white-space:nowrap;position:sticky;top:0;">'
        f'{c if c in ("Cabang", "TOTAL") else _pilar_label(c)}</th>'
        for c in display_cols
    )
    rows_html = ""
    for _, r in pilar_by_branch.iterrows():
        is_total = r["Cabang"] == "TOTAL"
        bg = "#eff6ff" if is_total else "#ffffff"
        fw = "700" if is_total else "400"
        cells = f'<td style="padding:7px 10px;font-weight:{fw};background:{bg};position:sticky;left:0;">{r["Cabang"]}</td>'
        for c in display_cols[1:]:
            cells += f'<td style="padding:7px 10px;text-align:right;font-weight:{fw};background:{bg};">{format_rupiah(r.get(c, 0))}</td>'
        rows_html += f"<tr>{cells}</tr>"
    return (
        '<div style="overflow-x:auto;max-height:420px;overflow-y:auto;border:1px solid #e5e7eb;border-radius:8px;">'
        f'<table style="border-collapse:collapse;width:100%;font-size:12px;">'
        f"<thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table></div>"
    )


def render_pilar_summary_table_html(pilar_summary: pd.DataFrame) -> str:
    """Tabel ringkas per Pilar: Omset, Gross Profit, Margin %, dan Qty (Qty relevan
    terutama untuk Penyewaan Corporate & Maintenance Corporate, tapi ditampilkan untuk
    semua pilar supaya konsisten)."""
    if pilar_summary.empty:
        return "<p>Tidak ada data untuk periode ini.</p>"
    cols = ["Pilar", "Omset", "Gross Profit", "Margin %", "Qty"]
    header_html = "".join(
        f'<th style="padding:8px 10px;text-align:{"left" if c == "Pilar" else "right"};'
        f'background:#1e3a8a;color:white;font-size:12px;">{c}</th>' for c in cols
    )
    rows_html = ""
    for _, r in pilar_summary.iterrows():
        omset = r["OmsetBulanIni"]
        gp = r.get("GrossProfitBulanIni", 0.0)
        qty = r.get("QtyBulanIni", 0.0)
        margin = (gp / omset) if omset else None
        show_qty = r["Pilar"] in _PILAR_SHOW_QTY
        qty_text = format_number(qty) if show_qty else "-"
        cells = (
            f'<td style="padding:7px 10px;">{PILAR_ICONS.get(r["Pilar"], "")} {_pilar_label(r["Pilar"])}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{format_rupiah(omset)}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{format_rupiah(gp)}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{format_percent(margin) if margin is not None else "-"}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{qty_text}</td>'
        )
        rows_html += f"<tr>{cells}</tr>"
    return (
        '<div style="overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;">'
        f'<table style="border-collapse:collapse;width:100%;font-size:12.5px;">'
        f"<thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table></div>"
    )


# --------------------------------------------------------------------------------------
# Auto-ekstrak Target & Scoreboard Marketing Corporate dari sheet "Scoreboard" (kalau ada)
# --------------------------------------------------------------------------------------

_SECTION_PATTERN = re.compile(r"SCOREBOARD\s+OMSET\s+(SERVICE|GADGET.*AKSESORIS|MARKETING\s+CORPORATE)", re.IGNORECASE)

_TARGET_DF_COLUMNS = ["CABANG", "TargetService", "TargetGadget", "TargetAll", "TargetCorp"]


def _empty_target_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_TARGET_DF_COLUMNS)


def _has_target_signal(df: pd.DataFrame) -> bool:
    """True kalau df target punya minimal satu angka target > 0 di salah satu kolom
    (TargetService/TargetGadget/TargetAll/TargetCorp). Dipakai untuk membedakan hasil
    auto-ekstrak yang BENERAN nemu target vs yang cuma nemu baris CABANG kosong (0 semua) -
    supaya file Target manual yang diupload user tidak diabaikan begitu saja."""
    if df is None or df.empty:
        return False
    numeric_cols = [c for c in _TARGET_DF_COLUMNS if c != "CABANG" and c in df.columns]
    if not numeric_cols:
        return False
    return bool(df[numeric_cols].fillna(0).to_numpy().sum() > 0)


def _read_scoreboard_sections(wb):
    """Cari sheet 'Scoreboard' dan kembalikan dict {section_key: list of rows (as tuples)}
    berdasarkan header section 'SCOREBOARD OMSET SERVICE/GADGET.../MARKETING CORPORATE'."""
    sheet_name = next((n for n in wb.sheetnames if n.strip().lower() == "scoreboard"), None)
    if sheet_name is None:
        return {}
    ws = wb[sheet_name]
    sections = {}
    current_key = None
    header = None
    for row in ws.iter_rows(values_only=True):
        first_cell = row[0] if row else None
        if first_cell and isinstance(first_cell, str):
            m = _SECTION_PATTERN.search(first_cell)
            if m:
                key = m.group(1).upper().replace(" ", "")
                current_key = key
                header = None
                sections[current_key] = []
                continue
        if current_key is None:
            continue
        if header is None:
            if first_cell and str(first_cell).strip().upper() in ("CABANG", "NAMA"):
                header = [str(c).strip().upper() if c else "" for c in row]
            continue
        if all(c is None for c in row):
            current_key = None
            continue
        sections[current_key].append((header, row))
    return sections


def extract_scoreboard_target(wb) -> pd.DataFrame:
    sections = _read_scoreboard_sections(wb)
    target_map = {}

    def _target_col_idx(header):
        for i, h in enumerate(header):
            if "TARGET" in h:
                return i
        return None

    key_to_col = {"SERVICE": "TargetService", "GADGET&AKSESORIS": "TargetGadget", "MARKETINGCORPORATE": "TargetCorp"}
    for key, col in key_to_col.items():
        rows = sections.get(key, [])
        for header, row in rows:
            t_idx = _target_col_idx(header)
            if t_idx is None or t_idx >= len(row):
                continue
            cabang = row[0]
            if not cabang or str(cabang).strip().upper() in _SALES_TOTAL_LABELS:
                continue
            val = row[t_idx]
            try:
                val = float(val) if val is not None else 0.0
            except (TypeError, ValueError):
                val = 0.0
            cabang_u = str(cabang).strip().upper()
            target_map.setdefault(cabang_u, {}).__setitem__(col, val)

    if not target_map:
        return _empty_target_df()

    rows = []
    for cabang, vals in target_map.items():
        r = {"CABANG": cabang}
        for col in ["TargetService", "TargetGadget", "TargetCorp"]:
            r[col] = vals.get(col, 0.0)
        r["TargetAll"] = r["TargetService"] + r["TargetGadget"]
        rows.append(r)
    return pd.DataFrame(rows)


def extract_scoreboard_target_all(main_dir: str) -> pd.DataFrame:
    files = sorted(f for f in os.listdir(main_dir) if f.lower().endswith(".xlsx")) if os.path.isdir(main_dir) else []
    frames = []
    for fname in files:
        fpath = os.path.join(main_dir, fname)
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
            df = extract_scoreboard_target(wb)
            if not df.empty:
                frames.append(df)
        except Exception:  # noqa: BLE001
            continue
    if not frames:
        return _empty_target_df()
    combined = pd.concat(frames, ignore_index=True)
    return combined.groupby("CABANG", as_index=False).sum(numeric_only=True)


def extract_scoreboard_snapshot_date(wb):
    sheet_name = next((n for n in wb.sheetnames if n.strip().lower() == "scoreboard"), None)
    if sheet_name is None:
        return None
    ws = wb[sheet_name]
    for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
        for cell in row:
            if isinstance(cell, datetime):
                return cell.date()
    return None


def extract_scoreboard_corporate(wb) -> pd.DataFrame:
    sections = _read_scoreboard_sections(wb)
    rows_data = sections.get("MARKETINGCORPORATE", [])
    if not rows_data:
        return pd.DataFrame()
    snapshot_date = extract_scoreboard_snapshot_date(wb) or date.today()
    rows = []
    for header, row in rows_data:
        cabang = row[0] if row else None
        if not cabang or str(cabang).strip().upper() in _SALES_TOTAL_LABELS:
            continue
        omset_idx = None
        for i, h in enumerate(header):
            if "OMSET" in h and "TARGET" not in h:
                omset_idx = i
                break
        if omset_idx is None or omset_idx >= len(row):
            continue
        val = row[omset_idx]
        try:
            val = float(val) if val is not None else 0.0
        except (TypeError, ValueError):
            val = 0.0
        rows.append({"Cabang": str(cabang).strip().upper(), "Tahun": snapshot_date.year, "Bulan": snapshot_date.month, "Omset": val})
    return pd.DataFrame(rows)


def extract_scoreboard_corporate_all(main_dir: str):
    files = sorted(f for f in os.listdir(main_dir) if f.lower().endswith(".xlsx")) if os.path.isdir(main_dir) else []
    frames = []
    latest_date = None
    for fname in files:
        fpath = os.path.join(main_dir, fname)
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
            df = extract_scoreboard_corporate(wb)
            snap = extract_scoreboard_snapshot_date(wb)
            if snap and (latest_date is None or snap > latest_date):
                latest_date = snap
            if not df.empty:
                frames.append(df)
        except Exception:  # noqa: BLE001
            continue
    if not frames:
        return pd.DataFrame(), None
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.groupby(["Cabang", "Tahun", "Bulan"], as_index=False)["Omset"].sum()
    return combined, latest_date


# --------------------------------------------------------------------------------------
# Ledger permanen (histori upload) - supaya progres harian & histori bulanan tidak
# hilang meski file lama diganti/dihapus dari sidebar
# --------------------------------------------------------------------------------------

_HISTORY_LOG_COLUMNS = ["Cabang", "Tanggal", "Tahun", "Bulan", "Kelompok", "Omset"]
_CORP_HISTORY_LOG_COLUMNS = ["Cabang", "Tahun", "Bulan", "Omset"]


def _read_log(path: str, cols: list) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame(columns=cols)
    try:
        df = pd.read_csv(path)
        if "Tanggal" in df.columns:
            df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        return df
    except Exception:  # noqa: BLE001
        return pd.DataFrame(columns=cols)


def _upsert_log(path: str, repo_path: str, new_df: pd.DataFrame, key_cols: list, cols: list):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    existing = _read_log(path, cols)
    if new_df is None or new_df.empty:
        return
    combined = pd.concat([existing, new_df[cols]], ignore_index=True) if not existing.empty else new_df[cols].copy()
    combined = combined.drop_duplicates(subset=key_cols, keep="last")
    combined.to_csv(path, index=False)
    if _GH_ENABLED:
        status = github_upload_file(path, repo_path, "Update ledger histori")
        if status == "error" and "_gh_warnings" in st.session_state:
            st.session_state["_gh_warnings"].add(f"Gagal backup {os.path.basename(path)} ke GitHub.")


def build_upload_log(df_main: pd.DataFrame):
    if df_main.empty:
        return
    daily = df_main.groupby(["Cabang", "Tanggal", "Tahun", "Bulan", "Kelompok"], as_index=False)["Omset"].sum()
    _upsert_log(HISTORY_LOG_PATH, "data/history_log.csv", daily, ["Cabang", "Tanggal", "Kelompok"], _HISTORY_LOG_COLUMNS)


def build_corp_upload_log(df_corp: pd.DataFrame):
    if df_corp is None or df_corp.empty:
        return
    monthly = df_corp.groupby(["Cabang", "Tahun", "Bulan"], as_index=False)["Omset"].sum()
    _upsert_log(CORP_HISTORY_LOG_PATH, "data/corp_history_log.csv", monthly, ["Cabang", "Tahun", "Bulan"], _CORP_HISTORY_LOG_COLUMNS)


def compute_corp_hari_ini(corp_log: pd.DataFrame, tanggal_acuan: date) -> pd.DataFrame:
    if corp_log.empty:
        return pd.DataFrame(columns=_CORP_HISTORY_LOG_COLUMNS)
    return corp_log[(corp_log["Tahun"] == tanggal_acuan.year) & (corp_log["Bulan"] == tanggal_acuan.month)]


# --------------------------------------------------------------------------------------
# Loader Target manual (dipakai kalau file yang diupload tidak punya kolom TARGET
# di sheet Scoreboard). Nilai TARGET yang diinput di sini juga diasumsikan sebagai total
# 1 KUARTAL PENUH (3 bulan), sama seperti target yang diekstrak otomatis dari file Excel.
# --------------------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_target_data(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes))
    raw.columns = [str(c).strip().upper() for c in raw.columns]
    if "CABANG" not in raw.columns:
        raise ValueError("Kolom CABANG tidak ditemukan di file Target.")
    df = pd.DataFrame()
    df["CABANG"] = raw["CABANG"].astype(str).str.strip().str.upper()
    # Pakai startswith (bukan exact match) supaya header dengan sufiks seperti
    # "TARGET SERVICE (PER KUARTAL)" tetap terdeteksi sebagai kolom "TARGET SERVICE".
    for src, dst in [("TARGET SERVICE", "TargetService"), ("TARGET GADGET", "TargetGadget"),
                      ("TARGET ALL", "TargetAll"), ("TARGET CORPORATE", "TargetCorp"), ("TARGET CORP", "TargetCorp")]:
        if dst in df.columns:
            continue
        match_col = next((c for c in raw.columns if c.startswith(src)), None)
        if match_col is not None:
            df[dst] = pd.to_numeric(raw[match_col], errors="coerce").fillna(0)
    for col in _TARGET_DF_COLUMNS:
        if col not in df.columns:
            df[col] = 0.0
    if (df["TargetAll"] == 0).all():
        df["TargetAll"] = df["TargetService"] + df["TargetGadget"]
    return df[_TARGET_DF_COLUMNS]


def make_target_template() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Target"
    ws.append(["Cabang", "Target Service", "Target Gadget", "Target All", "Target Corporate"])
    for b in BRANCH_ORDER[:3]:
        ws.append([b, 150000000, 180000000, 330000000, 45000000])
    for col, w in zip("ABCDE", [16, 16, 16, 14, 18]):
        ws.column_dimensions[col].width = w

    info = wb.create_sheet("Petunjuk")
    info.append(["Petunjuk pengisian Target"])
    info.append([""])
    info.append(["- Nilai target di sini adalah TOTAL untuk 1 KUARTAL PENUH (3 bulan), BUKAN target bulanan."])
    info.append(["- Kuartal MFlash: Jan-Mar, Apr-Jun, Jul-Sep, Okt-Des."])
    info.append(["- Dashboard otomatis membagi target ini dengan jumlah hari di kuartal berjalan untuk"])
    info.append(["  menghitung Expected Value & % Pencapaian harian."])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ========================================================================================
# UI
# ========================================================================================

os.makedirs(MAIN_DATA_DIR, exist_ok=True)
os.makedirs(ADS_DATA_DIR, exist_ok=True)
os.makedirs(WALKIN_DATA_DIR, exist_ok=True)

if "_gh_synced_v1" not in st.session_state:
    sync_data_from_github()
    st.session_state["_gh_synced_v1"] = True
if "_gh_warnings" not in st.session_state:
    st.session_state["_gh_warnings"] = set()

logo_col, title_col = st.columns([1, 4])
with logo_col:
    if LOGO_BASE64 and "PLACEHOLDER" not in LOGO_BASE64:
        st.image(io.BytesIO(base64.b64decode(LOGO_BASE64)), width=90)
with title_col:
    st.markdown("## Dashboard Omset MFlash")
    st.caption("Monitoring Omset, Iklan, Walk-in & 6 Pilar — 18 Cabang")

# ---------------------------- Sidebar: Kelola Data ----------------------------
with st.sidebar:
    st.markdown("### 📂 Kelola Data")

    st.markdown("**1️⃣ Data Omset Utama**")
    st.caption(
        "File master (Faktur Penjualan) ATAU file per-cabang (Rincian Faktur Penjualan). "
        "Khusus untuk tab 6 Pilar, gunakan file per-cabang (Rincian Faktur Penjualan) karena "
        "file master lama tidak punya kolom KATEGORI PILAR."
    )
    main_files = st.file_uploader(
        "Upload file Omset", type=["xlsx"], accept_multiple_files=True, key="upl_main",
    )
    if main_files:
        for uf in main_files:
            if len(os.listdir(MAIN_DATA_DIR)) >= MAX_MAIN_FILES:
                st.warning(f"Batas maksimal {MAX_MAIN_FILES} file Omset tercapai.")
                break
            fname = sanitize_filename(uf.name)
            fpath = os.path.join(MAIN_DATA_DIR, fname)
            with open(fpath, "wb") as f:
                f.write(uf.getbuffer())
            status = github_upload_file(fpath, f"data/main/{fname}", f"Upload {fname}")
            if status == "error":
                st.session_state["_gh_warnings"].add(f"Gagal backup {fname} ke GitHub.")
        st.success(f"{len(main_files)} file omset tersimpan.")

    existing_main = sorted(f for f in os.listdir(MAIN_DATA_DIR) if f.lower().endswith(".xlsx"))
    if existing_main:
        with st.expander(f"File Omset tersimpan ({len(existing_main)})"):
            for fname in existing_main:
                c1, c2 = st.columns([4, 1])
                c1.caption(fname)
                if c2.button("🗑️", key=f"del_main_{fname}"):
                    os.remove(os.path.join(MAIN_DATA_DIR, fname))
                    github_delete_file(f"data/main/{fname}", f"Hapus {fname}")
                    st.rerun()

    st.divider()
    st.markdown("**2️⃣ Data Iklan**")
    st.caption("Export Campaigns dari Meta Ads Manager (.xlsx).")
    ads_files = st.file_uploader(
        "Upload file Iklan", type=["xlsx"], accept_multiple_files=True, key="upl_ads",
    )
    if ads_files:
        for uf in ads_files:
            if len(os.listdir(ADS_DATA_DIR)) >= MAX_ADS_FILES:
                st.warning(f"Batas maksimal {MAX_ADS_FILES} file Iklan tercapai.")
                break
            fname = sanitize_filename(uf.name)
            fpath = os.path.join(ADS_DATA_DIR, fname)
            with open(fpath, "wb") as f:
                f.write(uf.getbuffer())
            status = github_upload_file(fpath, f"data/ads/{fname}", f"Upload {fname}")
            if status == "error":
                st.session_state["_gh_warnings"].add(f"Gagal backup {fname} ke GitHub.")
        st.success(f"{len(ads_files)} file iklan tersimpan.")

    existing_ads = sorted(f for f in os.listdir(ADS_DATA_DIR) if f.lower().endswith(".xlsx"))
    if existing_ads:
        with st.expander(f"File Iklan tersimpan ({len(existing_ads)})"):
            for fname in existing_ads:
                c1, c2 = st.columns([4, 1])
                c1.caption(fname)
                if c2.button("🗑️", key=f"del_ads_{fname}"):
                    os.remove(os.path.join(ADS_DATA_DIR, fname))
                    github_delete_file(f"data/ads/{fname}", f"Hapus {fname}")
                    st.rerun()

    st.divider()
    st.markdown("**3️⃣ Data Walk-in**")
    st.caption("Rincian Pengiriman Pesanan per cabang (.xlsx).")
    walkin_files = st.file_uploader(
        "Upload file Walk-in", type=["xlsx"], accept_multiple_files=True, key="upl_walkin",
    )
    if walkin_files:
        for uf in walkin_files:
            if len(os.listdir(WALKIN_DATA_DIR)) >= MAX_WALKIN_FILES:
                st.warning(f"Batas maksimal {MAX_WALKIN_FILES} file Walk-in tercapai.")
                break
            fname = sanitize_filename(uf.name)
            fpath = os.path.join(WALKIN_DATA_DIR, fname)
            with open(fpath, "wb") as f:
                f.write(uf.getbuffer())
            status = github_upload_file(fpath, f"data/walkin/{fname}", f"Upload {fname}")
            if status == "error":
                st.session_state["_gh_warnings"].add(f"Gagal backup {fname} ke GitHub.")
        st.success(f"{len(walkin_files)} file walk-in tersimpan.")

    existing_walkin = sorted(f for f in os.listdir(WALKIN_DATA_DIR) if f.lower().endswith(".xlsx"))
    if existing_walkin:
        with st.expander(f"File Walk-in tersimpan ({len(existing_walkin)})"):
            for fname in existing_walkin:
                c1, c2 = st.columns([4, 1])
                c1.caption(fname)
                if c2.button("🗑️", key=f"del_walkin_{fname}"):
                    os.remove(os.path.join(WALKIN_DATA_DIR, fname))
                    github_delete_file(f"data/walkin/{fname}", f"Hapus {fname}")
                    st.rerun()

    st.divider()
    st.markdown("**4️⃣ Target (opsional)**")
    st.caption("Hanya perlu diupload kalau file Omset tidak punya kolom TARGET di sheet Scoreboard. Nilai target adalah total per KUARTAL (3 bulan), bukan bulanan. File yang diupload di sini SELALU diprioritaskan kalau isinya ada angka target.")
    target_file = st.file_uploader("Upload file Target", type=["xlsx"], key="upl_target")
    if target_file:
        fpath = TARGET_DATA_PATH
        with open(fpath, "wb") as f:
            f.write(target_file.getbuffer())
        github_upload_file(fpath, "data/target.xlsx", "Upload target.xlsx")
        st.success("File Target tersimpan.")
    st.download_button("⬇️ Download Template Target", data=make_target_template(),
                        file_name="template_target.xlsx", key="dl_target_template")

    st.divider()
    st.markdown("**5️⃣ Marketing Corporate (opsional)**")
    st.caption("Hanya perlu diupload kalau file Omset tidak punya section Scoreboard Marketing Corporate.")
    corp_file = st.file_uploader("Upload file Corporate", type=["xlsx"], key="upl_corp")
    if corp_file:
        fpath = CORPORATE_DATA_PATH
        with open(fpath, "wb") as f:
            f.write(corp_file.getbuffer())
        github_upload_file(fpath, "data/corporate.xlsx", "Upload corporate.xlsx")
        st.success("File Corporate tersimpan.")
    st.download_button("⬇️ Download Template Corporate", data=make_corporate_template(),
                        file_name="template_corporate.xlsx", key="dl_corp_template")

    if not _GH_ENABLED:
        st.divider()
        st.caption("ℹ️ Auto-backup GitHub belum aktif. Lihat README untuk setup (opsional, supaya data tidak hilang saat app sleep/redeploy).")
    for w in st.session_state.get("_gh_warnings", set()):
        st.warning(w)

# ---------------------------- Load semua data ----------------------------
df_main, _main_load_errors = load_all_main_data(MAIN_DATA_DIR)
df_ads, _ads_load_errors = load_all_ads_data(ADS_DATA_DIR)
df_walkin, _walkin_load_errors = load_all_walkin_data(WALKIN_DATA_DIR)

_all_errors = _main_load_errors + _ads_load_errors + _walkin_load_errors
if _all_errors:
    with st.expander(f"⚠️ {len(_all_errors)} file gagal dimuat — klik untuk detail"):
        for e in _all_errors:
            st.caption(e)

# ---------------------------- Resolusi Target & Scoreboard Corporate ----------------------------
# Prioritas TARGET: file manual yang diupload user (kalau isinya beneran ada angka target)
# SELALU menang dibanding hasil auto-ekstrak dari sheet Scoreboard - supaya auto-ekstrak
# yang "nemu" baris CABANG tapi gagal membaca kolom TARGET (jadi 0 semua) tidak diam-diam
# menutupi/mengabaikan target manual yang sudah user upload dengan benar.
df_target_auto = extract_scoreboard_target_all(MAIN_DATA_DIR)

df_target_manual = _empty_target_df()
if os.path.exists(TARGET_DATA_PATH):
    try:
        with open(TARGET_DATA_PATH, "rb") as f:
            df_target_manual = load_target_data(f.read())
    except Exception:  # noqa: BLE001
        df_target_manual = _empty_target_df()

if _has_target_signal(df_target_manual):
    df_target = df_target_manual
elif _has_target_signal(df_target_auto):
    df_target = df_target_auto
else:
    df_target = _empty_target_df()

corp_scoreboard_df, corp_scoreboard_tanggal = extract_scoreboard_corporate_all(MAIN_DATA_DIR)

df_corp_manual = pd.DataFrame()
if os.path.exists(CORPORATE_DATA_PATH):
    try:
        with open(CORPORATE_DATA_PATH, "rb") as f:
            df_corp_manual = load_corporate_data(f.read())
    except Exception:  # noqa: BLE001
        df_corp_manual = pd.DataFrame()

df_corp = corp_scoreboard_df if not corp_scoreboard_df.empty else df_corp_manual

# PENTING: dashboard TIDAK boleh berhenti render total (st.stop()) hanya karena data
# Omset belum ada - kalau user cuma upload Walk-in/Iklan/Corporate dulu, tab-tab lain
# (Iklan, Walk-in) harus tetap bisa dipakai. Tab Ringkasan, Scoreboard, & 6 Pilar yang
# memang tergantung data Omset akan menampilkan status kosong secara wajar (sudah
# ditangani masing-masing lewat pengecekan .empty di dalam tab), bukan menghentikan
# seluruh app.
if df_main.empty:
    st.info(
        "ℹ️ Data Omset belum diupload — tab **Ringkasan**, **Scoreboard**, & **6 Pilar** akan "
        "kosong sampai file Omset diupload di sidebar. Tab **Iklan** & **Walk-in** tetap bisa "
        "dipakai kalau datanya sudah diupload."
    )

# ---------------------------- Ledger permanen ----------------------------
build_upload_log(df_main)
build_corp_upload_log(df_corp)
history_log = _read_log(HISTORY_LOG_PATH, _HISTORY_LOG_COLUMNS)
corp_history_log = _read_log(CORP_HISTORY_LOG_PATH, _CORP_HISTORY_LOG_COLUMNS)

# ---------------------------- Filter ----------------------------
st.markdown("---")
f1, f2 = st.columns([1, 2])
with f1:
    if not df_main.empty:
        max_date = df_main["Tanggal"].max().date()
    elif not df_walkin.empty:
        max_date = df_walkin["Tanggal"].max().date()
    else:
        max_date = date.today()
    tanggal_acuan = st.date_input("Tanggal Acuan", value=max_date, key="tanggal_acuan")
with f2:
    # Gabungkan opsi cabang dari SEMUA sumber data (Omset/Iklan/Walk-in/Corporate) -
    # bukan cuma df_main - supaya filter cabang tetap berfungsi walau data Omset belum
    # diupload sama sekali (mis. user baru upload Walk-in atau Iklan duluan).
    branch_pool = set()
    if not df_main.empty:
        branch_pool |= set(df_main["Cabang"].dropna().unique().tolist())
    if not df_ads.empty:
        branch_pool |= set(df_ads["Cabang"].dropna().unique().tolist())
    if not df_walkin.empty:
        branch_pool |= set(df_walkin["Cabang"].dropna().unique().tolist())
    if df_corp is not None and not df_corp.empty and "Cabang" in df_corp.columns:
        branch_pool |= set(df_corp["Cabang"].dropna().unique().tolist())
    all_branches = order_branches(branch_pool)
    selected_branches = st.multiselect("Filter Cabang (kosongkan = semua)", options=all_branches, default=[], key="filter_cabang")

selected_branches = selected_branches or None
periode_label = f"{BULAN_ID[tanggal_acuan.month - 1]} {tanggal_acuan.year} (s/d tanggal {tanggal_acuan.day})"

# Guard dengan `not df_x.empty` di semua baris (bukan cuma ads/walkin) - kalau data Omset
# belum diupload, df_main benar-benar kosong (0 kolom), jadi df_main["Cabang"] akan error
# (KeyError) kalau tidak dijaga, terutama saat user memilih cabang dari data Iklan/Walk-in
# yang sudah ada duluan.
df_main_f = df_main[df_main["Cabang"].isin(selected_branches)] if selected_branches and not df_main.empty else df_main
df_ads_f = df_ads[df_ads["Cabang"].isin(selected_branches)] if selected_branches and not df_ads.empty else df_ads
df_walkin_f = df_walkin[df_walkin["Cabang"].isin(selected_branches)] if selected_branches and not df_walkin.empty else df_walkin

# ---------------------------- Scoreboard computation ----------------------------
sb_service = _finalize_scoreboard(build_scoreboard(df_main, "Service", tanggal_acuan, selected_branches), df_target, "TargetService", tanggal_acuan)
sb_gadget = _finalize_scoreboard(build_scoreboard(df_main, "Gadget & Aksesoris", tanggal_acuan, selected_branches), df_target, "TargetGadget", tanggal_acuan)
sb_all = _finalize_scoreboard(build_scoreboard(df_main, "All", tanggal_acuan, selected_branches), df_target, "TargetAll", tanggal_acuan)
sb_corp = build_scoreboard_corporate_manual(df_corp, df_target, tanggal_acuan, selected_branches)


def _smm_row(sb: pd.DataFrame, name_label: str = "CABANG"):
    if sb.empty:
        return None
    match = sb[sb[name_label].astype(str).str.upper().isin(_SALES_TOTAL_LABELS)]
    return match.iloc[-1] if not match.empty else sb.iloc[-1]


def _pct_and_target_ratio(row):
    if row is None:
        return None
    return row.get("% PENCAPAIAN")


def _sd_value(row):
    if row is None:
        return 0.0
    return float(row.get("OMSET S/D HARI INI", 0.0) or 0.0)


smm_service = _smm_row(sb_service)
smm_gadget = _smm_row(sb_gadget)
smm_all = _smm_row(sb_all)
smm_corp = _smm_row(sb_corp)

omset_service_val = _sd_value(smm_service)
omset_gadget_val = _sd_value(smm_gadget)
omset_all_val = _sd_value(smm_all)
omset_corp_val = _sd_value(smm_corp)

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 Ringkasan", "🗂️ Scoreboard", "📣 Iklan", "🚶 Walk-in", "🏛️ 6 Pilar"]
)

# ==================================== TAB 1: RINGKASAN ====================================
with tab1:
    st.markdown(f"#### Ringkasan Omset — {periode_label}")
    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(render_kpi_card("Omset Service", omset_service_val, "#0f766e", "#134e4a", "🛠️", _pct_and_target_ratio(smm_service)), unsafe_allow_html=True)
    with k2:
        st.markdown(render_kpi_card("Penjualan Gadget & Aksesoris", omset_gadget_val, "#6d28d9", "#4c1d95", "📱", _pct_and_target_ratio(smm_gadget)), unsafe_allow_html=True)
    with k3:
        st.markdown(render_kpi_card("Omset All", omset_all_val, "#1d4ed8", "#1e3a8a", "💰", _pct_and_target_ratio(smm_all)), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**% Pencapaian per Kategori**")
    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown("<div style='text-align:center;font-size:13px;color:#374151;font-weight:600;'>🛠️ Omset Service</div>", unsafe_allow_html=True)
        # key eksplisit WAJIB: dua ring bisa menghasilkan JSON Plotly identik (mis. % sama
        # persis) sehingga auto-generated element ID Streamlit bertabrakan
        # (StreamlitDuplicateElementId) kalau key tidak diberikan.
        st.plotly_chart(render_progress_ring(_pct_and_target_ratio(smm_service)), use_container_width=True, key="ring_service")
    with r2:
        st.markdown("<div style='text-align:center;font-size:13px;color:#374151;font-weight:600;'>📱 Gadget & Aksesoris</div>", unsafe_allow_html=True)
        st.plotly_chart(render_progress_ring(_pct_and_target_ratio(smm_gadget)), use_container_width=True, key="ring_gadget")
    with r3:
        st.markdown("<div style='text-align:center;font-size:13px;color:#374151;font-weight:600;'>💰 Omset All (SMM)</div>", unsafe_allow_html=True)
        st.plotly_chart(render_progress_ring(_pct_and_target_ratio(smm_all)), use_container_width=True, key="ring_all")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Kontribusi terhadap Omset All**")
    pie_fig = render_contribution_pie(omset_service_val, omset_gadget_val, omset_corp_val)
    if pie_fig is not None:
        st.plotly_chart(pie_fig, use_container_width=True, key="pie_kontribusi")
    else:
        st.caption("Belum ada data untuk periode ini.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Progres Harian — Omset All (Aktual vs Target Pace Lurus)**")
    fig_prog = render_daily_progress_chart(df_main_f, df_target, "TargetAll", selected_branches, tanggal_acuan, "#1d4ed8")
    if fig_prog is not None:
        st.plotly_chart(fig_prog, use_container_width=True, key="daily_progress_chart")
    else:
        st.caption("Belum ada data untuk periode ini.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Histori Omset Bulanan**")
    hist_df = build_daily_history(df_main_f, "All", selected_branches)
    fig_hist = render_daily_history_chart(hist_df)
    if fig_hist is not None:
        st.plotly_chart(fig_hist, use_container_width=True, key="daily_history_chart")
    else:
        st.caption("Belum ada histori untuk ditampilkan.")

# ==================================== TAB 2: SCOREBOARD ====================================
with tab2:
    st.markdown(f"#### Scoreboard — {periode_label}")

    st.markdown("**🛠️ Omset Service**")
    st.markdown(render_scoreboard_html(sb_service), unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**📱 Penjualan Gadget & Aksesoris**")
    st.markdown(render_scoreboard_html(sb_gadget), unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**💰 Omset All**")
    st.markdown(render_scoreboard_html(sb_all), unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if not sb_corp.empty:
        st.markdown("**🏢 Marketing Corporate**")
        st.markdown(render_scoreboard_html(sb_corp), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📋 Insight & Rencana Perbaikan")
    all_sales_insights = generate_all_sales_insights(sb_service, sb_gadget, sb_all, sb_corp, selected_branches)
    if not all_sales_insights:
        st.success("Tidak ada catatan khusus — semua cabang dalam kondisi baik untuk periode ini.")
    else:
        for cat in ["Omset Service", "Penjualan Gadget & Aksesoris", "Marketing Corporate"]:
            cat_items = [x for x in all_sales_insights if x["category"] == cat]
            if not cat_items:
                continue
            with st.expander(f"{cat} ({len(cat_items)} catatan)", expanded=(cat == "Omset Service")):
                for item in cat_items:
                    st.markdown(render_structured_insight_card(item), unsafe_allow_html=True)

# ==================================== TAB 3: IKLAN ====================================
with tab3:
    st.markdown(f"#### Performa Iklan — {periode_label}")
    if df_ads_f.empty:
        st.info("Belum ada data Iklan. Upload file export Meta Ads Manager di sidebar.")
    else:
        ads_insights, ads_total_spend, ads_total_msg, ads_avg_cost = generate_ads_insights(df_ads_f)
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Spend", format_rupiah(ads_total_spend))
        m2.metric("Total Messaging Conversation", format_number(ads_total_msg))
        m3.metric("Rata-rata Cost / Conversation", format_rupiah(ads_avg_cost) if ads_avg_cost else "-")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Ringkasan per Cabang**")
        ads_by_branch = aggregate_ads_by_branch(df_ads_f)
        if not ads_by_branch.empty:
            fig_ads = px.bar(ads_by_branch, x="Cabang", y="Spend", text_auto=".2s", color_discrete_sequence=["#1d4ed8"])
            fig_ads.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig_ads, use_container_width=True, key="ads_spend_by_branch")

        st.markdown("---")
        st.markdown("#### 📋 Insight & Rekomendasi Iklan")
        if not ads_insights:
            st.success("Tidak ada catatan khusus untuk campaign iklan saat ini.")
        else:
            for ins in ads_insights:
                st.markdown(render_insight_card(ins["title"], ins["text"], ins["level"]), unsafe_allow_html=True)

# ==================================== TAB 4: WALK-IN ====================================
with tab4:
    st.markdown(f"#### Walk-in — {periode_label}")
    if df_walkin_f.empty:
        st.info("Belum ada data Walk-in. Upload file Rincian Pengiriman Pesanan di sidebar.")
    else:
        walkin_agg = aggregate_walkin_monthly(df_walkin_f)
        total_walkin = int(walkin_agg["TotalWalkin"].sum()) if not walkin_agg.empty else 0
        total_hari = walkin_agg["HariEfektif"].sum() if not walkin_agg.empty else 0
        rata2_harian = (total_walkin / total_hari) if total_hari else 0.0
        jumlah_cabang = df_walkin_f["Cabang"].nunique()

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Walk-in", format_number(total_walkin))
        m2.metric("Rata-rata Walk-in / Hari", format_decimal(rata2_harian))
        m3.metric("Jumlah Cabang", format_number(jumlah_cabang))

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Walk-in per Cabang (Bulan Berjalan)**")
        if not walkin_agg.empty:
            latest_ym = walkin_agg[["Tahun", "Bulan"]].apply(tuple, axis=1).max()
            latest_agg = walkin_agg[walkin_agg[["Tahun", "Bulan"]].apply(tuple, axis=1) == latest_ym]
            fig_walkin = px.bar(latest_agg, x="Cabang", y="TotalWalkin", text_auto=True, color_discrete_sequence=["#0f766e"])
            fig_walkin.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig_walkin, use_container_width=True, key="walkin_by_branch")

        walkin_insights = generate_walkin_insights(walkin_agg)
        st.markdown("---")
        st.markdown("#### 📋 Insight & Rekomendasi Walk-in")
        if not walkin_insights:
            st.success("Tidak ada catatan khusus untuk walk-in saat ini.")
        else:
            for item in walkin_insights:
                st.markdown(render_structured_insight_card(item), unsafe_allow_html=True)

# ==================================== TAB 5: 6 PILAR ====================================
with tab5:
    st.markdown(f"#### 6 Pilar MFlash — {periode_label}")
    st.caption(
        "Berdasarkan kolom KATEGORI PILAR di file per-cabang (Rincian Faktur Penjualan). "
        "File master lama tidak punya kolom ini, jadi omsetnya otomatis masuk "
        "'Belum Dikategorikan'. Setiap Pilar juga dilengkapi Gross Profit, dan Qty khusus "
        "untuk Pilar Penyewaan Corporate & Maintenance Corporate."
    )
    if df_main_f.empty or "Pilar" not in df_main_f.columns:
        st.info("Belum ada data Omset untuk periode ini.")
    else:
        pilar_summary = build_pilar_summary(df_main_f, tanggal_acuan, selected_branches)
        if pilar_summary.empty or pilar_summary["OmsetBulanIni"].sum() == 0:
            st.info("Belum ada omset untuk periode ini.")
        else:
            st.markdown("**Omset, Gross Profit & Qty per Pilar (Bulan Berjalan)**")
            active_rows = [row for _, row in pilar_summary.iterrows()]
            for start in range(0, len(active_rows), 4):
                cols = st.columns(4)
                for offset, row in enumerate(active_rows[start:start + 4]):
                    with cols[offset]:
                        qty_to_show = row["QtyBulanIni"] if row["Pilar"] in _PILAR_SHOW_QTY else None
                        st.markdown(
                            render_pilar_kpi_card(
                                row["Pilar"], row["OmsetBulanIni"], row["PctChange"],
                                gross_profit=row["GrossProfitBulanIni"], qty=qty_to_show,
                            ),
                            unsafe_allow_html=True,
                        )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Kontribusi per Pilar**")
            pie_df = pilar_summary[pilar_summary["OmsetBulanIni"] > 0].copy()
            if not pie_df.empty:
                pie_df["Label"] = pie_df["Pilar"].apply(_pilar_label)
                fig_pilar_pie = px.pie(
                    pie_df, names="Label", values="OmsetBulanIni", hole=0.45,
                    color="Pilar", color_discrete_map=PILAR_COLORS,
                )
                fig_pilar_pie.update_traces(textinfo="label+percent")
                fig_pilar_pie.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), showlegend=True)
                st.plotly_chart(fig_pilar_pie, use_container_width=True, key="pilar_contribution_pie")

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Ringkasan Omset, Gross Profit & Qty per Pilar**")
            st.markdown(render_pilar_summary_table_html(pilar_summary), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Omset per Cabang per Pilar (Bulan Berjalan)**")
            pilar_by_branch = build_pilar_by_branch(df_main_f, tanggal_acuan, selected_branches)
            if not pilar_by_branch.empty:
                chart_df = pilar_by_branch[pilar_by_branch["Cabang"] != "TOTAL"]
                fig_pilar_stack = go.Figure()
                for p in PILAR_ORDER:
                    if p in chart_df.columns and chart_df[p].sum() > 0:
                        fig_pilar_stack.add_trace(go.Bar(
                            name=_pilar_label(p), x=chart_df["Cabang"], y=chart_df[p],
                            marker_color=PILAR_COLORS.get(p, "#6b7280"),
                        ))
                fig_pilar_stack.update_layout(
                    barmode="stack", height=380, margin=dict(l=10, r=10, t=20, b=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig_pilar_stack, use_container_width=True, key="pilar_stacked_by_branch")

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("**Tabel Detail Cabang x Pilar (Omset)**")
                st.markdown(render_pilar_table_html(pilar_by_branch), unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("#### 📋 Insight Kualitas Data")
                pilar_insights = generate_pilar_insights(pilar_by_branch)
                if not pilar_insights:
                    st.success("Semua cabang sudah cukup rapi mengisi KATEGORI PILAR di faktur penjualan.")
                else:
                    for ins in pilar_insights:
                        st.markdown(render_insight_card(ins["title"], ins["text"], ins["level"]), unsafe_allow_html=True)

# ==================================== EXPORT LAPORAN ====================================
st.markdown("---")
st.markdown("### 📤 Export Laporan Presentasi CEO")
st.caption("Berisi 3 bagian: Penyajian Data, Evaluasi, dan Rencana Perbaikan — siap dipresentasikan.")

_report_ads_df = df_ads_f
_report_walkin_df = df_walkin_f
_report_ads_insights, _report_ads_spend, _report_ads_msg, _report_ads_avg_cost = generate_ads_insights(_report_ads_df)
_report_walkin_agg = aggregate_walkin_monthly(_report_walkin_df)
_report_walkin_total = int(_report_walkin_agg["TotalWalkin"].sum()) if not _report_walkin_agg.empty else 0
_report_walkin_hari = _report_walkin_agg["HariEfektif"].sum() if not _report_walkin_agg.empty else 0
_report_walkin_konversi = (_report_walkin_total / _report_walkin_hari) if _report_walkin_hari else 0.0
_report_walkin_insights = generate_walkin_insights(_report_walkin_agg)
_report_sales_insights = generate_all_sales_insights(sb_service, sb_gadget, sb_all, sb_corp, selected_branches)

_report_sections = _build_report_sections(
    _report_sales_insights, _report_ads_insights, _report_walkin_insights,
    omset_all_val, omset_service_val, omset_gadget_val,
    _report_ads_spend, _report_ads_msg, _report_walkin_total, _report_walkin_konversi,
)

_chart_data_list = []
_sb_all_no_total = sb_all[~sb_all["CABANG"].astype(str).str.upper().isin(_SALES_TOTAL_LABELS)] if not sb_all.empty else pd.DataFrame()
if not _sb_all_no_total.empty:
    _chart_data_list.append({
        "title": "Omset All per Cabang", "categories": _sb_all_no_total["CABANG"].tolist(),
        "series_name": "Omset", "values": _sb_all_no_total["OMSET S/D HARI INI"].tolist(),
    })

exp1, exp2 = st.columns(2)
with exp1:
    if st.button("🖨️ Generate PPTX", key="btn_gen_pptx", use_container_width=True):
        st.session_state["_pptx_report"] = generate_pptx_report(_report_sections, periode_label, _chart_data_list)
    if "_pptx_report" in st.session_state:
        st.download_button(
            "⬇️ Download Laporan (.pptx)", data=st.session_state["_pptx_report"],
            file_name=f"Laporan_Omset_MFlash_{tanggal_acuan.isoformat()}.pptx", key="dl_pptx", use_container_width=True,
        )
with exp2:
    if st.button("📄 Generate PDF", key="btn_gen_pdf", use_container_width=True):
        st.session_state["_pdf_report"] = generate_pdf_report(_report_sections, periode_label)
    if "_pdf_report" in st.session_state:
        st.download_button(
            "⬇️ Download Laporan (.pdf)", data=st.session_state["_pdf_report"],
            file_name=f"Laporan_Omset_MFlash_{tanggal_acuan.isoformat()}.pdf", key="dl_pdf", use_container_width=True,
        )
