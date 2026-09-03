"""Dashboard Omset MFlash

Dashboard Streamlit untuk memantau Omset, Iklan (Meta Ads), Walk-in,
6 Pilar MFlash, dan Kontribusi Marketing Corporate vs Sales Retail
di 18 cabang MFlash. Termasuk styling tabel Walk-in (kotak + warna),
export tabel Walk-in & Scoreboard ke JPG/PDF, insight otomatis, dan
export laporan lengkap ke PPTX/PDF. Scoreboard mengikuti persis format
& rumus pada sheet "Scoreboard" di file Excel master (target kuartalan,
EXPECTED VALUE berdasar hari berjalan dalam kuartal, % PENCAPAIAN =
S/D HARI INI dibagi EXPECTED VALUE).
"""

import base64
import calendar
import io
import os
import re
from datetime import datetime, date, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.table import Table as MplTable

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN

from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

LOGO_BASE64 = "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"


def _page_icon():
    try:
        raw = base64.b64decode(LOGO_BASE64)
        with open("/tmp/_mflash_icon.png", "wb") as f:
            f.write(raw)
        return "/tmp/_mflash_icon.png"
    except Exception:
        return "📊"


st.set_page_config(
    page_title="Dashboard Omset MFlash",
    page_icon=_page_icon(),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ========================= GitHub Auto-Backup =========================

def _get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


_GH_TOKEN = os.environ.get("GH_TOKEN", "") or _get_secret("GH_TOKEN", "")
_GH_REPO = _get_secret("GH_REPO", "") or os.environ.get("GH_REPO", "")
_GH_BRANCH = _get_secret("GH_BRANCH", "main") or "main"
_GH_ENABLED = bool(_GH_TOKEN and _GH_REPO)
_GH_API = "https://api.github.com"


def _gh_config():
    return _GH_TOKEN, _GH_REPO, _GH_BRANCH


def _gh_headers():
    return {
        "Authorization": f"token {_GH_TOKEN}",
        "Accept": "application/vnd.github+json",
    }


def github_get_file_sha(path: str):
    if not _GH_ENABLED:
        return None
    url = f"{_GH_API}/repos/{_GH_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=15)
        if r.status_code == 200:
            return r.json().get("sha")
    except Exception:
        pass
    return None


def github_upload_file(path: str, content_bytes: bytes, message: str = "auto-backup"):
    if not _GH_ENABLED:
        return False
    sha = github_get_file_sha(path)
    url = f"{_GH_API}/repos/{_GH_REPO}/contents/{path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": _GH_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(url, headers=_gh_headers(), json=payload, timeout=20)
        return r.status_code in (200, 201)
    except Exception:
        return False


def github_delete_file(path: str, message: str = "auto-delete"):
    if not _GH_ENABLED:
        return False
    sha = github_get_file_sha(path)
    if not sha:
        return True
    url = f"{_GH_API}/repos/{_GH_REPO}/contents/{path}"
    payload = {"message": message, "sha": sha, "branch": _GH_BRANCH}
    try:
        r = requests.delete(url, headers=_gh_headers(), json=payload, timeout=15)
        return r.status_code in (200,)
    except Exception:
        return False


def github_download_file(path: str, local_path: str):
    if not _GH_ENABLED:
        return False
    url = f"{_GH_API}/repos/{_GH_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=20)
        if r.status_code == 200:
            content = base64.b64decode(r.json()["content"])
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            with open(local_path, "wb") as f:
                f.write(content)
            return True
    except Exception:
        pass
    return False


def github_list_dir(path: str):
    if not _GH_ENABLED:
        return []
    url = f"{_GH_API}/repos/{_GH_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": _GH_BRANCH}, timeout=15)
        if r.status_code == 200:
            return [item["name"] for item in r.json() if item["type"] == "file"]
    except Exception:
        pass
    return []


def sync_data_from_github():
    if not _GH_ENABLED:
        return
    for sub in ["main", "ads", "walkin", "target", "corp", "log"]:
        remote_dir = f"data/{sub}"
        local_dir = os.path.join(DATA_DIR, sub)
        os.makedirs(local_dir, exist_ok=True)
        for fname in github_list_dir(remote_dir):
            local_path = os.path.join(local_dir, fname)
            if not os.path.exists(local_path):
                github_download_file(f"{remote_dir}/{fname}", local_path)


# ========================= Constants =========================

MAIN_SHEET_NAME = "Faktur Penjualan"
_FAKTUR_SHEET_NAME = "Rincian Faktur Penjualan"

DATA_DIR = "data"
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
ADS_DATA_DIR = os.path.join(DATA_DIR, "ads")
WALKIN_DATA_DIR = os.path.join(DATA_DIR, "walkin")
TARGET_DATA_DIR = os.path.join(DATA_DIR, "target")
CORP_DATA_DIR = os.path.join(DATA_DIR, "corp")
LOG_DIR = os.path.join(DATA_DIR, "log")

BRANCH_ORDER = [
    "KLENDER", "CEGER", "BINTARA", "RADJIMAN", "JATIMULYA", "DRAMAGA",
    "CONDET", "JATIBENING", "SAWANGAN", "WARBONG", "CINERE", "CIBINONG",
    "KARAWANG", "JATIWARINGIN", "CIKAMPEK", "CILANGKAP", "PEJATEN", "CIBUBUR",
]
_BRANCH_RANK = {b: i for i, b in enumerate(BRANCH_ORDER)}

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}
BULAN_MAP = {v.lower(): k for k, v in BULAN_ID.items()}
BULAN_ALIAS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5, "jun": 6,
    "jul": 7, "agu": 8, "aug": 8, "sep": 9, "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12,
}


def order_branches(names):
    return sorted(set(names), key=lambda b: _BRANCH_RANK.get(str(b).upper(), 999))


# ========================= Format helpers =========================

def format_rupiah(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "Rp 0"
    sign = "-" if v < 0 else ""
    return f"{sign}Rp {abs(v):,.0f}".replace(",", ".")


def format_number(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    return f"{v:,.0f}".replace(",", ".")


def format_decimal(v, digits=1) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    s = f"{v:,.{digits}f}"
    a, _, b = s.partition(".")
    a = a.replace(",", ".")
    return f"{a},{b}" if b else a


def format_percent(v, digits=1) -> str:
    try:
        v = float(v) * 100
    except (TypeError, ValueError):
        return "0%"
    return f"{format_decimal(v, digits)}%"


def sanitize_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name)).strip("_")


def branch_from_filename(fname: str):
    up = str(fname).upper()
    for b in BRANCH_ORDER:
        if b in up.replace(" ", ""):
            return b
    return None


def branch_from_sheetname(sheet: str):
    up = str(sheet).upper()
    for b in BRANCH_ORDER:
        if b in up:
            return b
    return None


_SERVICE_KEYWORDS = ["SERVICE", "JASA", "SPAREPART"]


def classify_kategori(v) -> str:
    """Klasifikasi Service vs Gadget & Aksesoris berdasar KATEGORI PENJUALAN."""
    if not v:
        return "Gadget & Aksesoris"
    up = str(v).strip().upper()
    for kw in _SERVICE_KEYWORDS:
        if kw in up:
            return "Service"
    return "Gadget & Aksesoris"


def _extract_filename_timestamp(fname: str):
    """Ekstrak timestamp YYMMDDHHMMSS di akhir nama file (format export MFlash)."""
    m = re.search(r"(\d{12})(?:\.\w+)?$", str(fname))
    if not m:
        m = re.search(r"_(\d{6,14})\.\w+$", str(fname))
        if m:
            return m.group(1)
        return None
    return m.group(1)


# ========================= 6 Pilar MFlash =========================

PILAR_ORDER = [
    "Service", "Handphone", "Laptop", "Aksesoris", "Voucher & Perdana", "Lainnya",
]
PILAR_ICONS = {
    "Service": "🛠️", "Handphone": "📱", "Laptop": "💻",
    "Aksesoris": "🎧", "Voucher & Perdana": "🎫", "Lainnya": "📦",
}
PILAR_COLORS = {
    "Service": "#0f766e", "Handphone": "#1d4ed8", "Laptop": "#6d28d9",
    "Aksesoris": "#c2410c", "Voucher & Perdana": "#be185d", "Lainnya": "#525252",
}
_PILAR_SHOW_QTY = {"Handphone", "Laptop", "Aksesoris", "Voucher & Perdana", "Lainnya"}


def _pilar_label(p: str) -> str:
    return f"{PILAR_ICONS.get(p, '')} {p}".strip()


def _find_pilar_column_index(col_idx: dict):
    for header, idx in col_idx.items():
        if "KATEGORI PILAR" in header.upper():
            return idx
    return None


def classify_pilar(v) -> str:
    if not v:
        return "Lainnya"
    up = str(v).strip().upper()
    if up.startswith("SERVICE"):
        return "Service"
    if up.startswith("HANDPHONE") or up.startswith("HP"):
        return "Handphone"
    if up.startswith("LAPTOP"):
        return "Laptop"
    if up.startswith("AKSESORIS"):
        return "Aksesoris"
    if up.startswith("VOUCHER") or up.startswith("PERDANA"):
        return "Voucher & Perdana"
    return "Lainnya"


def parse_bulan(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        iv = int(v)
        return iv if 1 <= iv <= 12 else None
    s = str(v).strip().lower()
    if s in BULAN_MAP:
        return BULAN_MAP[s]
    if s in BULAN_ALIAS:
        return BULAN_ALIAS[s]
    for k, num in BULAN_ALIAS.items():
        if s.startswith(k):
            return num
    try:
        return int(s)
    except ValueError:
        return None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return pd.to_datetime(v).date()
    except Exception:
        return None


def _to_float_or_none(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


# ========================= Marketing Corporate vs Sales Retail =========================

MARKETING_CORPORATE_NAMES = [
    "DICKY YUNIAWAN", "FAISAL ABDUL RAHMAN", "IQBAL SABARI SALES", "IQBAL SABARI",
    "IKBAL SABARI", "KOUTSAREZRA KANZA", "M SYAFAAT", "MUHAMMAD SYAFAAT",
    "RAID IMADUDIN FIRAS", "TEGAR SALES", "TEGAR PUTRA YANSA", "TEGAR PUTRA YANSYAH",
    "WAHYU JP JATIWARINGIN", "WAHYU JP (RADJIMAN)", "WAHYU JP", "SUPRIYADI",
    "PAOLO MAROLANZANO", "SOLEHUDIN", "RIFQI ADITYA", "M FARHAN ZAHRAN",
    "KAUKABAN AL AKWAN",
]
_MC_LABEL = "Marketing Corporate"
_RETAIL_LABEL = "Sales Retail"


def _find_penjual_column_index(col_idx: dict):
    for header, idx in col_idx.items():
        if header == "NAMA PENJUAL":
            return idx
    for header, idx in col_idx.items():
        if "DEFAULT PENJUAL" in header:
            return idx
    return None


def classify_penjual_kelompok(nama_penjual) -> str:
    if not nama_penjual:
        return _RETAIL_LABEL
    v = str(nama_penjual).strip().upper()
    if not v:
        return _RETAIL_LABEL
    for known in MARKETING_CORPORATE_NAMES:
        if known in v or v in known:
            return _MC_LABEL
    return _RETAIL_LABEL


# ========================= Loader Data Omset Utama =========================

def _extract_qty_gp(row_dict):
    qty = _to_float_or_none(row_dict.get("QTY")) or 0.0
    gp = _to_float_or_none(row_dict.get("GROSS PROFIT")) or 0.0
    return qty, gp


def _build_col_idx(header_row):
    col_idx = {}
    for i, h in enumerate(header_row):
        if h is None:
            continue
        key = str(h).strip().upper()
        if key and key not in col_idx:
            col_idx[key] = i
    return col_idx


def _load_faktur_sheet(path: str, cabang_hint=None) -> pd.DataFrame:
    """Loader untuk format baru: 'Rincian Faktur Penjualan' per cabang."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet_name = None
    for name in wb.sheetnames:
        if "RINCIAN FAKTUR" in name.upper() or "FAKTUR PENJUALAN" in name.upper():
            sheet_name = name
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)
    if header_row is None:
        return pd.DataFrame()
    col_idx = _build_col_idx(header_row)

    def gi(*names):
        for n in names:
            if n in col_idx:
                return col_idx[n]
        return None

    idx_cabang = gi("CABANG")
    idx_tgl = gi("TGL FAKTUR", "TANGGAL")
    idx_kategori = gi("KATEGORI PENJUALAN")
    idx_total = gi("TOTAL HARGA")
    idx_qty = gi("QTY")
    idx_gp = gi("GROSS PROFIT")
    idx_pilar = _find_pilar_column_index(col_idx)
    idx_penjual = _find_penjual_column_index(col_idx)

    cabang_fallback = cabang_hint or branch_from_filename(os.path.basename(path)) or branch_from_sheetname(sheet_name)

    records = []
    for row in rows_iter:
        if row is None:
            continue
        total = _to_float_or_none(row[idx_total]) if idx_total is not None and idx_total < len(row) else None
        if total is None:
            continue
        cabang = row[idx_cabang] if idx_cabang is not None and idx_cabang < len(row) else None
        cabang = str(cabang).strip().upper() if cabang else cabang_fallback
        tgl = to_date(row[idx_tgl]) if idx_tgl is not None and idx_tgl < len(row) else None
        kategori_raw = row[idx_kategori] if idx_kategori is not None and idx_kategori < len(row) else None
        pilar_raw = row[idx_pilar] if idx_pilar is not None and idx_pilar < len(row) else None
        penjual_raw = row[idx_penjual] if idx_penjual is not None and idx_penjual < len(row) else None
        qty = _to_float_or_none(row[idx_qty]) if idx_qty is not None and idx_qty < len(row) else 0.0
        gp = _to_float_or_none(row[idx_gp]) if idx_gp is not None and idx_gp < len(row) else 0.0
        records.append({
            "Cabang": cabang,
            "Tanggal": tgl,
            "Kelompok": classify_kategori(kategori_raw),
            "TotalHarga": total,
            "Qty": qty or 0.0,
            "GrossProfit": gp or 0.0,
            "Pilar": classify_pilar(pilar_raw),
            "Penjual": str(penjual_raw).strip() if penjual_raw else "",
            "KelompokPenjual": classify_penjual_kelompok(penjual_raw),
        })
    wb.close()
    return pd.DataFrame(records)


def _load_master_sheet(path: str) -> pd.DataFrame:
    """Loader untuk format lama: master file dengan sheet 'Faktur Penjualan'."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet_candidates = [s for s in wb.sheetnames if "FAKTUR" in s.upper() or s.upper().startswith("FP ")]
    if MAIN_SHEET_NAME in wb.sheetnames:
        sheet_candidates = [MAIN_SHEET_NAME] + [s for s in sheet_candidates if s != MAIN_SHEET_NAME]
    if not sheet_candidates:
        sheet_candidates = [wb.sheetnames[0]]

    frames = []
    for sheet_name in sheet_candidates:
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        header_row = next(rows_iter, None)
        if header_row is None:
            continue
        col_idx = _build_col_idx(header_row)
        if "TOTAL HARGA" not in col_idx:
            continue

        def gi(*names):
            for n in names:
                if n in col_idx:
                    return col_idx[n]
            return None

        idx_cabang = gi("CABANG")
        idx_tgl = gi("TGL FAKTUR", "TANGGAL")
        idx_kategori = gi("KATEGORI PENJUALAN")
        idx_total = gi("TOTAL HARGA")
        idx_qty = gi("QTY")
        idx_gp = gi("GROSS PROFIT")
        idx_pilar = _find_pilar_column_index(col_idx)
        idx_penjual = _find_penjual_column_index(col_idx)
        cabang_fallback = branch_from_sheetname(sheet_name)

        for row in rows_iter:
            if row is None:
                continue
            total = _to_float_or_none(row[idx_total]) if idx_total is not None and idx_total < len(row) else None
            if total is None:
                continue
            cabang = row[idx_cabang] if idx_cabang is not None and idx_cabang < len(row) else None
            cabang = str(cabang).strip().upper() if cabang else cabang_fallback
            tgl = to_date(row[idx_tgl]) if idx_tgl is not None and idx_tgl < len(row) else None
            kategori_raw = row[idx_kategori] if idx_kategori is not None and idx_kategori < len(row) else None
            pilar_raw = row[idx_pilar] if idx_pilar is not None and idx_pilar < len(row) else None
            penjual_raw = row[idx_penjual] if idx_penjual is not None and idx_penjual < len(row) else None
            qty = _to_float_or_none(row[idx_qty]) if idx_qty is not None and idx_qty < len(row) else 0.0
            gp = _to_float_or_none(row[idx_gp]) if idx_gp is not None and idx_gp < len(row) else 0.0
            frames.append({
                "Cabang": cabang,
                "Tanggal": tgl,
                "Kelompok": classify_kategori(kategori_raw),
                "TotalHarga": total,
                "Qty": qty or 0.0,
                "GrossProfit": gp or 0.0,
                "Pilar": classify_pilar(pilar_raw),
                "Penjual": str(penjual_raw).strip() if penjual_raw else "",
                "KelompokPenjual": classify_penjual_kelompok(penjual_raw),
            })
    wb.close()
    return pd.DataFrame(frames)


def _looks_like_ads_export(path: str) -> bool:
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        header = next(ws.iter_rows(values_only=True), None)
        wb.close()
        if not header:
            return False
        up = [str(h).upper() for h in header if h]
        return any("CAMPAIGN" in h or "AD SET" in h or "AMOUNT SPENT" in h for h in up)
    except Exception:
        return False


def _detect_main_file_kind(path: str):
    """Deteksi apakah file adalah format baru (per-cabang) atau lama (master)."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        names_up = [s.upper() for s in wb.sheetnames]
        wb.close()
        if any("RINCIAN FAKTUR" in n or "RINCIAN PENGIRIMAN" in n for n in names_up):
            if any("RINCIAN PENGIRIMAN" in n for n in names_up) and not any("RINCIAN FAKTUR" in n for n in names_up):
                return "walkin"
            return "faktur"
        if "FAKTUR PENJUALAN" in names_up or any(n.startswith("FP ") for n in names_up):
            return "master"
    except Exception:
        pass
    return None


def load_main_data(path: str) -> pd.DataFrame:
    kind = _detect_main_file_kind(path)
    if kind == "faktur":
        return _load_faktur_sheet(path)
    if kind == "master":
        return _load_master_sheet(path)
    try:
        return _load_faktur_sheet(path)
    except Exception:
        return pd.DataFrame()


def _dedupe_main_files():
    """Hapus file Omset duplikat/basi per cabang, sisakan yang timestamp-nya terbaru."""
    if not os.path.isdir(MAIN_DATA_DIR):
        return
    files = [f for f in os.listdir(MAIN_DATA_DIR) if f.lower().endswith((".xlsx", ".xls"))]
    groups = {}
    for f in files:
        branch = branch_from_filename(f) or "UNKNOWN"
        ts = _extract_filename_timestamp(f) or ""
        groups.setdefault(branch, []).append((ts, f))
    for branch, items in groups.items():
        if branch == "UNKNOWN" or len(items) <= 1:
            continue
        items.sort(key=lambda x: x[0])
        stale = items[:-1]
        for ts, fname in stale:
            fpath = os.path.join(MAIN_DATA_DIR, fname)
            try:
                os.remove(fpath)
                if _GH_ENABLED:
                    github_delete_file(f"data/main/{fname}", "auto-dedupe stale branch file")
            except Exception:
                pass


def load_all_main_data() -> pd.DataFrame:
    if not os.path.isdir(MAIN_DATA_DIR):
        return pd.DataFrame()
    frames = []
    for fname in sorted(os.listdir(MAIN_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls")):
            continue
        fpath = os.path.join(MAIN_DATA_DIR, fname)
        try:
            df = load_main_data(fpath)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    combined["Tahun"] = combined["Tanggal"].apply(lambda d: d.year if d else None)
    combined["Bulan"] = combined["Tanggal"].apply(lambda d: d.month if d else None)
    combined = combined.dropna(subset=["Cabang", "Tanggal"])
    return combined


# ========================= Loader Data Iklan (Meta Ads) =========================

_ADS_REQUIRED_COLS = ["Campaign name", "Amount spent", "Reach", "Impressions"]


def load_ads_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    if not any(c in df.columns for c in _ADS_REQUIRED_COLS):
        return pd.DataFrame()
    branch = branch_from_filename(os.path.basename(path))
    df["Cabang"] = branch
    rename_map = {}
    used_targets = set()

    def _claim(col, target):
        if target in used_targets:
            return
        rename_map[col] = target
        used_targets.add(target)

    for c in df.columns:
        cl = c.lower()
        if "amount spent" in cl:
            _claim(c, "AmountSpent")
        elif cl == "reach":
            _claim(c, "Reach")
        elif cl == "impressions":
            _claim(c, "Impressions")
        elif "link click" in cl or cl == "clicks (all)":
            _claim(c, "Clicks")
        elif "cpm" in cl:
            _claim(c, "CPM")
        elif "cpc" in cl:
            _claim(c, "CPC")
        elif "ctr" in cl:
            _claim(c, "CTR")
        elif cl == "results":
            _claim(c, "Results")
        elif "messaging conversation" in cl and "cost per" not in cl:
            _claim(c, "Results")
        elif cl == "campaign name":
            _claim(c, "CampaignName")
    df = df.rename(columns=rename_map)
    return df


def load_all_ads_data() -> pd.DataFrame:
    if not os.path.isdir(ADS_DATA_DIR):
        return pd.DataFrame()
    frames = []
    for fname in sorted(os.listdir(ADS_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
            continue
        fpath = os.path.join(ADS_DATA_DIR, fname)
        try:
            df = load_ads_data(fpath)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def aggregate_ads_by_branch(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    agg = {}
    if "AmountSpent" in df.columns:
        agg["AmountSpent"] = "sum"
    if "Reach" in df.columns:
        agg["Reach"] = "sum"
    if "Impressions" in df.columns:
        agg["Impressions"] = "sum"
    if "Clicks" in df.columns:
        agg["Clicks"] = "sum"
    if "Results" in df.columns:
        agg["Results"] = "sum"
    if not agg:
        return pd.DataFrame()
    out = df.groupby("Cabang").agg(agg).reset_index()
    return out


def _content_diagnosis(row):
    ctr = None
    if "Clicks" in row and "Impressions" in row and row.get("Impressions", 0):
        ctr = (row.get("Clicks", 0) or 0) / row["Impressions"] * 100
    if ctr is not None and ctr < 1:
        return "CTR rendah — konten iklan kurang menarik perhatian, coba ganti creative/thumbnail."
    return "CTR cukup baik — pertahankan gaya konten yang sedang berjalan."


def generate_ads_insights(agg_df: pd.DataFrame) -> list:
    insights = []
    if agg_df.empty:
        return insights
    for _, row in agg_df.iterrows():
        cabang = row["Cabang"]
        spent = row.get("AmountSpent", 0) or 0
        results = row.get("Results", 0) or 0
        cpr = (spent / results) if results else None
        if cpr and cpr > 50000:
            insights.append({
                "level": "bad",
                "category": cabang,
                "title": f"{cabang}: Cost per Result tinggi ({format_rupiah(cpr)})",
                "problem": "Biaya per hasil melebihi Rp 50.000, efisiensi iklan perlu dievaluasi.",
                "online": ["Uji creative baru (video pendek, testimoni)", "Perbaiki targeting audiens lookalike"],
                "offline": ["Latih tim CS merespon leads lebih cepat", "Evaluasi promo/penawaran yang ditawarkan"],
            })
        insights.append({
            "level": "warn",
            "category": cabang,
            "title": f"{cabang}: Diagnosa Konten",
            "problem": _content_diagnosis(row),
            "online": [], "offline": [],
        })
    return insights


def render_insight_card(ins: dict):
    colors = {"bad": "#dc2626", "warn": "#d97706", "good": "#16a34a"}
    c = colors.get(ins.get("level"), "#525252")
    st.markdown(
        f"""<div style="border-left:4px solid {c};background:#fafafa;padding:10px 14px;
        border-radius:6px;margin-bottom:8px;">
        <b>{ins.get('title','')}</b><br/><span style="color:#444;">{ins.get('problem','')}</span>
        </div>""",
        unsafe_allow_html=True,
    )


# ========================= Sales Insight Engine =========================

_SALES_TOTAL_LABELS = {
    "Omset All": "Omset All", "Service": "Omset Service",
    "Gadget & Aksesoris": "Penjualan Gadget & Aksesoris",
    "Marketing Corporate": "Kontribusi Marketing Corporate",
}

_CATEGORY_ACTION_PLANS = {
    "Service": {
        "online": [
            "Promo diskon service via WhatsApp Broadcast & Instagram Story",
            "Konten edukasi tips perawatan HP/laptop untuk menarik trafik service",
        ],
        "offline": [
            "Follow-up pelanggan lama untuk service berkala",
            "Training teknisi agar waktu pengerjaan lebih cepat & akurat",
        ],
    },
    "Gadget & Aksesoris": {
        "online": [
            "Bundling promo aksesoris di marketplace/social media",
            "Live selling produk gadget best-seller",
        ],
        "offline": [
            "Display produk baru di etalase depan toko",
            "Cross-selling aksesoris saat pelanggan ambil unit service",
        ],
    },
    "Marketing Corporate": {
        "online": [
            "Follow-up leads corporate via email/WA blast mingguan",
            "Buat konten studi kasus kerja sama korporat sebelumnya",
        ],
        "offline": [
            "Kunjungan langsung ke kantor/perusahaan sekitar cabang",
            "Ikuti pameran/expo bisnis lokal untuk jaringan corporate",
        ],
    },
}


def generate_sales_insights(df_branch_month: pd.DataFrame, kategori: str) -> list:
    insights = []
    if df_branch_month.empty or len(df_branch_month) < 2:
        return insights
    df_sorted = df_branch_month.sort_values(["Tahun", "Bulan"])
    if len(df_sorted) < 2:
        return insights
    last, prev = df_sorted.iloc[-1], df_sorted.iloc[-2]
    delta = last["Total"] - prev["Total"]
    if prev["Total"] and delta / prev["Total"] < -0.1:
        plan = _CATEGORY_ACTION_PLANS.get(kategori, {"online": [], "offline": []})
        insights.append({
            "level": "bad",
            "category": kategori,
            "title": f"{_SALES_TOTAL_LABELS.get(kategori, kategori)} turun {format_percent(abs(delta / prev['Total']))} dibanding bulan lalu",
            "problem": f"Dari {format_rupiah(prev['Total'])} menjadi {format_rupiah(last['Total'])}.",
            "online": plan["online"], "offline": plan["offline"],
        })
    return insights


def generate_all_sales_insights(df: pd.DataFrame) -> list:
    insights = []
    if df.empty:
        return insights
    for kategori in ["Service", "Gadget & Aksesoris"]:
        sub = df[df["Kelompok"] == kategori]
        if sub.empty:
            continue
        g = sub.groupby(["Tahun", "Bulan"])["TotalHarga"].sum().reset_index().rename(columns={"TotalHarga": "Total"})
        insights.extend(generate_sales_insights(g, kategori))
    return insights


def _render_online_offline_html(online: list, offline: list) -> str:
    def _list_html(items):
        if not items:
            return "<i>-</i>"
        return "<ul style='margin:4px 0 0 0;padding-left:18px;'>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"
    return f"""
    <div style="display:flex;gap:16px;margin-top:6px;">
        <div style="flex:1;"><b>💻 Online</b>{_list_html(online)}</div>
        <div style="flex:1;"><b>🏪 Offline</b>{_list_html(offline)}</div>
    </div>
    """


def render_structured_insight_card(ins: dict):
    colors = {"bad": "#dc2626", "warn": "#d97706", "good": "#16a34a"}
    bg = {"bad": "#fef2f2", "warn": "#fffbeb", "good": "#f0fdf4"}
    c = colors.get(ins.get("level"), "#525252")
    b = bg.get(ins.get("level"), "#fafafa")
    online_offline_html = _render_online_offline_html(ins.get("online", []), ins.get("offline", []))
    st.markdown(
        f"""<div style="border-left:4px solid {c};background:{b};padding:12px 16px;
        border-radius:8px;margin-bottom:10px;">
        <b style="font-size:1.02em;">{ins.get('title','')}</b><br/>
        <span style="color:#444;">{ins.get('problem','')}</span>
        {online_offline_html}
        </div>""",
        unsafe_allow_html=True,
    )


def render_kpi_card_text(label: str, value: str, color: str = "#1d4ed8", icon: str = "", sub: str = ""):
    sub_html = f"<div style='font-size:0.78em;color:#6b7280;margin-top:2px;'>{sub}</div>" if sub else ""
    st.markdown(
        f"""<div style="border:2px solid {color};border-radius:10px;padding:14px 16px;
        background:linear-gradient(180deg,#ffffff 0%,{color}0d 100%);text-align:center;">
        <div style="font-size:0.85em;color:#4b5563;font-weight:600;">{icon} {label}</div>
        <div style="font-size:1.6em;font-weight:800;color:{color};margin-top:4px;">{value}</div>
        {sub_html}
        </div>""",
        unsafe_allow_html=True,
    )


def render_kpi_card(label: str, value: str, color: str = "#1d4ed8", icon: str = ""):
    render_kpi_card_text(label, value, color, icon)


# ========================= Walk-in per Cabang =========================

_WALKIN_REQUIRED_COLS = ["TGL PENGIRIMAN", "NOMOR PENGIRIMAN"]


def _find_walkin_sheet(wb):
    for name in wb.sheetnames:
        if "PENGIRIMAN" in name.upper():
            return name
    return wb.sheetnames[0]


def load_walkin_data(path: str) -> pd.DataFrame:
    # Pakai pandas.read_excel (bukan openpyxl read_only) karena sejumlah file
    # export "Rincian Pengiriman Pesanan" MFlash punya metadata dimensi sheet yang
    # tidak akurat sehingga openpyxl read_only gagal mendeteksi baris data.
    try:
        sheet_name = None
        try:
            xls = pd.ExcelFile(path)
            for s in xls.sheet_names:
                if "PENGIRIMAN" in s.upper():
                    sheet_name = s
                    break
            if sheet_name is None:
                sheet_name = xls.sheet_names[0]
        except Exception:
            sheet_name = 0
        raw = pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    if raw is None or raw.empty:
        return pd.DataFrame()
    raw.columns = [str(c).strip().upper() for c in raw.columns]

    def gi(*names):
        for n in names:
            if n in raw.columns:
                return n
        for n in names:
            for c in raw.columns:
                if n in c:
                    return c
        return None

    col_tgl = gi("TGL PENGIRIMAN", "TANGGAL PENGIRIMAN", "TANGGAL")
    col_nomor = gi("NOMOR PENGIRIMAN PESANAN", "NOMOR PENGIRIMAN", "NO PENGIRIMAN", "NO. PENGIRIMAN")
    col_cabang = gi("CABANG")
    if col_tgl is None or col_nomor is None:
        return pd.DataFrame()

    cabang_fallback = branch_from_filename(os.path.basename(path)) or (
        branch_from_sheetname(sheet_name) if isinstance(sheet_name, str) else None
    )
    records = []
    for _, row in raw.iterrows():
        tgl = to_date(row[col_tgl])
        nomor = row[col_nomor]
        if tgl is None or nomor is None or (isinstance(nomor, float) and pd.isna(nomor)) or str(nomor).strip() == "":
            continue
        cabang = row[col_cabang] if col_cabang is not None else None
        cabang = str(cabang).strip().upper() if cabang and str(cabang).strip() else cabang_fallback
        records.append({"Cabang": cabang, "Tanggal": tgl, "NomorPengiriman": str(nomor).strip()})
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df["Tahun"] = df["Tanggal"].apply(lambda d: d.year)
    df["Bulan"] = df["Tanggal"].apply(lambda d: d.month)
    return df


def load_all_walkin_data() -> pd.DataFrame:
    if not os.path.isdir(WALKIN_DATA_DIR):
        return pd.DataFrame()
    frames = []
    for fname in sorted(os.listdir(WALKIN_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls")):
            continue
        fpath = os.path.join(WALKIN_DATA_DIR, fname)
        try:
            df = load_walkin_data(fpath)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.dropna(subset=["Cabang", "Tanggal"])
    combined = combined.drop_duplicates(subset=["Cabang", "NomorPengiriman"], keep="first")
    return combined


def aggregate_walkin_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """Riwayat bulanan (multi-bulan) untuk insight tren — SATU baris per (Cabang, Tahun, Bulan)."""
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
            "TotalWalkin": int(total), "HariEfektif": int(hari_efektif), "RataRataPerHari": rata2,
        })
    return pd.DataFrame(rows).sort_values(["Tahun", "Bulan", "Cabang"]).reset_index(drop=True)


def aggregate_walkin_current_period(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    """Total & rata-rata Walk-in per cabang untuk PERIODE BERJALAN SAJA, mengikuti
    Tanggal Acuan global (Tahun & Bulan tanggal_acuan, hari 1 s/d tanggal_acuan.day).
    Ini memastikan tabel tidak pernah menggabungkan lebih dari satu bulan."""
    cols = ["Cabang", "Tahun", "Bulan", "TotalWalkin", "HariEfektif", "RataRataPerHari"]
    if df.empty or tanggal_acuan is None:
        return pd.DataFrame(columns=cols)
    mask = (
        (df["Tahun"] == tanggal_acuan.year)
        & (df["Bulan"] == tanggal_acuan.month)
        & (df["Tanggal"].apply(lambda d: d.day) <= tanggal_acuan.day)
    )
    d = df[mask]
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame(columns=cols)
    hari_efektif = tanggal_acuan.day
    rows = []
    for cabang, g in d.groupby("Cabang"):
        total = g["NomorPengiriman"].nunique()
        rata2 = (total / hari_efektif) if hari_efektif else 0.0
        rows.append({
            "Cabang": cabang, "Tahun": tanggal_acuan.year, "Bulan": tanggal_acuan.month,
            "TotalWalkin": int(total), "HariEfektif": int(hari_efektif), "RataRataPerHari": rata2,
        })
    return pd.DataFrame(rows).reset_index(drop=True)


def _walkin_ordered(d: pd.DataFrame) -> pd.DataFrame:
    if d.empty:
        return d
    d = d.copy()
    branches_present = order_branches(d["Cabang"].tolist())
    d["Cabang"] = pd.Categorical(d["Cabang"], categories=branches_present, ordered=True)
    return d.sort_values("Cabang").reset_index(drop=True)


def _walkin_overall_avg(d: pd.DataFrame) -> float:
    if d.empty:
        return 0.0
    return float(d["RataRataPerHari"].mean())


def render_walkin_table_html(d: pd.DataFrame, periode_label: str = "") -> str:
    if d.empty:
        return "<i>Tidak ada data walk-in untuk periode ini.</i>"
    overall_avg = _walkin_overall_avg(d)
    header_label = f" — {periode_label}" if periode_label else ""
    rows_html = ""
    for _, r in d.iterrows():
        above = r["RataRataPerHari"] >= overall_avg
        bg = "#dcfce7" if above else "#fee2e2"
        badge = "🟢" if above else "🔴"
        rows_html += f"""<tr style="background:{bg};">
            <td style="padding:6px 10px;border:1px solid #d1d5db;font-weight:600;">{r['Cabang']}</td>
            <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_number(r['TotalWalkin'])}</td>
            <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_decimal(r['RataRataPerHari'])} {badge}</td>
        </tr>"""
    return f"""
    <div style="border:2px solid #0f766e;border-radius:10px;overflow:hidden;">
    <div style="background:#0f766e;color:white;padding:8px 12px;font-weight:700;">
        📋 Tabel Jumlah &amp; Rata-rata Walk-in per Cabang{header_label}
    </div>
    <table style="border-collapse:collapse;width:100%;font-size:0.92em;">
        <thead><tr style="background:#f0fdfa;">
            <th style="padding:6px 10px;border:1px solid #d1d5db;text-align:left;">Cabang</th>
            <th style="padding:6px 10px;border:1px solid #d1d5db;">Total Walk-in</th>
            <th style="padding:6px 10px;border:1px solid #d1d5db;">Rata-rata/Hari</th>
        </tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
    </div>
    """


def generate_walkin_table_image(d: pd.DataFrame, periode_label: str = "") -> bytes:
    title0 = "Tabel Jumlah & Rata-rata Walk-in per Cabang"
    if periode_label:
        title0 += f" — {periode_label}"
    if d.empty:
        fig, ax = plt.subplots(figsize=(7, 2.2))
        ax.axis("off")
        ax.set_title(title0, fontsize=12, fontweight="bold", color="#0f766e", pad=14)
        ax.text(0.5, 0.5, "Tidak ada data walk-in untuk periode ini.", ha="center", va="center", fontsize=10, color="#6b7280")
        buf = io.BytesIO()
        fig.savefig(buf, format="jpg", dpi=200, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.read()
    overall_avg = _walkin_overall_avg(d)
    n = len(d)
    fig_h = max(1.5, 0.5 * n + 1.2)
    fig, ax = plt.subplots(figsize=(7, fig_h))
    ax.axis("off")
    title = "Tabel Jumlah & Rata-rata Walk-in per Cabang"
    if periode_label:
        title += f" — {periode_label}"
    ax.set_title(title, fontsize=12, fontweight="bold", color="#0f766e", pad=14)

    col_labels = ["Cabang", "Total Walk-in", "Rata-rata/Hari"]
    cell_text = [[r["Cabang"], format_number(r["TotalWalkin"]), format_decimal(r["RataRataPerHari"])] for _, r in d.iterrows()]
    tbl = ax.table(cellText=cell_text, colLabels=col_labels, loc="center", cellLoc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1, 1.6)
    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor("#d1d5db")
        if row == 0:
            cell.set_facecolor("#0f766e")
            cell.set_text_props(color="white", fontweight="bold")
        else:
            above = d.iloc[row - 1]["RataRataPerHari"] >= overall_avg
            cell.set_facecolor("#dcfce7" if above else "#fee2e2")
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="jpg", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def generate_walkin_table_pdf(d: pd.DataFrame, periode_label: str = "") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Title"], fontSize=14, textColor=rl_colors.HexColor("#0f766e"))
    elements = [Paragraph(f"Tabel Jumlah & Rata-rata Walk-in per Cabang{' — ' + periode_label if periode_label else ''}", title_style), Spacer(1, 12)]

    if d.empty:
        elements.append(Paragraph("Tidak ada data walk-in untuk periode ini.", styles["Normal"]))
    else:
        overall_avg = _walkin_overall_avg(d)
        data = [["Cabang", "Total Walk-in", "Rata-rata/Hari"]]
        row_colors = []
        for _, r in d.iterrows():
            data.append([r["Cabang"], format_number(r["TotalWalkin"]), format_decimal(r["RataRataPerHari"])])
            row_colors.append(r["RataRataPerHari"] >= overall_avg)
        table = Table(data, colWidths=[6 * cm, 5 * cm, 5 * cm])
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0f766e")),
            ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#d1d5db")),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for i, above in enumerate(row_colors, start=1):
            bg = rl_colors.HexColor("#dcfce7") if above else rl_colors.HexColor("#fee2e2")
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), bg))
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)

    doc.build(elements)
    buf.seek(0)
    return buf.read()


_WALKIN_ACTION_PLAN = {
    "online": [
        "Promosi lokasi via Google Maps/Instagram Ads radius sekitar cabang",
        "Kampanye promo walk-in (diskon cek gratis, hari tertentu)",
    ],
    "offline": [
        "Spanduk/banner promo di depan toko & area sekitar",
        "Kerja sama dengan warga/komunitas sekitar (RT/RW, kampus, kantor)",
    ],
}


def generate_walkin_marketing_insights(d: pd.DataFrame) -> list:
    insights = []
    if d.empty:
        return insights
    overall_avg = _walkin_overall_avg(d)
    if overall_avg <= 0:
        return insights
    for _, r in d.iterrows():
        if r["RataRataPerHari"] < 0.85 * overall_avg:
            insights.append({
                "level": "bad",
                "category": r["Cabang"],
                "title": f"{r['Cabang']}: Walk-in di bawah rata-rata cabang lain",
                "problem": f"Rata-rata {format_decimal(r['RataRataPerHari'])}/hari, sementara rata-rata seluruh cabang {format_decimal(overall_avg)}/hari.",
                "online": _WALKIN_ACTION_PLAN["online"], "offline": _WALKIN_ACTION_PLAN["offline"],
            })
    return insights


def generate_walkin_insights(walkin_agg: pd.DataFrame) -> list:
    """Insight tren bulan-ke-bulan (dibanding bulan lalu), pakai riwayat multi-bulan."""
    insights = []
    if walkin_agg.empty:
        return insights
    for cabang, g in walkin_agg.groupby("Cabang"):
        g = g.sort_values(["Tahun", "Bulan"])
        if len(g) < 2:
            continue
        last, prev = g.iloc[-1], g.iloc[-2]
        if prev["RataRataPerHari"] and (last["RataRataPerHari"] - prev["RataRataPerHari"]) / prev["RataRataPerHari"] < -0.15:
            insights.append({
                "level": "warn",
                "category": cabang,
                "title": f"{cabang}: Walk-in turun dibanding bulan lalu",
                "problem": f"Rata-rata/hari dari {format_decimal(prev['RataRataPerHari'])} menjadi {format_decimal(last['RataRataPerHari'])}.",
                "online": _WALKIN_ACTION_PLAN["online"], "offline": _WALKIN_ACTION_PLAN["offline"],
            })
    return insights


# ========================= Corporate Manual Loader =========================

def load_corporate_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def make_corporate_template() -> bytes:
    df = pd.DataFrame({
        "Nama": ["Contoh Nama"], "Cabang": ["KLENDER"],
        "PERIODE BULAN LALU": [0], "PERIODE BULAN INI": [0], "HARI INI": [0],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Corporate")
    buf.seek(0)
    return buf.read()


# ========================= Target Loader =========================

def load_target_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def make_target_template() -> bytes:
    df = pd.DataFrame({
        "Cabang": BRANCH_ORDER,
        "Target Omset All (Kuartal)": [0] * len(BRANCH_ORDER),
        "Target Service (Kuartal)": [0] * len(BRANCH_ORDER),
        "Target Gadget & Aksesoris (Kuartal)": [0] * len(BRANCH_ORDER),
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Target")
    buf.seek(0)
    return buf.read()


# ========================= Scoreboard (mengikuti persis sheet Excel) =========================

SCOREBOARD_KATEGORI = ["Omset All", "Service", "Gadget & Aksesoris"]
_SCOREBOARD_KATEGORI_LABEL = {
    "Omset All": "SCOREBOARD OMSET ALL",
    "Service": "SCOREBOARD OMSET SERVICE",
    "Gadget & Aksesoris": "SCOREBOARD OMSET GADGET & AKSESORIS",
}

MONEY_COLS = [
    "OmsetSamurai", "OmsetHarian", "ExpectedValue", "HariIni", "SdHariIni",
    "GapHariIni", "TotalGap", "KejarPerhari", "PeriodeBulanLalu", "PeriodeBulanIni", "GapPeriode",
]


def _quarter_bounds(d: date):
    """Kembalikan (start, end, total_hari, hari_berjalan, sisa_hari) kuartal kalender
    yang memuat tanggal d — SAMA PERSIS dengan sheet 'Data Periode' (TOTAL HARI,
    TANGGAL HARI INI, SISA HARI)."""
    q = (d.month - 1) // 3
    start_month = q * 3 + 1
    start = date(d.year, start_month, 1)
    end_month = start_month + 2
    end_year = d.year
    if end_month > 12:
        end_month -= 12
        end_year += 1
    end = date(end_year, end_month, calendar.monthrange(end_year, end_month)[1])
    total_hari = (end - start).days + 1
    hari_berjalan = (d - start).days + 1
    sisa_hari = total_hari - hari_berjalan
    return start, end, total_hari, hari_berjalan, sisa_hari


def pencapaian_color(pct: float) -> str:
    if pct >= 1.0:
        return "#16a34a"
    if pct >= 0.85:
        return "#d97706"
    return "#dc2626"


def build_scoreboard(df_main: pd.DataFrame, tanggal_acuan: date, target_map: dict,
                      kategori: str, selected_branches=None) -> pd.DataFrame:
    """Bangun tabel scoreboard untuk satu kategori (Omset All/Service/Gadget & Aksesoris),
    memakai rumus yang identik dengan sheet Excel 'Scoreboard' + 'Data Periode':
      OMSET HARIAN   = OMSET SAMURAI / TOTAL HARI (hari dalam kuartal berjalan)
      EXPECTED VALUE = OMSET HARIAN * TANGGAL HARI INI (hari berjalan dalam kuartal)
      % PENCAPAIAN   = S/D HARI INI / EXPECTED VALUE
      TOTAL GAP      = OMSET SAMURAI - S/D HARI INI
      KEJAR/HARI     = TOTAL GAP / SISA HARI
    """
    cols = ["Cabang", "OmsetSamurai", "OmsetHarian", "ExpectedValue", "HariIni", "SdHariIni",
            "PctPencapaian", "GapHariIni", "TotalGap", "KejarPerhari",
            "PeriodeBulanLalu", "PeriodeBulanIni", "GapPeriode"]
    if tanggal_acuan is None:
        return pd.DataFrame(columns=cols)

    start, end, total_hari, hari_berjalan, sisa_hari = _quarter_bounds(tanggal_acuan)
    kat_targets = target_map.get(kategori, {}) if target_map else {}

    if df_main is None or df_main.empty:
        sub = pd.DataFrame(columns=["Cabang", "Tanggal", "Tahun", "Bulan", "TotalHarga", "Kelompok"])
    elif kategori == "Omset All":
        sub = df_main
    else:
        sub = df_main[df_main["Kelompok"] == kategori]

    branches = set(kat_targets.keys())
    if not sub.empty:
        branches |= set(sub["Cabang"].dropna().unique().tolist())
    if selected_branches:
        branches &= set(selected_branches)
        branches |= (set(kat_targets.keys()) & set(selected_branches))
    branches = order_branches(branches)
    if not branches:
        return pd.DataFrame(columns=cols)

    prev_month = tanggal_acuan.month - 1 or 12
    prev_year = tanggal_acuan.year if tanggal_acuan.month > 1 else tanggal_acuan.year - 1
    days_in_prev_month = calendar.monthrange(prev_year, prev_month)[1]
    days_elapsed_this_month = tanggal_acuan.day

    rows = []
    for cabang in branches:
        g = sub[sub["Cabang"] == cabang] if not sub.empty else sub
        target = float(kat_targets.get(cabang, 0.0) or 0.0)
        omset_harian = (target / total_hari) if total_hari else 0.0
        expected = omset_harian * hari_berjalan

        hari_ini = float(g[g["Tanggal"] == tanggal_acuan]["TotalHarga"].sum()) if not g.empty else 0.0
        sd_hari_ini = float(g[(g["Tanggal"] >= start) & (g["Tanggal"] <= tanggal_acuan)]["TotalHarga"].sum()) if not g.empty else 0.0
        pct = (sd_hari_ini / expected) if expected else 0.0
        gap_hari_ini = expected - sd_hari_ini
        total_gap = target - sd_hari_ini
        kejar_perhari = (total_gap / sisa_hari) if sisa_hari else 0.0

        bulan_lalu_total = float(g[(g["Tahun"] == prev_year) & (g["Bulan"] == prev_month)]["TotalHarga"].sum()) if not g.empty else 0.0
        periode_bulan_lalu = (bulan_lalu_total / days_in_prev_month) if days_in_prev_month else 0.0
        bulan_ini_total = float(g[(g["Tahun"] == tanggal_acuan.year) & (g["Bulan"] == tanggal_acuan.month) &
                                   (g["Tanggal"].apply(lambda d: d.day) <= tanggal_acuan.day)]["TotalHarga"].sum()) if not g.empty else 0.0
        periode_bulan_ini = (bulan_ini_total / days_elapsed_this_month) if days_elapsed_this_month else 0.0
        gap_periode = periode_bulan_ini - periode_bulan_lalu

        rows.append({
            "Cabang": cabang, "OmsetSamurai": target, "OmsetHarian": omset_harian,
            "ExpectedValue": expected, "HariIni": hari_ini, "SdHariIni": sd_hari_ini,
            "PctPencapaian": pct, "GapHariIni": gap_hari_ini, "TotalGap": total_gap,
            "KejarPerhari": kejar_perhari, "PeriodeBulanLalu": periode_bulan_lalu,
            "PeriodeBulanIni": periode_bulan_ini, "GapPeriode": gap_periode,
        })

    df_out = pd.DataFrame(rows)

    total_target = df_out["OmsetSamurai"].sum()
    total_omset_harian = df_out["OmsetHarian"].sum()
    total_expected = df_out["ExpectedValue"].sum()
    total_hari_ini = df_out["HariIni"].sum()
    total_sd = df_out["SdHariIni"].sum()
    total_pct = (total_sd / total_expected) if total_expected else 0.0
    total_gap_hari_ini = total_expected - total_sd
    total_gap = total_target - total_sd
    total_kejar = (total_gap / sisa_hari) if sisa_hari else 0.0
    total_bulan_lalu = df_out["PeriodeBulanLalu"].sum()
    total_bulan_ini = df_out["PeriodeBulanIni"].sum()
    total_row = pd.DataFrame([{
        "Cabang": "SMM", "OmsetSamurai": total_target, "OmsetHarian": total_omset_harian,
        "ExpectedValue": total_expected, "HariIni": total_hari_ini, "SdHariIni": total_sd,
        "PctPencapaian": total_pct, "GapHariIni": total_gap_hari_ini, "TotalGap": total_gap,
        "KejarPerhari": total_kejar, "PeriodeBulanLalu": total_bulan_lalu,
        "PeriodeBulanIni": total_bulan_ini, "GapPeriode": total_bulan_ini - total_bulan_lalu,
    }])
    df_out = pd.concat([df_out, total_row], ignore_index=True)
    return df_out


def _finalize_scoreboard(df_out: pd.DataFrame) -> pd.DataFrame:
    return df_out


# ========================= Auto-extract Target dari sheet Scoreboard =========================

_SECTION_MARKERS = {
    "SCOREBOARD OMSET ALL": "Omset All",
    "SCOREBOARD OMSET SERVICE": "Service",
    "SCOREBOARD OMSET GADGET & AKSESORIS": "Gadget & Aksesoris",
}


def _read_scoreboard_sections(path: str):
    """Pindai sheet 'Scoreboard' pada file yang di-upload, kembalikan dict:
    {kategori: {cabang: omset_samurai}} dan tanggal snapshot (jika ada)."""
    result = {v: {} for v in _SECTION_MARKERS.values()}
    snapshot_date = None
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return result, snapshot_date
    sheet_name = None
    for name in wb.sheetnames:
        if name.strip().upper() == "SCOREBOARD":
            sheet_name = name
            break
    if sheet_name is None:
        wb.close()
        return result, snapshot_date
    ws = wb[sheet_name]

    current_kategori = None
    header_map = None
    for row in ws.iter_rows(values_only=True):
        if row is None:
            continue
        row_vals = list(row)
        joined = " ".join(str(c).strip().upper() for c in row_vals if c)
        matched_kategori = None
        for marker, kat in _SECTION_MARKERS.items():
            if marker in joined:
                matched_kategori = kat
                break
        if matched_kategori:
            current_kategori = matched_kategori
            header_map = None
            for c in row_vals:
                if isinstance(c, (datetime, date)):
                    snapshot_date = c if isinstance(c, date) and not isinstance(c, datetime) else c.date()
            continue
        if current_kategori is None:
            continue
        up_vals = [str(c).strip().upper() if c is not None else None for c in row_vals]
        if "CABANG" in up_vals and "OMSET SAMURAI" in up_vals:
            header_map = {v: i for i, v in enumerate(up_vals) if v}
            continue
        if header_map is None:
            continue
        idx_cabang = header_map.get("CABANG")
        idx_target = header_map.get("OMSET SAMURAI")
        if idx_cabang is None or idx_target is None or idx_cabang >= len(row_vals):
            continue
        cabang_raw = row_vals[idx_cabang]
        if not cabang_raw:
            continue
        cabang = str(cabang_raw).strip().upper()
        if cabang not in BRANCH_ORDER:
            if cabang in ("SMM", "TOTAL", "GRAND TOTAL"):
                continue
            if cabang not in [m.upper() for m in _SECTION_MARKERS]:
                pass
        target_val = _to_float_or_none(row_vals[idx_target]) if idx_target < len(row_vals) else None
        if cabang in BRANCH_ORDER and target_val is not None:
            result[current_kategori][cabang] = target_val
    wb.close()
    return result, snapshot_date


def extract_scoreboard_target(paths) -> dict:
    """Gabungkan target dari beberapa file (ambil yang punya sheet Scoreboard)."""
    merged = {v: {} for v in _SECTION_MARKERS.values()}
    if isinstance(paths, str):
        paths = [paths]
    for p in paths:
        try:
            sections, _ = _read_scoreboard_sections(p)
        except Exception:
            continue
        for kat, d in sections.items():
            merged[kat].update(d)
    return merged


def extract_scoreboard_snapshot_date(paths):
    if isinstance(paths, str):
        paths = [paths]
    for p in paths:
        try:
            _, snap = _read_scoreboard_sections(p)
            if snap:
                return snap
        except Exception:
            continue
    return None


def extract_scoreboard_corporate(paths):
    """Best-effort: tidak dipakai untuk perhitungan utama — tab Kontribusi MC
    memakai data transaksi langsung (lebih akurat & real-time)."""
    return {}


# ========================= Render & Export Scoreboard =========================

def _fmt_scoreboard_cell(col: str, val) -> str:
    if col == "PctPencapaian":
        return format_percent(val)
    if col in MONEY_COLS:
        return format_rupiah(val)
    return str(val)


_SCOREBOARD_COL_ORDER = [
    ("OmsetSamurai", "OMSET SAMURAI"), ("OmsetHarian", "OMSET HARIAN"), ("ExpectedValue", "EXPECTED VALUE"),
    ("HariIni", "HARI INI"), ("SdHariIni", "S/D HARI INI"), ("PctPencapaian", "% PENCAPAIAN"),
    ("GapHariIni", "GAP HARI INI"), ("TotalGap", "TOTAL GAP"), ("KejarPerhari", "KEJAR TARGET/HARI"),
    ("PeriodeBulanLalu", "PERIODE BULAN LALU"), ("PeriodeBulanIni", "PERIODE BULAN INI"), ("GapPeriode", "GAP"),
]
_SCOREBOARD_GROUPS = [
    ("DETAIL TARGET", 3, "#0f766e"), ("DETAIL PENCAPAIAN", 3, "#1d4ed8"),
    ("DETAIL GAP", 3, "#b45309"), ("MONITORING PROGRESS RATA-RATA/HARI", 3, "#6d28d9"),
]


def render_scoreboard_html(df_score: pd.DataFrame, title: str, periode_label: str = "") -> str:
    if df_score.empty:
        return "<i>Belum ada data/target untuk kategori ini.</i>"

    group_header = ""
    for label, span, color in _SCOREBOARD_GROUPS:
        group_header += f'<th colspan="{span}" style="background:{color};color:white;padding:5px;border:1px solid #d1d5db;">{label}</th>'
    col_header = '<th style="padding:6px 8px;border:1px solid #d1d5db;background:#111827;color:white;">CABANG</th>'
    for _, label in _SCOREBOARD_COL_ORDER:
        col_header += f'<th style="padding:6px 8px;border:1px solid #d1d5db;background:#374151;color:white;font-size:0.82em;">{label}</th>'

    body_rows = ""
    for _, r in df_score.iterrows():
        is_total = str(r["Cabang"]) == "SMM"
        row_bg = "#f3f4f6" if is_total else "#ffffff"
        fw = "800" if is_total else "500"
        tds = f'<td style="padding:5px 8px;border:1px solid #d1d5db;font-weight:{fw};background:{row_bg};">{r["Cabang"]}</td>'
        for col, _label in _SCOREBOARD_COL_ORDER:
            val = r[col]
            txt = _fmt_scoreboard_cell(col, val)
            cell_bg = row_bg
            cell_color = "#111827"
            if col == "PctPencapaian":
                cell_color = pencapaian_color(val)
                tds_style_extra = "font-weight:700;"
            else:
                tds_style_extra = ""
            if col in ("GapHariIni", "TotalGap", "GapPeriode"):
                cell_color = "#16a34a" if val >= 0 else "#dc2626"
            tds += (f'<td style="padding:5px 8px;border:1px solid #d1d5db;background:{cell_bg};'
                    f'color:{cell_color};text-align:right;font-weight:{fw};{tds_style_extra}">{txt}</td>')
        body_rows += f"<tr>{tds}</tr>"

    header_label = f" — {periode_label}" if periode_label else ""
    return f"""
    <div style="border:2px solid #111827;border-radius:10px;overflow-x:auto;">
    <div style="background:#111827;color:white;padding:8px 12px;font-weight:700;">
        🏆 {title}{header_label}
    </div>
    <table style="border-collapse:collapse;width:100%;font-size:0.85em;min-width:1100px;">
        <thead>
            <tr><th style="border:1px solid #d1d5db;background:#111827;"></th>{group_header}</tr>
            <tr>{col_header}</tr>
        </thead>
        <tbody>{body_rows}</tbody>
    </table>
    </div>
    """


def generate_scoreboard_table_image(df_score: pd.DataFrame, title: str, periode_label: str = "") -> bytes:
    full_title = title + (f" — {periode_label}" if periode_label else "")
    if df_score.empty:
        fig, ax = plt.subplots(figsize=(10, 2.5))
        ax.axis("off")
        ax.set_title(full_title, fontsize=15, fontweight="bold", color="#111827", pad=18)
        ax.text(0.5, 0.5, "Belum ada data/target untuk kategori ini.", ha="center", va="center", fontsize=11, color="#6b7280")
        buf = io.BytesIO()
        fig.savefig(buf, format="jpg", dpi=180, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    n = len(df_score)
    fig_w = 20
    fig_h = max(2.2, 0.42 * n + 1.6)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")
    ax.set_title(full_title, fontsize=15, fontweight="bold", color="#111827", pad=18)

    col_labels = ["CABANG"] + [label for _, label in _SCOREBOARD_COL_ORDER]
    cell_text = []
    for _, r in df_score.iterrows():
        row = [r["Cabang"]] + [_fmt_scoreboard_cell(col, r[col]) for col, _ in _SCOREBOARD_COL_ORDER]
        cell_text.append(row)

    tbl = ax.table(cellText=cell_text, colLabels=col_labels, loc="center", cellLoc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.scale(1, 1.7)
    pct_col_idx = 1 + [c for c, _ in _SCOREBOARD_COL_ORDER].index("PctPencapaian")
    gap_cols_idx = [1 + [c for c, _ in _SCOREBOARD_COL_ORDER].index(c) for c in ("GapHariIni", "TotalGap", "GapPeriode")]
    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor("#d1d5db")
        if row == 0:
            cell.set_facecolor("#111827")
            cell.set_text_props(color="white", fontweight="bold")
            continue
        r_data = df_score.iloc[row - 1]
        is_total = str(r_data["Cabang"]) == "SMM"
        cell.set_facecolor("#f3f4f6" if is_total else "white")
        if col == pct_col_idx:
            cell.set_text_props(color=pencapaian_color(r_data["PctPencapaian"]), fontweight="bold")
        elif col in gap_cols_idx:
            gap_col_name = ("GapHariIni", "TotalGap", "GapPeriode")[gap_cols_idx.index(col)]
            v = r_data[gap_col_name]
            cell.set_text_props(color="#16a34a" if v >= 0 else "#dc2626")
        if is_total:
            cell.set_text_props(fontweight="bold")
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="jpg", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def generate_scoreboard_pdf(scoreboards: dict, periode_label: str = "") -> bytes:
    """scoreboards: {kategori_label(str): df_score}. Menghasilkan satu PDF landscape
    berisi seluruh kategori Scoreboard, siap dibagikan ke tim."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), topMargin=1 * cm, bottomMargin=1 * cm,
                             leftMargin=0.8 * cm, rightMargin=0.8 * cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Title"], fontSize=15, textColor=rl_colors.HexColor("#111827"))
    sub_style = ParagraphStyle("sub", parent=styles["Normal"], fontSize=9, textColor=rl_colors.HexColor("#4b5563"))
    elements = [Paragraph("Scoreboard Omset MFlash", title_style)]
    if periode_label:
        elements.append(Paragraph(periode_label, sub_style))
    elements.append(Spacer(1, 10))

    headers = ["CABANG"] + [label for _, label in _SCOREBOARD_COL_ORDER]
    for title, df_score in scoreboards.items():
        elements.append(Paragraph(title, ParagraphStyle("h2", parent=styles["Heading2"], fontSize=12,
                                                          textColor=rl_colors.HexColor("#111827"))))
        elements.append(Spacer(1, 4))
        if df_score.empty:
            elements.append(Paragraph("Belum ada data/target untuk kategori ini.", styles["Normal"]))
            elements.append(Spacer(1, 10))
            continue
        data = [headers]
        text_colors = []
        for _, r in df_score.iterrows():
            row = [r["Cabang"]] + [_fmt_scoreboard_cell(col, r[col]) for col, _ in _SCOREBOARD_COL_ORDER]
            data.append(row)
        n_cols = len(headers)
        col_w = (26 * cm) / n_cols
        table = Table(data, colWidths=[col_w] * n_cols, repeatRows=1)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#111827")),
            ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#d1d5db")),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, -1), 6.7),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        pct_idx = 1 + [c for c, _ in _SCOREBOARD_COL_ORDER].index("PctPencapaian")
        for i, (_, r) in enumerate(df_score.iterrows(), start=1):
            if str(r["Cabang"]) == "SMM":
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), rl_colors.HexColor("#f3f4f6")))
                style_cmds.append(("FONTNAME", (0, i), (-1, i), "Helvetica-Bold"))
            pct_color = rl_colors.HexColor(pencapaian_color(r["PctPencapaian"]))
            style_cmds.append(("TEXTCOLOR", (pct_idx, i), (pct_idx, i), pct_color))
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
        elements.append(Spacer(1, 16))

    doc.build(elements)
    buf.seek(0)
    return buf.read()


# ========================= Progress Ring, Pie, Daily Charts =========================

def render_progress_ring(label: str, pct: float, sub: str = ""):
    pct_display = max(0.0, pct)
    color = pencapaian_color(pct)
    fig = go.Figure(go.Pie(
        values=[min(pct_display, 1.0), max(0.0, 1.0 - min(pct_display, 1.0))],
        hole=0.72, marker=dict(colors=[color, "#e5e7eb"]), textinfo="none", sort=False, direction="clockwise",
    ))
    fig.update_layout(
        showlegend=False, margin=dict(l=0, r=0, t=0, b=0), height=180,
        annotations=[dict(text=f"<b>{format_percent(pct)}</b>", x=0.5, y=0.5, font_size=20, showarrow=False,
                           font_color=color)],
    )
    st.plotly_chart(fig, use_container_width=True, key=f"ring_{label}_{sub}")
    st.markdown(f"<div style='text-align:center;font-weight:600;'>{label}</div>", unsafe_allow_html=True)


def render_contribution_pie(labels, values, key):
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.4))
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=320)
    st.plotly_chart(fig, use_container_width=True, key=key)


def build_daily_progress(df_main: pd.DataFrame, tanggal_acuan: date, target_map: dict,
                          kategori: str, selected_branches=None) -> pd.DataFrame:
    if df_main is None or df_main.empty or tanggal_acuan is None:
        return pd.DataFrame(columns=["Tanggal", "Actual", "Target"])
    start, end, total_hari, hari_berjalan, sisa_hari = _quarter_bounds(tanggal_acuan)
    sub = df_main if kategori == "Omset All" else df_main[df_main["Kelompok"] == kategori]
    if selected_branches:
        sub = sub[sub["Cabang"].isin(selected_branches)]
    kat_targets = target_map.get(kategori, {}) if target_map else {}
    if selected_branches:
        total_target = sum(v for k, v in kat_targets.items() if k in selected_branches)
    else:
        total_target = sum(kat_targets.values())
    daily_target = (total_target / total_hari) if total_hari else 0.0
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    g = sub.groupby("Tanggal")["TotalHarga"].sum().reindex(
        pd.date_range(start, tanggal_acuan).date, fill_value=0
    ).reset_index()
    g.columns = ["Tanggal", "Actual"]
    g["ActualCum"] = g["Actual"].cumsum()
    g["TargetCum"] = [daily_target * (i + 1) for i in range(len(g))]
    return g


def render_daily_progress_chart(g: pd.DataFrame, key: str):
    if g.empty:
        st.info("Belum ada data untuk grafik progres harian.")
        return
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=g["Tanggal"], y=g["ActualCum"], mode="lines+markers", name="Realisasi Kumulatif",
                              line=dict(color="#1d4ed8", width=3)))
    fig.add_trace(go.Scatter(x=g["Tanggal"], y=g["TargetCum"], mode="lines", name="Pace Target",
                              line=dict(color="#dc2626", width=2, dash="dash")))
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=340, legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True, key=key)


def build_daily_history(log_df: pd.DataFrame, cabang=None, bulan=None, tahun=None) -> pd.DataFrame:
    if log_df is None or log_df.empty:
        return pd.DataFrame()
    d = log_df.copy()
    if cabang:
        d = d[d["Cabang"] == cabang]
    if bulan:
        d = d[d["Bulan"] == bulan]
    if tahun:
        d = d[d["Tahun"] == tahun]
    return d.sort_values("Tanggal")


def render_daily_history_chart(d: pd.DataFrame, key: str):
    if d.empty:
        st.info("Tidak ada riwayat untuk filter ini.")
        return
    fig = px.bar(d, x="Tanggal", y="OmsetHarian", color_discrete_sequence=["#0f766e"])
    fig.update_layout(margin=dict(l=10, r=10, t=20, b=10), height=320)
    st.plotly_chart(fig, use_container_width=True, key=key)


# ========================= 6 Pilar MFlash =========================

def build_pilar_summary(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    cols = ["Pilar", "Omset", "Qty", "GrossProfit"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    d = df
    if tanggal_acuan is not None:
        d = d[(d["Tahun"] == tanggal_acuan.year) & (d["Bulan"] == tanggal_acuan.month) &
              (d["Tanggal"].apply(lambda x: x.day) <= tanggal_acuan.day)]
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame(columns=cols)
    g = d.groupby("Pilar").agg(Omset=("TotalHarga", "sum"), Qty=("Qty", "sum"), GrossProfit=("GrossProfit", "sum")).reset_index()
    g["Pilar"] = pd.Categorical(g["Pilar"], categories=PILAR_ORDER, ordered=True)
    return g.sort_values("Pilar").reset_index(drop=True)


def build_pilar_by_branch(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    d = df
    if tanggal_acuan is not None:
        d = d[(d["Tahun"] == tanggal_acuan.year) & (d["Bulan"] == tanggal_acuan.month) &
              (d["Tanggal"].apply(lambda x: x.day) <= tanggal_acuan.day)]
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    if d.empty:
        return pd.DataFrame()
    piv = d.pivot_table(index="Cabang", columns="Pilar", values="TotalHarga", aggfunc="sum", fill_value=0)
    for p in PILAR_ORDER:
        if p not in piv.columns:
            piv[p] = 0
    piv = piv[PILAR_ORDER]
    piv["Total"] = piv.sum(axis=1)
    branches_present = order_branches(piv.index.tolist())
    piv = piv.reindex(branches_present)
    piv = piv.reset_index()
    return piv


def generate_pilar_insights(pilar_summary: pd.DataFrame) -> list:
    insights = []
    if pilar_summary.empty:
        return insights
    total = pilar_summary["Omset"].sum()
    if total <= 0:
        return insights
    for _, r in pilar_summary.iterrows():
        share = r["Omset"] / total
        if share < 0.03 and r["Pilar"] != "Lainnya":
            insights.append({
                "level": "warn", "category": r["Pilar"],
                "title": f"{_pilar_label(r['Pilar'])}: kontribusi masih kecil ({format_percent(share)})",
                "problem": f"Omset {_pilar_label(r['Pilar'])} baru {format_rupiah(r['Omset'])} dari total {format_rupiah(total)}.",
                "online": ["Promosikan kategori ini lebih intensif di media sosial"],
                "offline": ["Latih tim sales untuk cross-sell kategori ini ke pelanggan yang datang"],
            })
    return insights


def render_pilar_kpi_card(row):
    render_kpi_card_text(_pilar_label(row["Pilar"]), format_rupiah(row["Omset"]), PILAR_COLORS.get(row["Pilar"], "#111827"))


def render_pilar_table_html(piv: pd.DataFrame) -> str:
    if piv.empty:
        return "<i>Tidak ada data.</i>"
    header = "<th style='padding:5px 8px;border:1px solid #d1d5db;background:#111827;color:white;'>Cabang</th>"
    for p in PILAR_ORDER + ["Total"]:
        header += f"<th style='padding:5px 8px;border:1px solid #d1d5db;background:#374151;color:white;'>{p}</th>"
    body = ""
    for _, r in piv.iterrows():
        body += f"<tr><td style='padding:5px 8px;border:1px solid #d1d5db;font-weight:600;'>{r['Cabang']}</td>"
        for p in PILAR_ORDER + ["Total"]:
            body += f"<td style='padding:5px 8px;border:1px solid #d1d5db;text-align:right;'>{format_rupiah(r[p])}</td>"
        body += "</tr>"
    return f"""<div style="overflow-x:auto;"><table style="border-collapse:collapse;width:100%;font-size:0.85em;">
    <thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>"""


def render_pilar_summary_table_html(summary: pd.DataFrame) -> str:
    if summary.empty:
        return "<i>Tidak ada data.</i>"
    rows = ""
    total = summary["Omset"].sum()
    for _, r in summary.iterrows():
        share = (r["Omset"] / total) if total else 0
        rows += f"""<tr>
        <td style="padding:6px 10px;border:1px solid #d1d5db;font-weight:600;">{_pilar_label(r['Pilar'])}</td>
        <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_rupiah(r['Omset'])}</td>
        <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_percent(share)}</td>
        <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_number(r['Qty']) if r['Pilar'] in _PILAR_SHOW_QTY else '-'}</td>
        <td style="padding:6px 10px;border:1px solid #d1d5db;text-align:right;">{format_rupiah(r['GrossProfit'])}</td>
        </tr>"""
    return f"""<table style="border-collapse:collapse;width:100%;font-size:0.9em;">
    <thead><tr style="background:#111827;color:white;">
        <th style="padding:6px 10px;border:1px solid #d1d5db;text-align:left;">Pilar</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Omset</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">% Kontribusi</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Qty</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Gross Profit</th>
    </tr></thead><tbody>{rows}</tbody></table>"""


# ========================= Kontribusi Marketing Corporate vs Sales Retail =========================

_MC_CATEGORY_ORDER = ["All", "Service", "Gadget & Aksesoris"]
_MC_CATEGORY_LABELS = {"All": "Omset All", "Service": "Omset Service", "Gadget & Aksesoris": "Penjualan Gadget & Aksesoris"}
_MC_CATEGORY_ICONS = {"All": "💰", "Service": "🛠️", "Gadget & Aksesoris": "📱"}
_MC_CATEGORY_COLORS = {"All": "#1d4ed8", "Service": "#0f766e", "Gadget & Aksesoris": "#6d28d9"}


def _mc_filter_bulan_berjalan(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    if df is None or df.empty or tanggal_acuan is None:
        return pd.DataFrame()
    d = df[(df["Tahun"] == tanggal_acuan.year) & (df["Bulan"] == tanggal_acuan.month) &
           (df["Tanggal"].apply(lambda x: x.day) <= tanggal_acuan.day)]
    if selected_branches:
        d = d[d["Cabang"].isin(selected_branches)]
    return d


def build_mc_contribution_summary(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    d = _mc_filter_bulan_berjalan(df, tanggal_acuan, selected_branches)
    rows = []
    for kat in _MC_CATEGORY_ORDER:
        sub = d if kat == "All" else d[d["Kelompok"] == kat] if not d.empty else d
        total = float(sub["TotalHarga"].sum()) if not sub.empty else 0.0
        mc = float(sub[sub["KelompokPenjual"] == _MC_LABEL]["TotalHarga"].sum()) if not sub.empty else 0.0
        retail = total - mc
        pct_mc = (mc / total) if total else 0.0
        rows.append({"Kategori": kat, "OmsetTotal": total, "OmsetMC": mc, "OmsetRetail": retail, "PctMC": pct_mc})
    return pd.DataFrame(rows)


def build_mc_person_table(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    d = _mc_filter_bulan_berjalan(df, tanggal_acuan, selected_branches)
    if d.empty:
        return pd.DataFrame(columns=["Penjual", "Service", "Gadget & Aksesoris", "Total"])
    d = d[d["KelompokPenjual"] == _MC_LABEL]
    if d.empty:
        return pd.DataFrame(columns=["Penjual", "Service", "Gadget & Aksesoris", "Total"])
    piv = d.pivot_table(index="Penjual", columns="Kelompok", values="TotalHarga", aggfunc="sum", fill_value=0)
    for c in ["Service", "Gadget & Aksesoris"]:
        if c not in piv.columns:
            piv[c] = 0
    piv = piv[["Service", "Gadget & Aksesoris"]]
    piv["Total"] = piv.sum(axis=1)
    piv = piv.sort_values("Total", ascending=False).reset_index()
    total_row = pd.DataFrame([{
        "Penjual": "TOTAL", "Service": piv["Service"].sum(),
        "Gadget & Aksesoris": piv["Gadget & Aksesoris"].sum(), "Total": piv["Total"].sum(),
    }])
    return pd.concat([piv, total_row], ignore_index=True)


def build_retail_by_branch(df: pd.DataFrame, tanggal_acuan: date, selected_branches=None) -> pd.DataFrame:
    d = _mc_filter_bulan_berjalan(df, tanggal_acuan, selected_branches)
    if d.empty:
        return pd.DataFrame(columns=["Cabang", "Service", "Gadget & Aksesoris", "Total"])
    d = d[d["KelompokPenjual"] == _RETAIL_LABEL]
    if d.empty:
        return pd.DataFrame(columns=["Cabang", "Service", "Gadget & Aksesoris", "Total"])
    piv = d.pivot_table(index="Cabang", columns="Kelompok", values="TotalHarga", aggfunc="sum", fill_value=0)
    for c in ["Service", "Gadget & Aksesoris"]:
        if c not in piv.columns:
            piv[c] = 0
    piv = piv[["Service", "Gadget & Aksesoris"]]
    piv["Total"] = piv.sum(axis=1)
    branches_present = order_branches(piv.index.tolist())
    piv = piv.reindex(branches_present).fillna(0).reset_index()
    total_row = pd.DataFrame([{
        "Cabang": "TOTAL", "Service": piv["Service"].sum(),
        "Gadget & Aksesoris": piv["Gadget & Aksesoris"].sum(), "Total": piv["Total"].sum(),
    }])
    return pd.concat([piv, total_row], ignore_index=True)


def render_mc_contribution_card(kat: str, row):
    icon = _MC_CATEGORY_ICONS.get(kat, "")
    color = _MC_CATEGORY_COLORS.get(kat, "#111827")
    label = _MC_CATEGORY_LABELS.get(kat, kat)
    render_kpi_card_text(f"Kontribusi MC — {label}", format_rupiah(row["OmsetMC"]), color, icon,
                          sub=f"{format_percent(row['PctMC'])} dari {format_rupiah(row['OmsetTotal'])}")


def render_mc_split_donut(row, key):
    fig = go.Figure(go.Pie(labels=["Marketing Corporate", "Sales Retail"],
                            values=[row["OmsetMC"], row["OmsetRetail"]],
                            hole=0.5, marker=dict(colors=["#1d4ed8", "#9ca3af"])))
    fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=260)
    st.plotly_chart(fig, use_container_width=True, key=key)


def render_mc_person_table_html(mc_person: pd.DataFrame) -> str:
    if mc_person.empty:
        return "<i>Belum ada kontribusi Marketing Corporate pada periode ini.</i>"
    rows = ""
    for _, r in mc_person.iterrows():
        is_total = r["Penjual"] == "TOTAL"
        bg = "#f3f4f6" if is_total else "white"
        fw = "800" if is_total else "500"
        rows += f"""<tr style="background:{bg};">
        <td style="padding:5px 10px;border:1px solid #d1d5db;font-weight:{fw};">{r['Penjual']}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};">{format_rupiah(r['Service'])}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};">{format_rupiah(r['Gadget & Aksesoris'])}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};">{format_rupiah(r['Total'])}</td>
        </tr>"""
    return f"""<table style="border-collapse:collapse;width:100%;font-size:0.88em;">
    <thead><tr style="background:#1d4ed8;color:white;">
        <th style="padding:6px 10px;border:1px solid #d1d5db;text-align:left;">Nama (Marketing Corporate)</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Service</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Gadget & Aksesoris</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Total</th>
    </tr></thead><tbody>{rows}</tbody></table>"""


def render_retail_by_branch_table_html(retail_by_branch: pd.DataFrame) -> str:
    if retail_by_branch.empty:
        return "<i>Belum ada kontribusi Sales Retail pada periode ini.</i>"
    non_total = retail_by_branch[retail_by_branch["Cabang"] != "TOTAL"]
    avg_service = non_total["Service"].mean() if not non_total.empty else 0
    avg_gadget = non_total["Gadget & Aksesoris"].mean() if not non_total.empty else 0
    rows = ""
    for _, r in retail_by_branch.iterrows():
        is_total = r["Cabang"] == "TOTAL"
        bg = "#f3f4f6" if is_total else "white"
        fw = "800" if is_total else "500"
        svc_color = "#111827" if is_total else ("#16a34a" if r["Service"] >= avg_service else "#dc2626")
        gdg_color = "#111827" if is_total else ("#16a34a" if r["Gadget & Aksesoris"] >= avg_gadget else "#dc2626")
        rows += f"""<tr style="background:{bg};">
        <td style="padding:5px 10px;border:1px solid #d1d5db;font-weight:{fw};">{r['Cabang']}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};color:{svc_color};">{format_rupiah(r['Service'])}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};color:{gdg_color};">{format_rupiah(r['Gadget & Aksesoris'])}</td>
        <td style="padding:5px 10px;border:1px solid #d1d5db;text-align:right;font-weight:{fw};">{format_rupiah(r['Total'])}</td>
        </tr>"""
    return f"""<div style="border:2px solid #b45309;border-radius:10px;overflow:hidden;">
    <div style="background:#b45309;color:white;padding:8px 12px;font-weight:700;">🏬 Kontribusi Sales Retail per Cabang</div>
    <table style="border-collapse:collapse;width:100%;font-size:0.88em;">
    <thead><tr style="background:#fef3c7;">
        <th style="padding:6px 10px;border:1px solid #d1d5db;text-align:left;">Cabang</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Service</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Gadget & Aksesoris</th>
        <th style="padding:6px 10px;border:1px solid #d1d5db;">Total</th>
    </tr></thead><tbody>{rows}</tbody></table></div>"""


_MC_MARKETING_ACTION_PLAN = {
    "online": ["Follow-up leads corporate via email/WA blast mingguan", "Konten studi kasus kerja sama korporat"],
    "offline": ["Kunjungan langsung ke kantor/perusahaan sekitar cabang", "Ikuti pameran/expo bisnis lokal"],
}


def generate_mc_insights(mc_summary: pd.DataFrame, mc_person: pd.DataFrame) -> list:
    insights = []
    if mc_summary.empty:
        return insights
    row_all = mc_summary[mc_summary["Kategori"] == "All"]
    if not row_all.empty and float(row_all.iloc[0]["PctMC"]) < 0.05 and float(row_all.iloc[0]["OmsetTotal"]) > 0:
        insights.append({
            "level": "warn", "category": "Marketing Corporate",
            "title": f"Kontribusi Marketing Corporate masih kecil ({format_percent(row_all.iloc[0]['PctMC'])})",
            "problem": f"Dari total Omset All {format_rupiah(row_all.iloc[0]['OmsetTotal'])}, MC baru menyumbang {format_rupiah(row_all.iloc[0]['OmsetMC'])}.",
            "online": _MC_MARKETING_ACTION_PLAN["online"], "offline": _MC_MARKETING_ACTION_PLAN["offline"],
        })
    return insights


def generate_retail_branch_insights(retail_by_branch: pd.DataFrame) -> list:
    insights = []
    if retail_by_branch.empty:
        return insights
    d = retail_by_branch[retail_by_branch["Cabang"] != "TOTAL"]
    if d.empty:
        return insights
    for kat_col in ["Service", "Gadget & Aksesoris"]:
        avg = d[kat_col].mean()
        if avg <= 0:
            continue
        top = d.loc[d[kat_col].idxmax()]
        insights.append({
            "level": "good", "category": kat_col,
            "title": f"{top['Cabang']}: Retail terbaik di kategori {kat_col}",
            "problem": f"Kontribusi {format_rupiah(top[kat_col])}, di atas rata-rata cabang ({format_rupiah(avg)}).",
            "online": [], "offline": [],
        })
        plan = _CATEGORY_ACTION_PLANS.get(kat_col, {"online": [], "offline": []})
        for _, r in d[d[kat_col] < 0.6 * avg].iterrows():
            insights.append({
                "level": "bad", "category": kat_col,
                "title": f"{r['Cabang']}: Retail {kat_col} di bawah rata-rata",
                "problem": f"Kontribusi hanya {format_rupiah(r[kat_col])}, jauh di bawah rata-rata cabang ({format_rupiah(avg)}).",
                "online": plan["online"], "offline": plan["offline"],
            })
    return insights


# ========================= Ledger Riwayat Harian =========================

_HISTORY_LOG_COLUMNS = ["Tanggal", "Tahun", "Bulan", "Cabang", "Kelompok", "OmsetHarian"]
_LOG_PATH = os.path.join(LOG_DIR, "omset_harian_log.csv")


def _read_log() -> pd.DataFrame:
    if not os.path.exists(_LOG_PATH):
        return pd.DataFrame(columns=_HISTORY_LOG_COLUMNS)
    try:
        d = pd.read_csv(_LOG_PATH, parse_dates=["Tanggal"])
        d["Tanggal"] = d["Tanggal"].dt.date
        return d
    except Exception:
        return pd.DataFrame(columns=_HISTORY_LOG_COLUMNS)


def _upsert_log(new_rows: pd.DataFrame):
    if new_rows is None or new_rows.empty:
        return
    os.makedirs(LOG_DIR, exist_ok=True)
    existing = _read_log()
    combined = pd.concat([existing, new_rows], ignore_index=True)
    combined = combined.drop_duplicates(subset=["Tanggal", "Cabang", "Kelompok"], keep="last")
    combined.to_csv(_LOG_PATH, index=False)
    if _GH_ENABLED:
        try:
            with open(_LOG_PATH, "rb") as f:
                github_upload_file("data/log/omset_harian_log.csv", f.read(), "update ledger")
        except Exception:
            pass


def build_upload_log(df_main: pd.DataFrame):
    if df_main is None or df_main.empty:
        return
    g_all = df_main.groupby(["Cabang", "Tanggal"])["TotalHarga"].sum().reset_index()
    g_all["Kelompok"] = "Omset All"
    g_kat = df_main.groupby(["Cabang", "Tanggal", "Kelompok"])["TotalHarga"].sum().reset_index()
    g = pd.concat([g_all.rename(columns={"TotalHarga": "OmsetHarian"}),
                    g_kat.rename(columns={"TotalHarga": "OmsetHarian"})], ignore_index=True)
    g["Tahun"] = g["Tanggal"].apply(lambda d: d.year)
    g["Bulan"] = g["Tanggal"].apply(lambda d: d.month)
    _upsert_log(g[_HISTORY_LOG_COLUMNS])


def build_corp_upload_log(*args, **kwargs):
    pass


def compute_corp_hari_ini(*args, **kwargs):
    return 0.0


# ========================= Export Laporan Lengkap (PPTX & PDF) =========================

def _pptx_add_title_slide(prs, title, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    if subtitle and len(slide.placeholders) > 1:
        slide.placeholders[1].text = subtitle
    return slide


def _pptx_add_table_slide(prs, title, df: pd.DataFrame, money_cols=None):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    if df.empty:
        return slide
    money_cols = money_cols or []
    rows, cols = df.shape[0] + 1, df.shape[1]
    left, top, width, height = Inches(0.4), Inches(1.4), Inches(9.2), Inches(0.35 * min(rows, 14))
    table_shape = slide.shapes.add_table(min(rows, 15), cols, left, top, width, height)
    table = table_shape.table
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i, (_, r) in enumerate(df.head(14).iterrows(), start=1):
        for j, col in enumerate(df.columns):
            v = r[col]
            table.cell(i, j).text = format_rupiah(v) if col in money_cols else str(v)
    return slide


def _build_report_sections(df_main, scoreboards: dict, periode_label: str):
    return scoreboards


def generate_pptx_report(df_main, scoreboards: dict, periode_label: str) -> bytes:
    prs = Presentation()
    _pptx_add_title_slide(prs, "Laporan Omset MFlash", periode_label)
    for title, df_score in scoreboards.items():
        _pptx_add_table_slide(prs, title, df_score, money_cols=[c for c, _ in _SCOREBOARD_COL_ORDER if c in MONEY_COLS])
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


def generate_pdf_report(df_main, scoreboards: dict, periode_label: str) -> bytes:
    return generate_scoreboard_pdf(scoreboards, periode_label)


# ========================= UI ==========================

for _d in [MAIN_DATA_DIR, ADS_DATA_DIR, WALKIN_DATA_DIR, TARGET_DATA_DIR, CORP_DATA_DIR, LOG_DIR]:
    os.makedirs(_d, exist_ok=True)

if _GH_ENABLED and not st.session_state.get("_gh_synced_v1"):
    sync_data_from_github()
    st.session_state["_gh_synced_v1"] = True

if not st.session_state.get("_main_deduped_v1"):
    _dedupe_main_files()
    st.session_state["_main_deduped_v1"] = True

_logo_col1, _logo_col2 = st.columns([1, 6])
with _logo_col1:
    try:
        st.image(base64.b64decode(LOGO_BASE64), width=90)
    except Exception:
        pass
with _logo_col2:
    st.markdown("## 📊 Dashboard Omset MFlash")
    st.caption("Omset • Iklan • Walk-in • 6 Pilar MFlash • Kontribusi Marketing Corporate")

# ---- Sidebar uploads ----
with st.sidebar:
    st.markdown("### 📁 Upload Data")

    with st.expander("💰 Data Omset (Faktur Penjualan)", expanded=False):
        up_main = st.file_uploader("Upload file Omset (.xlsx)", type=["xlsx", "xls"],
                                    accept_multiple_files=True, key="up_main")
        if up_main:
            for f in up_main:
                fpath = os.path.join(MAIN_DATA_DIR, sanitize_filename(f.name))
                with open(fpath, "wb") as out:
                    out.write(f.getbuffer())
                if _GH_ENABLED:
                    with open(fpath, "rb") as rf:
                        github_upload_file(f"data/main/{sanitize_filename(f.name)}", rf.read(), "upload omset")
            _dedupe_main_files()
            st.success(f"{len(up_main)} file omset diunggah.")
        existing_main = sorted(os.listdir(MAIN_DATA_DIR)) if os.path.isdir(MAIN_DATA_DIR) else []
        for fname in existing_main:
            c1, c2 = st.columns([4, 1])
            c1.caption(fname)
            if c2.button("🗑️", key=f"del_main_{fname}"):
                os.remove(os.path.join(MAIN_DATA_DIR, fname))
                if _GH_ENABLED:
                    github_delete_file(f"data/main/{fname}")
                st.rerun()

    with st.expander("📢 Data Iklan (Meta Ads)", expanded=False):
        up_ads = st.file_uploader("Upload export Meta Ads (.xlsx)", type=["xlsx", "xls"],
                                   accept_multiple_files=True, key="up_ads")
        if up_ads:
            for f in up_ads:
                fpath = os.path.join(ADS_DATA_DIR, sanitize_filename(f.name))
                with open(fpath, "wb") as out:
                    out.write(f.getbuffer())
                if _GH_ENABLED:
                    with open(fpath, "rb") as rf:
                        github_upload_file(f"data/ads/{sanitize_filename(f.name)}", rf.read(), "upload ads")
            st.success(f"{len(up_ads)} file iklan diunggah.")
        existing_ads = sorted(os.listdir(ADS_DATA_DIR)) if os.path.isdir(ADS_DATA_DIR) else []
        for fname in existing_ads:
            c1, c2 = st.columns([4, 1])
            c1.caption(fname)
            if c2.button("🗑️", key=f"del_ads_{fname}"):
                os.remove(os.path.join(ADS_DATA_DIR, fname))
                if _GH_ENABLED:
                    github_delete_file(f"data/ads/{fname}")
                st.rerun()

    with st.expander("🚶 Data Walk-in", expanded=False):
        up_walkin = st.file_uploader("Upload rincian pengiriman pesanan (.xlsx)", type=["xlsx", "xls"],
                                      accept_multiple_files=True, key="up_walkin")
        if up_walkin:
            for f in up_walkin:
                fpath = os.path.join(WALKIN_DATA_DIR, sanitize_filename(f.name))
                with open(fpath, "wb") as out:
                    out.write(f.getbuffer())
                if _GH_ENABLED:
                    with open(fpath, "rb") as rf:
                        github_upload_file(f"data/walkin/{sanitize_filename(f.name)}", rf.read(), "upload walkin")
            st.success(f"{len(up_walkin)} file walk-in diunggah.")
        existing_walkin = sorted(os.listdir(WALKIN_DATA_DIR)) if os.path.isdir(WALKIN_DATA_DIR) else []
        for fname in existing_walkin:
            c1, c2 = st.columns([4, 1])
            c1.caption(fname)
            if c2.button("🗑️", key=f"del_walkin_{fname}"):
                os.remove(os.path.join(WALKIN_DATA_DIR, fname))
                if _GH_ENABLED:
                    github_delete_file(f"data/walkin/{fname}")
                st.rerun()

    with st.expander("🎯 Target Omset (opsional)", expanded=False):
        st.caption("Jika tidak diupload, target akan otomatis dibaca dari sheet 'Scoreboard' pada file Omset di atas (jika ada).")
        st.download_button("⬇️ Unduh Template Target", data=make_target_template(),
                            file_name="template_target.xlsx", key="dl_target_template")
        up_target = st.file_uploader("Upload Target (.xlsx)", type=["xlsx", "xls"], key="up_target")
        if up_target:
            fpath = os.path.join(TARGET_DATA_DIR, sanitize_filename(up_target.name))
            with open(fpath, "wb") as out:
                out.write(up_target.getbuffer())
            if _GH_ENABLED:
                with open(fpath, "rb") as rf:
                    github_upload_file(f"data/target/{sanitize_filename(up_target.name)}", rf.read(), "upload target")
            st.success("File target diunggah.")
        existing_target = sorted(os.listdir(TARGET_DATA_DIR)) if os.path.isdir(TARGET_DATA_DIR) else []
        for fname in existing_target:
            c1, c2 = st.columns([4, 1])
            c1.caption(fname)
            if c2.button("🗑️", key=f"del_target_{fname}"):
                os.remove(os.path.join(TARGET_DATA_DIR, fname))
                if _GH_ENABLED:
                    github_delete_file(f"data/target/{fname}")
                st.rerun()

    with st.expander("🤝 Data Corporate (manual, opsional)", expanded=False):
        st.download_button("⬇️ Unduh Template Corporate", data=make_corporate_template(),
                            file_name="template_corporate.xlsx", key="dl_corp_template")
        up_corp = st.file_uploader("Upload Corporate (.xlsx)", type=["xlsx", "xls"], key="up_corp")
        if up_corp:
            fpath = os.path.join(CORP_DATA_DIR, sanitize_filename(up_corp.name))
            with open(fpath, "wb") as out:
                out.write(up_corp.getbuffer())
            if _GH_ENABLED:
                with open(fpath, "rb") as rf:
                    github_upload_file(f"data/corp/{sanitize_filename(up_corp.name)}", rf.read(), "upload corp")
            st.success("File corporate diunggah.")


# ---- Load all data ----
df_main = load_all_main_data()
df_ads = load_all_ads_data()
df_walkin = load_all_walkin_data()

if df_main is not None and not df_main.empty:
    build_upload_log(df_main)

# ---- Resolve target (dedicated file > auto-extract from Scoreboard sheet) ----
target_map = {v: {} for v in _SECTION_MARKERS.values()}
_target_files = [os.path.join(TARGET_DATA_DIR, f) for f in os.listdir(TARGET_DATA_DIR)] if os.path.isdir(TARGET_DATA_DIR) else []
_target_files = [f for f in _target_files if f.lower().endswith((".xlsx", ".xls"))]
if _target_files:
    tdf = load_target_data(_target_files[-1])
    if not tdf.empty and "Cabang" in tdf.columns:
        col_map = {
            "Omset All": "Target Omset All (Kuartal)",
            "Service": "Target Service (Kuartal)",
            "Gadget & Aksesoris": "Target Gadget & Aksesoris (Kuartal)",
        }
        for kat, col in col_map.items():
            if col in tdf.columns:
                for _, r in tdf.iterrows():
                    cab = str(r["Cabang"]).strip().upper()
                    val = _to_float_or_none(r[col])
                    if cab in BRANCH_ORDER and val:
                        target_map[kat][cab] = val

_main_files_all = [os.path.join(MAIN_DATA_DIR, f) for f in os.listdir(MAIN_DATA_DIR)] if os.path.isdir(MAIN_DATA_DIR) else []
_main_files_all = [f for f in _main_files_all if f.lower().endswith((".xlsx", ".xls"))]
_auto_targets = extract_scoreboard_target(_main_files_all) if _main_files_all else {v: {} for v in _SECTION_MARKERS.values()}
for kat in target_map:
    for cab, val in _auto_targets.get(kat, {}).items():
        target_map[kat].setdefault(cab, val)

_snapshot_date = extract_scoreboard_snapshot_date(_main_files_all) if _main_files_all else None

if df_main is None or df_main.empty:
    st.info("Belum ada data Omset yang diunggah. Silakan upload file di sidebar untuk mulai memakai dashboard.")

# ---- Filter row ----
_default_tanggal = _snapshot_date or (df_main["Tanggal"].max() if df_main is not None and not df_main.empty else date.today())
if isinstance(_default_tanggal, datetime):
    _default_tanggal = _default_tanggal.date()

filt_col1, filt_col2 = st.columns([1, 3])
with filt_col1:
    tanggal_acuan = st.date_input("📅 Tanggal Acuan", value=_default_tanggal, key="tanggal_acuan")
with filt_col2:
    all_branches_present = order_branches(df_main["Cabang"].dropna().unique().tolist()) if df_main is not None and not df_main.empty else BRANCH_ORDER
    selected_branches = st.multiselect("🏬 Filter Cabang (kosongkan = semua)", options=all_branches_present, default=[], key="branch_filter")
selected_branches = selected_branches or None

periode_label = f"{BULAN_ID.get(tanggal_acuan.month, '')} {tanggal_acuan.year} (s/d tanggal {tanggal_acuan.day})"

# ---- Build scoreboards for all 3 categories (dipakai di beberapa tab) ----
scoreboards = {}
for kat in SCOREBOARD_KATEGORI:
    scoreboards[_SCOREBOARD_KATEGORI_LABEL[kat]] = build_scoreboard(df_main, tanggal_acuan, target_map, kat, selected_branches)

log_df = _read_log()

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🏠 Ringkasan", "🏆 Scoreboard", "📢 Iklan", "🚶 Walk-in", "🧩 6 Pilar", "🤝 Kontribusi MC",
])

# ========================= TAB 1: RINGKASAN =========================
with tab1:
    st.subheader(f"Ringkasan — {periode_label}")
    sb_all = scoreboards[_SCOREBOARD_KATEGORI_LABEL["Omset All"]]
    sb_service = scoreboards[_SCOREBOARD_KATEGORI_LABEL["Service"]]
    sb_gadget = scoreboards[_SCOREBOARD_KATEGORI_LABEL["Gadget & Aksesoris"]]

    if not sb_all.empty:
        total_row = sb_all[sb_all["Cabang"] == "SMM"].iloc[0]
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            render_kpi_card_text("Omset S/D Hari Ini", format_rupiah(total_row["SdHariIni"]), "#1d4ed8", "💰")
        with k2:
            render_kpi_card_text("Target Kuartal", format_rupiah(total_row["OmsetSamurai"]), "#0f766e", "🎯")
        with k3:
            render_kpi_card_text("% Pencapaian", format_percent(total_row["PctPencapaian"]),
                                  pencapaian_color(total_row["PctPencapaian"]), "📈")
        with k4:
            render_kpi_card_text("Kejar Target/Hari", format_rupiah(total_row["KejarPerhari"]), "#dc2626", "⏱️")

        st.markdown("#### Progres Pencapaian")
        r1, r2, r3 = st.columns(3)
        with r1:
            render_progress_ring("Omset All", total_row["PctPencapaian"], "all")
        if not sb_service.empty:
            with r2:
                tr_s = sb_service[sb_service["Cabang"] == "SMM"].iloc[0]
                render_progress_ring("Service", tr_s["PctPencapaian"], "service")
        if not sb_gadget.empty:
            with r3:
                tr_g = sb_gadget[sb_gadget["Cabang"] == "SMM"].iloc[0]
                render_progress_ring("Gadget & Aksesoris", tr_g["PctPencapaian"], "gadget")

        st.markdown("#### Grafik Progres Harian (Realisasi vs Pace Target) — Kuartal Berjalan")
        g = build_daily_progress(df_main, tanggal_acuan, target_map, "Omset All", selected_branches)
        render_daily_progress_chart(g, "daily_progress_ringkasan")

        if df_main is not None and not df_main.empty:
            d_period = df_main[(df_main["Tahun"] == tanggal_acuan.year) & (df_main["Bulan"] == tanggal_acuan.month) &
                                (df_main["Tanggal"].apply(lambda x: x.day) <= tanggal_acuan.day)]
            if selected_branches:
                d_period = d_period[d_period["Cabang"].isin(selected_branches)]
            if not d_period.empty:
                st.markdown("#### Kontribusi Omset: Service vs Gadget & Aksesoris vs Marketing Corporate")
                svc = float(d_period[d_period["Kelompok"] == "Service"]["TotalHarga"].sum())
                gdg = float(d_period[d_period["Kelompok"] == "Gadget & Aksesoris"]["TotalHarga"].sum())
                mc = float(d_period[d_period["KelompokPenjual"] == _MC_LABEL]["TotalHarga"].sum())
                retail = float(d_period["TotalHarga"].sum()) - mc
                cc1, cc2 = st.columns(2)
                with cc1:
                    render_contribution_pie(["Service", "Gadget & Aksesoris"], [svc, gdg], "pie_kategori")
                with cc2:
                    render_contribution_pie(["Marketing Corporate", "Sales Retail"], [mc, retail], "pie_mc_retail")

        st.markdown("#### Riwayat Pencapaian Harian")
        hc1, hc2, hc3 = st.columns(3)
        with hc1:
            hist_cabang = st.selectbox("Cabang", ["Semua"] + all_branches_present, key="hist_cabang")
        with hc2:
            hist_bulan = st.selectbox("Bulan", ["Semua"] + list(BULAN_ID.values()), key="hist_bulan")
        with hc3:
            tahun_opts = sorted(log_df["Tahun"].unique().tolist()) if not log_df.empty else [date.today().year]
            hist_tahun = st.selectbox("Tahun", ["Semua"] + tahun_opts, key="hist_tahun")
        d_hist = build_daily_history(
            log_df[log_df["Kelompok"] == "Omset All"] if not log_df.empty else log_df,
            cabang=None if hist_cabang == "Semua" else hist_cabang,
            bulan=None if hist_bulan == "Semua" else list(BULAN_ID.values()).index(hist_bulan) + 1,
            tahun=None if hist_tahun == "Semua" else hist_tahun,
        )
        render_daily_history_chart(d_hist, "daily_history_ringkasan")
    else:
        st.info("Belum ada target/data untuk menampilkan ringkasan. Upload data Omset & Target di sidebar.")

# ========================= TAB 2: SCOREBOARD =========================
with tab2:
    st.subheader(f"Scoreboard — {periode_label}")
    st.caption(
        "Mengikuti persis rumus pada sheet Excel 'Scoreboard' & 'Data Periode': "
        "OMSET HARIAN = OMSET SAMURAI ÷ TOTAL HARI kuartal, EXPECTED VALUE = OMSET HARIAN × hari berjalan "
        "dalam kuartal, % PENCAPAIAN = S/D HARI INI ÷ EXPECTED VALUE."
    )
    if _snapshot_date:
        st.caption(f"Tanggal snapshot terdeteksi dari file Scoreboard: {_snapshot_date}")

    exp_c1, exp_c2, exp_c3 = st.columns([1, 1, 4])
    with exp_c1:
        st.download_button(
            "📄 Export Scoreboard (PDF)",
            data=generate_scoreboard_pdf(scoreboards, periode_label),
            file_name=f"Scoreboard_MFlash_{sanitize_filename(periode_label)}.pdf",
            mime="application/pdf",
            key="dl_scoreboard_pdf",
        )
    with exp_c2:
        _sb_zip_target = scoreboards[_SCOREBOARD_KATEGORI_LABEL["Omset All"]]
        st.download_button(
            "🖼️ Export Scoreboard All (JPG)",
            data=generate_scoreboard_table_image(_sb_zip_target, "SCOREBOARD OMSET ALL", periode_label),
            file_name=f"Scoreboard_OmsetAll_{sanitize_filename(periode_label)}.jpg",
            mime="image/jpeg",
            key="dl_scoreboard_jpg_all",
        )

    for kat in SCOREBOARD_KATEGORI:
        title = _SCOREBOARD_KATEGORI_LABEL[kat]
        df_score = scoreboards[title]
        st.markdown(render_scoreboard_html(df_score, title, periode_label), unsafe_allow_html=True)
        if not df_score.empty:
            jc1, jc2 = st.columns([1, 5])
            with jc1:
                st.download_button(
                    f"🖼️ JPG — {title}",
                    data=generate_scoreboard_table_image(df_score, title, periode_label),
                    file_name=f"{sanitize_filename(title)}_{sanitize_filename(periode_label)}.jpg",
                    mime="image/jpeg",
                    key=f"dl_jpg_{kat}",
                )
        st.markdown("<br/>", unsafe_allow_html=True)

# ========================= TAB 3: IKLAN =========================
with tab3:
    st.subheader("Performa Iklan (Meta Ads)")
    if df_ads is None or df_ads.empty:
        st.info("Belum ada data iklan yang diunggah.")
    else:
        d_ads = df_ads[df_ads["Cabang"].isin(selected_branches)] if selected_branches else df_ads
        agg = aggregate_ads_by_branch(d_ads)
        if not agg.empty:
            m1, m2, m3 = st.columns(3)
            with m1:
                render_kpi_card_text("Total Spend", format_rupiah(agg.get("AmountSpent", pd.Series([0])).sum()), "#1d4ed8", "💸")
            with m2:
                render_kpi_card_text("Total Reach", format_number(agg.get("Reach", pd.Series([0])).sum()), "#0f766e", "👥")
            with m3:
                render_kpi_card_text("Total Impressions", format_number(agg.get("Impressions", pd.Series([0])).sum()), "#6d28d9", "👁️")
            if "AmountSpent" in agg.columns:
                fig = px.bar(agg, x="Cabang", y="AmountSpent", text_auto=".2s", color_discrete_sequence=["#1d4ed8"])
                fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True, key="ads_bar")
            st.markdown("#### Insight & Rekomendasi")
            for ins in generate_ads_insights(agg):
                render_insight_card(ins)
        else:
            st.info("Kolom data iklan tidak lengkap untuk agregasi.")

# ========================= TAB 4: WALK-IN =========================
with tab4:
    st.subheader(f"Walk-in per Cabang — {periode_label}")
    if df_walkin is None or df_walkin.empty:
        st.info("Belum ada data walk-in yang diunggah.")
    else:
        d_current = aggregate_walkin_current_period(df_walkin, tanggal_acuan, selected_branches)
        d_current = _walkin_ordered(d_current)

        total_walkin = int(d_current["TotalWalkin"].sum()) if not d_current.empty else 0
        avg_walkin = _walkin_overall_avg(d_current)
        n_cabang = d_current["Cabang"].nunique() if not d_current.empty else 0

        m1, m2, m3 = st.columns(3)
        with m1:
            render_kpi_card_text("Total Walk-in", format_number(total_walkin), "#0f766e", "🚶")
        with m2:
            render_kpi_card_text("Rata-rata Walk-in/Hari", format_decimal(avg_walkin), "#1d4ed8", "📊")
        with m3:
            render_kpi_card_text("Jumlah Cabang", format_number(n_cabang), "#6d28d9", "🏬")

        if not d_current.empty:
            fig = px.bar(d_current, x="Cabang", y="TotalWalkin", text_auto=True, color_discrete_sequence=["#0f766e"])
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True, key="walkin_bar")

        st.markdown(render_walkin_table_html(d_current, periode_label), unsafe_allow_html=True)

        if not d_current.empty:
            dc1, dc2 = st.columns(2)
            with dc1:
                st.download_button("🖼️ Download Tabel (JPG)", data=generate_walkin_table_image(d_current, periode_label),
                                    file_name=f"Walkin_{sanitize_filename(periode_label)}.jpg", mime="image/jpeg",
                                    key="dl_walkin_jpg")
            with dc2:
                st.download_button("📄 Download Tabel (PDF)", data=generate_walkin_table_pdf(d_current, periode_label),
                                    file_name=f"Walkin_{sanitize_filename(periode_label)}.pdf", mime="application/pdf",
                                    key="dl_walkin_pdf")

        st.markdown("#### 💡 Insight & Rekomendasi Program Marketing")
        for ins in generate_walkin_marketing_insights(d_current):
            render_structured_insight_card(ins)

        st.markdown("#### 📈 Insight Tren Bulanan (Dibanding Bulan Lalu)")
        walkin_agg_hist = aggregate_walkin_monthly(df_walkin if not selected_branches else df_walkin[df_walkin["Cabang"].isin(selected_branches)])
        tren_insights = generate_walkin_insights(walkin_agg_hist)
        if tren_insights:
            for ins in tren_insights:
                render_structured_insight_card(ins)
        else:
            st.caption("Belum ada tren penurunan signifikan dibanding bulan lalu.")

# ========================= TAB 5: 6 PILAR MFLASH =========================
with tab5:
    st.subheader(f"6 Pilar MFlash — {periode_label}")
    if df_main is None or df_main.empty:
        st.info("Belum ada data Omset yang diunggah.")
    else:
        pilar_summary = build_pilar_summary(df_main, tanggal_acuan, selected_branches)
        if pilar_summary.empty:
            st.info("Tidak ada data untuk periode ini.")
        else:
            kpi_cols = st.columns(len(PILAR_ORDER))
            for i, p in enumerate(PILAR_ORDER):
                row = pilar_summary[pilar_summary["Pilar"] == p]
                if not row.empty:
                    with kpi_cols[i]:
                        render_pilar_kpi_card(row.iloc[0])

            pc1, pc2 = st.columns([1, 1])
            with pc1:
                fig = px.pie(pilar_summary, names="Pilar", values="Omset", hole=0.4,
                              color="Pilar", color_discrete_map=PILAR_COLORS)
                fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True, key="pilar_pie")
            with pc2:
                st.markdown(render_pilar_summary_table_html(pilar_summary), unsafe_allow_html=True)

            st.markdown("#### Detail Omset per Pilar per Cabang")
            piv = build_pilar_by_branch(df_main, tanggal_acuan, selected_branches)
            if not piv.empty:
                fig2 = go.Figure()
                for p in PILAR_ORDER:
                    fig2.add_trace(go.Bar(name=p, x=piv["Cabang"], y=piv[p], marker_color=PILAR_COLORS.get(p)))
                fig2.update_layout(barmode="stack", height=380, margin=dict(l=10, r=10, t=20, b=10),
                                    legend=dict(orientation="h", y=1.15))
                st.plotly_chart(fig2, use_container_width=True, key="pilar_stack")
                st.markdown(render_pilar_table_html(piv), unsafe_allow_html=True)

            st.markdown("#### Insight Kualitas Data")
            for ins in generate_pilar_insights(pilar_summary):
                render_structured_insight_card(ins)

# ========================= TAB 6: KONTRIBUSI MARKETING CORPORATE =========================
with tab6:
    st.subheader(f"🤝 Kontribusi Marketing Corporate — {periode_label}")
    if df_main is None or df_main.empty:
        st.info("Belum ada data Omset yang diunggah.")
    else:
        mc_summary = build_mc_contribution_summary(df_main, tanggal_acuan, selected_branches)
        mc_person = build_mc_person_table(df_main, tanggal_acuan, selected_branches)

        if not mc_summary.empty:
            cols = st.columns(len(_MC_CATEGORY_ORDER))
            for i, kat in enumerate(_MC_CATEGORY_ORDER):
                row = mc_summary[mc_summary["Kategori"] == kat]
                if not row.empty:
                    with cols[i]:
                        render_mc_contribution_card(kat, row.iloc[0])

            dc = st.columns(len(_MC_CATEGORY_ORDER))
            for i, kat in enumerate(_MC_CATEGORY_ORDER):
                row = mc_summary[mc_summary["Kategori"] == kat]
                if not row.empty:
                    with dc[i]:
                        st.caption(_MC_CATEGORY_LABELS[kat])
                        render_mc_split_donut(row.iloc[0], f"mc_donut_{kat}")

        st.markdown("#### Kontribusi per Orang (Marketing Corporate)")
        st.markdown(render_mc_person_table_html(mc_person), unsafe_allow_html=True)
        if not mc_person.empty:
            d_bar = mc_person[mc_person["Penjual"] != "TOTAL"]
            if not d_bar.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Service", x=d_bar["Penjual"], y=d_bar["Service"], marker_color="#0f766e"))
                fig.add_trace(go.Bar(name="Gadget & Aksesoris", x=d_bar["Penjual"], y=d_bar["Gadget & Aksesoris"], marker_color="#6d28d9"))
                fig.update_layout(barmode="stack", height=360, margin=dict(l=10, r=10, t=20, b=10),
                                   legend=dict(orientation="h", y=1.15))
                st.plotly_chart(fig, use_container_width=True, key="mc_person_bar")

        st.markdown("#### Insight Marketing Corporate")
        for ins in generate_mc_insights(mc_summary, mc_person):
            render_structured_insight_card(ins)

        st.markdown("---")
        retail_by_branch = build_retail_by_branch(df_main, tanggal_acuan, selected_branches)
        st.markdown(render_retail_by_branch_table_html(retail_by_branch), unsafe_allow_html=True)
        if not retail_by_branch.empty:
            d_bar2 = retail_by_branch[retail_by_branch["Cabang"] != "TOTAL"]
            if not d_bar2.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Service", x=d_bar2["Cabang"], y=d_bar2["Service"], marker_color="#0f766e"))
                fig.add_trace(go.Bar(name="Gadget & Aksesoris", x=d_bar2["Cabang"], y=d_bar2["Gadget & Aksesoris"], marker_color="#6d28d9"))
                fig.update_layout(barmode="stack", height=360, margin=dict(l=10, r=10, t=20, b=10),
                                   legend=dict(orientation="h", y=1.15))
                st.plotly_chart(fig, use_container_width=True, key="retail_branch_bar")

        st.markdown("#### Insight Sales Retail per Cabang")
        for ins in generate_retail_branch_insights(retail_by_branch):
            render_structured_insight_card(ins)

# ========================= EXPORT LAPORAN LENGKAP =========================
st.markdown("---")
st.markdown("### 📤 Export Laporan Lengkap")
ex1, ex2 = st.columns(2)
with ex1:
    st.download_button(
        "📊 Export ke PowerPoint (PPTX)",
        data=generate_pptx_report(df_main, scoreboards, periode_label),
        file_name=f"Laporan_MFlash_{sanitize_filename(periode_label)}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="dl_full_pptx",
    )
with ex2:
    st.download_button(
        "📄 Export ke PDF",
        data=generate_pdf_report(df_main, scoreboards, periode_label),
        file_name=f"Laporan_MFlash_{sanitize_filename(periode_label)}.pdf",
        mime="application/pdf",
        key="dl_full_pdf",
    )
