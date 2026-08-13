"""
Dashboard Omset MFlash
=======================
Menampilkan omset untuk 3 kategori:
  1. Omset All
  2. Omset Service
  3. Omset Gadget & Aksesoris
(+ Omset Marketing Corporate di tab Ringkasan, kalau datanya diupload)

Ada 2 tampilan:
  - Tab "Ringkasan": total omset & tren, bisa difilter Tahun / Bulan / Cabang.
  - Tab "Scoreboard": tabel per cabang gaya scoreboard (Target / Expected Value /
    Pencapaian / Gap / Kejar Target Per Hari / rata-rata omset bulan lalu vs bulan ini),
    dihitung dari Target yang diupload + tanggal acuan ("hari ini").

Semua data diisi lewat tombol upload (tidak perlu edit source code / repo GitHub).

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
import streamlit as st

# --------------------------------------------------------------------------------------
# Konfigurasi umum
# --------------------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard Omset MFlash", layout="wide", page_icon="📊")

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
os.makedirs(MAIN_DATA_DIR, exist_ok=True)
CORPORATE_DATA_PATH = os.path.join(DATA_DIR, "corporate_data.xlsx")
TARGET_DATA_PATH = os.path.join(DATA_DIR, "target_data.xlsx")

MAIN_SHEET_NAME = "Faktur Penjualan"
# Kolom wajib untuk hitung Omset Service vs Gadget & Aksesoris:
#   - Omset Service = baris dengan KATEGORI BARANG "Jasa" atau "Sparepart"
#   - Omset Gadget & Aksesoris = Omset All - Omset Service (semua baris lainnya)
CORE_COLUMNS = ["TGL FAKTUR", "KATEGORI BARANG", "TOTAL HARGA"]
REQUIRED_COLUMNS = ["CABANG"] + CORE_COLUMNS
MAX_MAIN_FILES = 50

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

KATEGORI_TABS = ["Omset All", "Omset Service", "Omset Gadget & Aksesoris"]


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
# Loader data Marketing Corporate (Tahun, Bulan, Cabang, Omset)
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
    info.append(["- Tahun: angka 4 digit, contoh 2026"])
    info.append(["- Bulan: nama bulan dalam Bahasa Indonesia, contoh Januari, Februari, ... Desember"])
    info.append(["- Cabang: nama cabang, harus konsisten dengan nama cabang di data utama"])
    info.append(["- Omset: angka total omset marketing corporate untuk cabang & bulan tersebut"])
    info.append(["- Satu baris = satu kombinasi Tahun + Bulan + Cabang"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Loader data Target (untuk tab Scoreboard)
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
    return df[["Cabang", "PeriodeMulai", "PeriodeSelesai", "TargetService", "TargetGadget", "TargetAll", "TargetCorporate"]]


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
# Perhitungan Scoreboard
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

    return pd.concat([result, pd.DataFrame([total])], ignore_index=True)


def build_scoreboard_corporate(df_corp: pd.DataFrame, df_target: pd.DataFrame,
                                branches: list, tanggal_acuan: date) -> pd.DataFrame:
    """Versi Scoreboard untuk Omset Marketing Corporate, yang datanya bulanan (Tahun, Bulan, Omset),
    bukan harian. 'HARI INI' tidak tersedia (data tidak punya tanggal spesifik)."""
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


MONEY_COLS = [
    "OMSET SAMURAI", "OMSET HARIAN (DR TARGET)", "EXPECTED VALUE", "HARI INI", "S/D HARI INI",
    "GAP VS EXPECTED", "TOTAL GAP SAMURAI", "KEJAR TARGET PERHARI",
    "PERIODE BULAN LALU", "PERIODE BULAN INI", "GAP",
]


def style_scoreboard(df: pd.DataFrame):
    """Dipakai untuk fallback (mis. export). Untuk tampilan utama pakai render_scoreboard_html."""
    display_df = df.drop(columns=["SISA HARI"]).copy()

    def color_pct(v):
        if pd.isna(v):
            return ""
        if v >= 1:
            return "background-color:#c6efce;color:#006100"
        if v >= 0.8:
            return "background-color:#ffeb9c;color:#9c6500"
        return "background-color:#ffc7ce;color:#9c0006"

    def color_pos_good(v):
        if pd.isna(v):
            return ""
        return "background-color:#c6efce;color:#006100" if v >= 0 else "background-color:#ffc7ce;color:#9c0006"

    def color_pos_bad(v):
        if pd.isna(v):
            return ""
        return "background-color:#ffc7ce;color:#9c0006" if v > 0 else "background-color:#c6efce;color:#006100"

    sty = display_df.style
    sty = sty.map(color_pct, subset=["% PENCAPAIAN"])
    sty = sty.map(color_pos_good, subset=["GAP VS EXPECTED", "GAP"])
    sty = sty.map(color_pos_bad, subset=["TOTAL GAP SAMURAI"])
    fmt = {c: format_number for c in MONEY_COLS}
    fmt["% PENCAPAIAN"] = format_percent
    sty = sty.format(fmt, na_rep="-")
    sty = sty.set_properties(**{"text-align": "right"}, subset=MONEY_COLS + ["% PENCAPAIAN"])
    return sty


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
        if v >= 0.8:
            return "background-color:#ffeb9c;color:#9c6500"
        return "background-color:#ffc7ce;color:#9c0006"
    if col in ("GAP VS EXPECTED", "GAP"):
        return "background-color:#c6efce;color:#006100" if v >= 0 else "background-color:#ffc7ce;color:#9c0006"
    if col == "TOTAL GAP SAMURAI":
        return "background-color:#ffc7ce;color:#9c0006" if v > 0 else "background-color:#c6efce;color:#006100"
    return ""


def render_scoreboard_html(title: str, df: pd.DataFrame, accent: str = "#1f2937") -> str:
    cols = ["CABANG"] + [c for _, group_cols, _ in _SCOREBOARD_GROUPS for c in group_cols]

    header_group_html = f'<th rowspan="2" style="background:{accent};color:white;padding:8px;text-align:left;">CABANG</th>'
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
    for _, row in df.iterrows():
        is_total = row["CABANG"] == "SMM"
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


# --------------------------------------------------------------------------------------
# Sidebar: upload data
# --------------------------------------------------------------------------------------

st.sidebar.title("📥 Data")

st.sidebar.caption(
    "Upload file data per cabang (masing-masing file = data 1 cabang, kolom minimal: "
    "TGL FAKTUR, KATEGORI BARANG, TOTAL HARGA). Bisa pilih banyak file sekaligus, sampai 50 file."
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

corp_upload = st.sidebar.file_uploader("Upload file Data Marketing Corporate", type=["xlsx"], key="corp_upload")
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

st.sidebar.markdown("---")

target_upload = st.sidebar.file_uploader("Upload file Data Target (untuk tab Scoreboard)", type=["xlsx"], key="target_upload")
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

st.sidebar.caption(
    "File yang sudah diupload akan otomatis dipakai lagi setiap kali dashboard dibuka, "
    "sampai diganti/dihapus. Upload ulang dengan nama file sama = update (bukan duplikat)."
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

df_target = pd.DataFrame(columns=["Cabang", "PeriodeMulai", "PeriodeSelesai", "TargetService", "TargetGadget", "TargetAll"])
if os.path.exists(TARGET_DATA_PATH):
    with open(TARGET_DATA_PATH, "rb") as f:
        target_bytes = f.read()
    try:
        df_target = load_target_data(target_bytes)
    except ValueError as e:
        st.warning(f"Data Target tidak dipakai: {e}")

# --------------------------------------------------------------------------------------
# Filter (dipakai di tab Ringkasan)
# --------------------------------------------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.title("🔎 Filter Ringkasan")

tahun_options = sorted(set(df_main["Tahun"]).union(set(df_corp["Tahun"])), reverse=True)
bulan_options = list(range(1, 13))
cabang_options = sorted(set(df_main["Cabang"]).union(set(df_corp["Cabang"])).union(set(df_target["Cabang"])))

sel_tahun = st.sidebar.multiselect("Tahun", tahun_options, default=tahun_options)
sel_bulan = st.sidebar.multiselect("Bulan", bulan_options, default=bulan_options, format_func=lambda m: BULAN_ID[m - 1])
sel_cabang = st.sidebar.multiselect("Cabang", cabang_options, default=cabang_options)

st.sidebar.markdown("---")
st.sidebar.title("🎯 Tanggal Acuan (Scoreboard)")
tanggal_acuan = st.sidebar.date_input("Dianggap sebagai 'Hari Ini'", value=date.today())

if not sel_tahun or not sel_bulan or not sel_cabang:
    st.warning("Pilih minimal satu Tahun, Bulan, dan Cabang di sidebar.")
    st.stop()

f_main = df_main[df_main["Tahun"].isin(sel_tahun) & df_main["Bulan"].isin(sel_bulan) & df_main["Cabang"].isin(sel_cabang)]
f_corp = df_corp[df_corp["Tahun"].isin(sel_tahun) & df_corp["Bulan"].isin(sel_bulan) & df_corp["Cabang"].isin(sel_cabang)]

tab_ringkasan, tab_scoreboard = st.tabs(["📈 Ringkasan", "🏆 Scoreboard"])

# --------------------------------------------------------------------------------------
# TAB 1: Ringkasan
# --------------------------------------------------------------------------------------

def render_kpi_card(label: str, value: float, color1: str, color2: str, icon: str) -> str:
    return f"""
    <div style="background:linear-gradient(135deg,{color1},{color2});border-radius:14px;
                padding:16px 18px;color:white;box-shadow:0 2px 8px rgba(0,0,0,.12);height:100%;">
      <div style="font-size:26px;line-height:1;">{icon}</div>
      <div style="font-size:13px;opacity:.9;margin-top:8px;">{label}</div>
      <div style="font-size:22px;font-weight:700;margin-top:2px;">{format_rupiah(value)}</div>
    </div>
    """


with tab_ringkasan:
    omset_all = f_main["Omset"].sum()
    omset_service = f_main.loc[f_main["Kelompok"] == "Service", "Omset"].sum()
    omset_gadget = f_main.loc[f_main["Kelompok"] == "Gadget & Aksesoris", "Omset"].sum()
    omset_corporate = f_corp["Omset"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_kpi_card("Omset All", omset_all, "#111827", "#374151", "📊"), unsafe_allow_html=True)
    c2.markdown(render_kpi_card("Omset Service", omset_service, "#0f766e", "#14b8a6", "🛠️"), unsafe_allow_html=True)
    c3.markdown(render_kpi_card("Omset Gadget & Aksesoris", omset_gadget, "#6d28d9", "#a78bfa", "📱"), unsafe_allow_html=True)
    c4.markdown(render_kpi_card("Omset Marketing Corporate", omset_corporate, "#b45309", "#f59e0b", "🤝"), unsafe_allow_html=True)

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.caption(
        "Lihat persentase pencapaian target per cabang (dengan warna hijau/kuning/merah) di tab **🏆 Scoreboard**."
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

    bcol1, bcol2 = st.columns([2, 1])
    with bcol1:
        branch_long = by_branch.melt(id_vars="Cabang", var_name="Kategori", value_name="Omset")
        fig_branch = px.bar(branch_long, x="Cabang", y="Omset", color="Kategori", barmode="group")
        fig_branch.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Omset (Rp)")
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
        "Omset Marketing Corporate dihitung dari file terpisah yang diupload (lihat sidebar)."
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

    if df_target.empty:
        st.info(
            "Belum ada data Target. Download template Target di sidebar, isi Cabang / Periode Mulai / "
            "Periode Selesai / Target Service / Target Gadget & Aksesoris / Target Marketing Corporate, "
            "lalu upload. Tanpa data ini, kolom Target / Expected Value / % Pencapaian / Gap / Kejar Target "
            "tidak bisa dihitung — tabel di bawah tetap tampil dengan kolom lain saja."
        )

    branches_sb = sel_cabang if sel_cabang else cabang_options

    kategori_map = [
        ("SCOREBOARD OMSET ALL", df_main, "TargetAll", "#111827"),
        ("SCOREBOARD OMSET SERVICE", df_main[df_main["Kelompok"] == "Service"], "TargetService", "#0f766e"),
        ("SCOREBOARD OMSET GADGET & AKSESORIS", df_main[df_main["Kelompok"] == "Gadget & Aksesoris"], "TargetGadget", "#6d28d9"),
    ]

    for title, subset, target_col, accent in kategori_map:
        board = build_scoreboard(subset, df_target, target_col, branches_sb, tanggal_acuan)
        st.markdown(render_scoreboard_html(title, board, accent), unsafe_allow_html=True)

    st.markdown(
        render_scoreboard_html(
            "SCOREBOARD OMSET MARKETING CORPORATE",
            build_scoreboard_corporate(df_corp, df_target, branches_sb, tanggal_acuan),
            "#b45309",
        ),
        unsafe_allow_html=True,
    )
    st.caption(
        "Data Marketing Corporate bersifat bulanan (bukan harian), jadi kolom **HARI INI** tidak tersedia "
        "(ditampilkan \"-\")."
    )

    st.markdown(
        "**Cara baca warna:** hijau = sudah di atas ekspektasi / target tercapai, "
        "kuning = mendekati (80-100% dari ekspektasi), merah = di bawah ekspektasi / masih ada gap."
    )
