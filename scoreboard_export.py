"""Gambar Scoreboard (format sama dengan sheet 'Scoreboard' di Dashboard 6 Pilar)
untuk dibagikan ke semua cabang. Data diambil dari nilai hasil hitung (cache)
sheet Scoreboard file master di data/main, lalu digambar ulang dengan Pillow:
kotak TANGGAL/SISA HARI, judul, logo, header berwarna, warna % pencapaian, baris
total kuning - meniru tampilan Excel yang biasa di-capture manual.
"""
import io
import os
import datetime as _dt

from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------- warna
NAVY = (68, 84, 106)          # header kolom target / kotak tanggal
GREEN = (112, 173, 71)        # detail pencapaian
ORANGE = (237, 125, 49)       # detail gap
BLUE = (68, 114, 196)         # monitoring
WHITE = (255, 255, 255)
TEXT = (38, 38, 38)
GRID = (205, 205, 205)
YELLOW_TOTAL = (255, 230, 153)
OK_BG, OK_FG = (198, 239, 206), (0, 97, 0)
MID_BG, MID_FG = (255, 235, 156), (156, 87, 0)
BAD_BG, BAD_FG = (255, 199, 206), (156, 0, 6)

GROUP_DEFS = [  # (teks default, warna teks, warna header kolom)
    ("DETAIL TARGET", NAVY, NAVY),
    ("DETAIL PENCAPAIAN", GREEN, GREEN),
    ("DETAIL GAP", ORANGE, ORANGE),
    ("MONITORING PROGRESS KENAIKAN OMSET RATA2 PERHARI", BLUE, BLUE),
]

SCALE = 2  # gambar dibuat 2x supaya tajam saat di-zoom di HP


def _font(bold=False, size=11):
    import matplotlib
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    path = os.path.join(matplotlib.get_data_path(), "fonts", "ttf", name)
    return ImageFont.truetype(path, int(size * SCALE))


# ---------------------------------------------------------------- baca Excel
def _s(v):
    return str(v).strip() if v is not None else ""


def read_scoreboard_sections(path: str):
    """Kembalikan list section: {title, tanggal, sisa, groups, headers, cols, rows, kind}."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Scoreboard"]
    grid = {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=400, max_col=14, values_only=True), 1):
        grid[i] = list(row) + [None] * (14 - len(row))
    wb.close()
    col = lambda r, c: grid.get(r, [None] * 14)[ord(c) - 65]
    out = []
    for t in sorted(grid):
        title = next((_s(v) for v in grid[t] if isinstance(v, str) and "SCOREBOARD" in v.upper()), None)
        if not title or "RETAIL" in title.upper():
            continue
        hdr = next((r for r in range(t + 1, t + 8) if _s(col(r, "B")).upper() == "CABANG"), None)
        if hdr is None:
            continue
        cols = [c for c in "BCDEFGHIJKLMN" if c == "B" or _s(col(hdr, c))]
        # header kolom yang di-merge (mis. S/D HARI INI di F:G) -> sel G kosong -> dilewati
        rows = []
        r = hdr + 1
        while r < hdr + 60 and r in grid:
            name = _s(col(r, "B"))
            if name and name.upper() not in ("N/A", "-"):
                rows.append([col(r, c) for c in cols])
            if name.upper() in ("SMM", "HEAD OF CORPORATE"):
                break
            r += 1
        if not rows:
            continue
        tanggal = col(t, "C")
        sisa = col(t + 1, "C")
        groups = []
        for c in cols[1:]:
            g = _s(col(hdr - 1, c))
            groups.append(g)
        out.append({
            "title": title.upper(),
            "tanggal": tanggal,
            "sisa": sisa,
            "headers": [_s(col(hdr, c)) for c in cols],
            "groups": groups,
            "rows": rows,
        })
    return out


# ---------------------------------------------------------------- format angka
def _fmt_num(v):
    if v is None or v == "":
        return ""
    if isinstance(v, str):
        return v
    try:
        v = round(v)
        if v == 0:
            return "0"
        return f"{v:,.0f}".replace(",", ".")
    except Exception:
        return str(v)


def _fmt_pct(v):
    if v is None or v == "" or isinstance(v, str):
        return "" if v is None else str(v)
    return f"{v * 100:.2f}%".replace(".", ",")


def _wrap(draw, text, font, width):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if draw.textlength(test, font=font) <= width or not cur:
            cur = test
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


# ---------------------------------------------------------------- gambar 1 tabel
def _col_role(header: str):
    h = header.upper()
    if "%" in h:
        return "pct"
    if h == "GAP":
        return "gap"
    if "S/D HARI NI" == h:
        return "behind"
    return "num"


def _group_of(idx, headers):
    """Index grup (0 target, 1 pencapaian, 2 gap, 3 monitoring) per kolom data."""
    h = headers[idx].upper()
    if h in ("OMSET SAMURAI", "EXPECTED VALUE") or "OMSET HARIAN" in h:
        return 0
    if h in ("HARI INI", "S/D HARI INI") or "%" in h:
        return 1
    if h in ("S/D HARI NI", "TOTAL GAP SAMURAI", "KEJAR TARGET PERHARI"):
        return 2
    return 3


def render_section(sec, logo: Image.Image = None) -> Image.Image:
    S = SCALE
    headers = sec["headers"]
    n = len(headers)
    w_first, w_col = 150, 100
    widths = [w_first] + [w_col] * (n - 1)
    total_w = sum(widths)
    pad = 10
    W = (total_w + 2 * pad) * S
    top_h, gap_h, grp_h, hdr_h, row_h = 52, 14, 34, 54, 24
    H = (top_h + gap_h + grp_h + hdr_h + row_h * len(sec["rows"]) + pad * 2) * S
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    f_small_b = _font(True, 9)
    f_val = _font(False, 9.5)
    f_val_b = _font(True, 9.5)
    f_title = _font(True, 17)
    f_hdr = _font(True, 8.5)
    f_grp = _font(True, 10)

    x0, y = pad * S, pad * S

    # --- kotak tanggal / sisa hari + judul + logo
    box_w = w_first * S
    d.rectangle([x0, y, x0 + box_w, y + top_h * S], fill=NAVY)
    tgl = sec["tanggal"]
    if isinstance(tgl, (_dt.datetime, _dt.date)):
        tgl = tgl.strftime("%d/%m/%Y")
    sisa = sec["sisa"]
    sisa = f"{int(sisa)}" if isinstance(sisa, (int, float)) else _s(sisa)
    for k, (lab, val) in enumerate((("TANGGAL", tgl), ("SISA HARI", sisa))):
        yy = y + (8 + k * 20) * S
        d.text((x0 + box_w / 2, yy + 6 * S), lab, font=f_small_b, fill=WHITE, anchor="mm")
        d.text((x0 + box_w + 18 * S, yy + 6 * S), _s(val), font=f_small_b, fill=NAVY, anchor="lm")
    d.text((x0 + total_w * S / 2 + 40 * S, y + top_h * S / 2), sec["title"], font=f_title, fill=NAVY, anchor="mm")
    if logo is not None:
        lh = int(44 * S)
        lw = int(logo.width * lh / logo.height)
        lg = logo.resize((lw, lh))
        img.paste(lg, (x0 + total_w * S - lw - 6 * S, y + int((top_h * S - lh) / 2)), lg if lg.mode == "RGBA" else None)
    y += (top_h + gap_h) * S

    # --- baris grup (DETAIL TARGET / PENCAPAIAN / GAP / MONITORING)
    xs = [x0]
    for wv in widths:
        xs.append(xs[-1] + wv * S)
    groups = [_group_of(i, headers) for i in range(1, n)]
    i = 1
    while i < n:
        g = groups[i - 1]
        j = i
        while j < n and groups[j - 1] == g:
            j += 1
        label = next((sec["groups"][k - 1] for k in range(i, j) if sec["groups"][k - 1]), GROUP_DEFS[g][0])
        x_a, x_b = xs[i], xs[j]
        d.rectangle([x_a, y, x_b, y + grp_h * S], fill=WHITE, outline=GRID, width=S)
        lines = _wrap(d, label, f_grp, (x_b - x_a) - 8 * S)
        lh_ = 13 * S
        y_c = y + grp_h * S / 2 - (len(lines) - 1) * lh_ / 2
        for k, ln in enumerate(lines):
            d.text(((x_a + x_b) / 2, y_c + k * lh_), ln, font=f_grp, fill=GROUP_DEFS[g][1], anchor="mm")
        i = j
    y += grp_h * S

    # --- header kolom
    for c in range(n):
        bg = NAVY if c == 0 else GROUP_DEFS[groups[c - 1]][2]
        d.rectangle([xs[c], y, xs[c + 1], y + hdr_h * S], fill=bg, outline=WHITE, width=S)
        lines = _wrap(d, headers[c].upper(), f_hdr, (xs[c + 1] - xs[c]) - 8 * S)
        lh_ = 11 * S
        y_c = y + hdr_h * S / 2 - (len(lines) - 1) * lh_ / 2
        for k, ln in enumerate(lines):
            d.text(((xs[c] + xs[c + 1]) / 2, y_c + k * lh_), ln, font=f_hdr, fill=WHITE, anchor="mm")
    y += hdr_h * S

    # --- baris data
    roles = [None] + [_col_role(h) for h in headers[1:]]
    for ri, row in enumerate(sec["rows"]):
        is_total = _s(row[0]).upper() in ("SMM", "HEAD OF CORPORATE")
        for c in range(n):
            v = row[c]
            bg, fg = (YELLOW_TOTAL if is_total else WHITE), TEXT
            if c == 0:
                txt = _s(v)
            elif roles[c] == "pct":
                txt = _fmt_pct(v)
                if isinstance(v, (int, float)):
                    if v >= 1:
                        bg, fg = OK_BG, OK_FG
                    elif v >= 0.85:
                        bg, fg = MID_BG, MID_FG
                    else:
                        bg, fg = BAD_BG, BAD_FG
            else:
                txt = _fmt_num(v)
                if roles[c] == "behind" and isinstance(v, (int, float)) and not is_total:
                    bg, fg = (BAD_BG, BAD_FG) if v > 0 else (OK_BG, OK_FG)
                elif roles[c] == "gap" and isinstance(v, (int, float)):
                    if v > 0:
                        bg, fg = OK_BG, OK_FG
                    elif v < 0:
                        bg, fg = BAD_BG, BAD_FG
                    else:
                        bg, fg = MID_BG, MID_FG
            d.rectangle([xs[c], y, xs[c + 1], y + row_h * S], fill=bg, outline=GRID, width=1)
            font = f_val_b if (is_total or c == 0) else f_val
            d.text(((xs[c] + xs[c + 1]) / 2, y + row_h * S / 2), txt, font=font, fill=fg, anchor="mm")
        y += row_h * S
    return img


def _stack(images, gap=30):
    if not images:
        img = Image.new("RGB", (800, 200), WHITE)
        ImageDraw.Draw(img).text((400, 100), "Tidak ada data Scoreboard", fill=TEXT, anchor="mm", font=_font(True, 14))
        return img
    W = max(i.width for i in images)
    H = sum(i.height for i in images) + gap * SCALE * (len(images) - 1) + 14 * SCALE
    out = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(out)
    d.rectangle([0, 0, W, 10 * SCALE], fill=NAVY)  # pita atas seperti di Excel
    y = 14 * SCALE
    for im in images:
        out.paste(im, (0, y))
        y += im.height + gap * SCALE
    return out


GROUPS = {
    "Omset (All, Service, Gadget)": lambda t: any(k in t for k in ("OMSET ALL", "OMSET SERVICE", "GADGET")),
    "Marketing Corporate (Omset, GP, Pengadaan)": lambda t: "MARKETING CORPORATE" in t,
    "Pilar (Pengadaan, Maintenance, Sewa, ISP)": lambda t: ("MARKETING CORPORATE" not in t) and any(
        k in t for k in ("PENGADAAN CORPORATE", "MAINTENANCE", "SEWA", "ISP")),
}


def build_images(path: str, logo_path: str = None):
    """Return dict {nama grup: Image} dari file master Excel."""
    secs = read_scoreboard_sections(path)
    logo = None
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image.open(logo_path).convert("RGBA")
            bb = logo.getchannel("A").getbbox()
            if bb:
                logo = logo.crop(bb)
        except Exception:
            logo = None
    out = {}
    for name, pred in GROUPS.items():
        imgs = [render_section(s, logo) for s in secs if pred(s["title"])]
        if imgs:
            out[name] = _stack(imgs)
    return out


def to_jpg(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return buf.getvalue()


def to_pdf(images) -> bytes:
    images = [im.convert("RGB") for im in images]
    buf = io.BytesIO()
    if images:
        images[0].save(buf, format="PDF", save_all=True, append_images=images[1:], resolution=200)
    return buf.getvalue()
