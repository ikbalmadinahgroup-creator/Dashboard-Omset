"""Scoreboard Marketing Corporate (omset & gross profit) per nama marketing.
Pencapaian dihitung langsung dari sheet Faktur Penjualan (SUMIFS kolom AP = nama
penjual di Accurate) - tidak bergantung pivot Mark.Corporate (pivot sering tidak
ter-refresh sehingga GETPIVOTDATA = 0). Nama di kolom B Scoreboard (nama panggilan)
dipetakan ke nama-nama di Accurate lewat MC_ALIASES."""
import re, html
from xlsx_formula import unshare, CELL, translate, esc, unesc, c2n

MC_ALIASES = {
    'WAHYU JP': ['WAHYU JP', 'WAHYU JP (RADJIMAN)', 'WAHYU JP JATIWARINGIN'],
    'KOUTSAREZRA KANZA': ['KOUTSAREZRA KANZA'],
    'SUPRIYADI': ['SUPRIYADI'],
    'DICKY': ['DICKY YUNIAWAN', 'DICKY'],
    'FAISAL': ['FAISAL ABDUL RAHMAN'],
    'TEGAR': ['TEGAR PUTRA YANSA', 'TEGAR PUTRA YANSYAH', 'TEGAR SALES'],
    'RIFQI': ['RIFQI ADITYA'],
    'SOLEHUDIN': ['SOLEHUDIN'],
    'ROLAN': ['PAOLO MAROLANZANO'],
    'ZAHRAN': ['M FARHAN ZAHRAN'],
    'MUIS': ['NUR MUIS'],
    'IQBAL SABARI': ['IQBAL SABARI', 'IQBAL SABARI SALES', 'IKBAL SABARI'],
    'M SYAFAAT': ['M SYAFAAT', 'MUHAMMAD SYAFAAT'],
    'RAID IMADUDIN FIRAS': ['RAID IMADUDIN FIRAS'],
    'KAUKABAN AL AKWAN': ['KAUKABAN AL AKWAN'],
}
FP = "'Faktur Penjualan'!"
RNG = {c: f"{FP}${c}$2:${c}$95212" for c in ('AN', 'AP', 'AT', 'B', 'AR')}
SKIP = {'', 'N/A', '-'}

def aliases(label):
    l = (label or '').strip().upper()
    if l in SKIP: return None
    return MC_ALIASES.get(l, [l])

def _sst(sst):
    return [html.unescape(re.sub('<[^>]+>', '', s)) for s in re.findall(r'<si>(.*?)</si>', sst, re.S)]

def fill(x, sst, set_cells):
    """x = XML sheet Scoreboard. Return (x_baru, ringkasan list)."""
    x = unshare(x)
    ss = _sst(sst)
    cells = {}
    for m in CELL.finditer(x):
        ref = m.group(1) + m.group(2); a = m.group(3); inner = m.group(5) or ''
        v = re.search(r'<v>(.*?)</v>', inner); f = re.search(r'<f[^>]*>(.*?)</f>', inner, re.S)
        txt = None
        if v is not None:
            txt = ss[int(v.group(1))] if 't="s"' in a else html.unescape(v.group(1))
        else:
            it = re.search(r'<t[^>]*>(.*?)</t>', inner, re.S)
            if it: txt = html.unescape(it.group(1))
        cells[ref] = (re.sub(r'\s*t="[^"]*"', '', a), txt, html.unescape(f.group(1)) if f else None)
    def find_title(word, excl=None):
        for ref, (_, t, _) in cells.items():
            if t and word in t.upper() and 'MARKETING CORPORATE' in t.upper() and (excl is None or excl not in t.upper()):
                return int(re.sub('[A-Z]', '', ref))
        return None
    out = {}; info = []
    t_omset = find_title('SCOREBOARD OMSET', 'PENGADAAN')
    for kind, word, col, excl in (('omset', 'SCOREBOARD OMSET', 'AN', 'PENGADAAN'), ('gp', 'GROSS PROFIT', 'AT', None),
                                  ('pengadaan', 'PENGADAAN', 'AN', None)):
        t = find_title(word, excl)
        if t is None: continue
        r = t + 5; rows = []
        while r < t + 40:
            lab = cells.get(f'B{r}', ('', None, None))[1]
            if lab and str(lab).strip().upper() == 'HEAD OF CORPORATE': break
            rows.append(r); r += 1
        tot = r
        if not rows or tot >= t + 40: continue
        for r in rows:
            st = lambda c: cells.get(f'{c}{r}', ('', None, None))[0]
            lab = cells.get(f'B{r}', ('', None, None))[1]
            if kind == 'gp':
                # label GP diambil dari baris omset yang dirujuk target C (mis. C88*13%)
                cf = cells.get(f'C{r}', ('', None, None))[2] or ''
                mm = re.match(r'\$?C\$?(\d+)\*', cf)
                if mm:
                    src = int(mm.group(1)); lab = cells.get(f'B{src}', ('', None, None))[1]
                    out[f'B{r}'] = f'<c r="B{r}"{st("B")} t="str"><f>B{src}</f><v>{html.escape(str(lab or ""), quote=False)}</v></c>'
            if kind == 'pengadaan' and t_omset:
                src = r - (t - t_omset); lab = cells.get(f'B{src}', ('', None, None))[1]
                out[f'B{r}'] = f'<c r="B{r}"{st("B")} t="str"><f>B{src}</f><v>{html.escape(str(lab or ""), quote=False)}</v></c>'
            al = aliases(lab)
            arr = '{' + ','.join('"' + a.replace('"', '""') + '"' for a in al) + '}' if al else None
            def F(fx): return html.escape(fx, quote=False)
            def put(c, fx): out[f'{c}{r}'] = f'<c r="{c}{r}"{st(c)}><f>{F(fx)}</f></c>'
            extra = f',{RNG["AR"]},"*PENGADAAN*"' if kind == 'pengadaan' else ''
            if arr:
                put('F', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}{extra}))')
                prev = f',{RNG["B"]},">="&(EOMONTH($C${t},-2)+1),{RNG["B"]},"<="&EOMONTH($C${t},-1)'
                cur = f',{RNG["B"]},">="&(EOMONTH($C${t},-1)+1),{RNG["B"]},"<="&$C${t}'
                put('L', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}{extra}{prev}))/DAY(EOMONTH($C${t},-1))')
                put('M', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}{extra}{cur}))/DAY($C${t})')
            else:
                for c in 'FLM': out[f'{c}{r}'] = f'<c r="{c}{r}"{st(c)}><v>0</v></c>'
            put('H', f'IFERROR(F{r}/E{r},0)'); put('I', f'E{r}-F{r}'); put('J', f'C{r}-F{r}'); put('N', f'M{r}-L{r}')
            info.append((kind, r, lab, al))
        r0, r1 = rows[0], rows[-1]
        stt = lambda c: cells.get(f'{c}{tot}', ('', None, None))[0]
        def putt(c, fx): out[f'{c}{tot}'] = f'<c r="{c}{tot}"{stt(c)}><f>{html.escape(fx, quote=False)}</f></c>'
        putt('F', f'SUM(F{r0}:F{r1})'); putt('L', f'SUM(L{r0}:L{r1})'); putt('M', f'SUM(M{r0}:M{r1})')
        putt('H', f'IFERROR(F{tot}/E{tot},0)'); putt('I', f'E{tot}-F{tot}'); putt('J', f'C{tot}-F{tot}'); putt('N', f'M{tot}-L{tot}')
    if out: x = set_cells(x, out)
    return x, info


PENG_TITLE = 'SCOREBOARD OMSET PENGADAAN MARKETING CORPORATE'

def ensure_pengadaan_block(x, sst, drawing=None):
    """Tambahkan tabel 'SCOREBOARD OMSET PENGADAAN MARKETING CORPORATE' (salinan format
    tabel Omset Marketing Corporate) di bawah tabel terakhir, kalau belum ada.
    Return (x, drawing, added)."""
    ss = _sst(sst)
    def txt(inner, a):
        v = re.search(r'<v>(.*?)</v>', inner or '')
        if v is None:
            t = re.search(r'<t[^>]*>(.*?)</t>', inner or '', re.S)
            return html.unescape(t.group(1)) if t else None
        return ss[int(v.group(1))] if 't="s"' in a else html.unescape(v.group(1))
    t_om = None
    for m in CELL.finditer(x):
        t = txt(m.group(5), m.group(3))
        if t and 'MARKETING CORPORATE' in t.upper() and 'PENGADAAN' in t.upper():
            return x, drawing, False
        if t and 'SCOREBOARD OMSET' in t.upper() and 'MARKETING CORPORATE' in t.upper() and t_om is None:
            t_om = int(m.group(2))
    if t_om is None:
        return x, drawing, False
    x = unshare(x)
    # baris total (HEAD OF CORPORATE)
    tot = None
    for m in CELL.finditer(x):
        if m.group(1) == 'B' and int(m.group(2)) > t_om:
            t = txt(m.group(5), m.group(3))
            if t and t.strip().upper() == 'HEAD OF CORPORATE':
                tot = int(m.group(2)); break
    last = max(int(r) for r in re.findall(r'<row r="(\d+)"', x))
    t_new = last + 4; d = t_new - t_om
    rows = {int(m.group(1)): m for m in re.finditer(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', x, re.S)}
    new_rows = []
    for r in range(t_om, tot + 1):
        if r not in rows: continue
        m = rows[r]; attrs = re.sub(r'\s*spans="[^"]*"', '', m.group(2)); body = m.group(3) or ''
        cells = []
        for c in CELL.finditer(body):
            col = c.group(1); a = c.group(3); inner = c.group(5) or ''
            if c2n(col) > 14: continue
            ref = f'{col}{r + d}'
            fm = re.search(r'<f[^>]*>(.*?)</f>', inner, re.S)
            if r == t_om and txt(inner, a) and 'MARKETING CORPORATE' in txt(inner, a).upper():
                a2 = re.sub(r'\s*t="[^"]*"', '', a)
                cells.append(f'<c r="{ref}"{a2} t="inlineStr"><is><t>{PENG_TITLE}</t></is></c>'); continue
            if col == 'C' and r == t_om:
                a2 = re.sub(r'\s*t="[^"]*"', '', a); cells.append(f'<c r="{ref}"{a2}><f>$C$4</f></c>'); continue
            if col == 'C' and t_om + 5 <= r <= tot:   # target: kosong (diisi user), total = SUM
                a2 = re.sub(r'\s*t="[^"]*"', '', a)
                if r == tot: cells.append(f'<c r="{ref}"{a2}><f>SUM(C{t_om+5+d}:C{tot-1+d})</f></c>')
                else: cells.append(f'<c r="{ref}"{a2}/>')
                continue
            if fm:
                a2 = re.sub(r'\s*t="[^"]*"', '', a)
                cells.append(f'<c r="{ref}"{a2}><f>{esc(translate(unesc(fm.group(1)), d, 0))}</f></c>')
            else:
                cells.append(re.sub(r'^<c r="[A-Z]+\d+"', f'<c r="{ref}"', c.group(0)))
        new_rows.append(f'<row r="{r + d}"{attrs}>' + ''.join(cells) + '</row>')
    x = x.replace('</sheetData>', ''.join(new_rows) + '</sheetData>', 1)
    # merge & conditional formatting
    def shift_ref(ref): return re.sub(r'(\d+)', lambda mm: str(int(mm.group(1)) + d), ref)
    mg = [mm for mm in re.findall(r'<mergeCell ref="([^"]+)"/>', x)
          if all(t_om <= int(n) <= tot for n in re.findall(r'\d+', mm))]
    if mg:
        add = ''.join(f'<mergeCell ref="{shift_ref(mm)}"/>' for mm in mg)
        x = re.sub(r'<mergeCells count="(\d+)">', lambda mm: f'<mergeCells count="{int(mm.group(1)) + len(mg)}">', x, 1)
        x = x.replace('</mergeCells>', add + '</mergeCells>', 1)
    cfs = [mm for mm in re.finditer(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', x, re.S)
           if all(t_om <= int(n) <= tot for n in re.findall(r'\d+', mm.group(1)))]
    pr = [5000]
    def repri(sx):
        def f(mm): pr[0] += 1; return f'priority="{pr[0]}"'
        return re.sub(r'priority="\d+"', f, sx)
    add = ''.join(f'<conditionalFormatting sqref="{shift_ref(mm.group(1))}">{repri(mm.group(2))}</conditionalFormatting>' for mm in cfs)
    if add:
        j = x.rindex('</conditionalFormatting>') + len('</conditionalFormatting>'); x = x[:j] + add + x[j:]
    x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"/>', lambda mm: f'<dimension ref="{mm.group(1)}:{mm.group(2)}{tot + d}"/>', x, 1)
    # logo di baris judul (salin gambar baris judul omset MC)
    if drawing:
        anc = [a for a in re.findall(r'<xdr:oneCellAnchor>.*?</xdr:oneCellAnchor>', drawing, re.S)
               if f'<xdr:row>{t_om - 1}</xdr:row>' in a]
        if anc:
            mx = max(int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', drawing))
            a = anc[0].replace(f'<xdr:row>{t_om - 1}</xdr:row>', f'<xdr:row>{t_new - 1}</xdr:row>', 1)
            a = re.sub(r'<xdr:cNvPr id="\d+" name="[^"]*"', f'<xdr:cNvPr id="{mx + 1}" name="Picture {mx + 1}"', a)
            a = re.sub(r'<a16:creationId[^>]*/>', '', a)
            drawing = drawing.replace('</xdr:wsDr>', a + '</xdr:wsDr>')
    return x, drawing, True
