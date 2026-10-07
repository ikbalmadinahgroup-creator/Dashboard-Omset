"""Scoreboard Marketing Corporate (omset & gross profit) per nama marketing.
Pencapaian dihitung langsung dari sheet Faktur Penjualan (SUMIFS kolom AP = nama
penjual di Accurate) - tidak bergantung pivot Mark.Corporate (pivot sering tidak
ter-refresh sehingga GETPIVOTDATA = 0). Nama di kolom B Scoreboard (nama panggilan)
dipetakan ke nama-nama di Accurate lewat MC_ALIASES."""
import re, html
from xlsx_formula import unshare, CELL

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
RNG = {c: f"{FP}${c}$2:${c}$95212" for c in ('AN', 'AP', 'AT', 'B')}
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
        cells[ref] = (re.sub(r'\s*t="[^"]*"', '', a), txt, html.unescape(f.group(1)) if f else None)
    def find_title(word):
        for ref, (_, t, _) in cells.items():
            if t and word in t.upper() and 'MARKETING CORPORATE' in t.upper():
                return int(re.sub('[A-Z]', '', ref))
        return None
    out = {}; info = []
    for kind, word, col in (('omset', 'SCOREBOARD OMSET', 'AN'), ('gp', 'GROSS PROFIT', 'AT')):
        t = find_title(word)
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
            al = aliases(lab)
            arr = '{' + ','.join('"' + a.replace('"', '""') + '"' for a in al) + '}' if al else None
            def F(fx): return html.escape(fx, quote=False)
            def put(c, fx): out[f'{c}{r}'] = f'<c r="{c}{r}"{st(c)}><f>{F(fx)}</f></c>'
            if arr:
                put('F', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}))')
                prev = f',{RNG["B"]},">="&(EOMONTH($C${t},-2)+1),{RNG["B"]},"<="&EOMONTH($C${t},-1)'
                cur = f',{RNG["B"]},">="&(EOMONTH($C${t},-1)+1),{RNG["B"]},"<="&$C${t}'
                put('L', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}{prev}))/DAY(EOMONTH($C${t},-1))')
                put('M', f'SUM(SUMIFS({RNG[col]},{RNG["AP"]},{arr}{cur}))/DAY($C${t})')
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
