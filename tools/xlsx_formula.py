import re
def c2n(c):
    n=0
    for ch in c: n=n*26+ord(ch)-64
    return n
def n2c(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
REF=r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
TOK=re.compile(r'(?P<str>"(?:[^"]|"")*")|(?P<pre>(?:\'(?:[^\']|\'\')+\'|[A-Za-z_][\w\.]*)!)?(?<![A-Za-z0-9_.\$])(?P<r1>\$?[A-Z]{1,3}\$?\d+)(?::(?P<r2>\$?[A-Z]{1,3}\$?\d+))?(?![\w(])')
REFRE=re.compile('^'+REF+'$')
def map_refs(f, fn):
    """fn(prefix_or_None, (cabs,col,rabs,row)) -> new ref string"""
    def rep(m):
        if m.group('str'): return m.group(0)
        pre=m.group('pre')
        out=(pre or '')+fn(pre, REFRE.match(m.group('r1')).groups())
        if m.group('r2'): out+=':'+fn(pre, REFRE.match(m.group('r2')).groups())
        return out
    return TOK.sub(rep,f)
def mk(ca,col,ra,row): return f"{ca}{col}{ra}{row}"
def translate(f, dr, dc):
    def fn(pre,g):
        ca,col,ra,row=g
        if not ca: col=n2c(c2n(col)+dc)
        if not ra: row=str(int(row)+dr)
        return mk(ca,col,ra,row)
    return map_refs(f,fn)
def shift_formula(f, k, own, target="'Detail Data Penjualan'!"):
    """insert column at k; shift refs to target sheet (own=True: unprefixed refs are target)"""
    def fn(pre,g):
        ca,col,ra,row=g
        if (pre is None and own) or (pre==target):
            if c2n(col)>=k: col=n2c(c2n(col)+1)
        return mk(ca,col,ra,row)
    return map_refs(f,fn)
def shift_cell(ref,k):
    m=re.match(r'(\$?)([A-Z]+)(\$?)(\d+)$',ref)
    ca,col,ra,row=m.groups()
    if c2n(col)>=k: col=n2c(c2n(col)+1)
    return f"{ca}{col}{ra}{row}"
def shift_sqref(s,k):
    return ' '.join(':'.join(shift_cell(p,k) for p in part.split(':')) for part in s.split())
import html
def unesc(s): return html.unescape(s)
def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
CELL=re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)',re.S)
FRE=re.compile(r'<f([^>]*?)(?:/>|>(.*?)</f>)',re.S)
def unshare(x, only=None):
    """expand shared formulas into explicit ones. only: predicate on master text"""
    masters={}
    for m in CELL.finditer(x):
        inner=m.group(5) or ''
        fm=FRE.search(inner)
        if fm and 't="shared"' in fm.group(1) and fm.group(2):
            si=re.search(r'si="(\d+)"',fm.group(1)).group(1)
            masters[si]=(c2n(m.group(1)),int(m.group(2)),unesc(fm.group(2)))
    def rep(m):
        inner=m.group(5)
        if not inner: return m.group(0)
        fm=FRE.search(inner)
        if not fm or 't="shared"' not in fm.group(1): return m.group(0)
        si=re.search(r'si="(\d+)"',fm.group(1)).group(1)
        mc,mr,mt=masters[si]
        if only and not only(mt): return m.group(0)
        f=translate(mt,int(m.group(2))-mr,c2n(m.group(1))-mc)
        attrs=re.sub(r'\s*(t="shared"|ref="[^"]*"|si="\d+")','',fm.group(1))
        newf=f'<f{attrs}>{esc(f)}</f>'
        return f'<c r="{m.group(1)}{m.group(2)}"{m.group(3)}>'+inner[:fm.start()]+newf+inner[fm.end():]+'</c>'
    return CELL.sub(rep,x)
def cells(x):
    d={}
    for m in CELL.finditer(x):
        inner=m.group(5) or ''
        fm=FRE.search(inner); v=re.search(r'<v>(.*?)</v>',inner)
        d[m.group(1)+m.group(2)]=(unesc(fm.group(2)) if fm and fm.group(2) else None, v.group(1) if v else None, m.group(3))
    return d
