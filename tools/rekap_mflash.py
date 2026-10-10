#!/usr/bin/env python3
"""Rekap faktur penjualan semua cabang MFlash ke Dashboard 6 Pilar (XML surgery, tanpa openpyxl full-load).
Usage: python3 rekap_mflash.py --dashboard DASH.xlsx --out OUT.xlsx --workdir DIR BRANCH1.xlsx BRANCH2.xlsx ...
"""
import argparse, zipfile, re, html, datetime, os, sys, time, shutil, subprocess, signal, warnings, collections
import openpyxl
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from style_tools import ensure_styles, remap_sheet
import mc_scoreboard
warnings.filterwarnings('ignore')

CABANG = {'001':'KLENDER','002':'CEGER','003':'BINTARA','004':'RADJIMAN','005':'JATIMULYA','006':'DRAMAGA',
          '007':'CONDET','008':'JATIBENING','009':'SAWANGAN','010':'WARBONG','011':'CINERE','012':'CIBINONG',
          '013':'KARAWANG','014':'JATIWARINGIN','015':'CIKAMPEK','016':'CILANGKAP','017':'PEJATEN','018':'CIBUBUR',
          '019':'CIMANGGIS'}
PILAR = ['SERVICE','PENJUALAN RITEL','PENGADAAN CORPORATE','MAINTENANCE CORPORATE','CICILAN SYARIAH','SEWA']
PILAR_SHEET = 'Omset Kategori Pilar'
STYLE_IDS = [245,87,88,89,91,92,93,94,248,249,95,218,250,251,254,255,256,258,11,16]
# Urutan kolom data di sheet Faktur Penjualan (B, D, F, ..., AR)
DASH_COLS = ['TGL FAKTUR','NO FAKTUR','KATEGORI PELANGGAN','ID PELANGGAN','NAMA CUSTOMER','NAMA ADMIN',
             'KATEGORI PENJUALAN','KATEGORI BARANG','NAMA TEKNISI','NAMA TEKNISI (FINAL)','YANG MENYERAHKAN/MENJUAL',
             'KERUSAKAN UTAMA','MERK UNIT','TIPE UNIT','KODE BARANG','NAMA BARANG','HARGA BELI','QTY','@HARGA',
             'TOTAL HARGA','NAMA DEFAULT PENJUAL PELANGGAN FAKTUR PENJUALAN','KATEGORI PILAR']
EPOCH = datetime.datetime(1899,12,30)
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'

def col(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
def cidx(c): return sum((ord(ch)-64)*26**i for i,ch in enumerate(reversed(c)))

# ---------- Rekap per Sales (Mark.Corporate) ----------
# Nama yang berbeda di Accurate tapi ORANG YANG SAMA (dikonfirmasi user 1 Okt 2026).
# Pivot & tabel per cabang tetap pakai nama asli (varian nama = penanda cabang);
# tabel ini menjumlahkan per orang.
SALES_GROUPS = [
    ('WAHYU JP', ['WAHYU JP','WAHYU JP (RADJIMAN)','WAHYU JP JATIWARINGIN']),
    ('DICKY YUNIAWAN', ['DICKY YUNIAWAN','DICKY']),
    ('FAISAL ABDUL RAHMAN', ['FAISAL ABDUL RAHMAN']),
    ('IQBAL SABARI', ['IQBAL SABARI','IQBAL SABARI SALES','IKBAL SABARI']),
    ('KOUTSAREZRA KANZA', ['KOUTSAREZRA KANZA']),
    ('M SYAFAAT', ['M SYAFAAT','MUHAMMAD SYAFAAT']),
    ('RAID IMADUDIN FIRAS', ['RAID IMADUDIN FIRAS']),
    ('TEGAR PUTRA YANSA', ['TEGAR PUTRA YANSA','TEGAR PUTRA YANSYAH','TEGAR SALES']),
    ('SUPRIYADI', ['SUPRIYADI']),
    ('PAOLO MAROLANZANO', ['PAOLO MAROLANZANO']),
    ('SOLEHUDIN', ['SOLEHUDIN']),
    ('RIFQI ADITYA', ['RIFQI ADITYA']),
    ('M FARHAN ZAHRAN', ['M FARHAN ZAHRAN']),
    ('KAUKABAN AL AKWAN', ['KAUKABAN AL AKWAN']),
    ('NUR MUIS', ['NUR MUIS']),
]
SALES_MARK = 'REKAP PER SALES'

def set_cells(xml, cells):
    """Sisipkan/timpa sel (dict ref->xml <c>) ke sheetData, urut baris & kolom."""
    a=xml.index('<sheetData>')+len('<sheetData>'); b=xml.index('</sheetData>'); body=xml[a:b]
    rows={int(m.group(1)):m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)',body,re.S)}
    byrow={}
    for ref,cx in cells.items():
        r=int(re.sub('[A-Z]','',ref)); byrow.setdefault(r,{})[ref]=cx
    for r,cs in byrow.items():
        if r in rows:
            row=rows[r]
            if row.endswith('/>'): head=row[:-2]+'>'; inner=''
            else:
                head=re.match(r'<row [^>]*>',row).group(0); inner=row[len(head):-len('</row>')]
            head=re.sub(r'\s+spans="[^"]*"','',head)
            ex={m.group(1):m.group(0) for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>.*?</c>)',inner,re.S)}
        else:
            head=f'<row r="{r}">'; ex={}
        ex.update(cs)
        rows[r]=head+''.join(v for k,v in sorted(ex.items(),key=lambda t:cidx(re.sub(r'\d','',t[0]))))+'</row>'
    nb=''.join(v for k,v in sorted(rows.items()))
    return xml[:a]+nb+xml[b:]

def add_sales_table(x, sst_lookup=None):
    # Selalu ditulis ulang (idempoten): setelah file disimpan Excel, teks berubah jadi
    # shared string & id gaya di-renumber, jadi deteksi teks tidak bisa diandalkan.
    existed = SALES_MARK in x
    FP="'Faktur Penjualan'!"; AN,AP,AS,AT=[f"{FP}${c}$2:${c}$95212" for c in ('AN','AP','AS','AT')]
    def tcell(ref,s,t): return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{html.escape(t,quote=False)}</t></is></c>'
    def fcell(ref,s,f): return f'<c r="{ref}" s="{s}"><f>{html.escape(f,quote=False)}</f></c>'
    def ncell(ref,s,v): return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
    c={}
    c['S3']=tcell('S3','218',SALES_MARK+' (NAMA YANG SAMA DIGABUNG)')
    # Baris 4: judul kelompok; baris 5: header. Omset per bulan T:V + total W,
    # Gross Profit per bulan X:Z + total AA, nama Accurate AB.
    c['T4']=tcell('T4','251','OMSET'); c['X4']=tcell('X4','251','GROSS PROFIT')
    hdr=[('S','250','SALES'),('W','251','TOTAL OMSET'),('AA','251','TOTAL GROSS PROFIT'),('AB','250','NAMA DI ACCURATE')]
    for col_,st_,t in hdr: c[f'{col_}5']=tcell(f'{col_}5',st_,t)
    for col_,m in zip('TUV',(7,8,9)): c[f'{col_}5']=ncell(f'{col_}5','254',m)
    for col_,m in zip('XYZ',(7,8,9)): c[f'{col_}5']=ncell(f'{col_}5','254',m)
    r=6
    for name,al in SALES_GROUPS:
        arr='{'+','.join('"'+a+'"' for a in al)+'}'
        c[f'S{r}']=tcell(f'S{r}','255',name)
        for col_ in 'TUV':
            c[f'{col_}{r}']=fcell(f'{col_}{r}','11',f'SUM(SUMIFS({AN},{AP},{arr},{AS},{col_}$5))')
        c[f'W{r}']=fcell(f'W{r}','16',f'SUM(T{r}:V{r})')
        for col_ in 'XYZ':
            c[f'{col_}{r}']=fcell(f'{col_}{r}','11',f'SUM(SUMIFS({AT},{AP},{arr},{AS},{col_}$5))')
        c[f'AA{r}']=fcell(f'AA{r}','16',f'SUM(X{r}:Z{r})')
        c[f'AB{r}']=tcell(f'AB{r}','258',', '.join(al))
        r+=1
    c[f'S{r}']=tcell(f'S{r}','256','TOTAL')
    for col_ in ['T','U','V','W','X','Y','Z','AA']: c[f'{col_}{r}']=fcell(f'{col_}{r}','16',f'SUM({col_}6:{col_}{r-1})')
    c[f'S{r+1}']=tcell(f'S{r+1}','258','Header bulan (angka bulan di baris 5) otomatis mengikuti kuartal. Pivot & tabel per cabang tetap memakai nama asli Accurate.')
    x=set_cells(x,c)
    cols='<col min="19" max="19" width="24" customWidth="1"/><col min="20" max="27" width="15" customWidth="1"/><col min="28" max="28" width="60" customWidth="1"/>'
    m_=re.search(r'<cols>(.*?)</cols>',x,re.S)
    if m_:
        # sisipkan lebar kolom S..Y (19..25) dengan urutan benar; potong entri lama yang bertumpuk
        keep=[]
        for c_ in re.findall(r'<col [^>]*/>',m_.group(1)):
            mn=int(re.search(r'min="(\d+)"',c_).group(1)); mx=int(re.search(r'max="(\d+)"',c_).group(1))
            if mx<19 or mn>28: keep.append((mn,c_)); continue
            if mn<19: keep.append((mn,re.sub(r'max="\d+"','max="18"',c_)))
            if mx>28: keep.append((29,re.sub(r'min="\d+"','min="29"',c_)))
        for c_ in re.findall(r'<col [^>]*/>',cols): keep.append((int(re.search(r'min="(\d+)"',c_).group(1)),c_))
        x=x[:m_.start()]+'<cols>'+''.join(c_ for _,c_ in sorted(keep,key=lambda t:t[0]))+'</cols>'+x[m_.end():]
    x=re.sub(r'<dimension ref="[^"]*"/>','<dimension ref="B3:AB40"/>',x,1)
    return x, not existed

# ---------- Sheet "Omset Bulanan" (dibangun ulang tiap rekap) ----------
BULAN_NAMA={1:'JANUARI',2:'FEBRUARI',3:'MARET',4:'APRIL',5:'MEI',6:'JUNI',7:'JULI',8:'AGUSTUS',9:'SEPTEMBER',10:'OKTOBER',11:'NOVEMBER',12:'DESEMBER'}
HARI_NAMA=['SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU','MINGGU']
def rebuild_omset_bulanan(x, qs, maxd):
    """Hapus isi lama, buat kotak per bulan (kuartal berjalan s/d bulan data terakhir):
    judul bulan, header HARI | TANGGAL | 18 cabang | TOTAL, omset per tanggal (SUMIFS ke
    Faktur Penjualan), baris TOTAL per bulan. Gaya mengikuti kotak buatan user (Mei/Juni 2025)."""
    FP="'Faktur Penjualan'!"; AN,A,B=[f"{FP}${c}$2:${c}$95212" for c in ('AN','A','B')]
    cab=list(CABANG.values()); c0=4; cl=c0+len(cab)-1; ct=cl+1   # D .. U, V
    LC=col(cl); TC=col(ct)
    rows=[]; merges=[]; r=4
    def row(rn,cells,ht=None):
        rows.append(f'<row r="{rn}"'+(f' ht="{ht}" customHeight="1"' if ht else '')+'>'+''.join(cells)+'</row>')
    def t(ref,s,v): return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{html.escape(v,quote=False)}</t></is></c>'
    def n(ref,s,v): return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
    def f(ref,s,fx): return f'<c r="{ref}" s="{s}"><f>{html.escape(fx,quote=False)}</f></c>'
    m=qs.month
    while m<=maxd.month and m<qs.month+3:
        y=qs.year; first=datetime.datetime(y,m,1)
        nd=(datetime.datetime(y+(m==12),(m%12)+1,1)-first).days
        row(r,[t(f'B{r}','245',f'{BULAN_NAMA[m]} {y}')]+[f'<c r="{col(k)}{r}" s="245"/>' for k in range(3,ct+1)],24)
        row(r+1,[f'<c r="{col(k)}{r+1}" s="245"/>' for k in range(2,ct+1)],24)
        merges.append(f'B{r}:{TC}{r+1}')
        h=r+3
        row(h,[t(f'B{h}','87','HARI'),t(f'C{h}','88','TANGGAL')]+[t(f'{col(c0+i)}{h}','89',cb) for i,cb in enumerate(cab)]+[t(f'{TC}{h}','89','TOTAL')],19)
        d0=h+1
        for i in range(nd):
            rr=d0+i; d=first+datetime.timedelta(days=i)
            cells=[t(f'B{rr}','91',HARI_NAMA[d.weekday()]),n(f'C{rr}','92',(d-EPOCH).days)]
            cells+=[f(f'{col(c0+k)}{rr}','93',f'SUMIFS({AN},{A},{col(c0+k)}${h},{B},$C{rr})') for k in range(len(cab))]
            cells.append(f(f'{TC}{rr}','94',f'SUM(D{rr}:{LC}{rr})'))
            row(rr,cells)
        tr=d0+nd; dl=tr-1
        row(tr,[t(f'B{tr}','248','TOTAL'),f'<c r="C{tr}" s="249"/>']+[f(f'{col(k)}{tr}','95',f'SUM({col(k)}{d0}:{col(k)}{dl})') for k in range(c0,ct+1)],19)
        merges.append(f'B{tr}:C{tr}')
        r=tr+4; m+=1
    a=x.index('<sheetData>'); b=x.index('</sheetData>')+len('</sheetData>')
    x=x[:a]+'<sheetData>'+''.join(rows)+'</sheetData>'+x[b:]
    x=re.sub(r'<mergeCells[^>]*>.*?</mergeCells>|<mergeCells[^>]*/>','',x,flags=re.S)
    mc=f'<mergeCells count="{len(merges)}">'+''.join(f'<mergeCell ref="{m_}"/>' for m_ in merges)+'</mergeCells>'
    x=x.replace('</sheetData>','</sheetData>'+mc,1)
    x=re.sub(r'<conditionalFormatting.*?</conditionalFormatting>','',x,flags=re.S)
    x=re.sub(r'<cols>.*?</cols>',f'<cols><col min="2" max="2" width="12" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/><col min="{c0}" max="{cl}" width="15.5" customWidth="1"/><col min="{ct}" max="{ct}" width="19" customWidth="1"/></cols>',x,1,flags=re.S)
    x=re.sub(r'<dimension ref="[^"]*"/>',f'<dimension ref="B4:{TC}{r}"/>',x,1)
    x=re.sub(r'<sheetView ([^>]*?)topLeftCell="[^"]*"',r'<sheetView \1',x,1)
    x=re.sub(r'<selection [^>]*/>','<selection activeCell="B4" sqref="B4"/>',x,1)
    return x

def wrap_getpivot(formula):
    """Bungkus setiap GETPIVOTDATA(...) dengan IFERROR(...,0) (kurung di dalam
    string nama spt "WAHYU JP (RADJIMAN)" diabaikan). Lewati yang sudah dibungkus
    atau yang menunjuk pivot #REF!."""
    out=''; i=0; n=0
    while True:
        j=formula.find('GETPIVOTDATA(',i)
        if j<0: out+=formula[i:]; break
        k=j+len('GETPIVOTDATA('); depth=1; inq=False
        while k<len(formula) and depth:
            ch=formula[k]
            if ch=='"': inq=not inq
            elif not inq and ch=='(': depth+=1
            elif not inq and ch==')': depth-=1
            k+=1
        call=formula[j:k]
        if formula[max(0,j-8):j]=='IFERROR(' or '#REF!' in call: out+=formula[i:k]
        else: out+=formula[i:j]+f'IFERROR({call},0)'; n+=1
        i=k
    return out,n

def wrap_getpivot_sheet(x):
    tot=0
    def fx(m):
        nonlocal tot
        body=html.unescape(m.group(2)); nb,n=wrap_getpivot(body); tot+=n
        return m.group(1)+html.escape(nb,quote=False)+m.group(3)
    x=re.sub(r'(<f(?: [^>]*)?>)([^<]*GETPIVOTDATA[^<]*)(</f>)',fx,x)
    return x,tot

def sheet_paths(z):
    wb=z.read('xl/workbook.xml').decode('utf8'); rels=z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid={m.group(1):m.group(2) for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"',rels)}
    rid.update({m.group(2):m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"',rels)})
    out=collections.OrderedDict()
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"',wb):
        t=rid[m.group(2)].lstrip('/'); t=t if t.startswith('xl/') else 'xl/'+t
        out[html.unescape(m.group(1))]=t
    return out

def read_branches(files):
    data=[]; summary=[]
    seen=set()
    for f in sorted(files, key=lambda f: re.search(r'penjualan_(\d{3})',os.path.basename(f)).group(1)):
        code=re.search(r'penjualan_(\d{3})',os.path.basename(f)).group(1)
        if code not in CABANG: sys.exit(f'Kode cabang {code} tidak dikenal: {f}')
        if code in seen: sys.exit(f'Kode cabang {code} dobel')
        seen.add(code)
        ws=openpyxl.load_workbook(f,read_only=True).active; ws.reset_dimensions()
        rows=list(ws.iter_rows(values_only=True)); hdr=[str(h).strip().upper() if h is not None else None for h in rows[0]]
        # Petakan kolom berdasarkan NAMA header -> bisa baca format omset (ada kolom
        # pemisah) maupun format walk-in (tanpa pemisah, + kolom Nomor/Tanggal
        # Pengiriman Pesanan). Isi kedua export sama (dicek 1 Okt 2026).
        pos={}
        for i,h in enumerate(hdr):
            if h is None: continue
            key='KATEGORI PILAR' if h.startswith('KATEGORI PILAR') else h
            pos.setdefault(key,i)
        missing_cols=[c for c in DASH_COLS if c not in pos and c!='KATEGORI PILAR']
        if missing_cols: sys.exit(f'Kolom tidak ditemukan di {f}: {missing_cols}')
        rs=[r for r in rows[1:] if any(v is not None for v in r)]
        out_rows=[]
        for r in rs:
            o=[None]*44; o[0]=CABANG[code]
            for k,cname in enumerate(DASH_COLS):
                i=pos.get(cname)
                o[1+2*k]=r[i] if i is not None and i<len(r) else None
            if not isinstance(o[1],datetime.datetime): sys.exit(f'Tanggal tidak valid di {f}: {o[1]!r}')
            out_rows.append(o)
        data.extend(out_rows)
        ds=[r[1] for r in out_rows]
        summary.append((code,CABANG[code],len(out_rows),min(ds).date(),max(ds).date(),sum((r[39] or 0) for r in out_rows)))
    missing=set(CABANG)-seen
    return data,summary,missing

def build(dash, data, out):
    zin=zipfile.ZipFile(dash); sp=sheet_paths(zin)
    fpath=sp['Faktur Penjualan']
    sh=zin.read(fpath).decode('utf8')
    a=sh.index('<sheetData>')+len('<sheetData>'); b=sh.index('</sheetData>'); body=sh[a:b]
    rowre=re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)',re.S)
    rows={int(m.group(1)):m.group(0) for m in rowre.finditer(body)}
    row1=rows[1]; row2=rows[2]
    sty={}
    for m in re.finditer(r'<c r="([A-Z]+)2"(?: s="(\d+)")?',row2): sty[m.group(1)]=m.group(2) or '0'
    N=len(data)
    # Timpa SEMUA baris data lama (termasuk kalau data lama lebih panjang dari data baru);
    # hanya baris sentinel jauh di bawah (>= 1.000.000) yang dipertahankan.
    old_last=max([k for k in rows if k<1000000] or [1])
    end=max(75000,N+1,old_last)
    if N+1>95212: print('PERINGATAN: data melebihi baris 95212 (range SUMIFS Detail Data Penjualan)!')
    sst=zin.read('xl/sharedStrings.xml').decode('utf8')
    sis=re.findall(r'<si>(.*?)</si>',sst,flags=re.S); smap={}
    for i,s in enumerate(sis):
        m=re.fullmatch(r'<t(?: xml:space="preserve")?>(.*?)</t>',s,flags=re.S)
        if m: smap.setdefault(html.unescape(m.group(1)),i)
    new_si=[]
    def sidx(t):
        if t in smap: return smap[t]
        i=len(sis)+len(new_si); smap[t]=i
        new_si.append(f'<si><t xml:space="preserve">{html.escape(t,quote=False)}</t></si>'); return i
    COLS=[col(i) for i in range(1,45)]
    def cell(ref,st,v):
        if v is None or v=='': return f'<c r="{ref}" s="{st}"/>'
        if isinstance(v,datetime.datetime):
            d=v-EPOCH; x=d.days+d.seconds/86400
            return f'<c r="{ref}" s="{st}"><v>{int(x) if x==int(x) else x}</v></c>'
        if isinstance(v,(int,float)):
            if isinstance(v,float) and v.is_integer(): v=int(v)
            return f'<c r="{ref}" s="{st}"><v>{v!r}</v></c>'
        return f'<c r="{ref}" s="{st}" t="s"><v>{sidx(str(v))}</v></c>'
    out_rows=[]
    for rn in range(2,end+1):
        r=data[rn-2] if rn-2<N else [None]*44
        cs=''.join(cell(f'{COLS[j]}{rn}',sty.get(COLS[j],'0'),r[j]) for j in range(44))
        out_rows.append(f'<row r="{rn}" spans="1:47">{cs}<c r="AS{rn}" s="{sty.get("AS","0")}"><f>MONTH(B{rn})</f></c><c r="AT{rn}" s="{sty.get("AT","0")}"><f>AN{rn}-AH{rn}</f></c></row>')
    tail=''.join(v for k,v in sorted(rows.items()) if k>end)
    newsh=sh[:a]+row1+''.join(out_rows)+tail+sh[b:]
    newsh=re.sub(r'<pane [^>]*/>','<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',newsh,1)
    newsh=re.sub(r'<selection pane="bottomLeft"[^>]*/>','<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>',newsh,1)
    sst=sst.replace('</sst>',''.join(new_si)+'</sst>')
    sst=re.sub(r'uniqueCount="\d+"',f'uniqueCount="{len(sis)+len(new_si)}"',sst,1)
    sst=re.sub(r' count="\d+"','',sst,1)
    rep={fpath:newsh,'xl/sharedStrings.xml':sst}
    # Data Periode C6 = hari ke-berapa dalam kuartal (tanggal faktur terakhir)
    maxd=max(r[1] for r in data); qs=datetime.datetime(maxd.year,3*((maxd.month-1)//3)+1,1)
    day=(maxd-qs).days+1
    if 'Data Periode' in sp:
        p=sp['Data Periode']; x=zin.read(p).decode('utf8')
        x2,n=re.subn(r'(<c r="C6"[^>]*>)<v>[^<]*</v>(</c>)',rf'\g<1><v>{day}</v>\g<2>',x,1)
        if n: rep[p]=x2
        else: print('PERINGATAN: Data Periode!C6 bukan angka biasa, tidak diubah')
    sst_now=sst
    def _sst_txt(i):
        m=re.findall(r'<si>(.*?)</si>',sst_now,flags=re.S)
        return re.sub('<[^>]+>','',m[int(i)]) if int(i)<len(m) else ''
    if 'Detail Data Penjualan' in sp:
        p=sp['Detail Data Penjualan']; x=zin.read(p).decode('utf8')
        sis_all=re.findall(r'<si>(.*?)</si>',sst_now,flags=re.S)
        tot_cols=[c for c,v in re.findall(r'<c r="([A-Z]+)8"[^>]*?t="s"[^>]*><v>(\d+)</v>',x)
                  if re.sub('<[^>]+>','',sis_all[int(v)]).strip().upper()=='TOTAL']
        # Kolom cabang pertama tiap blok = header KLENDER terdekat di kiri kolom TOTAL;
        # TOTAL = SUM(KLENDER .. kolom sebelum TOTAL) -> otomatis ikut cabang baru (mis. CIMANGGIS).
        kl_cols=[cidx(c) for c,v in re.findall(r'<c r="([A-Z]+)8"[^>]*?t="s"[^>]*><v>(\d+)</v>',x)
                 if re.sub('<[^>]+>','',sis_all[int(v)]).strip().upper()=='KLENDER']
        nfix=0
        for tc in tot_cols:
            ti=cidx(tc); kl=[k for k in kl_cols if k<ti]
            first=col(max(kl)) if kl else col(ti-len(CABANG)); last=col(ti-1)
            # total harian (baris 9..100) & total target (baris 6) = SUM semua cabang
            def fx(m):
                nonlocal nfix
                r=m.group(2); nfix+=1
                return f'{m.group(1)}SUM({first}{r}:{last}{r})</f>'
            x=re.sub(r'(<c r="%s(?:[6-9]|\d\d+)"[^>]*><f[^>]*>)SUM\(\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?\2\)</f>'%tc, fx, x)
        # Target Pejaten & Cibubur (baris 6) sempat menunjuk sel Expected Value
        # Cikampek/Cilangkap (E23/E24 dst) -> arahkan ke baris Pejaten/Cibubur.
        tfix={'T6':('Scoreboard!E23','Scoreboard!C25'),'U6':('Scoreboard!E24','Scoreboard!C26'),
              'AO6':('Scoreboard!E48','Scoreboard!C50'),'AP6':('Scoreboard!E49','Scoreboard!C51'),
              'BK6':('Scoreboard!E73','Scoreboard!C75'),'BL6':('Scoreboard!E74','Scoreboard!C76')}
        nt=0
        for cell,(old,new) in tfix.items():
            x,k=re.subn(r'(<c r="%s"[^>]*><f>)%s(</f>)'%(cell,re.escape(old)),rf'\g<1>{new}\g<2>',x); nt+=k
        # Tanggal baris 9..100 (kolom C) = tanggal kuartal berjalan (92 hari maks)
        qd=0
        for i in range(92):
            rr=9+i; dser=(qs+datetime.timedelta(days=i)-EPOCH).days
            x,k=re.subn(r'(<c r="C%d"[^>]*>)<v>[^<]*</v>(</c>)'%rr,rf'\g<1><v>{dser}</v>\g<2>',x,1); qd+=k
        print(f'  Detail Data Penjualan: tanggal C9:C100 diset mulai {qs:%d %b %Y} ({qd} sel)')
        rep[p]=x
        print(f'  Detail Data Penjualan: {nfix} rumus TOTAL diset ke semua cabang ({", ".join(tot_cols)}); {nt} rumus target Pejaten/Cibubur diperbaiki')
    if 'Scoreboard' in sp:
        p=sp['Scoreboard']; x=zin.read(p).decode('utf8')
        serial=(maxd-EPOCH).days
        x2,n=re.subn(r'(<c r="C4"[^>]*>)<v>[^<]*</v>(</c>)',rf'\g<1><v>{serial}</v>\g<2>',x,1)
        if n:
            x2=re.sub(r'(<c r="C6"[^>]*>)<f>DAY\(C4\)\+31\+31</f>',r"\g<1><f>'Data Periode'!$C$6</f>",x2,1)
            rep[p]=x2; print(f'  Scoreboard C4 (TANGGAL) = {maxd.date()}')
        else: print('PERINGATAN: Scoreboard!C4 bukan angka biasa, tidak diubah')
    if 'Scoreboard' in sp:
        p=sp['Scoreboard']; x,ng=wrap_getpivot_sheet(rep.get(p) or zin.read(p).decode('utf8'))
        if ng: rep[p]=x; print(f'  Scoreboard: {ng} GETPIVOTDATA dibungkus IFERROR(...,0)')
        # Tabel Omset Pengadaan Marketing Corporate (sekali tambah, format = tabel Omset MC)
        drw=None
        relp=p.replace('worksheets/','worksheets/_rels/')+'.rels'
        if relp in zin.namelist():
            mm=re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"',zin.read(relp).decode('utf8'))
            if mm: drw='xl/drawings/'+mm.group(1)
        dx=(rep.get(drw) or zin.read(drw).decode('utf8')) if drw else None
        x,dx2,addp=mc_scoreboard.ensure_pengadaan_block(rep.get(p) or zin.read(p).decode('utf8'),sst_now,dx)
        if addp:
            rep[p]=x
            if drw: rep[drw]=dx2
            print('  Scoreboard: tabel OMSET PENGADAAN MARKETING CORPORATE ditambahkan')
        # Marketing Corporate & GP: pencapaian per nama marketing (SUMIFS ke Faktur, bukan pivot)
        x,mcinfo=mc_scoreboard.fill(rep.get(p) or zin.read(p).decode('utf8'),sst_now,set_cells)
        if mcinfo:
            rep[p]=x
            print('  Scoreboard Marketing Corporate: '+'; '.join(f"{lab}={'/'.join(al) if al else '-'}" for k,r,lab,al in mcinfo if k=='omset'))
    if 'Omset Bulanan' in sp:
        p=sp['Omset Bulanan']; rep[p]=rebuild_omset_bulanan(zin.read(p).decode('utf8'),qs,maxd)
        print(f'  Omset Bulanan: dibangun ulang {qs:%b}..{maxd:%b %Y} ({len(CABANG)} cabang, per tanggal + total bulan)')
    months=[qs.month,qs.month+1,qs.month+2]
    def set_month_hdr(x,cells):
        for ref,mv in zip(cells,months):
            x=re.sub(r'(<c r="%s"[^>]*>)<v>[^<]*</v>(</c>)'%ref,rf'\g<1><v>{mv}</v>\g<2>',x,1)
        return x
    if 'Mark.Corporate' in sp:
        p=sp['Mark.Corporate']; x,added_sales=add_sales_table(rep.get(p) or zin.read(p).decode('utf8'))
        if added_sales: print('  Mark.Corporate: tabel Rekap per Sales ditambahkan (S3:Y22)')
        # GETPIVOTDATA untuk sales yang belum ada transaksinya -> #REF! (merembet ke
        # Scoreboard/Dashboard). Bungkus IFERROR(...,0).
        x,ng=wrap_getpivot_sheet(x)
        if ng: print(f'  Mark.Corporate: {ng} GETPIVOTDATA dibungkus IFERROR(...,0)')
        rep[p]=set_month_hdr(set_month_hdr(x,['T5','U5','V5']),['X5','Y5','Z5'])
    if PILAR_SHEET in sp:
        p=sp[PILAR_SHEET]; rep[p]=set_month_hdr(rep.get(p) or zin.read(p).decode('utf8'),['C30','D30','E30'])
    # Id gaya hardcode (dari file asli) -> id gaya yang setara di workbook ini
    # (Excel me-renumber gaya saat file disimpan ulang; id di luar jangkauan = file 'repair').
    st_now=rep.get('xl/styles.xml') or zin.read('xl/styles.xml').decode('utf8')
    st_new,smap=ensure_styles(st_now,STYLE_IDS)
    if st_new!=st_now: rep['xl/styles.xml']=st_new
    for nm in ('Omset Bulanan','Mark.Corporate'):
        if nm in sp and sp[nm] in rep:
            x=rep[sp[nm]]
            if nm=='Mark.Corporate':
                nrow=6+len(SALES_GROUPS)+1
                def _rs(m):
                    c_,r_,sv=m.group(2),int(m.group(3)),int(m.group(4))
                    if 19<=cidx(c_)<=28 and 3<=r_<=nrow: sv=smap.get(sv,sv)
                    return f'{m.group(1)}s="{sv}"'
                x=re.sub(r'(<c r="([A-Z]+)(\d+)"[^>]*?\s)s="(\d+)"',_rs,x)
            else:
                x=remap_sheet(x,smap)
            rep[sp[nm]]=x
    # <col max> tidak boleh > 16384 (kolom XFD) -> kalau lebih, Excel 'repair' file
    for nm,pth in sp.items():
        if nm=='Faktur Penjualan': continue
        x=rep.get(pth) or zin.read(pth).decode('utf8')
        def _cl(m):
            mn,mx=int(m.group(2)),int(m.group(3))
            if mn>16384: return ''
            return m.group(1)+f'min="{mn}" max="{min(mx,16384)}"'+m.group(4)
        x2=re.sub(r'(<col )min="(\d+)" max="(\d+)"([^>]*/>)',_cl,x)
        if x2!=x: rep[pth]=x2; print(f'  {nm}: lebar kolom > XFD dirapikan')
    print(f'  Header bulan (Omset Kategori Pilar C30:E30, Rekap per Sales T5:V5) = {months}')
    wbx=zin.read('xl/workbook.xml').decode('utf8')
    wbx=re.sub(r'<calcPr([^>]*?)\s*fullCalcOnLoad="1"','<calcPr\\1',wbx)
    wbx=re.sub(r'<calcPr([^>]*?)/>',r'<calcPr\1 fullCalcOnLoad="1"/>',wbx,1)
    ct=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',zin.read('[Content_Types].xml').decode('utf8'))
    rels=re.sub(r'<Relationship [^>]*Target="/?(?:xl/)?calcChain.xml"[^>]*/>','',zin.read('xl/_rels/workbook.xml.rels').decode('utf8'))
    added=False
    if PILAR_SHEET not in sp:
        wbx,rels,ct,st,sheet_xml,newpath=add_pilar_sheet(zin,wbx,rels,ct)
        rep['xl/styles.xml']=st; added=True
    rep.update({'xl/workbook.xml':wbx,'[Content_Types].xml':ct,'xl/_rels/workbook.xml.rels':rels})
    zout=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        if it.filename=='xl/calcChain.xml': continue
        d=rep.get(it.filename); zout.writestr(it, d.encode('utf8') if d is not None else zin.read(it.filename))
    if added: zout.writestr(newpath, sheet_xml.encode('utf8'))
    zout.close()
    return day, maxd, added

def add_pilar_sheet(zin,wbx,rels,ct):
    st=zin.read('xl/styles.xml').decode('utf8')
    m=re.search(r'<numFmts count="(\d+)">',st)
    ids=[int(x) for x in re.findall(r'numFmtId="(\d+)"',re.search(r'<numFmts.*?</numFmts>',st,re.S).group(0))]
    nid=max(ids+[199])+1
    st=st.replace(m.group(0),f'<numFmts count="{int(m.group(1))+1}">',1)
    st=st.replace('</numFmts>',f'<numFmt numFmtId="{nid}" formatCode="&quot;BULAN &quot;0"/></numFmts>',1)
    m=re.search(r'<cellXfs count="(\d+)">',st); base=int(m.group(1))
    al='<alignment horizontal="{}" vertical="center"{}/>'
    def xf(num,font,fill,h,wrap=False):
        return f'<xf numFmtId="{num}" fontId="{font}" fillId="{fill}" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1">'+al.format(h,' wrapText="1"' if wrap else '')+'</xf>'
    new=[xf(0,6,2,'center',True),xf(0,6,3,'center',True),xf(10,0,0,'center'),xf(10,6,3,'center'),xf(nid,6,2,'center'),
         xf(0,0,0,'left'),xf(0,6,3,'left'),xf(0,0,7,'center'),xf(0,0,0,'left'),xf(3,0,0,'center'),xf(3,6,3,'center'),xf(0,11,0,'left')]
    st=st.replace(m.group(0),f'<cellXfs count="{base+len(new)}">',1)
    i=st.index('</cellXfs>'); st=st[:i]+''.join(new)+st[i:]
    H,HT,PCT,PCTT,MH,LBL,LBLT,INP,NOTE,NUM,NUMT,TITLE=[str(base+k) for k in range(len(new))]
    FP="'Faktur Penjualan'!"; AN,A,AR,AS=[f"{FP}${c}$2:${c}$95212" for c in ('AN','A','AR','AS')]
    rows={}
    def put(ref,s,v=None,f=None):
        r=int(re.sub('[A-Z]','',ref)); c=re.sub(r'\d','',ref)
        if f is not None: x=f'<c r="{ref}" s="{s}"><f>{html.escape(f,quote=False)}</f></c>'
        elif isinstance(v,(int,float)): x=f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
        else: x=f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{html.escape(v,quote=False)}</t></is></c>'
        rows.setdefault(r,[]).append((c,x))
    cols=list('CDEFGHIJ'); CAB=list(CABANG.values()); crit='IF($C$3="SEMUA",">0",$C$3)'
    put('B1',TITLE,'PENCAPAIAN OMSET PER KATEGORI PILAR')
    put('B2',NOTE,'Sumber: sheet Faktur Penjualan — omset = TOTAL HARGA (kol. AN), dikelompokkan per KATEGORI PILAR (kol. AR). Otomatis ikut update saat data faktur diganti.')
    put('B3',LBL,'FILTER BULAN'); put('C3',INP,'SEMUA'); put('D3',NOTE,'← isi SEMUA untuk seluruh periode, atau angka bulan (contoh: 9 = September)')
    put('B5',H,'CABANG')
    for j,p in enumerate(PILAR): put(f'{cols[j]}5',H,p)
    put('I5',H,'BELUM ADA KATEGORI PILAR'); put('J5',HT,'TOTAL OMSET')
    for k,c in enumerate(CAB):
        r=6+k; put(f'B{r}',LBL,c)
        for j in range(6): put(f'{cols[j]}{r}',NUM,f=f'SUMIFS({AN},{A},$B{r},{AR},{cols[j]}$5,{AS},{crit})')
        put(f'I{r}',NUM,f=f'J{r}-SUM(C{r}:H{r})'); put(f'J{r}',NUMT,f=f'SUMIFS({AN},{A},$B{r},{AS},{crit})')
    TR=6+len(CAB)   # baris TOTAL (19 cabang -> 25), maks 21 cabang sebelum tabel bulanan (baris 28)
    put(f'B{TR}',LBLT,'TOTAL')
    for c in cols: put(f'{c}{TR}',NUMT,f=f'SUM({c}6:{c}{TR-1})')
    put(f'B{TR+1}',LBL,'% KONTRIBUSI')
    for c in cols: put(f'{c}{TR+1}',PCTT if c=='J' else PCT,f=f'IF($J${TR}=0,0,{c}{TR}/$J${TR})')
    put('B28',TITLE,'OMSET KATEGORI PILAR PER BULAN (ALL CABANG)')
    put('B29',NOTE,'Angka bulan di header (7, 8, 9) bisa diganti untuk periode berikutnya.')
    put('B30',H,'KATEGORI PILAR')
    for c,mo in zip('CDE',(7,8,9)): put(f'{c}30',MH,mo)
    put('F30',HT,'TOTAL'); put('G30',HT,'% KONTRIBUSI')
    for k,p in enumerate(PILAR+['BELUM ADA KATEGORI PILAR']):
        r=31+k; put(f'B{r}',LBL,p)
        for c in 'CDE':
            put(f'{c}{r}',NUM,f=(f'SUMIFS({AN},{AR},$B{r},{AS},{c}$30)' if k<6 else f'SUMIFS({AN},{AS},{c}$30)-SUM({c}31:{c}36)'))
        put(f'F{r}',NUMT,f=f'SUM(C{r}:E{r})'); put(f'G{r}',PCT,f=f'IF($F$38=0,0,F{r}/$F$38)')
    put('B38',LBLT,'TOTAL')
    for c in 'CDEF': put(f'{c}38',NUMT,f=f'SUM({c}31:{c}37)')
    put('G38',PCTT,f='SUM(G31:G37)')
    sd=''.join(f'<row r="{r}"'+(' ht="34" customHeight="1"' if r in (5,30) else '')+'>'+''.join(x for _,x in sorted(rows[r],key=lambda t:cidx(t[0])))+'</row>' for r in sorted(rows))
    sheet=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="16"/>'
      '<cols><col min="1" max="1" width="2.5" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="10" width="19" customWidth="1"/></cols>'
      f'<sheetData>{sd}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    existing=set(n for n in zin.namelist() if n.startswith('xl/worksheets/sheet'))
    k=1
    while f'xl/worksheets/sheet{k}.xml' in existing: k+=1
    newpath=f'xl/worksheets/sheet{k}.xml'
    sid=max(int(x) for x in re.findall(r'<sheet [^>]*sheetId="(\d+)"',wbx))+1
    rid='rIdPilar1'
    entry=f'<sheet name="{PILAR_SHEET}" sheetId="{sid}" r:id="{rid}"/>'
    m=re.search(r'<sheet name="Dashboard"[^>]*/>',wbx)
    wbx=wbx.replace(m.group(0),m.group(0)+entry,1) if m else wbx.replace('</sheets>',entry+'</sheets>',1)
    rels=rels.replace('</Relationships>',f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{k}.xml"/></Relationships>')
    ct=ct.replace('</Types>',f'<Override PartName="/{newpath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    return wbx,rels,ct,st,sheet,newpath

def recalc_values(path, workdir):
    cp=os.path.join(workdir,'recalc_copy.xlsx'); shutil.copy(path,cp); before=os.path.getmtime(cp)
    p=subprocess.Popen([sys.executable,RECALC,cp,'900'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,start_new_session=True)
    t0=time.time()
    while time.time()-t0<900:
        time.sleep(10)
        if os.path.getmtime(cp)!=before:
            s1=os.path.getsize(cp); time.sleep(15)
            if os.path.getsize(cp)==s1: break
    else: sys.exit('LibreOffice recalc timeout')
    try: os.killpg(p.pid,signal.SIGKILL)
    except Exception: pass
    subprocess.run('pkill -9 -f soffice.bin',shell=True)
    print(f'  recalc selesai {time.time()-t0:.0f}s')
    wb=openpyxl.load_workbook(cp,read_only=True,data_only=True); vals={}
    for n in wb.sheetnames:
        if n=='Faktur Penjualan': continue
        d={}
        for row in wb[n].iter_rows():
            for c in row:
                if c.value is not None and hasattr(c,'coordinate'): d[c.coordinate]=c.value
        vals[n]=d
    return vals

def enc(v):
    if isinstance(v,bool): return ' t="b"',str(int(v))
    if isinstance(v,datetime.datetime):
        d=v-EPOCH; return '',repr(d.days+d.seconds/86400+d.microseconds/8.64e10)
    if isinstance(v,(int,float)): return '',repr(v)
    if isinstance(v,str) and v in('#REF!','#DIV/0!','#N/A','#VALUE!','#NAME?','#NUM!','#NULL!'): return ' t="e"',v
    return ' t="str"',html.escape(str(v),quote=False)

def patch(src,out,vals,data):
    cellre=re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)>(<f[^>]*?(?:/>|>.*?</f>))(?:<v>[^<]*</v>)?</c>',re.S)
    zin=zipfile.ZipFile(src); sp=sheet_paths(zin); byp={v:k for k,v in sp.items()}
    zout=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d=zin.read(it.filename); name=byp.get(it.filename)
        if name:
            V=vals.get(name,{})
            def rep(m):
                c,r,attrs,f=m.groups(); attrs=re.sub(r'\s+t="[^"]*"','',attrs)
                if name=='Faktur Penjualan':
                    i=int(r)-2
                    if c=='AS': v=data[i][1].month if i<len(data) else 1
                    elif c=='AT':
                        v=((data[i][39] or 0)-(data[i][33] or 0)) if i<len(data) else 0
                        if isinstance(v,float) and v.is_integer(): v=int(v)
                    else: return m.group(0)
                else:
                    v=V.get(c+r)
                    if v is None: return f'<c r="{c}{r}"{attrs}>{f}</c>'
                t,s=enc(v); return f'<c r="{c}{r}"{attrs}{t}>{f}<v>{s}</v></c>'
            d=cellre.sub(rep,d.decode('utf8')).encode('utf8')
        elif re.fullmatch(r'xl/pivotCache/pivotCacheDefinition\d+\.xml',it.filename):
            t=d.decode('utf8'); t=re.sub(r'\s+refreshOnLoad="[^"]*"','',t)
            d=t.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" ',1).encode('utf8')
        zout.writestr(it,d)
    zout.close()

if __name__=='__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('--dashboard',required=True); ap.add_argument('--out',required=True)
    ap.add_argument('--workdir',required=True); ap.add_argument('branches',nargs='+'); a=ap.parse_args()
    os.makedirs(a.workdir,exist_ok=True)
    data,summary,missing=read_branches(a.branches)
    print('Cabang:'); [print(f'  {c} {n:13s} {k:6d} baris  {d1}..{d2}  Rp {t:,.0f}') for c,n,k,d1,d2,t in summary]
    if missing: print('PERINGATAN: file cabang tidak ada untuk', sorted(CABANG[m] for m in missing))
    print('Total baris', len(data))
    stage=os.path.join(a.workdir,'stage1.xlsx')
    day,maxd,added=build(a.dashboard,data,stage)
    print(f'Tanggal faktur terakhir {maxd.date()} -> Data Periode!C6 = {day}; sheet pilar baru ditambahkan: {added}')
    vals=recalc_values(stage,a.workdir)
    errs=[(n,k,v) for n,d in vals.items() for k,v in d.items() if isinstance(v,str) and v.startswith('#')]
    print('Sel error setelah recalc:',len(errs),collections.Counter((n,v) for n,k,v in errs))
    patch(stage,a.out,vals,data)
    D=vals.get('Detail Data Penjualan',{})
    print('Detail row5 (TTL PENCAPAIAN):'); [print(f'  {D.get(c+"8")}: {D.get(c+"5")}') for c in 'DEFGHIJKLMNOPQRSTUV']
    P=vals.get(PILAR_SHEET,{})
    print('Pilar (seluruh periode):'); [print(f'  {P.get(f"B{r}")}: {P.get(f"F{r}")}') for r in range(31,39)]
    print('OUT',a.out)
