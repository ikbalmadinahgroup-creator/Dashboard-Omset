#!/usr/bin/env python3
"""Rekap faktur penjualan 18 cabang MFlash ke Dashboard 6 Pilar (XML surgery, tanpa openpyxl full-load).
Usage: python3 rekap_mflash.py --dashboard DASH.xlsx --out OUT.xlsx --workdir DIR BRANCH1.xlsx BRANCH2.xlsx ...
"""
import argparse, zipfile, re, html, datetime, os, sys, time, shutil, subprocess, signal, warnings, collections
import openpyxl
warnings.filterwarnings('ignore')

CABANG = {'001':'KLENDER','002':'CEGER','003':'BINTARA','004':'RADJIMAN','005':'JATIMULYA','006':'DRAMAGA',
          '007':'CONDET','008':'JATIBENING','009':'SAWANGAN','010':'WARBONG','011':'CINERE','012':'CIBINONG',
          '013':'KARAWANG','014':'JATIWARINGIN','015':'CIKAMPEK','016':'CILANGKAP','017':'PEJATEN','018':'CIBUBUR'}
PILAR = ['SERVICE','PENJUALAN RITEL','PENGADAAN CORPORATE','MAINTENANCE CORPORATE','CICILAN SYARIAH','SEWA']
PILAR_SHEET = 'Omset Kategori Pilar'
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
    N=len(data); end=max(75000,N+1)
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
        nfix=0
        for tc in tot_cols:
            ti=cidx(tc); first=col(ti-18); last=col(ti-1)
            # total harian (baris 9..100) & total target (baris 6) = SUM 18 cabang
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
        rep[p]=x
        print(f'  Detail Data Penjualan: {nfix} rumus TOTAL diset ke 18 cabang ({", ".join(tot_cols)}); {nt} rumus target Pejaten/Cibubur diperbaiki')
    if 'Scoreboard' in sp:
        p=sp['Scoreboard']; x=zin.read(p).decode('utf8')
        serial=(maxd-EPOCH).days
        x2,n=re.subn(r'(<c r="C4"[^>]*>)<v>[^<]*</v>(</c>)',rf'\g<1><v>{serial}</v>\g<2>',x,1)
        if n: rep[p]=x2; print(f'  Scoreboard C4 (TANGGAL) = {maxd.date()}')
        else: print('PERINGATAN: Scoreboard!C4 bukan angka biasa, tidak diubah')
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
    put('B24',LBLT,'TOTAL')
    for c in cols: put(f'{c}24',NUMT,f=f'SUM({c}6:{c}23)')
    put('B25',LBL,'% KONTRIBUSI')
    for c in cols: put(f'{c}25',PCTT if c=='J' else PCT,f=f'IF($J$24=0,0,{c}24/$J$24)')
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
