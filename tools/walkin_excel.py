"""Buat Dashboard Walk-in MFlash (Excel) dari folder data/walkin.
Usage: python3 walkin_excel.py <folder data/walkin> <output.xlsx> [target_per_hari=22]"""
import sys,glob,math,datetime,warnings; warnings.filterwarnings('ignore')
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.formatting.rule import CellIsRule,FormulaRule
from openpyxl.utils import get_column_letter as L

SRC,OUT=sys.argv[1],sys.argv[2]; TARGET=float(sys.argv[3]) if len(sys.argv)>3 else 22
CAB={'001':'KLENDER','002':'CEGER','003':'BINTARA','004':'RADJIMAN','005':'JATIMULYA','006':'DRAMAGA','007':'CONDET','008':'JATIBENING','009':'SAWANGAN','010':'WARBONG','011':'CINERE','012':'CIBINONG','013':'KARAWANG','014':'JATIWARINGIN','015':'CIKAMPEK','016':'CILANGKAP','017':'PEJATEN','018':'CIBUBUR'}
BULAN={1:'JANUARI',2:'FEBRUARI',3:'MARET',4:'APRIL',5:'MEI',6:'JUNI',7:'JULI',8:'AGUSTUS',9:'SEPTEMBER',10:'OKTOBER',11:'NOVEMBER',12:'DESEMBER'}
HARI=['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu']

# ---------- data ----------
recs=[]; missing=[]
for code,name in CAB.items():
    fd=glob.glob(f'{SRC}/rincian_pengiriman_pesanan_{code}*'); ff=glob.glob(f'{SRC}/rincian_faktur_penjualan_{code}*')
    if not fd or not ff: missing.append(name); continue
    do=pd.read_excel(sorted(fd)[-1]); fk=pd.read_excel(sorted(ff)[-1])
    if 'KECAMATAN' not in do.columns: do['KECAMATAN']=None
    c=[x for x in fk.columns if 'Pengiriman' in str(x) and ('Nomor' in str(x) or 'No' in str(x).split())][0]
    dov=do.dropna(subset=['NOMOR PENGIRIMAN PESANAN']).groupby('NOMOR PENGIRIMAN PESANAN').agg(tgl=('TGL PENGIRIMAN','min'),st=('STATUS PENGERJAAN','first'),kat=('KATEGORI PENJUALAN','first'),kec=('KECAMATAN','first')).reset_index()
    fv=fk.dropna(subset=['NO FAKTUR']).groupby('NO FAKTUR').agg(has=(c,lambda s:s.notna().any()),tgl=('TGL FAKTUR','min'),kat=('KATEGORI PENJUALAN','first')).reset_index()
    for r in dov.itertuples(): recs.append((name,pd.Timestamp(r.tgl).normalize(),'DO',str(r.st).upper(),str(r.kat).upper(),r.kec))
    for r in fv[~fv.has].itertuples(): recs.append((name,pd.Timestamp(r.tgl).normalize(),'LANGSUNG','',str(r.kat).upper(),None))
df=pd.DataFrame(recs,columns=['cab','tgl','src','st','kat','kec'])
last=df.tgl.max(); qm=3*((last.month-1)//3)+1
qs=pd.Timestamp(last.year,qm,1); qe=(qs+pd.DateOffset(months=3))-pd.Timedelta(days=1)
df=df[(df.tgl>=qs)&(df.tgl<=last)]
dates=pd.date_range(qs,qe); months=[qm,qm+1,qm+2]
cabs=[c for c in CAB.values() if c not in missing]
cnt={s:df[df.src==s].groupby(['tgl','cab']).size() for s in ('DO','LANGSUNG')}

# ---------- styles ----------
AR='Arial'
fT=Font(name=AR,bold=True,size=14,color='0F766E'); fH=Font(name=AR,bold=True,color='FFFFFF',size=10)
fB=Font(name=AR,size=10); fBb=Font(name=AR,size=10,bold=True); fIn=Font(name=AR,size=10,color='0000FF'); fNote=Font(name=AR,size=9,italic=True,color='6B7280')
pH=PatternFill('solid',fgColor='0F766E'); pT=PatternFill('solid',fgColor='134E4A'); pSum=PatternFill('solid',fgColor='E6F4F1'); pIn=PatternFill('solid',fgColor='FFF2CC')
C=Alignment(horizontal='center',vertical='center',wrap_text=True); Lw=Alignment(horizontal='left',vertical='top',wrap_text=True)
thin=Side(style='thin',color='D1D5DB'); BR=Border(left=thin,right=thin,top=thin,bottom=thin)
N0='#,##0;-#,##0;"-"'; N1='#,##0;-#,##0;0'  # rata-rata: bilangan bulat (dibulatkan ke atas)

wb=Workbook()
# ---------- Info ----------
wi=wb.active; wi.title='Info'
wi['B2']='DASHBOARD WALK-IN MFLASH'; wi['B2'].font=fT
rows=[('Periode kuartal',f'{qs:%d %b %Y} – {qe:%d %b %Y}',None),
      ('Tanggal data terakhir',last.to_pydatetime(),'dd mmm yyyy'),
      ('Hari berjalan (s/d data terakhir)','=C5-DATE(%d,%d,1)+1'%(qs.year,qs.month),'0'),
      ('Target walk-in / hari / cabang',TARGET,'0.0')]
for i,(k,v,f) in enumerate(rows,4):
    wi[f'B{i}']=k; wi[f'C{i}']=v; wi[f'B{i}'].font=fBb; wi[f'C{i}'].font=fB
    if f: wi[f'C{i}'].number_format=f
wi['C7'].font=fIn; wi['C7'].fill=pIn
wi['B9']='Cara hitung walk-in'; wi['B9'].font=fBb
notes=['1 nomor Pengiriman Pesanan (DO) = 1 walk-in, di tanggal pesanan dibuat. Semua status dihitung, termasuk CANCEL / PENDING / CLAIM.',
       'Faktur yang punya nomor DO tidak dihitung lagi (konsumen yang sama saat ambil unit).',
       'Faktur tanpa DO = 1 walk-in per nomor faktur (konsumen datang & langsung beli), di tanggal faktur.',
       'Sumber: file Rincian Pengiriman Pesanan + Rincian Faktur Penjualan (format walk-in) per cabang. Angka harian diisi dari file export (biru); total, rata-rata & status pakai rumus.',
       'Target / hari (sel kuning C7) bisa diubah — sheet Rekap & Strategi ikut menyesuaikan status.']
for i,n in enumerate(notes,10): wi[f'B{i}']='• '+n; wi[f'B{i}'].font=fB; wi[f'B{i}'].alignment=Lw; wi.merge_cells(f'B{i}:H{i}'); wi.row_dimensions[i].height=30
if missing: wi['B16']='File belum ada untuk: '+', '.join(missing); wi['B16'].font=Font(name=AR,color='DC2626',bold=True)
wi.column_dimensions['A'].width=2; wi.column_dimensions['B'].width=36; wi.column_dimensions['C'].width=26
for c in 'DEFGH': wi.column_dimensions[c].width=14
HARI_BERJALAN="Info!$C$6"; TGT="Info!$C$7"; LAST="Info!$C$5"

# ---------- Detail ----------
wd=wb.create_sheet('Detail Walk-in')
n=len(cabs); FIRST=9; LASTROW=FIRST+len(dates)-1
sections=[('WALK-IN ALL',3),('PESANAN (DO)',3+n+3),('BELI LANGSUNG (FAKTUR TANPA DO)',3+2*(n+3))]
def cols(start): return {'tgl':start,'cab':{cb:start+1+i for i,cb in enumerate(cabs)},'tot':start+1+n}
S=[cols(s) for _,s in sections]
for (title,st),sc in zip(sections,S):
    tl=L(sc['tgl']); tt=L(sc['tot'])
    wd.cell(1,st,f'DATA WALK-IN — {title}').font=fT
    wd.merge_cells(f'{tl}1:{tt}1')
    labels=['MULAI DATA','TOTAL WALK-IN','HARI AKTIF','RATA-RATA / HARI','TARGET / HARI','SELISIH / HARI']
    for k,lab in enumerate(labels,2):
        c=wd.cell(k,sc['tgl'],lab); c.font=fBb; c.alignment=C; c.fill=pSum; c.border=BR
    for cb,ci in list(sc['cab'].items())+[('TOTAL',sc['tot'])]:
        col=L(ci); rng=f'{col}{FIRST}:{col}{LASTROW}'
        drng=f'${L(sc["tgl"])}${FIRST}:${L(sc["tgl"])}${LASTROW}'
        wd.cell(2,ci,f'=IFERROR(_xlfn.MINIFS({drng},{rng},">0"),"")' if cb!='TOTAL' else f'=MIN({L(sc["cab"][cabs[0]])}2:{L(sc["cab"][cabs[-1]])}2)').number_format='dd mmm yy'
        wd.cell(3,ci,f'=SUM({rng})').number_format=N0
        wd.cell(4,ci,f'=IF({col}2="",0,MAX(0,{LAST}-{col}2+1))').number_format=N0
        wd.cell(5,ci,f'=IF({col}4=0,0,ROUNDUP({col}3/{col}4,0))').number_format=N1
        if st==3:
            wd.cell(6,ci,f'={TGT}' if cb!='TOTAL' else f'={TGT}*{n}').number_format=N1
            wd.cell(7,ci,f'={col}5-{col}6').number_format=N1
        for k in range(2,8):
            cc=wd.cell(k,ci); cc.font=fBb if cb=='TOTAL' else fB; cc.alignment=C; cc.border=BR; cc.fill=pSum
        h=wd.cell(8,ci,cb); h.font=fH; h.fill=pT if cb=='TOTAL' else pH; h.alignment=C; h.border=BR
    h=wd.cell(8,sc['tgl'],'TANGGAL'); h.font=fH; h.fill=pH; h.alignment=C; h.border=BR
    if st==3:
        wd.conditional_formatting.add(f'{L(sc["cab"][cabs[0]])}7:{L(sc["cab"][cabs[-1]])}7',CellIsRule(operator='lessThan',formula=['0'],font=Font(name=AR,bold=True,color='DC2626')))
        wd.conditional_formatting.add(f'{L(sc["cab"][cabs[0]])}7:{L(sc["cab"][cabs[-1]])}7',CellIsRule(operator='greaterThanOrEqual',formula=['0'],font=Font(name=AR,bold=True,color='16A34A')))
for r,d in enumerate(dates,FIRST):
    has=d<=last
    for si,sc in enumerate(S):
        c=wd.cell(r,sc['tgl'],d.to_pydatetime()); c.number_format='dd mmm yyyy'; c.font=fB; c.alignment=C; c.border=BR
        for cb,ci in sc['cab'].items():
            if si==0:
                v=f'={L(S[1]["cab"][cb])}{r}+{L(S[2]["cab"][cb])}{r}' if has else None; font=fB
            else:
                src='DO' if si==1 else 'LANGSUNG'
                v=int(cnt[src].get((d,cb),0)) if has else None; font=fIn
            cc=wd.cell(r,ci,v); cc.font=font; cc.number_format=N0; cc.alignment=C; cc.border=BR
        t=wd.cell(r,sc['tot'],f'=SUM({L(sc["cab"][cabs[0]])}{r}:{L(sc["cab"][cabs[-1]])}{r})' if has else None)
        t.font=fBb; t.number_format=N0; t.alignment=C; t.border=BR
for sc in S:
    wd.column_dimensions[L(sc['tgl'])].width=17
    for ci in list(sc['cab'].values())+[sc['tot']]: wd.column_dimensions[L(ci)].width=12.5
    wd.column_dimensions[L(sc['tot']+1)].width=3
wd.column_dimensions['A'].width=2; wd.column_dimensions['B'].width=2
wd.row_dimensions[8].height=30; wd.freeze_panes='D9'

# ---------- Rekap Bulanan ----------
wr=wb.create_sheet('Rekap Bulanan')
wr['B1']='REKAP WALK-IN PER BULAN'; wr['B1'].font=fT
wr['B2']='Rata-rata / hari bulan berjalan dihitung sampai tanggal data terakhir. Status dibanding target di sheet Info.'; wr['B2'].font=fNote
hdr=['CABANG']+[f'{BULAN[m]}' for m in months]+['TOTAL Q']+[f'RATA2/HARI {BULAN[m][:3]}' for m in months]+['RATA2/HARI Q','TARGET/HARI','SELISIH/HARI','TAMBAHAN WALK-IN/BULAN','STATUS','% BELI LANGSUNG','% CANCEL (DARI DO)','MULAI DATA','HARI AKTIF']
for j,h in enumerate(hdr,2):
    c=wr.cell(4,j,h); c.font=fH; c.fill=pH; c.alignment=C; c.border=BR
wr.row_dimensions[4].height=42
tcol=L(S[0]['tgl']); dr=f"'Detail Walk-in'!${tcol}${FIRST}:${tcol}${LASTROW}"
def days_in(m,r):  # hari aktif di bulan m (mulai dari tgl data pertama cabang)
    y=qs.year; start=f'MAX(DATE({y},{m},1),Q{r})'; end=f'EOMONTH(DATE({y},{m},1),0)'
    return f'MAX(0,MIN({end},{LAST})-{start}+1)'
stats={}
for i,cb in enumerate(cabs):
    r=5+i; wr.cell(r,2,cb).font=fBb
    sub=df[df.cab==cb]; ndo=(sub.src=='DO').sum()
    stats[cb]=dict(lgs=(sub.src=='LANGSUNG').mean() if len(sub) else 0, cancel=sub.st.str.contains('CANCEL').sum()/ndo if ndo else 0)
    vcol=L(S[0]['cab'][cb]); vr=f"'Detail Walk-in'!${vcol}${FIRST}:${vcol}${LASTROW}"
    for k,m in enumerate(months):
        y=qs.year
        wr.cell(r,3+k,f'=SUMPRODUCT(({dr}>=DATE({y},{m},1))*({dr}<=EOMONTH(DATE({y},{m},1),0))*N(+{vr}))').number_format=N0
    wr.cell(r,6,f'=SUM(C{r}:E{r})').number_format=N0
    for k,m in enumerate(months):
        col=L(3+k); wr.cell(r,7+k,f'=IF({days_in(m,r)}=0,0,ROUNDUP({col}{r}/{days_in(m,r)},0))').number_format=N1
    wr.cell(r,10,f'=IF(R{r}=0,0,ROUNDUP(F{r}/R{r},0))').number_format=N1
    wr.cell(r,17,f"='Detail Walk-in'!{vcol}2").number_format='dd mmm yy'
    wr.cell(r,18,f"='Detail Walk-in'!{vcol}4").number_format=N0
    wr.cell(r,11,f'={TGT}').number_format=N1
    wr.cell(r,12,f'=J{r}-K{r}').number_format=N1
    wr.cell(r,13,f'=IF(L{r}<0,-L{r}*30,0)').number_format=N0
    wr.cell(r,14,f'=IF(J{r}<K{r},"DI BAWAH TARGET","ON TARGET")')
    wr.cell(r,15,round(stats[cb]['lgs'],4)).number_format='0.0%'
    wr.cell(r,16,round(stats[cb]['cancel'],4)).number_format='0.0%'
    for j in range(2,19):
        c=wr.cell(r,j); c.border=BR; c.alignment=C if j>2 else Alignment(vertical='center'); 
        if j not in (2,): c.font=fIn if j in (15,16) else fB
rT=5+len(cabs)
wr.cell(rT,2,'TOTAL / RATA-RATA').font=fH
for j in range(3,7): wr.cell(rT,j,f'=SUM({L(j)}5:{L(j)}{rT-1})').number_format=N0
for j in range(7,11): wr.cell(rT,j,f'=SUM({L(j)}5:{L(j)}{rT-1})').number_format=N1
wr.cell(rT,11,f'={TGT}*{len(cabs)}').number_format=N1; wr.cell(rT,12,f'=J{rT}-K{rT}').number_format=N1
wr.cell(rT,14,f'=COUNTIF(N5:N{rT-1},"DI BAWAH TARGET")&" cabang di bawah target"')
for j in range(2,19):
    c=wr.cell(rT,j); c.fill=pT; c.font=fH; c.alignment=C; c.border=BR
wr.conditional_formatting.add(f'N5:N{rT-1}',CellIsRule(operator='equal',formula=['"DI BAWAH TARGET"'],font=Font(name=AR,bold=True,color='DC2626'),fill=PatternFill('solid',fgColor='FEE2E2')))
wr.conditional_formatting.add(f'N5:N{rT-1}',CellIsRule(operator='equal',formula=['"ON TARGET"'],font=Font(name=AR,bold=True,color='16A34A'),fill=PatternFill('solid',fgColor='DCFCE7')))
wr.conditional_formatting.add(f'L5:L{rT-1}',CellIsRule(operator='lessThan',formula=['0'],font=Font(name=AR,bold=True,color='DC2626')))
wr.cell(rT+2,2,'Rata-rata dibulatkan ke atas (tanpa koma) dan dihitung per HARI AKTIF (sejak tanggal data pertama cabang, supaya cabang baru tidak terlihat rendah karena hari sebelum buka). Tambahan walk-in / bulan = kekurangan rata-rata per hari × 30 hari. % beli langsung & % cancel dari data export (biru).').font=fNote
wr.column_dimensions['A'].width=2; wr.column_dimensions['B'].width=18
for j in range(3,19): wr.column_dimensions[L(j)].width=13
wr.column_dimensions['N'].width=18; wr.freeze_panes='C5'

# ---------- Strategi ----------
ws=wb.create_sheet('Strategi Walk-in')
ndays=(last-qs).days+1
dd=pd.date_range(qs,last); dwn=pd.Series(dd.dayofweek).value_counts()
prev_m,cur_m=months[1] if last.month==months[2] else months[0], last.month
def mavg(sub,m):
    s=max(pd.Timestamp(qs.year,m,1),sub.tgl.min()); e=min(pd.Timestamp(qs.year,m,1)+pd.offsets.MonthEnd(0),last); dn=(e-s).days+1
    return (sub[(sub.tgl>=s)&(sub.tgl<=e)].shape[0]/dn) if dn>0 else 0
rows=[]
for cb in cabs:
    sub=df[df.cab==cb]; avg=math.ceil(round(len(sub)/max((last-max(sub.tgl.min(),qs)).days+1,1),9))
    cur=math.ceil(round(mavg(sub,cur_m),9)); prev=math.ceil(round(mavg(sub,prev_m),9)) if prev_m!=cur_m else cur
    rows.append((cb,avg,cur,prev,sub))
below=[r for r in rows if r[1]<TARGET]; watch=[r for r in rows if r[1]>=TARGET and r[2]<TARGET]
ws['B1']=f'STRATEGI MENINGKATKAN WALK-IN — CABANG DI BAWAH {TARGET:g} / HARI'; ws['B1'].font=fT
ws['B2']=(f'Dasar: rata-rata walk-in per hari aktif {qs:%d %b} – {last:%d %b %Y} (cabang baru dihitung sejak data pertama). {len(below)} cabang di bawah target'
          +(f'; {len(watch)} cabang lain sudah di atas target secara kuartal tapi bulan {BULAN[cur_m].title()} turun di bawah target (perlu dijaga).' if watch else '.')
          +' Rekomendasi disusun dari pola data tiap cabang; angka dari file export, strategi adalah usulan untuk dieksekusi & dievaluasi cabang.')
ws['B2'].font=fNote; ws['B2'].alignment=Lw; ws.merge_cells('B2:K2'); ws.row_dimensions[2].height=42
H=['CABANG','RATA2/HARI Q',f'RATA2/HARI {BULAN[cur_m][:3]}','TREN vs BULAN LALU','KURANG / HARI','TAMBAHAN / BULAN','TEMUAN DATA','FOKUS MASALAH','STRATEGI ONLINE','STRATEGI OFFLINE','TARGET 30 HARI']
for j,h in enumerate(H,2):
    c=ws.cell(4,j,h); c.font=fH; c.fill=pH; c.alignment=C; c.border=BR
ws.row_dimensions[4].height=32
def pct(x): return f'{x*100:.0f}%'
def build(cb,avg,cur,prev,sub,flag_watch=False):
    ndo=(sub.src=='DO').sum(); lg=(sub.src=='LANGSUNG').mean(); canc=sub.st.str.contains('CANCEL').sum()/ndo if ndo else 0
    dow=sub.groupby(sub.tgl.dt.dayofweek).size().reindex(range(7),fill_value=0)/dwn.reindex(range(7)).values
    weak=dow.idxmin(); strong=dow.idxmax()
    lap=sub.kat.str.contains('LAPTOP').mean()
    kec=sub[sub.src=='DO'].kec.dropna().astype(str).str.strip().str.upper()
    kec=kec[kec.str.len()>2]; kec_ok=len(kec)/max(ndo,1)
    import difflib
    vc=kec.value_counts(normalize=True); merged={}
    for k,v in vc.items():
        hit=next((m for m in merged if difflib.SequenceMatcher(None,k,m).ratio()>=0.8),None)
        merged[hit or k]=merged.get(hit or k,0)+v
    top=pd.Series(merged).sort_values(ascending=False).head(2)
    start=sub.tgl.min()
    f=[];m=[];on=[];off=[]
    trend=cur-prev
    f.append(f'Hari tersepi {HARI[weak]} ({math.ceil(dow[weak])}/hari), teramai {HARI[strong]} ({math.ceil(dow[strong])}/hari).')
    f.append(f'Beli langsung {pct(lg)} dari walk-in; cancel {pct(canc)} dari DO; service laptop {pct(lap)}.')
    if len(top) and kec_ok>0.5: f.append('Area utama konsumen: '+', '.join(f'{k.title()} {v*100:.0f}%' for k,v in top.items())+'.')
    else: f.append(f'Kolom KECAMATAN hanya terisi rapi {pct(kec_ok)} DO — area asal konsumen belum bisa dibaca.')
    if start>qs+pd.Timedelta(days=7): f.append(f'Data baru mulai {start:%d %b %Y} (cabang baru).')
    if avg<10:
        m.append('Awareness rendah — jumlah kunjungan jauh di bawah cabang lain.')
        on+=['Lengkapi & optimasi Google Business Profile (foto toko, jam buka, WA, layanan) dan minta review dari setiap konsumen selesai service.',
             'Iklan Meta/Google radius 3–5 km dari cabang dengan pesan "cek kerusakan gratis" + tombol WhatsApp.']
        off+=['Program pembukaan/re-launch: promo cek gratis & diskon jasa minggu pertama, spanduk & brosur di perumahan, sekolah/kampus, dan pasar sekitar.',
              'Kerja sama komunitas/RT-RW dan kantor sekitar (titip brosur, kupon diskon karyawan).']
    if trend<=-2:
        m.append(f'Turun {abs(trend):.0f} walk-in/hari di {BULAN[cur_m].title()} dibanding {BULAN[prev_m].title()}.')
        off.append('Cek penyebab penurunan bulan ini: jam buka, jumlah teknisi/admin, stok sparepart populer, kompetitor baru di sekitar.')
        on.append('Pastikan iklan & posting media sosial cabang tetap jalan tiap minggu (cek apakah ada iklan yang berhenti bulan ini).')
    elif trend>=1:
        m.append(f'Sudah naik {trend:.0f}/hari di {BULAN[cur_m].title()} — pertahankan aktivitas yang sedang jalan.')
    if canc>=0.15:
        m.append(f'Cancel tinggi ({pct(canc)} DO) — banyak konsumen datang tapi batal.')
        off.append('Follow-up WA H+1 untuk semua DO cancel (tawarkan opsi sparepart lebih murah / cicilan / garansi); catat alasan cancel.')
    if lg<0.05:
        m.append('Hampir tidak ada pembelian langsung (walk-in hampir semua lewat pesanan service).')
        off.append('Perkuat penjualan aksesoris impulsif: display tempered glass, charger, casing di dekat kasir & etalase depan; bundling saat serah terima unit.')
    if lap>=0.2:
        on.append('Konten & iklan khusus service laptop (ganti baterai, upgrade SSD/RAM, bersih fan) — segmen laptop sudah kuat di cabang ini.')
    elif lap<=0.12:
        on.append('Kenalkan layanan service laptop di cabang ini (konten before-after, promo cek laptop gratis) — porsinya masih kecil.')
    on.append(f'Jadwalkan promo khusus hari {HARI[weak]} (mis. diskon jasa / voucher aksesoris) dan umumkan H-1 di media sosial & status WA.')
    if len(top) and kec_ok>0.5:
        on.append('Targetkan iklan lokasi ke area '+' & '.join(k.title() for k in top.index)+' (asal konsumen terbanyak) dan uji 1 area baru di sekitarnya.')
    else:
        off.append('Admin wajib isi KECAMATAN di setiap DO supaya area asal konsumen bisa dipakai untuk target iklan.')
    off.append('Program referral: konsumen yang bawa teman dapat voucher aksesoris / diskon jasa berikutnya.')
    if flag_watch: m.insert(0,f'Rata-rata kuartal sudah ≥ {TARGET:g}, tapi {BULAN[cur_m].title()} turun di bawah target.')
    gap=max(TARGET-avg,0) if not flag_watch else max(TARGET-cur,0)
    step=min(gap, max(2, math.ceil(gap*0.3)))
    tgt=f'Naik ke {min(TARGET,cur+step):.0f} walk-in/hari dalam 30 hari (+{step:.0f}/hari), lalu evaluasi mingguan.'
    return avg,cur,trend,gap,f,m,on,off,tgt
r=5
allrows=[(x,False) for x in sorted(below,key=lambda t:t[1])]+[(x,True) for x in watch]
for (cb,avg,cur,prev,sub),w in allrows:
    avg,cur,trend,gap,f,m,on,off,tgt=build(cb,avg,cur,prev,sub,w)
    vals=[cb+(' (WASPADA)' if w else ''),avg,cur,trend,gap,gap*30,'\n'.join('• '+x for x in f),'\n'.join('• '+x for x in m) or '• Selisih ke target relatif kecil.','\n'.join('• '+x for x in on),'\n'.join('• '+x for x in off),tgt]
    for j,v in enumerate(vals,2):
        c=ws.cell(r,j,v); c.border=BR; c.font=fBb if j==2 else fB
        c.alignment=Lw if j>=8 else C
        if j in (3,4,6): c.number_format=N1
        if j==5: c.number_format='+0;-0;0'
        if j==7: c.number_format=N0
    ws.row_dimensions[r].height=max(150, 16*max(len(on),len(off),len(f))*2.2)
    r+=1
r+=1
ws.cell(r,2,'STRATEGI UMUM SEMUA CABANG').font=fT; r+=1
umum=[('Online','Google Business Profile tiap cabang: foto terbaru, jam buka akurat, balas semua review, minta review setelah service selesai.'),
      ('Online','Iklan Meta/Google per cabang dengan radius 3–5 km dan tombol WhatsApp; ukur jumlah chat yang jadi walk-in (tanyakan "tahu dari mana" saat buat DO).'),
      ('Online','Konten rutin before-after perbaikan (LCD, baterai, mati total) & testimoni konsumen di Instagram/TikTok cabang.'),
      ('Offline','Follow-up semua DO cancel & pending lewat WhatsApp dalam 24 jam — setiap konsumen yang batal adalah walk-in yang bisa kembali.'),
      ('Offline','Promo hari sepi per cabang (lihat kolom Temuan Data) supaya kunjungan lebih merata sepanjang minggu.'),
      ('Offline','Kerja sama B2B lokal: kantor, sekolah, kampus, komunitas ojol (paket service karyawan/anggota).'),
      ('Offline','Standar etalase aksesoris di depan toko & bundling saat serah terima unit untuk menaikkan pembelian langsung.'),
      ('Evaluasi','Pantau walk-in harian di sheet Detail Walk-in & app Streamlit tiap minggu; bandingkan dengan target di sheet Info.')]
for k,(a,b) in enumerate(umum):
    ws.cell(r,2,a).font=fBb; c=ws.cell(r,3,b); c.font=fB; c.alignment=Lw; ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=12); ws.row_dimensions[r].height=22; r+=1
widths={'A':2,'B':20,'C':11,'D':11,'E':11,'F':10,'G':11,'H':40,'I':34,'J':48,'K':48,'L':26}
for k,v in widths.items(): ws.column_dimensions[k].width=v
ws.freeze_panes='C5'
wb.move_sheet('Info',offset=3)
wb.active=0
wb.save(OUT)
print('saved',OUT,'last',last.date(),'below',[b[0] for b in below],'watch',[w[0] for w in watch],'missing',missing)
