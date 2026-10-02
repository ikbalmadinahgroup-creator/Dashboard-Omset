"""Masukkan file walk-in cabang (Rincian Pengiriman Pesanan + Rincian Faktur Penjualan
format walk-in) ke data/walkin. File lama cabang yang sama DIGANTI hanya jika datanya
berada di KUARTAL yang sama dengan file baru (export kumulatif per kuartal); file
kuartal sebelumnya tetap disimpan sebagai riwayat.
Usage: python3 tools/walkin_update.py <file1.xlsx> <file2.xlsx> ...  (jalankan dari root repo)"""
import sys,os,re,glob,shutil,subprocess,warnings
import pandas as pd
warnings.filterwarnings('ignore')
def quarter_of(path):
    d=pd.read_excel(path)
    col='TGL PENGIRIMAN' if 'TGL PENGIRIMAN' in d.columns else 'TGL FAKTUR'
    t=pd.to_datetime(d[col],errors='coerce').dropna()
    if t.empty: return None
    m=t.max(); return (m.year,(m.month-1)//3+1)
for f in sys.argv[1:]:
    base=re.sub(r'^[0-9a-f]{8}-','',os.path.basename(f))
    kind='pengiriman' if 'pengiriman_pesanan' in base else 'faktur'
    code=re.search(r'_(\d{3})mflash([a-z]+)',base).group(0)
    q=quarter_of(f)
    for old in glob.glob(f'data/walkin/rincian_{"pengiriman_pesanan" if kind=="pengiriman" else "faktur_penjualan"}{code}*'):
        if os.path.basename(old)==base: continue
        if quarter_of(old)==q:
            subprocess.run(['git','rm','-q',old]);
            if os.path.exists(old): os.remove(old)
            print('ganti', os.path.basename(old))
        else:
            print('simpan (kuartal lain)', os.path.basename(old))
    shutil.copy(f,'data/walkin/'+base); print('tambah', base, q)
