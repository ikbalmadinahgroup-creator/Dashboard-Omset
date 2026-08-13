# Dashboard Omset MFlash

Dashboard Streamlit untuk 4 KPI omset:

1. Omset All
2. Omset Service
3. Omset Gadget & Aksesoris
4. Omset Marketing Corporate

Filter: Tahun, Bulan, Cabang. Semua data diisi lewat tombol upload di sidebar — tidak perlu edit apapun di kode / repo GitHub.

## Cara menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Sumber data

- **File data utama** (wajib): file Excel yang punya sheet bernama `Faktur Penjualan` dengan kolom `CABANG`, `TGL FAKTUR`, `KATEGORI PENJUALAN`, `TOTAL HARGA` — sama seperti file "Dashboard 6 Pilar MFlash" yang sudah dipakai selama ini. Dari sini dihitung otomatis:
  - **Omset All** = total semua transaksi
  - **Omset Service** = transaksi dengan `KATEGORI PENJUALAN` diawali "SERVICE"
  - **Omset Gadget & Aksesoris** = transaksi dengan `KATEGORI PENJUALAN` diawali "PENJUALAN"

- **File Marketing Corporate** (opsional, terpisah): karena data corporate di file master tidak punya tabel tanggal+cabang yang rapi (hanya pivot per nama sales), dashboard ini pakai file kecil terpisah dengan kolom `Tahun`, `Bulan`, `Cabang`, `Omset` — satu baris per kombinasi bulan+cabang. Download template-nya lewat tombol di sidebar, isi, lalu upload.

Kedua file yang diupload otomatis tersimpan di folder `data/` dan dipakai lagi setiap dashboard dibuka, sampai diganti dengan upload baru.

## Deploy

Bisa langsung dipush ke GitHub lalu di-deploy ke Streamlit Community Cloud seperti biasa — upload data dilakukan lewat aplikasi (sidebar), bukan lewat commit ke repo.
