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

- **File data cabang** (wajib, bisa banyak file sekaligus — sampai 50 file): tiap file berisi data 1 cabang, dengan kolom `TGL FAKTUR`, `KATEGORI PENJUALAN`, `TOTAL HARGA` (sama seperti sheet `Faktur Penjualan` di file "Dashboard 6 Pilar MFlash" yang sudah dipakai selama ini — nama sheet-nya boleh apa saja, yang penting kolomnya ada).
  - Kolom `CABANG` sebaiknya ada di tiap file. Kalau tidak ada, nama cabang otomatis ditebak dari nama file (mis. `FP Klender.xlsx` atau `Klender - Agustus.xlsx` -> `KLENDER`).
  - Upload ulang file dengan nama yang sama akan menggantikan (update) data file itu, bukan menambah duplikat. Upload file baru dengan nama berbeda akan ditambahkan ke dataset.
  - Semua file digabung otomatis, lalu dari situ dihitung:
    - **Omset All** = total semua transaksi
    - **Omset Service** = transaksi dengan `KATEGORI PENJUALAN` diawali "SERVICE"
    - **Omset Gadget & Aksesoris** = transaksi dengan `KATEGORI PENJUALAN` diawali "PENJUALAN"
  - Daftar file yang tersimpan & tombol hapus per-file / hapus semua ada di sidebar.

- **File Marketing Corporate** (opsional, terpisah): karena data corporate di file master tidak punya tabel tanggal+cabang yang rapi (hanya pivot per nama sales), dashboard ini pakai file kecil terpisah dengan kolom `Tahun`, `Bulan`, `Cabang`, `Omset` — satu baris per kombinasi bulan+cabang. Download template-nya lewat tombol di sidebar, isi, lalu upload.

File yang diupload otomatis tersimpan di folder `data/` dan dipakai lagi setiap dashboard dibuka, sampai diganti/dihapus.

## Deploy

Bisa langsung dipush ke GitHub lalu di-deploy ke Streamlit Community Cloud seperti biasa — upload data dilakukan lewat aplikasi (sidebar), bukan lewat commit ke repo.
