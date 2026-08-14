# Dashboard Omset MFlash

Dashboard Streamlit untuk memantau omset & performa 18 cabang MFlash: Omset All, Omset
Service, Omset Gadget & Aksesoris, Marketing Corporate, performa Iklan Meta Ads, dan
Walk-in (jumlah order servis) per cabang — lengkap dengan export laporan presentasi.

## Cara pakai

1. Install dependency:
   ```bash
   pip install -r requirements.txt
   ```
2. Jalankan:
   ```bash
   streamlit run app.py
   ```
3. Upload data lewat sidebar:
   - **Upload file data cabang** — file per-cabang, atau langsung file master yang punya
     sheet `Faktur Penjualan` + `Scoreboard` (target & scoreboard Marketing Corporate
     otomatis kebaca dari situ). Bisa banyak file sekaligus, sampai 50 file.
   - **Data Marketing Corporate (manual)** dan **Data Target (manual)** — opsional, hanya
     diperlukan kalau file di atas tidak punya sheet `Scoreboard`. Template tersedia di sidebar.
   - **Data Iklan Meta Ads** — hasil export "Campaigns" dari Meta Ads Manager (Export >
     Excel). Sampai 50 file.
   - **Data Walk-in (Pengiriman Pesanan)** — file "Rincian Pengiriman Pesanan" per cabang
     dari sistem service. Sampai 60 file (1 file = 1 cabang).

Semua file yang sudah diupload disimpan dan otomatis dipakai lagi setiap dashboard dibuka.
Upload ulang dengan nama file sama = update, bukan duplikat.

## Tab Ringkasan

- Kartu KPI (Omset All / Service / Gadget & Aksesoris / Marketing Corporate) dengan badge
  % pencapaian.
- Progress ring % pencapaian per kategori (🟢 ≥100% · 🟡 85-99,9% · 🔴 <85%).
- **Kontribusi terhadap Omset All** — donut chart Service / Penjualan Retail / Marketing
  Corporate.
- Tren omset bulanan & omset per cabang, dengan filter Tahun/Bulan/Cabang di sidebar.

## Tab Scoreboard

- Tabel scoreboard per cabang (Omset All / Service / Gadget & Aksesoris / Marketing
  Corporate) gaya scoreboard: Detail Target, Detail Pencapaian, Detail Gap, Monitoring
  Progress Rata-rata/Hari.
- **Riwayat Pencapaian Harian** — dibangun dari ledger permanen, mencatat satu snapshot
  per cabang setiap kali file diupload, jadi histori tidak hilang walau harinya sudah lewat.
- **Insight & Rekomendasi Perbaikan** — dikelompokkan per kategori (Omset Service /
  Penjualan Gadget & Aksesoris / Marketing Corporate). Tiap kartu menampilkan pernyataan
  masalah di atas, lalu rencana aksi **🌐 Online** dan **🏬 Offline** sebagai bullet list
  berdampingan supaya mudah dibaca (bukan satu paragraf panjang).
- **To-Do List Evaluasi Cabang** — checklist interaktif, dikelompokkan per kategori. Buka
  "Lihat rencana aksi" pada tiap item untuk detail Online/Offline-nya.
- Marketing Corporate **Hari Ini** dihitung dari selisih S/D Hari Ini vs snapshot upload
  sebelumnya (ledger delta), dengan estimasi Periode Bulan Ini (ditandai `*`) pada upload
  pertama.

## Tab Iklan (Meta Ads)

- KPI Total Spend, Messaging Conversations, Rata-rata Cost per Messaging, Campaign Aktif.
- Performa per cabang (bar chart & tabel) dan detail per campaign.
- **Insight & Rekomendasi Improvement** — diagnosis KONTEN (creative/hook/CTA) vs FUNNEL
  (kecepatan respon admin) berdasarkan perbandingan CTR terhadap median.

## Tab Walk-in Cabang

- Upload file **Rincian Pengiriman Pesanan** per cabang (sampai 60 file).
- **Total Walk-in** per bulan per cabang = jumlah **NOMOR PENGIRIMAN PESANAN yang unik**.
- **Rata-rata Walk-in per Hari** = Total Walk-in ÷ Hari Efektif.
- Filter Cabang / Tahun / Bulan, KPI, bar chart, tabel detail, dan insight penurunan
  walk-in bulan-ke-bulan per cabang.

## Export Laporan untuk Presentasi (PPTX & PDF)

Di bagian paling bawah dashboard (di luar tab, jadi selalu tampil) ada dua tombol:
**📊 Download PowerPoint (.pptx)** dan **📄 Download PDF (.pdf)**. Laporan ini otomatis
dibangun dari data yang sedang diupload dan berisi, untuk masing-masing dari 3 domain
(Omset & Scoreboard, Iklan Meta Ads, Walk-in Cabang):

- **Penyajian Data** — tabel ringkasan KPI per cabang + grafik batang.
- **Evaluasi** — daftar temuan (penurunan omset/walk-in, campaign iklan kurang efisien, dsb).
- **Rencana Perbaikan** — rencana aksi Online & Offline per kategori (dirangkum sekali per
  kategori, tidak diulang per cabang) untuk siap dipresentasikan ke CEO.

Laporan dibuat ulang otomatis setiap dashboard dibuka/data berubah, jadi selalu sinkron
dengan angka yang tampil di layar.

## Tampilan & Branding

- Ikon aplikasi (favicon/tab browser) memakai logo MFlash, bukan ikon bawaan Streamlit.
- Menu hamburger, tombol "Deploy", dan footer "Made with Streamlit" disembunyikan supaya
  tampilan bersih untuk dipakai presentasi ke pihak luar.

## Aturan perhitungan

- **Omset Service** = baris transaksi dengan `KATEGORI BARANG` = "Jasa" atau "Sparepart".
- **Omset Gadget & Aksesoris** = Omset All dikurangi Omset Service.
- **% Pencapaian** = S/D Hari Ini ÷ Expected Value (target harian × hari berjalan periode).
- **Total Walk-in** = jumlah `NOMOR PENGIRIMAN PESANAN` unik per cabang per bulan.
- Warna threshold konsisten di semua ring/badge/tabel: 🟢 ≥100% · 🟡 85-99,9% · 🔴 <85%.

## Struktur file

```
dashboard_mflash/
├── app.py                  # aplikasi utama
├── requirements.txt
├── README.md
├── .gitignore
├── assets/
│   ├── logo_small.png      # logo MFlash (juga di-embed sebagai base64 di app.py)
│   └── logo_b64.txt
└── data/                   # dibuat otomatis, isi file upload & ledger (tidak di-commit)
    ├── main/                    # file data cabang
    ├── ads/                     # file export Meta Ads
    ├── walkin/                  # file Rincian Pengiriman Pesanan
    ├── corporate_data.xlsx      # data Marketing Corporate manual (opsional)
    ├── target_data.xlsx         # data Target manual (opsional)
    ├── history_log.csv          # ledger riwayat omset harian
    └── history_log_corp.csv     # ledger riwayat S/D Hari Ini Marketing Corporate
```
