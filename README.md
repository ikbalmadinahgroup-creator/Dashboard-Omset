# Dashboard Omset MFlash

Dashboard Streamlit untuk memantau omset & performa 18 cabang MFlash: Omset All, Omset
Service, Omset Gadget & Aksesoris, Marketing Corporate, performa Iklan Meta Ads, dan
Walk-in (jumlah order servis) per cabang.

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
  Corporate, dengan Corporate = jumlah S/D Hari Ini semua nama sales, dan Retail = total
  Gadget & Aksesoris dikurangi total Corporate.
- Tren omset bulanan & omset per cabang, dengan filter Tahun/Bulan/Cabang di sidebar.

## Tab Scoreboard

- Tabel scoreboard per cabang (Omset All / Service / Gadget & Aksesoris / Marketing
  Corporate) gaya scoreboard: Detail Target, Detail Pencapaian, Detail Gap, Monitoring
  Progress Rata-rata/Hari.
- **Riwayat Pencapaian Harian** — dibangun dari ledger permanen (`data/history_log.csv`,
  `data/history_log_corp.csv`) yang mencatat satu snapshot per cabang setiap kali file
  diupload (ditandai tanggal dari sheet Scoreboard). Jadi histori tidak hilang walau
  harinya sudah lewat, dan hanya menampilkan tanggal yang benar-benar pernah diupload
  (bar abu-abu = belum ada upload di tanggal itu).
- **Insight & Rekomendasi Perbaikan** — otomatis mendeteksi penurunan rata-rata omset
  harian & pencapaian di bawah 85%, dengan rencana aksi konkret per kategori (🌐 Online /
  🏬 Offline), bukan saran generik.
- **To-Do List Evaluasi Cabang** — checklist interaktif dari insight di atas, urut dari
  yang paling mendesak.
- Marketing Corporate **Hari Ini** dihitung dari selisih S/D Hari Ini vs snapshot upload
  sebelumnya (ledger delta). Kalau belum ada upload sebelumnya untuk dibandingkan, dipakai
  estimasi dari Periode Bulan Ini (ditandai `*`).

## Tab Iklan (Meta Ads)

- KPI Total Spend, Messaging Conversations, Rata-rata Cost per Messaging, Campaign Aktif.
- Performa per cabang (bar chart & tabel) dan detail per campaign.
- **Insight & Rekomendasi Improvement** — dibandingkan terhadap rata-rata Cost per
  Messaging tertimbang dan CTR median, dengan diagnosis: CTR di bawah rata-rata → evaluasi
  KONTEN (creative/hook/CTA); CTR sudah baik → evaluasi FUNNEL (kecepatan respon admin,
  auto-reply).

## Tab Walk-in Cabang

- Upload file **Rincian Pengiriman Pesanan** per cabang (sampai 60 file, 1 file = 1 cabang;
  nama cabang ditebak otomatis dari nama file).
- **Total Walk-in** per bulan per cabang = jumlah **NOMOR PENGIRIMAN PESANAN yang unik**
  (bukan jumlah baris) — satu nomor pengiriman bisa muncul di beberapa baris kalau
  order-nya berisi beberapa item, tetap dihitung sebagai satu walk-in.
- **Rata-rata Walk-in per Hari** = Total Walk-in ÷ Hari Efektif, di mana Hari Efektif untuk
  bulan yang sudah lewat penuh = jumlah hari kalender bulan itu, dan untuk bulan paling
  baru di data (bulan berjalan) = jumlah hari yang sudah tercatat sampai tanggal terakhir.
- Filter Cabang / Tahun / Bulan, KPI, bar chart Total Walk-in & Rata-rata per Hari, dan
  tabel detail per cabang per bulan.

## Aturan perhitungan

- **Omset Service** = baris transaksi dengan `KATEGORI BARANG` = "Jasa" atau "Sparepart".
- **Omset Gadget & Aksesoris** = Omset All dikurangi Omset Service (semua baris lainnya).
- **% Pencapaian** = S/D Hari Ini ÷ Expected Value (target harian × hari berjalan periode).
- **Total Walk-in** = jumlah `NOMOR PENGIRIMAN PESANAN` unik per cabang per bulan (dari
  kolom `TGL PENGIRIMAN`).
- Warna threshold dipakai konsisten di semua ring/badge/tabel: 🟢 ≥100% · 🟡 85-99,9% ·
  🔴 <85%.

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
