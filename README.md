# Dashboard Omset MFlash

Dashboard Streamlit untuk monitoring Omset, Iklan, dan Walk-in 18 cabang MFlash.

## Cara Menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Kalau Sidebar "Kelola Data" (Upload File) Tidak Muncul

Penyebab paling umum: CSS kustom yang menyembunyikan menu/toolbar bawaan Streamlit ikut menyembunyikan tombol panah buka/tutup sidebar. Versi app.py ini sudah diperbaiki — hanya `#MainMenu` dan `footer` yang disembunyikan, toolbar/header **tidak** disentuh sama sekali, dan sidebar diset default terbuka (`initial_sidebar_state="expanded"`).

Kalau setelah redeploy versi ini sidebar masih belum muncul:
1. Coba **hard refresh** browser (Ctrl+Shift+R / Cmd+Shift+R) — browser mungkin masih menyimpan cache halaman versi lama.
2. Cek di pojok kiri atas halaman, ada panah kecil **"»"** — klik untuk membuka sidebar.
3. Pastikan file `app.py` yang di-deploy di Streamlit Cloud memang sudah file yang terbaru (cek tab "Manage app" → lihat log deploy, pastikan tidak ada error saat build).
4. Kalau halaman terlihat "macet"/tidak merespons klik sama sekali, coba buka di jendela browser baru (mode incognito) untuk memastikan bukan masalah cache/extension browser.

## Struktur Tab

1. **📊 Ringkasan** — KPI card Omset Service / Gadget & Aksesoris / All, kontribusi per kategori (pie chart), progress ring % pencapaian, grafik progres harian (aktual vs target pace lurus), dan riwayat pencapaian harian dari ledger upload.
2. **🎯 Scoreboard** — Tabel scoreboard berwarna (Service, Gadget & Aksesoris, All, Marketing Corporate) lengkap dengan target, expected value, gap, dan kejar target per hari. Di bawahnya ada **Insight & Rekomendasi Perbaikan** dan **To Do List**, keduanya dikelompokkan rapi per kategori (Omset Service / Penjualan Gadget & Aksesoris / Marketing Corporate) dengan rencana aksi 🌐 Online dan 🏬 Offline terpisah supaya mudah dibaca.
3. **📢 Iklan** — Ringkasan spend, Messaging Conversation, cost per conversation per campaign, plus insight otomatis (campaign boros, campaign mahal, campaign efisien untuk di-scale-up, evaluasi konten vs funnel).
4. **🚶 Walk-in** — Total walk-in & rata-rata per hari per cabang per bulan, plus insight otomatis untuk cabang yang trafiknya turun signifikan.

## Data Omset Utama — Dua Format Didukung

Dashboard secara otomatis mendeteksi format file yang diupload di slot "1️⃣ Data Omset Utama", tidak perlu memilih format secara manual:

- **Format lama (file master)** — satu file berisi sheet `Faktur Penjualan` (data transaksi seluruh cabang) dan sheet `Scoreboard` (untuk auto-ekstrak Target & Scoreboard Marketing Corporate). Cocok untuk upload satu file besar.
- **Format baru (per cabang)** — file export "Rincian Faktur Penjualan", satu file per cabang. Cabang diambil otomatis dari nama file. Bisa upload sampai 18 file (atau lebih) sekaligus. Format ini **tidak** punya sheet Scoreboard, sehingga:
  - Target harus diisi manual lewat template Target di sidebar (opsional, hanya jika ingin pencapaian vs target).
  - Data Marketing Corporate harus diisi manual lewat template Corporate di sidebar (opsional).

Kedua format bisa digabung — misalnya beberapa cabang masih pakai format lama dan cabang lain sudah pakai format baru, dashboard akan menggabungkan datanya secara otomatis.

**Aturan perhitungan (berlaku untuk kedua format):**
- **Omset Service** = jumlah `TOTAL HARGA` dari baris dengan `KATEGORI BARANG` = JASA atau SPAREPART.
- **Omset Gadget & Aksesoris** = jumlah `TOTAL HARGA` dari baris dengan kategori barang selain JASA/SPAREPART (HP, laptop, aksesoris, kartu perdana, parfum, dll).
- **Omset All** = Omset Service + Omset Gadget & Aksesoris.

## Export Laporan untuk Presentasi (PPTX / PDF)

Di bagian bawah dashboard ada tombol **📤 Export Laporan untuk Presentasi** yang menghasilkan file PowerPoint (.pptx) atau PDF siap dipresentasikan ke CEO, berisi untuk masing-masing kategori (Omset, Iklan, Walk-in):
- **Penyajian Data** — ringkasan KPI dalam bentuk tabel/slide.
- **Evaluasi** — daftar temuan/masalah dari insight engine.
- **Rencana Perbaikan** — rencana aksi 🌐 Online dan 🏬 Offline.

Laporan mengikuti filter Tanggal Acuan & Cabang yang sedang aktif di dashboard.

## ⚠️ GitHub Auto-Backup (supaya data tidak hilang saat app sleep/restart)

Streamlit Community Cloud menjalankan app di container yang disk-nya **sementara** — begitu app di-sleep (tidak dipakai lama) atau di-restart/redeploy, semua file yang diupload lewat dashboard (data omset, iklan, walk-in, target, corporate, ledger histori) akan **hilang** karena hanya tersimpan di disk lokal container, bukan di repo Git.

Untuk mencegah ini, dashboard mendukung auto-backup setiap file yang diupload langsung ke sebuah repo GitHub lewat GitHub API. Setiap kali app start, dashboard otomatis menarik kembali file-file itu dari GitHub sebelum menampilkan data.

### Cara setup:

1. Buat repo GitHub baru (boleh **private**) khusus untuk menyimpan data dashboard ini, misal `mflash-dashboard-data`.
2. Buat **Personal Access Token (PAT)** di GitHub: Settings → Developer settings → Personal access tokens → Fine-grained tokens (atau classic token dengan scope `repo`). Beri akses ke repo yang dibuat di langkah 1.
3. Di Streamlit Community Cloud, buka pengaturan app → **Secrets**, lalu tambahkan:

   ```toml
   [github]
   token = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
   repo = "username-anda/mflash-dashboard-data"
   branch = "main"
   ```

4. Save & reboot app. Kalau secrets sudah benar, sidebar akan menampilkan "☁️ Auto-backup ke GitHub: **aktif**".

### Catatan:
- Ukuran file per-upload dibatasi ~90MB oleh GitHub API (biasanya cukup untuk file Excel omset).
- Data yang di-backup: semua file di folder `data/main`, `data/ads`, `data/walkin`, file target/corporate manual, dan ledger histori (`history_log.csv`, `corp_history_log.csv`).
- Kalau backup gagal (misal token salah/kadaluarsa), dashboard tetap jalan seperti biasa tapi akan menampilkan peringatan di sidebar — data tetap berisiko hilang saat app sleep sampai backup diperbaiki.
- Fitur ini **opsional** — kalau secrets tidak diisi, dashboard tetap berfungsi normal tanpa backup (sama seperti sebelumnya).

## Catatan Lain

- "Zzzz - This app has gone to sleep" adalah perilaku normal hosting gratis Streamlit Community Cloud setelah tidak diakses beberapa waktu — cukup klik "Yes, get this app back up!", bukan bug.
- Ledger histori (`history_log.csv`, `corp_history_log.csv`) mencatat snapshot omset setiap kali file diupload, supaya grafik "Riwayat Pencapaian Harian" tetap ada meski file lama sudah diganti/dihapus.
