# Dashboard Omset MFlash

Dashboard Streamlit untuk memantau omset 18 cabang MFlash: **Omset All**, **Omset Service**,
**Omset Gadget & Aksesoris**, **Omset Marketing Corporate**, dan performa **Iklan Meta Ads**.
Semua data diisi lewat tombol upload di sidebar — tidak perlu edit apapun di GitHub/kode.
Logo MFlash sudah ter-embed langsung di dalam `app.py` (base64), jadi selalu tampil tanpa
tergantung file gambar terpisah.

## Cara menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cara pakai

1. **Upload file data cabang** di sidebar (bisa banyak file sekaligus, sampai 50 file). Kalau file
   yang diupload adalah file master (punya sheet `Faktur Penjualan` + `Scoreboard`, seperti file yang
   biasa dipakai sekarang), Target, Expected Value, dan Scoreboard Marketing Corporate (per nama sales)
   **otomatis terbaca** — tidak perlu isi apa-apa lagi.
2. Kalau file yang diupload **tidak** punya sheet `Scoreboard`, download template **Data Target**
   dan/atau **Data Marketing Corporate** di sidebar, isi, lalu upload — ini jadi fallback manual.
3. File yang sudah diupload tersimpan otomatis dan tetap terpakai tiap kali dashboard dibuka lagi.
   Upload ulang dengan nama file sama = update, upload dengan nama beda = ditambahkan.
4. **Setiap kali Anda upload file baru, itu otomatis tercatat sebagai satu titik histori** (lihat
   "Riwayat Pencapaian Harian" di bawah) — jadi disarankan upload file terbaru secara rutin (misalnya
   tiap hari) supaya histori pencapaian harian makin lengkap.

## Tab Ringkasan

- 4 kartu KPI (Omset All / Service / Gadget & Aksesoris / Marketing Corporate) dengan badge
  % pencapaian terhadap ekspektasi pace saat ini.
- **Progress Pencapaian**: 3 lingkaran (ring) % pencapaian untuk Omset All, Service, dan
  Gadget & Aksesoris. Di bawah tiap ring ditampilkan juga persentase dari total target periode
  yang sudah tercapai.
- **🧩 Kontribusi terhadap Omset All**: donut chart komposisi Service / Penjualan Retail /
  Marketing Corporate, dengan label persentase di tiap slice + rincian angka di sampingnya.
  Marketing Corporate = penjumlahan S/D Hari Ini semua nama sales; Penjualan Retail = Omset
  Gadget & Aksesoris dikurangi total pencapaian Marketing Corporate (supaya tidak dobel hitung).
- Tren omset per bulan & omset per cabang (grafik dengan angka di atas tiap bar + tabel dengan
  gradasi warna), difilter oleh Tahun / Bulan / Cabang di sidebar.

## Tab Scoreboard

Tabel gaya scoreboard per cabang untuk Omset All, Service, dan Gadget & Aksesoris, dengan kolom:

| Kelompok | Kolom |
|---|---|
| Detail Target | Omset Samurai (Target), Omset Harian, Expected Value |
| Detail Pencapaian | Hari Ini, S/D Hari Ini, % Pencapaian |
| Detail Gap | Gap vs Expected, Total Gap Samurai, Kejar Target Perhari |
| Monitoring Rata-rata/Hari | Periode Bulan Lalu, Periode Bulan Ini, Gap |

Setiap tabel juga punya expander **Progress Harian** — grafik garis omset kumulatif aktual vs
target pace lurus (garis putus-putus), dengan garis vertikal menandai tanggal acuan.

### 📅 Riwayat Pencapaian Harian

Chart ini dibangun dari **ledger permanen** (`data/history_log.csv`), bukan dari seluruh tanggal
transaksi di file. Setiap kali Anda upload file, aplikasi mencatat SATU snapshot per cabang,
ditandai dengan tanggal yang sama untuk semua cabang di file itu (diambil dari sel "TANGGAL" di
sheet Scoreboard). Konsekuensinya:

- Histori hanya menampilkan hari-hari yang **benar-benar pernah diupload** (bar berwarna), bukan
  seluruh rentang tanggal transaksi yang kebetulan ada di dalam file — jadi kalau Anda upload file
  di tanggal 12 tapi datanya berisi riwayat transaksi dari tanggal 1, chart tidak akan menampilkan
  tanggal 1 sebagai titik histori terpisah.
- Histori ini **tetap ada** walau file lama sudah diganti/dihapus, karena tersimpan permanen di
  `data/history_log.csv`, terpisah dari file Excel yang diupload.
- Filter tersedia: Kategori, Tahun, Bulan, Cabang, dan rentang tanggal manual.
- Warna bar: 🟢 ≥100% · 🟡 85-99,9% · 🔴 <85% · abu-abu = belum ada upload di tanggal itu.

**Scoreboard Marketing Corporate** ditampilkan **per nama sales** (bukan per cabang), diambil apa
adanya (snapshot) dari sheet Scoreboard file yang diupload — sheet ini memang tidak punya kolom Hari
Ini untuk section per-sales ini. Kolom **Hari Ini** dihitung dari selisih S/D Hari Ini dibanding
upload sebelumnya begitu tersedia (minimal 2x upload di tanggal berbeda). Sebelum itu (upload
pertama, belum ada pembanding), dipakai **estimasi** dari kolom "Periode Bulan Ini" (rata-rata omset
harian bulan berjalan, sudah ada di sheet Excel) supaya kolomnya tidak kosong — ditandai tanda
bintang (*) dengan keterangan di bawah tabel. Kalau sheet Scoreboard tidak ada, otomatis dipakai
data manual bulanan (fallback, kolom Hari Ini juga "-" karena datanya bulanan).

**Tanggal Acuan** ("dianggap sebagai Hari Ini") di sidebar defaultnya adalah tanggal transaksi
terakhir di data yang diupload, supaya kolom "Hari Ini" tidak pernah kosong. Bisa diganti manual
untuk melihat posisi di tanggal lain.

**Cara baca warna** (ring, badge KPI, cell % Pencapaian, dan bar Riwayat Harian semua pakai standar
yang sama): 🟢 hijau = ≥100% (target tercapai), 🟡 kuning = 85%-99,9% (mendekati), 🔴 merah = <85%
(masih ada gap).

### 💡 Insight & Rekomendasi Perbaikan

Di bagian bawah tab ini, dihitung otomatis untuk Omset Service, Penjualan (Gadget & Aksesoris), dan
Marketing Corporate — membandingkan rata-rata omset harian bulan ini vs bulan lalu per cabang/sales,
plus % Pencapaian terhadap Expected Value:

- 🚨 Turun >20% dibanding bulan lalu, atau pencapaian <70% → perlu evaluasi segera.
- ⚠️ Turun 5-20% dibanding bulan lalu, atau pencapaian 70-84,9% → pantau & tambah aktivasi.

Setiap insight dilengkapi **rencana aksi konkret** yang beda-beda per kategori (bukan saran generik
yang sama untuk semua cabang), dipecah jadi 🌐 **Online** dan 🏬 **Offline**:

- **Omset Service**: online = refresh promo servis di IG/FB, iklan Meta Ads Click-to-WhatsApp radius
  cabang, broadcast reminder servis ke database lama; offline = cek sparepart & jadwal teknisi,
  banner promo servis, cek gratis untuk tarik walk-in, upsell servis tambahan oleh CS.
- **Penjualan Gadget & Aksesoris**: online = posting katalog & harga terbaru, iklan Meta Ads produk
  best-seller, broadcast flash sale; offline = cek stok best-seller, display etalase, sebar
  brosur ke area sekitar, cross-sell aksesoris oleh tim sales.
- **Marketing Corporate**: online = follow-up leads yang belum closing, update portofolio di
  LinkedIn, siapkan penawaran bundling/kontrak volume; offline = kunjungan langsung ke calon
  klien, harga khusus untuk pembelian volume, ikut event/komunitas bisnis lokal.

### ✅ To-Do List Evaluasi Cabang

Checklist interaktif di bawah insight — setiap item bisa dicentang sebagai penanda "sudah
ditindaklanjuti". Daftarnya diturunkan otomatis dari insight di atas (item level 🚨/⚠️ saja, item
✅ yang sudah bagus tidak masuk to-do), diurutkan dari paling mendesak, maksimal 25 item. Status
centang tersimpan selama sesi browser masih terbuka (reset kalau halaman di-refresh total atau
data baru diupload).

## Tab Iklan (Meta Ads)

Upload file export **Campaigns** dari Meta Ads Manager (Ads Manager > Export > Excel) lewat sidebar
("📣 Data Iklan Meta Ads" — bisa banyak file sekaligus, sampai 50 file, mis. per periode). Cabang
ditebak otomatis dari nama campaign (contoh: `12 Cinere - MATOT !!!` → CINERE, `10 Warung Bongkok -
L - MATOT !!!` → WARBONG). Campaign yang namanya tidak diawali kode cabang (mis. campaign umum/brand)
dikelompokkan sebagai **LAINNYA**.

Kalau file Meta Ads ini ke-upload di tempat yang salah (misalnya di uploader "data cabang" alih-alih
"Data Iklan Meta Ads"), aplikasi otomatis mendeteksi dan melewatinya secara diam-diam (tidak muncul
sebagai warning) — cukup upload ulang di uploader yang benar.

Isi tab ini:

- KPI: Total Spend, total Messaging Conversations Started, rata-rata Cost per Messaging Conversation
  (tertimbang total spend ÷ total messaging conversation), jumlah campaign berstatus Active.
- Grafik Cost per Messaging Conversation dan jumlah Messaging Conversation per cabang (dengan angka
  di atas tiap bar).
- Tabel performa per cabang & detail per campaign, diurutkan dari Cost per Messaging paling efisien.
- **Insight & Rekomendasi Improvement** otomatis, dihitung dari data yang diupload:
  - 🚨 Campaign yang sudah spend signifikan tapi 0 Messaging Conversation → disarankan evaluasi/pause.
  - ⚠️ Campaign dengan Cost per Messaging > 1,5x rata-rata semua campaign → disarankan ganti
    creative/audience atau turunkan budget.
  - ✅ Campaign dengan Cost per Messaging < 0,7x rata-rata (dan volume cukup) → kandidat scale up budget.
  - ⚠️ Campaign berstatus Inactive tapi masih tercatat sisa spend kecil → disarankan cek ulang di
    Ads Manager untuk memastikan benar-benar berhenti.
  - **Diagnosa dari sisi konten**: setiap rekomendasi di atas juga membandingkan CTR campaign dengan
    CTR median semua campaign. Kalau CTR di bawah rata-rata → kemungkinan masalah ada di KONTEN
    (thumbnail/video 3 detik pertama, judul, CTA kurang jelas). Kalau CTR sudah bagus tapi Cost per
    Messaging masih tinggi → kemungkinan bukan di kontennya, tapi di funnel setelah klik (kecepatan
    respon admin, auto-reply WhatsApp/Messenger).

## Aturan perhitungan (penting)

- **Omset Service** = jumlah baris transaksi dengan kolom `KATEGORI BARANG` = "Jasa" atau "Sparepart".
- **Omset Gadget & Aksesoris** = Omset All − Omset Service (semua baris selain Jasa/Sparepart).
- **Omset All** = seluruh baris transaksi (Service + Gadget & Aksesoris).
- Urutan cabang mengikuti urutan dashboard asli (bukan abjad): Klender, Ceger, Bintara, Radjiman,
  Jatimulya, Dramaga, Condet, Jatibening, Sawangan, Warbong, Cinere, Cibinong, Karawang,
  Jatiwaringin, Cikampek, Cilangkap, Pejaten, Cibubur.

## Struktur file

```
app.py                  # Aplikasi Streamlit utama (logo sudah ter-embed sebagai base64 di dalamnya)
requirements.txt         # Dependensi Python
data/                    # Dibuat otomatis saat pertama upload (diabaikan git)
  main/                    # File-file data cabang (sampai 50 file)
  ads/                     # File export Meta Ads (sampai 50 file)
  corporate_data.xlsx      # Data Marketing Corporate manual (fallback, opsional)
  target_data.xlsx         # Data Target manual (fallback, opsional)
  history_log.csv          # Ledger permanen omset harian per cabang (dari tiap upload)
  history_log_corp.csv     # Ledger permanen S/D Hari Ini per sales Marketing Corporate
```
