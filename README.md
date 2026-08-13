# Dashboard Omset MFlash

Dashboard Streamlit dengan 2 tampilan:

- **Tab Ringkasan** — kartu KPI berwarna untuk Omset All, Omset Service, Omset Gadget & Aksesoris, Omset Marketing Corporate, masing-masing dengan badge **% pencapaian terhadap target**, + tren bulanan + tabel omset per cabang, bisa difilter Tahun / Bulan / Cabang.
- **Tab Scoreboard** — tabel per cabang gaya scoreboard berwarna (mengikuti contoh existing: header hijau/oranye/biru per grup kolom, sel hijau/kuning/merah sesuai pencapaian) untuk 4 kategori: **Omset All, Omset Service, Omset Gadget & Aksesoris, Omset Marketing Corporate** — lengkap dengan Target, Expected Value, % Pencapaian, Gap, Kejar Target Per Hari, dan rata-rata omset bulan lalu vs bulan ini.

Urutan cabang di semua tabel/grafik mengikuti urutan dashboard existing (Klender, Ceger, Bintara, Radjiman, ... Cibubur), bukan abjad.

Semua data diisi lewat tombol upload di sidebar — tidak perlu edit apapun di kode / repo GitHub.

## Cara menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Sumber data

1. **File data cabang** (wajib, bisa banyak file sekaligus — sampai 50 file): tiap file berisi data 1 cabang, dengan kolom `TGL FAKTUR`, `KATEGORI BARANG`, `TOTAL HARGA` (nama sheet boleh apa saja, yang penting kolomnya ada).
   - Kolom `CABANG` sebaiknya ada. Kalau tidak ada, nama cabang ditebak dari nama file (mis. `FP Klender.xlsx` atau `Klender - Agustus.xlsx` -> `KLENDER`).
   - Upload ulang file dengan nama sama = update data file itu (bukan duplikat). Nama file baru = ditambahkan ke dataset.
   - Dari sini dihitung: **Omset All** (semua baris transaksi), **Omset Service** (baris dengan `KATEGORI BARANG` = "Jasa" atau "Sparepart"), **Omset Gadget & Aksesoris** (Omset All dikurangi Omset Service). Sudah dicek cocok persis dengan angka Scoreboard Service di file aslinya.

2. **File Data Marketing Corporate** (opsional, terpisah, dipakai di tab Ringkasan): kolom `Tahun`, `Bulan`, `Cabang`, `Omset`. Template ada di sidebar.

3. **Target & Expected Value — otomatis kalau memungkinkan.** Kalau file yang diupload di poin 1 punya sheet **"Scoreboard"** (seperti file master "Dashboard 6 Pilar MFlash" yang biasa dipakai), Target (OMSET SAMURAI) & periodenya otomatis dibaca dari situ — **tidak perlu input manual**. Kalau tidak ada sheet itu, upload **File Data Target** terpisah (opsional): kolom `Cabang`, `Periode Mulai`, `Periode Selesai`, `Target Service`, `Target Gadget & Aksesoris`, `Target Marketing Corporate`. Target Omset All dihitung otomatis (Service + Gadget & Aksesoris). Template ada di sidebar. File Target manual ini selalu diprioritaskan kalau ada.
   - Kalau target berganti tiap periode (mis. tiap 3 bulan), tambahkan baris baru per periode — sistem otomatis pakai baris yang periodenya mencakup "Tanggal Acuan".

## Cara baca kolom Scoreboard

Mengikuti gaya scoreboard existing:

- **OMSET SAMURAI** = target untuk periode berjalan (dari file Target).
- **OMSET HARIAN (DR TARGET)** = target ÷ jumlah hari dalam periode.
- **EXPECTED VALUE** = omset harian target × jumlah hari yang sudah berjalan sampai Tanggal Acuan.
- **HARI INI** = omset di tanggal acuan.
- **S/D HARI INI** = omset kumulatif dari awal periode sampai tanggal acuan.
- **% PENCAPAIAN** = S/D Hari Ini ÷ Expected Value.
- **GAP VS EXPECTED** = S/D Hari Ini − Expected Value (hijau = di atas ekspektasi, merah = di bawah).
- **TOTAL GAP SAMURAI** = Target − S/D Hari Ini (sisa yang harus dikejar; hijau kalau sudah tercapai/negatif).
- **KEJAR TARGET PERHARI** = Total Gap Samurai ÷ sisa hari periode.
- **PERIODE BULAN LALU / PERIODE BULAN INI** = rata-rata omset per hari bulan kalender lalu vs bulan berjalan (month-to-date), independen dari data Target.
- **GAP** = Periode Bulan Ini − Periode Bulan Lalu.

Tanpa file Target, kolom Target/Expected Value/%Pencapaian/Gap/Kejar Target akan kosong ("-"), tapi kolom lain (Hari Ini, S/D Hari Ini, Periode Bulan Lalu/Ini, Gap rata-rata) tetap terisi dari data transaksi.

Khusus tabel **Scoreboard Omset Marketing Corporate**: karena datanya bulanan (bukan harian), kolom **HARI INI** selalu "-", dan **S/D HARI INI** dihitung dari jumlah bulan (bukan hari) yang sudah masuk dalam periode target.

**Warna:** hijau = sudah tercapai / di atas ekspektasi, kuning = mendekati (80-100% dari ekspektasi), merah = di bawah ekspektasi / masih ada gap — sama seperti skema warna di file scoreboard yang sudah ada.

File yang diupload otomatis tersimpan di folder `data/` dan dipakai lagi setiap dashboard dibuka, sampai diganti/dihapus lewat sidebar.

## Deploy

Bisa langsung dipush ke GitHub lalu di-deploy ke Streamlit Community Cloud seperti biasa — upload data dilakukan lewat aplikasi (sidebar), bukan lewat commit ke repo.
