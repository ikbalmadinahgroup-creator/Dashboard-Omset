# Dashboard Omset MFlash

Dashboard Streamlit untuk memantau omset 18 cabang MFlash: **Omset All**, **Omset Service**,
**Omset Gadget & Aksesoris**, dan **Omset Marketing Corporate**. Semua data diisi lewat tombol
upload di sidebar — tidak perlu edit apapun di GitHub/kode.

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

## Tab Ringkasan

- 4 kartu KPI (Omset All / Service / Gadget & Aksesoris / Marketing Corporate) dengan badge
  % pencapaian terhadap ekspektasi pace saat ini.
- **Progress Pencapaian**: 3 lingkaran (ring) % pencapaian untuk Omset All, Service, dan
  Gadget & Aksesoris — mengikuti gaya "Progress Top Leader". Di bawah tiap ring ditampilkan juga
  persentase dari total target periode yang sudah tercapai.
- Tren omset per bulan & omset per cabang (grafik + tabel dengan gradasi warna), difilter oleh
  Tahun / Bulan / Cabang di sidebar.

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

**Scoreboard Marketing Corporate** ditampilkan **per nama sales** (bukan per cabang), diambil apa
adanya (snapshot) dari sheet Scoreboard file yang diupload — karena data ini bulanan/statis, bukan
transaksi harian, jadi tidak berubah walau Tanggal Acuan diganti. Kalau sheet Scoreboard tidak ada,
otomatis dipakai data manual bulanan (fallback).

**Tanggal Acuan** ("dianggap sebagai Hari Ini") di sidebar defaultnya adalah tanggal transaksi
terakhir di data yang diupload, supaya kolom "Hari Ini" tidak pernah kosong. Bisa diganti manual
untuk melihat posisi di tanggal lain.

**Cara baca warna:** hijau = sudah di atas ekspektasi/target tercapai, kuning = mendekati
(80-100% dari ekspektasi), merah = di bawah ekspektasi/masih ada gap.

## Aturan perhitungan (penting)

- **Omset Service** = jumlah baris transaksi dengan kolom `KATEGORI BARANG` = "Jasa" atau "Sparepart".
- **Omset Gadget & Aksesoris** = Omset All − Omset Service (semua baris selain Jasa/Sparepart).
- **Omset All** = seluruh baris transaksi (Service + Gadget & Aksesoris).
- Urutan cabang mengikuti urutan dashboard asli (bukan abjad): Klender, Ceger, Bintara, Radjiman,
  Jatimulya, Dramaga, Condet, Jatibening, Sawangan, Warbong, Cinere, Cibinong, Karawang,
  Jatiwaringin, Cikampek, Cilangkap, Pejaten, Cibubur.

## Struktur file

```
app.py               # Aplikasi Streamlit utama
requirements.txt      # Dependensi Python
data/                 # Dibuat otomatis saat pertama upload (diabaikan git)
  main/                # File-file data cabang (sampai 50 file)
  corporate_data.xlsx   # Data Marketing Corporate manual (fallback, opsional)
  target_data.xlsx      # Data Target manual (fallback, opsional)
```
