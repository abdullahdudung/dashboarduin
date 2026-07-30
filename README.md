# Dashboard E-Semesta UIN Jakarta — Streamlit

Dashboard terintegrasi UIN Syarif Hidayatullah Jakarta berbasis **Streamlit** untuk kebutuhan monitoring pimpinan, akademik, dan capaian IKU.

Versi saat ini mencakup dashboard eksekutif, akademik, IKU, serta kerangka dashboard lain yang masih dalam tahap pengembangan.

## Fitur utama

### 1. Dashboard Landing Page

Menampilkan ringkasan eksekutif:

- total mahasiswa aktif;
- mahasiswa berdasarkan tahun angkatan;
- jumlah peminat PMB;
- jumlah lulusan;
- yield rate mahasiswa baru;
- distribusi mahasiswa aktif per fakultas;
- komposisi mahasiswa aktif per jenjang;
- beberapa kartu indikator sampel sementara.

Filter yang tersedia:

- Tahun.

### 2. Dashboard Akademik

Dashboard Akademik memiliki tab:

- **Program Studi**
- **Mahasiswa**
- **PMB**
- **Kelulusan**
- **Kebutuhan Khusus**
- **Daerah Tertinggal**

Filter yang tersedia:

- Tahun Akademik;
- Tahun Angkatan;
- Jenjang;
- Fakultas;
- Program Studi;
- Jalur Masuk;
- Propinsi.

#### Tab Program Studi

Menampilkan informasi akreditasi program studi:

- jumlah program studi;
- jumlah akreditasi aktif;
- jumlah peringkat Unggul;
- akreditasi yang berakhir dalam satu tahun;
- akreditasi kedaluwarsa;
- komposisi peringkat akreditasi;
- peringkat akreditasi per fakultas;
- satu tabel terpadu akreditasi dan masa berlaku.

#### Tab Mahasiswa

Menampilkan:

- total mahasiswa;
- jumlah mahasiswa berdasarkan seluruh status yang tersedia pada data, seperti Aktif, Lulus, Cuti, Tidak Aktif, DO/Putus Studi, Mengundurkan Diri, Meninggal, Skorsing, Pindah/Transfer, dan status lain;
- 20 program studi dengan mahasiswa terbanyak;
- ringkasan mahasiswa per program studi dan status;
- komposisi jenis kelamin;
- peta persebaran asal propinsi mahasiswa.

#### Tab PMB

Menampilkan:

- jumlah peminat;
- jumlah lulus seleksi;
- jumlah daftar ulang;
- yield rate;
- funnel PMB;
- ringkasan PMB per program studi, jenjang, fakultas, dan jalur masuk.

#### Tab Kelulusan

Menampilkan:

- jumlah lulusan;
- jumlah lulus tepat waktu;
- jumlah lulusan tepat waktu dengan IPK minimal 3,25;
- persentase kelulusan tepat waktu;
- komposisi ketepatan waktu;
- kategori ketepatan waktu dan IPK;
- persentase kelulusan tepat waktu per fakultas;
- detail data yudisium.

#### Tab Kebutuhan Khusus

Menampilkan:

- jumlah mahasiswa berkebutuhan khusus;
- mahasiswa berkebutuhan khusus berstatus aktif;
- jumlah program studi;
- jenis kebutuhan khusus;
- distribusi mahasiswa berdasarkan kebutuhan khusus.

#### Tab Daerah Tertinggal

Menampilkan:

- jumlah mahasiswa dari kabupaten daerah tertinggal;
- jumlah kabupaten yang teridentifikasi;
- jumlah propinsi;
- ringkasan persebaran mahasiswa dari daerah tertinggal.

### 3. Dashboard IKU

Dashboard IKU memiliki tab:

- IKU Akademik;
- IKU SDM;
- IKU Keuangan;
- IKU Aset;
- IKU Riset;
- IKU Alumni.

Indikator yang sudah tersedia pada tab **IKU Akademik**:

- IKU-01-01 Persentase peningkatan mahasiswa pada PTK;
- IKU-01-05 Persentase lulusan pesantren yang ditampung;
- IKU-01-07 Persentase mahasiswa baru dari daerah tertinggal;
- IKU-02-01 Persentase peningkatan mahasiswa berkebutuhan khusus;
- IKU-04-1 Persentase mahasiswa lulus tepat waktu dengan IPK minimal 3,25;
- IKU-04-3 Persentase kelulusan tepat waktu;
- IKU-35-54 Yield Rate Mahasiswa Baru;
- kesiapan Dashboard Profil Mahasiswa untuk IKU-48-06.

Indikator yang belum memiliki sumber data resmi ditampilkan sebagai **Data belum tersedia** dan tidak diisi dengan angka asumsi.

### 4. Dashboard lain

Halaman berikut sudah tersedia sebagai kerangka dan masih menampilkan konten sedang disiapkan:

- Dashboard SDM;
- Dashboard Keuangan;
- Dashboard Aset;
- Dashboard Riset;
- Dashboard Alumni;
- Prestasi Mahasiswa;
- Paperless (PLO).

Setiap halaman memiliki filter Tahun pada area konten dashboard.

## Export laporan

Pada setiap tab Dashboard Akademik tersedia fitur:

- **Export Excel**
- **Export PDF**

File Excel berisi:

- sheet `Ringkasan`;
- sheet `Data`.

File PDF berisi:

- judul laporan;
- periode atau filter aktif;
- ringkasan indikator;
- tabel data.

## Sumber data

Tempatkan file berikut di folder `data/`:

```text
data/
├── v_mahasiswa.xls
├── difabel.xls
├── vs_yudisium_mahasiswa.xls
├── vs_rekap_pmb.xls
└── akreditasiprodi.xls
```

Keterangan:

- `v_mahasiswa.xls` — data utama mahasiswa, status, fakultas, program studi, jenjang, jalur masuk, propinsi, dan kabupaten/kota.
- `difabel.xls` — data mahasiswa berkebutuhan khusus.
- `vs_yudisium_mahasiswa.xls` — data kelulusan, tanggal yudisium, IPK, semester, dan status tepat waktu. Tanggal yudisium diperlakukan sebagai tanggal lulus.
- `vs_rekap_pmb.xls` — data peminat, lulus seleksi, daftar ulang, program studi, jenjang, dan jalur seleksi.
- `akreditasiprodi.xls` — data peringkat akreditasi, nomor SK, tanggal SK, dan masa berlaku akreditasi program studi.

## Struktur proyek

```text
dashboarduin/
├── app.py
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml
└── data/
    ├── v_mahasiswa.xls
    ├── difabel.xls
    ├── vs_yudisium_mahasiswa.xls
    ├── vs_rekap_pmb.xls
    └── akreditasiprodi.xls
```

## Requirements

Isi `requirements.txt`:

```text
streamlit>=1.46,<2
pandas>=2.2,<3
plotly>=6,<7
xlrd>=2.0.1,<3
openpyxl>=3.1,<4
reportlab>=4.2,<5
```

## Menjalankan di Windows 11 / Antigravity IDE

Buka terminal pada folder proyek, lalu jalankan:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m streamlit run app.py
```

Buka aplikasi melalui:

```text
http://localhost:8501
```

Untuk menghentikan aplikasi:

```text
Ctrl + C
```

## Menjalankan setelah update kode

Jika hanya mengganti `app.py`:

```powershell
python -m streamlit run app.py
```

Jika `requirements.txt` ikut berubah:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Jika tampilan lama masih muncul, lakukan refresh browser:

```text
Ctrl + F5
```

## Catatan penting

- Filter Tahun Akademik saat ini digunakan sebagai konteks periode. Data sumber utama belum memiliki kolom semester akademik yang lengkap.
- Jika tahun PMB yang dipilih belum tersedia, dashboard menggunakan tahun PMB terbaru dan menampilkan peringatan.
- Peta propinsi menggunakan pencocokan nama propinsi dengan koordinat representatif.
- Daerah tertinggal mengacu pada daftar kabupaten daerah tertinggal periode 2025–2029 yang sudah dimasukkan ke dalam kode.
- Data pribadi mahasiswa harus disimpan pada repository privat dan hanya dapat diakses oleh pihak yang berwenang.
- Jangan mengunggah file `.streamlit/secrets.toml`, `.env`, token, password, atau kredensial ke GitHub.

## Repository

```text
https://github.com/abdullahdudung/dashboarduin
```
