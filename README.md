Apk Pencatatan Surat Masuk dan Keluar

## Ringkasan

Aplikasi GUI sederhana berbasis Tkinter untuk mencatat surat masuk dan keluar, menyimpan data ke SQLite, serta impor/ekspor Excel.

# Apk Pencatatan Surat Masuk dan Keluar

Profesional, ringan, dan mudah digunakan — aplikasi desktop sederhana berbasis Tkinter untuk mencatat,
mengelola, mengimpor, dan mengekspor surat masuk dan surat keluar.

## Table of Contents

- [Highlights](#highlights)
- [Quick start](#quick-start)
- [Requirements](#requirements)
- [Configuration](#configuration)
- [Usage](#usage)
- [Project structure](#project-structure)
- [Development & testing](#development--testing)
- [Packaging / Distribution](#packaging--distribution)
- [Contributing](#contributing)
- [License](#license)
- [Changelog](#changelog)

## Highlights

- Simple CRUD for incoming/outgoing letters (nomor, tanggal, pihak, perihal, penanggung, catatan)
- Import / Export Excel (.xlsx) using openpyxl
- Local persistence with SQLite (no external DB required)
- Small, focused codebase — easy to inspect and modify

## Quick start

1. Clone or copy this repository to your machine.
2. (Optional) Create and activate a virtual environment:

```bash
python -m venv .venv
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
# Windows cmd
.\.venv\Scripts\activate.bat
# Git Bash / WSL
source .venv/Scripts/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

# Apk Pencatatan Surat Masuk dan Keluar

Dokumentasi resmi dalam Bahasa Indonesia untuk aplikasi desktop ringan yang dibuat dengan
Tkinter. Aplikasi ini dirancang untuk membantu pencatatan surat masuk dan keluar secara sederhana,
dengan fitur impor/ekspor Excel dan penyimpanan lokal menggunakan SQLite.

## Daftar Isi

- [Sorotan](#sorotan)
- [Persiapan cepat](#persiapan-cepat)
- [Kebutuhan Sistem](#kebutuhan-sistem)
- [Konfigurasi](#konfigurasi)
- [Panduan Penggunaan](#panduan-penggunaan)
- [Struktur Proyek](#struktur-proyek)
- [Pengembangan & Pengujian](#pengembangan--pengujian)
- [Distribusi / Packaging](#distribusi--packaging)
- [Kontribusi](#kontribusi)
- [Troubleshooting](#troubleshooting)
- [Lisensi](#lisensi)
- [Changelog](#changelog)

## Sorotan

- CRUD sederhana untuk surat masuk dan keluar (nomor, tanggal, pihak, perihal, penanggung, catatan).
- Impor dan ekspor Excel (.xlsx) menggunakan `openpyxl`.
- Penyimpanan lokal dengan SQLite — tidak perlu server eksternal.
- Basis kode kecil dan mudah dimodifikasi.

## Persiapan cepat

1. Salin atau clone repository ini ke mesin Anda.
2. (Opsional) Buat dan aktifkan virtual environment:

```bash
python -m venv .venv
# PowerShell
.\.venv\Scripts\Activate.ps1
# cmd
.\.venv\Scripts\activate.bat
# Git Bash / WSL
source .venv/Scripts/activate
```

3. Install dependensi:

```bash
pip install -r requirements.txt
```

4. Jalankan aplikasi:

```bash
python Apk_Surat_Masuk_Keluar.py
```

Jendela aplikasi akan muncul. Gunakan menu di sisi kiri untuk beralih antara "Surat Masuk" dan "Surat Keluar".

## Kebutuhan Sistem

- Python 3.8 atau lebih baru
- Paket: lihat `requirements.txt` (minimal `openpyxl` dan `Pillow`)

## Konfigurasi

Lokasi Database

- Aplikasi menyimpan file database SQLite per pengguna di folder Local AppData agar data pengguna
  terpisah dan tidak ikut ter-commit ke repository.

Lokasi default pada Windows:

```
C:\Users\<username>\AppData\Local\DataBase-Surat Masuk Keluar\DataBase.db
```

- Aplikasi memanfaatkan variabel lingkungan `LOCALAPPDATA` bila tersedia; bila tidak, akan
  menggunakan `<home>/AppData/Local` sebagai fallback.
- Folder database akan dibuat otomatis jika belum ada.
- Jangan mengunggah atau men-commit file database ke GitHub; `.gitignore` sudah dikonfigurasi untuk
  mengabaikan file DB lokal.

Logging

- Logging dasar diaktifkan dan dikirim ke konsol. Periksa output konsol untuk mendapatkan pesan
  diagnostik saat terjadi masalah pada impor/ekspor atau akses berkas.

## Panduan Penggunaan

- Menambah surat: isi kolom di bagian "Input / Edit Surat" lalu klik "Tambah / Update".
- Mengedit surat: klik ganda baris pada tabel untuk mengisi form, ubah data, lalu klik "Tambah / Update".
- Menghapus surat: pilih baris lalu klik "Hapus Terpilih". Hati-hati — saat ini belum ada fitur undo;
  lakukan backup DB jika perlu.
- Pencarian: gunakan kotak pencarian untuk memfilter berdasarkan nomor, pihak, perihal, atau penanggung.
- Impor Excel: gunakan tombol "Import Excel"; file .xlsx diharapkan memiliki header. Aplikasi akan
  mencoba memetakan kolom dan melewati baris yang tidak valid.
- Ekspor Excel: gunakan tombol "Ekspor Excel" untuk menyimpan baris yang terlihat ke file .xlsx.

Catatan tentang tanggal

- Form meminta format tanggal `DD-MM-YYYY`. Saat ini tanggal disimpan sebagai string; disarankan untuk
  menyimpan tanggal dalam format ISO (`YYYY-MM-DD`) di perbaikan selanjutnya agar penyortiran lebih akurat.

## Struktur Proyek

Berkas dan folder penting:

- `Apk_Surat_Masuk_Keluar.py` — aplikasi utama (GUI + wrapper DB)
- `requirements.txt` — daftar dependensi Python
- `README.md` — dokumentasi (file ini)
- `.gitignore` — pola berkas yang akan diabaikan oleh Git (termasuk DB lokal)
- `LICENSE` — lisensi proyek
- `CHANGELOG.md`, `version.txt` — catatan rilis dan versi
- `icon/`, `img/` — aset gambar/logo

## Pengembangan & Pengujian

- Ikuti panduan kontribusi pada `CONTRIBUTING.md`.
- Static check (opsional):

```bash
pip install flake8
flake8 Apk_Surat_Masuk_Keluar.py
```

- Untuk pengujian unit, saya sarankan memisahkan logika database menjadi `db.py`, lalu menulis
  test yang menggunakan SQLite in-memory (`sqlite3.connect(':memory:')`). Saya dapat membantu
  membuat kerangka test jika Anda mau.

## Distribusi / Packaging

- Untuk membuat executable Windows (EXE) gunakan `pyinstaller`:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed Apk_Surat_Masuk_Keluar.py
```

- Jika Anda membundel aset (logo), pastikan menyertakannya dalam opsi PyInstaller atau menyesuaikan
  path pada kode agar dapat menemukan aset saat dijalankan dari executable.

## Kontribusi

Terima kasih bila Anda ingin berkontribusi. Hal penting:

- Jangan men-commit file database lokal (`*.db`, `DataBase.db`). File tersebut sudah di-ignore.
- Gunakan branch fitur dan pesan commit yang jelas.
- Tambahkan test untuk fitur baru bila memungkinkan.

Lihat juga `CONTRIBUTING.md` untuk panduan lengkap.

## Troubleshooting

- Aplikasi tidak mau jalan: periksa versi Python dan pastikan dependensi ter-install.
- Impor Excel gagal: buka file di Excel atau LibreOffice untuk verifikasi header dan urutan kolom.
- Aplikasi gagal membuat DB: periksa izin pada `%LOCALAPPDATA%` dan pastikan ruang disk mencukupi.

## Lisensi

Proyek ini dirilis di bawah lisensi MIT — lihat berkas `LICENSE` untuk detail.

## Changelog

Lihat `CHANGELOG.md` untuk histori rilis.

## Dukungan

Jika menemukan masalah atau ingin fitur baru, silakan buka issue di GitHub repository atau hubungi pemelihara.

---

Ingin saya lanjutkan ke salah satu dari berikut?

- Memecah kode menjadi `db.py` dan `ui.py` untuk testabilitas,
- Menambahkan unit test untuk lapisan DB (surat),
- Menambahkan workflow GitHub Actions untuk linting (flake8) dan tes ringan.

Balas dengan pilihan Anda dan saya akan lanjutkan.
# Apk-Surat-Masuk-keluar
# Apk-Surat-Masuk-keluar
