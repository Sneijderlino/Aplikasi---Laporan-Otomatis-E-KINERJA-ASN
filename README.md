<p align="center">
  <img src="/img/Walpaper.png" alt="" width="900"/>
</p>

# 🏢 APLIKASI LAPORAN E-KINERJA

[![License](https://img.shields.io/github/license/Sneijderlino/Aplikasi-Laporan-eKINERJA?style=for-the-badge&color=2ecc71)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.11+-3776AB.svg?style=for-the-badge&logo=python)](https://www.python.org/)
[![CI](https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/actions/workflows/python-app.yml/badge.svg)](https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/actions)
[![OS Support](https://img.shields.io/badge/OS-Windows%20%7C%20Linux%20%7C%20macOS-informational?style=for-the-badge)](https://www.python.org/)

_Aplikasi Desktop Manajemen Laporan Otomatis E-KINERJA Berbasis Python & Tkinter._

## 🎯 Tujuan Proyek

_Aplikasi Laporan E-KINERJA_ dikembangkan untuk memudahkan proses pencatatan, pengelolaan, dan pelaporan data E-KINERJA. Aplikasi ini menyimpan data ASN secara persisten dan mampu mengekspor laporan ke format Excel dan PDF.

## ✨ Fitur Utama

- _Modern GUI:_ Antarmuka pengguna grafis (GUI) yang modern dan responsif menggunakan _CustomTkinter_ dengan tema yang menarik.
- _Penyimpanan Data Terpadu:_ Menggunakan file **laporan_asn.xlsx** (Excel) sebagai database utama dan **identitas_asn.json** untuk menyimpan data identitas ASN dengan validasi data.
- _Identitas Persisten:_ Data Nama, NIP, Jabatan, dan Unit Kerja ASN disimpan dan dikunci secara otomatis dengan enkripsi data sensitif.
- _Dukungan Multimedia:_ Menyertakan _Pillow_ untuk upload, preview, dan optimasi foto di form input.
- _Export Multi-Format:_ Mampu mengekspor laporan ke format _Excel (.xlsx)_ dan _PDF (.pdf)_ dengan template yang dapat disesuaikan.
- _Operasi CRUD Lengkap:_ Mendukung operasi Tambah, Ubah, dan Hapus (CRUD) laporan dengan validasi data.
- _Backup & Restore:_ Fitur backup otomatis dan kemampuan restore data.

## ⚙ Prasyarat Sistem & Dependensi

### A. Dependensi Python

Aplikasi ini membutuhkan library berikut untuk fungsionalitas penuh:

| Library         | Kegunaan                                                       | Versi Min. |
| :-------------- | :------------------------------------------------------------- | :--------- |
| _pandas_        | Pengolahan dan penyimpanan data ke/dari Excel.                 | 2.1.0      |
| _openpyxl_      | Mesin backend untuk Pandas & fitur penyisipan gambar ke Excel. | 3.1.2      |
| _customtkinter_ | Framework GUI modern dengan tema yang menarik.                 | 5.2.0      |
| _Pillow_        | Manajemen dan optimasi gambar.                                 | 10.0.0     |
| _fpdf_          | Pembuatan laporan dalam format PDF.                            | 1.7.2      |
| _xlsxwriter_    | Pemformatan Excel lanjutan dan templating.                     | 3.1.2      |

### B. Prasyarat Sistem

- _Python 3.11+_
- _Git_
- Minimum 4GB RAM
- 500MB ruang disk
- Resolusi layar minimum 1366x768

---

## 🚀 Instalasi dan Penggunaan

### Langkah 1: Klon Repositori

```bash
git clone https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA.git
cd Aplikasi-Laporan-eKINERJA
```

---

### Langkah 2: Siapkan Lingkungan Virtual (Virtual Environment)

Langkah 2: Siapkan Lingkungan Virtual (Virtual Environment)
Sangat disarankan untuk mengisolasi proyek ini.

```bash
# Buat venv
python3 -m venv venv
# Aktifkan venv (Linux/macOS)
source venv/bin/activate
# Aktifkan venv (Windows PowerShell)
# .\venv\Scripts\activate
```

---

### Langkah 3: Instal Dependensi

```bash
Pastikan Anda berada di lingkungan virtual yang sudah aktif, lalu instal semua library:
pip install -r requirements.txt
```

---

### Langkah 4: Menjalankan Aplikasi

```bash
Jalankan skrip utama:
python E-kinerjaASN.py
```

---

## 📥 Cara Download Aplikasi (Versi Rilis)

Kamu juga dapat langsung _mengunduh file aplikasi yang sudah jadi_ melalui halaman rilis GitHub.

### 🪄 Langkah 1: Buka Halaman Rilis

1. Klik tab **[Releases](https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/releases)** pada repository ini.
2. Pilih versi rilis terbaru (v1.0).

### 💾 Langkah 2: Unduh File

- Cari bagian _“Assets”_ di bawah rilis tersebut.
- Klik file .zip ATAU .tar.gz untuk mengunduhnya ke perangkatmu.

### 🖥 Langkah 3: Ekstrak & Jalankan

- Ekstrak file .zip ke folder yang kamu inginkan.
- Jika file .exe tersedia → klik 2× untuk langsung membuka aplikasi.
- Jika file .py (source code):

  ```bash
  # Pastikan Python 3.8+ sudah terpasang
  python laporan_app.py

  ```

### ✅ Tips:

- Gunakan Windows 10/11 atau Linux dengan Python 3.8+

- Jika diperlukan, install dependensi dari requirements.txt:
- pip install -r requirements.txt

<p align="center">
  <img src="https://img.shields.io/badge/Made%20with-Python-blue?style=for-the-badge&logo=python" alt="Python Badge"/>
  <img src="https://img.shields.io/badge/Status-Active-success?style=for-the-badge" alt="Status Active"/>
  <img src="https://img.shields.io/github/stars/Sneijderlino/youtube-downloader-pro?style=for-the-badge" alt="GitHub Stars"/>
  <img src="https://img.shields.io/github/forks/Sneijderlino/youtube-downloader-pro?style=for-the-badge" alt="GitHub Forks"/>
</p>

---

<h3 align="center">📜 Lisensi</h3>

<p align="center">
  Proyek ini dilisensikan di bawah <a href="LICENSE">MIT License</a>.<br>
  Bebas digunakan, dimodifikasi, dan dibagikan selama mencantumkan kredit.
</p>

---

<h3 align="center">💬 Dukungan & Kontribusi</h3>

<p align="center">
  💡 Temukan bug atau ingin menambahkan fitur baru?<br>
  Silakan buka <a href="https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/issues">Issues</a> atau buat <a href="https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/pulls">Pull Request</a>.<br><br>
  ⭐ Jangan lupa beri bintang jika proyek ini bermanfaat!
</p>

---

<p align="center">
  Dibuat dengan ❤ oleh <a href="https://www.tiktok.com/@sneijderlino_official?is_from_webapp=1&sender_device=pc">Sneijderlino</a><br>
  <em>“Code. Create. Conquer.”</em>
</p>
