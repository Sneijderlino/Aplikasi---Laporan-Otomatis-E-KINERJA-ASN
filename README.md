<p align="center">
  <img src="/img/Walpaper.png" alt="" width="900"/>
</p>

# 🏢 APLIKASI EKINERJA ASN - LAPORAN KEJADIAN

[![License](https://img.shields.io/github/license/[USER]/[REPO]?style=for-the-badge&color=2ecc71)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.8+-3776AB.svg?style=for-the-badge&logo=python)](https://www.python.org/)
[![OS Support](https://img.shields.io/badge/OS-Windows%20%7C%20Linux%20%7C%20macOS-informational?style=for-the-badge)](https://www.python.org/)

_Aplikasi Desktop Manajemen Laporan Otomatis E-KINERJA ASN Berbasis Python & Tkinter._

## 🎯 Tujuan Proyek

Aplikasi _Aplikasi Laporan Otomatis E-KINERJA ASN_ dikembangkan untuk memudahkan proses pencatatan, pengelolaan, dan pelaporan data E-KINERJA. Aplikasi ini menyimpan data ASN secara persisten dan mampu mengekspor laporan ke format Excel dan PDF.

## ✨ Fitur Utama

- _GUI Profesional:_ Antarmuka pengguna grafis (GUI) yang bersih dan terstruktur menggunakan _Tkinter_ dan theme clam.
- _Penyimpanan Data:_ Menggunakan file **laporan_asn.xlsx** (Excel) sebagai database utama dan **identitas_asn.json** untuk menyimpan data identitas ASN.
- _Identitas Persisten:_ Data Nama, NIP, Jabatan, dan Unit Kerja ASN disimpan dan dikunci secara otomatis setelah entri pertama.
- _Dukungan Multimedia:_ Menyertakan _Pillow_ untuk upload dan preview foto di form input.
- _Export Serbaguna:_ Mampu mengekspor laporan ke format _Excel (.xlsx)_ dan _PDF (.pdf)_ (tunggal dan banyak file).
- _Operasi CRUD:_ Mendukung operasi Tambah, Ubah, dan Hapus (CRUD) laporan.

## ⚙ Prasyarat Sistem & Dependensi

### A. Dependensi Python

Aplikasi ini membutuhkan library berikut. Beberapa di antaranya bersifat _opsional_ (tapi sangat disarankan) untuk fungsionalitas penuh.

| Library    | Kegunaan                                                       | Wajib?   |
| :--------- | :------------------------------------------------------------- | :------- |
| _pandas_   | Pengolahan dan penyimpanan data ke/dari Excel.                 | YA       |
| _openpyxl_ | Mesin backend untuk Pandas & fitur penyisipan gambar ke Excel. | YA       |
| _Pillow_   | Manajemen gambar (upload & preview foto).                      | Opsional |
| _fpdf_     | Pembuatan laporan dalam format PDF.                            | Opsional |

### B. Prasyarat Sistem

- _Python 3.8+_
- _Git_

---

## 🚀 Instalasi dan Penggunaan

### Langkah 1: Klon Repositori

```bash
git clone https://github.com/Sneijderlino/Aplikasi---Laporan-Otomatis-E-KINERJA-ASN.git
cd Laporan-Otomatis-E-KINERJA-ASN
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

## Langkah 4: Menjalankan Aplikasi

```bash
Jalankan skrip utama:
python laporan_app.py
```

---

---

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
  Silakan buka <a href="https://github.com/Sneijderlino/youtube-downloader-pro/issues">Issues</a> atau buat <a href="https://github.com/Sneijderlino/youtube-downloader-pro/pulls">Pull Request</a>.<br><br>
  ⭐ Jangan lupa beri bintang jika proyek ini bermanfaat!
</p>

---

<p align="center">
  Dibuat dengan ❤ oleh <a href="https://github.com/Sneijderlino">Sneijderlino</a><br>
  <em>“Code. Create. Conquer.”</em>
</p>

