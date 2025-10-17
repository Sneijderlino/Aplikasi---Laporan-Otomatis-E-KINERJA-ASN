<p align="center">
  <img src="/Logo/logo.ico" alt="YouTube Downloader Pro Banner" width="900"/>
</p>



# 🏢 APLIKASI EKINERJA ASN - LAPORAN KEJADIAN

[![License](https://img.shields.io/github/license/[USER]/[REPO]?style=for-the-badge&color=2ecc71)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.8+-3776AB.svg?style=for-the-badge&logo=python)](https://www.python.org/)
[![OS Support](https://img.shields.io/badge/OS-Windows%20%7C%20Linux%20%7C%20macOS-informational?style=for-the-badge)](https://www.python.org/)

*Aplikasi Desktop Manajemen Laporan Otomatis E-KINERJA ASN Berbasis Python & Tkinter.*

## 🎯 Tujuan Proyek

Aplikasi *Aplikasi Laporan Otomatis E-KINERJA ASN* dikembangkan untuk memudahkan proses pencatatan, pengelolaan, dan pelaporan data E-KINERJA. Aplikasi ini menyimpan data ASN secara persisten dan mampu mengekspor laporan ke format Excel dan PDF.

## ✨ Fitur Utama

-   *GUI Profesional:* Antarmuka pengguna grafis (GUI) yang bersih dan terstruktur menggunakan *Tkinter* dan theme clam.
-   *Penyimpanan Data:* Menggunakan file **laporan_asn.xlsx** (Excel) sebagai database utama dan **identitas_asn.json** untuk menyimpan data identitas ASN.
-   *Identitas Persisten:* Data Nama, NIP, Jabatan, dan Unit Kerja ASN disimpan dan dikunci secara otomatis setelah entri pertama.
-   *Dukungan Multimedia:* Menyertakan *Pillow* untuk upload dan preview foto di form input.
-   *Export Serbaguna:* Mampu mengekspor laporan ke format *Excel (.xlsx)* dan *PDF (.pdf)* (tunggal dan banyak file).
-   *Operasi CRUD:* Mendukung operasi Tambah, Ubah, dan Hapus (CRUD) laporan.

## ⚙ Prasyarat Sistem & Dependensi

### A. Dependensi Python

Aplikasi ini membutuhkan library berikut. Beberapa di antaranya bersifat *opsional* (tapi sangat disarankan) untuk fungsionalitas penuh.

| Library | Kegunaan | Wajib? |
| :--- | :--- | :--- |
| *pandas* | Pengolahan dan penyimpanan data ke/dari Excel. | YA |
| *openpyxl* | Mesin backend untuk Pandas & fitur penyisipan gambar ke Excel. | YA |
| *Pillow* | Manajemen gambar (upload & preview foto). | Opsional |
| *fpdf* | Pembuatan laporan dalam format PDF. | Opsional |

### B. Prasyarat Sistem

-   *Python 3.8+*
-   *Git*

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

# Aplikasi---Laporan-Otomatis-E-KINERJA-ASN
