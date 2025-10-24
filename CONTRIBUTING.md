# Contributing to Aplikasi Laporan E-KINERJA

Terima kasih telah tertarik untuk berkontribusi! Berikut panduan singkat agar kontribusi Anda cepat ditinjau dan mudah digabungkan.

## Cara berkontribusi

1. Fork repository ini.
2. Buat branch fitur/bugfix dari `main`: `git checkout -b feat/nama-fitur`.
3. Tulis code dan tambahkan test jika relevan.
4. Jalankan `pip install -r requirements.txt` lalu cek bahwa skrip utama dapat di-import tanpa error.
5. Commit perubahan dengan pesan jelas (conventional commit style direkomendasikan):
   - `feat:`, `fix:`, `chore:`, `docs:`
6. Push ke branch Anda dan buat Pull Request.

## Style dan kualitas
- Gunakan `snake_case` untuk nama file Python.
- Ikuti PEP8; disarankan menggunakan `flake8`.
- Hindari men-commit file besar/biner yang tidak perlu (mis. dataset). Gunakan `.gitignore`.

## Issue & PR
- Jelaskan langkah reproduksi untuk bug.
- Tulis deskripsi dan alasan untuk fitur baru.
- Sertakan screenshot jika berkaitan dengan UI.

Terima kasih! 🎉