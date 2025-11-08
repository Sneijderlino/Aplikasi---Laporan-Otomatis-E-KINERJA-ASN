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

4. Run the app:

```bash
python Apk_Surat_Masuk_Keluar.py
```

The app window should open. Use the left menu to switch between "Surat Masuk" and "Surat Keluar".

## Requirements

- Python 3.8+
- See `requirements.txt` (this project uses at least `openpyxl` and `Pillow`)

## Configuration

Database location

- By design the application stores the SQLite database in the user's Local AppData folder to keep
  per-user data separate and avoid committing the database to source control.

Default path on Windows:

```
C:\Users\<username>\AppData\Local\DataBase-Surat Masuk Keluar\DataBase.db
```

- The app uses the `LOCALAPPDATA` environment variable when present; otherwise it falls back to
  `<home>/AppData/Local`.
- Database directory is created automatically if missing. Do not commit or upload the DB file to GitHub.

Logging

- Basic logging is enabled and printed to the console. Check console output for diagnostics when
  troubleshooting import/export or file access issues.

## Usage

- Add a letter: fill the fields in the "Input / Edit Surat" form and click "Tambah / Update".
- Edit: double-click a row in the table to prefill the form, make changes, then click "Tambah / Update".
- Delete: select a row and click "Hapus Terpilih" (be cautious — currently no undo; consider backing up DB).
- Search/Reset: use the search box to filter the table by nomor/pihak/perihal/penanggung.
- Import: `Import Excel` expects an .xlsx with headers (ID optional). The app attempts to map columns
  and skip invalid rows.
- Export: `Ekspor Excel` will save the currently visible rows into an .xlsx file.

Notes on dates

- The UI expects date input in `DD-MM-YYYY`. Internally the app stores date strings; future improvement
  suggestion: convert/save dates in ISO (YYYY-MM-DD) for reliable sorting/filtering.

## Project structure

Important files and folders:

- `Apk_Surat_Masuk_Keluar.py` — main application (GUI + DB wrapper class)
- `requirements.txt` — Python dependencies
- `README.md` — this file
- `.gitignore` — patterns for files that should not be committed (DB, virtualenv, temporary files)
- `LICENSE` — project license
- `CHANGELOG.md`, `version.txt` — release history and version meta
- `icon/`, `img/` — images and assets

## Development & testing

- Follow the style guidelines in `CONTRIBUTING.md`.
- To run a quick static check (flake8):

```bash
pip install flake8
flake8 Apk_Surat_Masuk_Keluar.py
```

- For unit testing, consider extracting the DB logic into `db.py` and adding tests that use
  an in-memory SQLite DB (`sqlite3.connect(':memory:')`). I can help with a test scaffold.

## Packaging / Distribution

- To create a Windows executable, use `pyinstaller`:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed Apk_Surat_Masuk_Keluar.py
```

- Note: when packaging, ensure static assets (logo) are included and update code to reference bundled paths.

## Contributing

Please see `CONTRIBUTING.md` for contribution guidelines. A few highlights:

- Do not commit the local SQLite DB (it's ignored via `.gitignore`).
- Use feature branches and descriptive commit messages.
- Add tests for new functionality where appropriate.

## Troubleshooting

- If the app fails to start: check Python version and that dependencies from `requirements.txt` are installed.
- If Excel import fails for a file: open the file in Excel/LibreOffice and verify headers and column order.
- If the app cannot create the DB: verify permissions on `%LOCALAPPDATA%` and that disk space is available.

## License

This project is released under the MIT License — see `LICENSE` for details.

## Changelog

See `CHANGELOG.md` for release notes and version history.

## Support / Contact

If you encounter issues or want enhancements, please open an issue on GitHub or contact the maintainer.

---

If you'd like, I can also:

- split the code into `db.py`/`ui.py` for better testability,
- add unit tests for the DB layer, or
- prepare a small GitHub Actions workflow for linting and basic checks.

Tell me which of those you'd like next and I will implement them.
# Apk-Surat-Masuk-keluar
