import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import pandas as pd
from pandas import Series

# --- Deteksi Library Opsional ---

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False
    Image = None
    ImageTk = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

# ---------------- KONFIGURASI FILE ----------------

# Nama file data
EXCEL_FILE = "laporan_asn.xlsx"
IDENT_FILE = "identitas_asn.json"

# Kolom untuk DataFrame/Excel
COLUMNS = [
    "Nama",
    "NIP",
    "Jabatan",
    "Unit Kerja",
    "Tanggal",
    "Waktu",
    "Uraian Kejadian",
    "Waktu Kebakaran",
    "Kerusakan",
    "Tindakan",
    "Foto",
    "Generated At"
]