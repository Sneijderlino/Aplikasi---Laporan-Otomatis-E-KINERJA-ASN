import os
from utils import resource_path

# Application Info
APP_NAME = "Aplikasi E-Kinerja ASN"
APP_VERSION = "1.0.0"

# Base path untuk data aplikasi
BASE_PATH = os.path.join(os.path.expanduser("~"), "Documents", "DATABASE_E-KINERJA")

# Resource paths
LOGO_PATH = resource_path(os.path.join("Logo", "logo.png"))

# File paths
EXCEL_FILE = os.path.join(BASE_PATH, "laporan_asn.xlsx")
IDENT_FILE = os.path.join(BASE_PATH, "identitas_asn.json")
CRED_FILE = os.path.join(BASE_PATH, "credentials.json")

# Kolom untuk data laporan
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
    "Foto 1", 
    "Foto 2", 
    "Foto 3", 
    "Generated At"
]