import os
import json
import shutil
import re
from pathlib import Path
from utils import resource_path

# Application Info
APP_NAME = "Aplikasi E-Kinerja ASN"
APP_VERSION = "1.0.0"

# Storage configuration
# Environment overrides
ENV_OVERRIDE = "E_KINERJA_DATA_DIR"  # if set, this path is used
ENV_PREFER_SHARED = "E_KINERJA_PREFER_SHARED"  # if set to truthy, prefer machine-wide ProgramData

# Legacy path that was previously hardcoded in this repo; we'll attempt to migrate from it if found.
LEGACY_LENOVO_PATH = Path(r"C:\Users\LENOVO\DATABASE_E-KINERJA")


def _is_truthy_env(name: str) -> bool:
    v = os.getenv(name)
    if not v:
        return False
    return v.lower() in ("1", "true", "yes", "on")


def _find_existing_base_path():
    """Search common locations on this machine for an existing app data folder.

    Return Path if found and non-empty, else None.
    """
    candidates = []
    try:
        if os.name == 'nt':
            progdata = os.getenv('PROGRAMDATA')
            if progdata:
                candidates.append(Path(progdata) / APP_NAME)
            # check all user profiles under C:\Users for an existing LocalAppData app folder
            try:
                users_root = Path(os.path.splitdrive(Path.home())[0] + os.sep) / 'Users'
                if users_root.exists():
                    for p in users_root.iterdir():
                        cand = p / 'AppData' / 'Local' / APP_NAME
                        if cand.exists():
                            candidates.append(cand)
            except Exception:
                pass
            candidates.append(LEGACY_LENOVO_PATH)
        else:
            candidates.append(Path('/var') / APP_NAME)
            candidates.append(Path(os.getenv('XDG_DATA_HOME') or (Path.home() / '.local' / 'share')) / APP_NAME)

        for c in candidates:
            try:
                if c.exists() and any(c.iterdir()):
                    return c
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_app_base_path(prefer_shared: bool = False) -> Path:
    """Determine a writable application data folder.

    Order of precedence:
    1. ENV_OVERRIDE if set
    2. If prefer_shared True or ENV_PREFER_SHARED set -> ProgramData on Windows or /var on *nix
    3. Per-user LocalAppData on Windows or XDG_DATA_HOME on *nix
    4. Fallback to current working directory
    """
    # 1) explicit override
    env = os.getenv(ENV_OVERRIDE)
    if env:
        base = Path(env)
    else:
        # If an existing installation is present on this machine, prefer that so multiple users
        # on the same PC share the same BASE_PATH. This allows a new user to store their
        # data under the same application folder without needing to configure anything.
        existing = _find_existing_base_path()
        if existing is not None:
            base = existing
        else:
            # decide shared vs per-user
            prefer_shared = prefer_shared or _is_truthy_env(ENV_PREFER_SHARED)
            if os.name == 'nt':
                if prefer_shared:
                    progdata = os.getenv('PROGRAMDATA') or r"C:\ProgramData"
                    base = Path(progdata) / APP_NAME
                else:
                    localapp = os.getenv('LOCALAPPDATA') or (Path.home() / 'AppData' / 'Local')
                    base = Path(localapp) / APP_NAME
            else:
                # non-windows
                if prefer_shared:
                    base = Path('/var') / APP_NAME
                else:
                    base = Path(os.getenv('XDG_DATA_HOME') or (Path.home() / '.local' / 'share')) / APP_NAME

    # try to create the folder
    try:
        base.mkdir(parents=True, exist_ok=True)
    except Exception:
        # fallback to cwd
        base = Path.cwd() / APP_NAME
        try:
            base.mkdir(parents=True, exist_ok=True)
        except Exception:
            base = Path.cwd()

    # If legacy data exists and base is different, attempt migration (safe copy then remove)
    try:
        if LEGACY_LENOVO_PATH.exists() and LEGACY_LENOVO_PATH.resolve() != base.resolve():
            # Only migrate if legacy has files and target is empty (to avoid overwriting)
            legacy_has = any(LEGACY_LENOVO_PATH.iterdir())
            target_has = any(base.iterdir()) if base.exists() else False
            if legacy_has and not target_has:
                try:
                    shutil.copytree(LEGACY_LENOVO_PATH, base, dirs_exist_ok=True)
                    # attempt to remove legacy (do not raise if fails)
                    try:
                        shutil.rmtree(LEGACY_LENOVO_PATH)
                    except Exception:
                        pass
                except Exception:
                    # if copy fails, ignore and continue
                    pass
    except Exception:
        pass

    return base


# default base path for the app (per-user LocalAppData by default)
BASE_PATH = get_app_base_path()


def get_user_dir(username):
    """Return a per-user folder under BASE_PATH and create it if missing.

    The folder will be: BASE_PATH/<username>
    If creation fails the function returns BASE_PATH as a fallback.
    """
    # normalize username to a safe folder name
    safe_name = _sanitize_username(str(username))
    user_dir = Path(BASE_PATH) / safe_name
    try:
        user_dir.mkdir(parents=True, exist_ok=True)
        return str(user_dir)
    except Exception:
        # fallback to BASE_PATH so callers still have a writable path if possible
        return str(BASE_PATH)


def _sanitize_username(name: str) -> str:
    """Return a filesystem-safe username (remove or replace illegal characters)."""
    # remove control chars and replace any non-alnum, dot, underscore, or hyphen with underscore
    name = re.sub(r'[\x00-\x1f<>:"/\\|?*]+', '_', name)
    name = re.sub(r'[^A-Za-z0-9._-]+', '_', name)
    return name or 'user'

# Resource paths
LOGO_PATH = resource_path(os.path.join("Logo", "logo.png"))

# File paths
def get_user_excel_file(username):
    """Get Excel file path for a specific user"""
    user_dir = get_user_dir(username)
    return os.path.join(user_dir, f"laporan_{username}.xlsx")

def get_user_identity_path(username):
    """Get identity file path for a specific user"""
    user_dir = get_user_dir(username)
    return os.path.join(user_dir, f"identitas_asn_{username}.json")

CRED_FILE = os.path.join(BASE_PATH, "credentials.json")

# Create an empty credentials file if missing (safe default)
if not os.path.exists(CRED_FILE):
    try:
        with open(CRED_FILE, 'w', encoding='utf-8') as _f:
            json.dump({}, _f)
    except Exception:
        # If we cannot create the file, let the application handle subsequent errors.
        pass

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


# Default user-editable settings (display labels and report subtitle)
DEFAULT_USER_SETTINGS = {
    "report_subtitle": "Koordinasi dengan Kepala Regu terkait informasi kejadian kebakaran",
    # Mapping from internal column key -> display label
    "field_labels": {
        "Uraian Kejadian": "Uraian Kejadian",
        "Waktu Kebakaran": "Waktu Kebakaran",
        "Kerusakan": "Kerusakan",
        "Tindakan": "Tindakan"
    }
}

# Default kop (header) lines used in PDF reports. Per-user override supported.
DEFAULT_USER_SETTINGS.setdefault('kop_lines', [
    "Dinas Satuan Polisi Pamong Praja Dan Pemadam Kebakaran",
    "Kabupaten Maluku Barat Daya",
    "Kota Tiakur",
])

def get_user_settings_path(username):
    """Return path for per-user settings JSON file."""
    user_dir = get_user_dir(username)
    return os.path.join(user_dir, f"settings_{username}.json")

def load_user_settings(username):
    path = get_user_settings_path(username)
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # merge with defaults
                out = DEFAULT_USER_SETTINGS.copy()
                out.update(data or {})
                # ensure field_labels contains defaults
                fl = out.get('field_labels', {})
                for k, v in DEFAULT_USER_SETTINGS['field_labels'].items():
                    fl.setdefault(k, v)
                out['field_labels'] = fl
                return out
    except Exception:
        pass
    return DEFAULT_USER_SETTINGS.copy()

def save_user_settings(username, settings):
    path = get_user_settings_path(username)
    dir_name = os.path.dirname(path)
    if dir_name and not os.path.exists(dir_name):
        os.makedirs(dir_name, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)