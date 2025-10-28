import os
import json
import hashlib
import binascii
import time
from config import CRED_FILE

# Simple PBKDF2 password hashing utilities

def _ensure_users_file():
    dir_name = os.path.dirname(CRED_FILE)
    if dir_name and not os.path.exists(dir_name):
        os.makedirs(dir_name, exist_ok=True)
    if not os.path.exists(CRED_FILE):
        with open(CRED_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": {}}, f, ensure_ascii=False, indent=2)


def load_users():
    _ensure_users_file()
    try:
        with open(CRED_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("users", {})
    except Exception:
        return {}


def save_users(users_dict):
    _ensure_users_file()
    try:
        with open(CRED_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": users_dict}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise


def _hash_password(password, salt=None, iterations=100000):
    if salt is None:
        salt = os.urandom(16)
    if isinstance(salt, str):
        salt = binascii.unhexlify(salt)
    dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, iterations)
    return {
        'salt': binascii.hexlify(salt).decode('ascii'),
        'hash': binascii.hexlify(dk).decode('ascii'),
        'iterations': iterations
    }


def add_user(username, password, full_name=None):
    username = str(username).strip()
    if not username:
        raise ValueError("Username kosong")
    if not password or len(password) < 6:
        raise ValueError("Password harus minimal 6 karakter")

    users = load_users()
    if username in users:
        raise ValueError("User sudah terdaftar")

    hp = _hash_password(password)
    users[username] = {
        'full_name': full_name or "",
        'created_at': time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()),
        'salt': hp['salt'],
        'hash': hp['hash'],
        'iterations': hp['iterations']
    }
    save_users(users)
    return True


def verify_user(username, password):
    username = str(username).strip()
    users = load_users()
    u = users.get(username)
    if not u:
        return False
    try:
        expected = u.get('hash')
        salt = u.get('salt')
        iterations = int(u.get('iterations', 100000))
        new = _hash_password(password, salt=salt, iterations=iterations)
        return new['hash'] == expected
    except Exception:
        return False


def user_exists(username):
    users = load_users()
    return username in users


def ensure_has_any_user():
    users = load_users()
    return bool(users)

def verify_full_name(username, full_name):
    users = load_users()
    user = users.get(username)
    if not user:
        return False
    return user.get('full_name', '').strip().lower() == full_name.strip().lower()

def reset_password(username, new_password):
    if not new_password or len(new_password) < 6:
        raise ValueError("Password baru harus minimal 6 karakter")
        
    users = load_users()
    if username not in users:
        raise ValueError("User tidak ditemukan")
        
    hp = _hash_password(new_password)
    users[username]['salt'] = hp['salt']
    users[username]['hash'] = hp['hash']
    users[username]['iterations'] = hp['iterations']
    save_users(users)
    return True
