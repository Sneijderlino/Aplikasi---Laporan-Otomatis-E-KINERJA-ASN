import os
import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def ensure_directory_exists(directory):
    """Create directory if it doesn't exist"""
    if not os.path.exists(directory):
        os.makedirs(directory)

def get_file_extension(filename):
    """Get file extension from filename"""
    return os.path.splitext(filename)[1].lower()

def is_image_file(filename):
    """Check if file is an image based on extension"""
    valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp']
    return get_file_extension(filename) in valid_extensions

def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    # Replace invalid characters with underscore
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename