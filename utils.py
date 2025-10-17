import os
import json
from datetime import datetime
import pandas as pd
from tkinter import messagebox, filedialog
from pandas import Series

# Impor dari config.py
from config import (
    EXCEL_FILE, IDENT_FILE, COLUMNS, 
    PIL_AVAILABLE, Image, ImageTk, FPDF, 
    OPENPYXL_AVAILABLE, load_workbook, XLImage, get_column_letter
)


# ---------------- Utility: date parsing/formatting ------------------------

def parse_date_flexible(date_str):
    """Mencoba memparsing string tanggal dengan berbagai format."""
    if not date_str:
        return None
    s = str(date_str).strip()
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s.split()[0]).date()
    except Exception:
        return None

def to_ddmmyyyy(date_or_str):
    """Mengubah tanggal menjadi format DD-MM-YYYY."""
    d = parse_date_flexible(date_or_str)
    return d.strftime("%d-%m-%Y") if d else ""

# ---------------- Excel helpers ------------------------

def ensure_excel():
    """Memastikan file Excel ada, jika tidak, membuatnya."""
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        try:
            df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        except Exception:
             messagebox.showerror("Error", "Gagal membuat file Excel. Pastikan 'openpyxl' terinstal.")

def load_data():
    """Memuat data dari file Excel."""
    ensure_excel()
    try:
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        for c in COLUMNS:
            if c not in df.columns:
                df[c] = ""
        return df.reindex(columns=COLUMNS).copy()
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membaca data: {e}\nPastikan file Excel tidak dibuka.")
        return pd.DataFrame(columns=COLUMNS)

def save_data(df):
    """Menyimpan DataFrame ke file Excel."""
    try:
        df = df.reindex(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan data: {e}\nPastikan file Excel tidak dibuka.")

def insert_image_into_excel_last_row(foto_path):
    """Menyisipkan gambar ke kolom 'Foto' pada baris terakhir Excel."""
    if not (PIL_AVAILABLE and OPENPYXL_AVAILABLE):
        return
    if not foto_path or not os.path.exists(foto_path):
        return
        
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        if "Foto" not in headers:
            wb.close()
            return
        
        target_row = ws.max_row
        if target_row < 2:
            wb.close()
            return

        col_idx = headers.index("Foto") + 1
        
        img = Image.open(foto_path)
        max_w = 160
        w, h = img.size
        
        if w > max_w:
            ratio = max_w / float(w)
            try:
                img = img.resize((max_w, int(h * ratio)), Image.Resampling.LANCZOS)
            except AttributeError:
                img = img.resize((max_w, int(h * ratio)), Image.LANCZOS)
        
        tmp_path = os.path.join(os.path.dirname(EXCEL_FILE), "__tmp_img.png")
        img.save(tmp_path)

        xi = XLImage(tmp_path)
        xi.anchor = f"{get_column_letter(col_idx)}{target_row}"
        ws.add_image(xi)
        
        wb.save(EXCEL_FILE)
        wb.close()
        
        try:
            os.remove(tmp_path)
        except Exception:
            pass
            
    except Exception as e:
        print("Warning: gagal sisip gambar ke Excel:", e) 

# ---------------- PDF helpers ------------------------

def _write_report_to_pdf_page(pdf: "FPDF", row: pd.Series):
    """Menulis konten laporan untuk satu halaman PDF."""
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Laporan - koordinasi dengan kepala regu terkait informasi kejadian kebakaran", ln=True, align="C")
    pdf.ln(4)
    pdf.set_font("Arial", "", 11)
    
    # Data Identitas
    for field in ["Nama", "NIP", "Jabatan", "Unit Kerja", "Tanggal", "Waktu"]:
        pdf.cell(36, 7, f"{field}:", 0)
        val = str(row.get(field, "") or "")
        if field == "Tanggal":
            val = to_ddmmyyyy(val)
        pdf.cell(0, 7, val, ln=True)
    
    pdf.ln(3)
    
    # Kunci kolom laporan: Uraian Kejadian, Waktu Kebakaran, Kerusakan, Tindakan
    for field in ["Uraian Kejadian", "Waktu Kebakaran", "Kerusakan", "Tindakan"]:
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 7, f"{field}:", ln=True)
        pdf.set_font("Arial", "", 11)
        text = str(row.get(field, "") or "")
        pdf.multi_cell(0, 6, text)
        pdf.ln(2)
        
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 7, "Foto Bukti:", ln=True)
    foto_path = row.get("Foto", "")
    
    if pd.notna(foto_path) and foto_path and os.path.exists(foto_path):
        try:
            pdf.ln(3)
            pdf.image(foto_path, w=120)
            pdf.ln(4)
        except Exception:
            pdf.set_font("Arial", "", 11)
            pdf.cell(0, 7, "[Gagal menampilkan foto. Pastikan format JPG/PNG.]", ln=True)
    else:
        pdf.set_font("Arial", "", 11)
        pdf.cell(0, 7, "[Tidak ada foto]", ln=True)
        
    pdf.ln(6)
    pdf.set_font("Arial", "I", 9)
    pdf.cell(0, 7, f"Dibuat: {row.get('Generated At', '')}", ln=True, align="R") 

def export_all_to_single_pdf(df, out_path):
    """Export semua data laporan ke satu file PDF."""
    if FPDF is None:
        messagebox.showwarning("Missing lib", "Install 'fpdf' (pip install fpdf) untuk export PDF.")
        return
    
    pdf = FPDF()
    for idx, row in df.iterrows():
        pdf.add_page()
        _write_report_to_pdf_page(pdf, row)
        
    try:
        pdf.output(out_path)
        messagebox.showinfo("Sukses", f"PDF semua laporan dibuat:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membuat PDF: {e}")

def export_each_row_to_pdf_files(df, out_dir):
    """Export setiap baris laporan ke file PDF terpisah."""
    if FPDF is None:
        messagebox.showwarning("Missing lib", "Install 'fpdf' (pip install fpdf) untuk export PDF.")
        return
    
    os.makedirs(out_dir, exist_ok=True)
    created = []
    
    for idx, row in df.iterrows():
        # Membuat nama file yang aman
        nama_safe = "".join(c for c in str(row.get("Nama","")) if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "") or f"row{idx}"
        tanggal_safe = str(row.get("Tanggal","")).replace(":", "-").replace("/", "-")
        filename = os.path.join(out_dir, f"laporan_{nama_safe}_{tanggal_safe}_{idx+1}.pdf")
        
        pdf = FPDF()
        pdf.add_page()
        _write_report_to_pdf_page(pdf, row)
        
        try:
            pdf.output(filename)
            created.append(filename)
        except Exception:
            pass
            
    if created:
        messagebox.showinfo("Sukses", f"{len(created)} file PDF dibuat di:\n{out_dir}")
    else:
        messagebox.showwarning("Hasil", "Tidak ada file PDF berhasil dibuat.") 

# ---------------- small UI helpers ------------------------

def style_button_hover(btn, normal_bg=None, hover_bg=None, active_bg=None):
    """Menambahkan efek hover pada widget tk.Button."""
    try:
        # ttk.Button tidak perlu fungsi ini
        if isinstance(btn, type(object)): # Ganti dengan class ttk.Button yang benar jika perlu
             pass
    except Exception:
        pass
    
    try:
        current = btn.cget("bg")
    except Exception:
        current = normal_bg or "#f0f0f0"
        
    normal = normal_bg or current
    hover = hover_bg or normal
    active = active_bg or hover
    
    try:
        btn.config(bg=normal, activebackground=active, cursor="hand2")
    except Exception:
        pass
        
    def on_enter(e):
        try:
            btn['bg'] = hover
        except Exception:
            pass
            
    def on_leave(e):
        try:
            btn['bg'] = normal
        except Exception:
            pass
            
    def on_press(e):
        try:
            btn['bg'] = active
        except Exception:
            pass
            
    def on_release(e):
        try:
            btn['bg'] = hover
        except Exception:
            pass
            
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    btn.bind("<ButtonPress-1>", on_press)
    btn.bind("<ButtonRelease-1>", on_release) 

# ---------------- identity persistence ------------------------

def load_identitas():
    """Memuat identitas ASN yang tersimpan dari JSON."""
    if os.path.exists(IDENT_FILE):
        try:
            with open(IDENT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_identitas(data):
    """Menyimpan identitas ASN ke file JSON."""
    try:
        with open(IDENT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Gagal menyimpan identitas:", e)

def hapus_identitas_file():
    """Menghapus file identitas ASN yang tersimpan."""
    try:
        if os.path.exists(IDENT_FILE):
            os.remove(IDENT_FILE)
    except Exception as e:
        print("Gagal hapus identitas:", e)
