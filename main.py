import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import pandas as pd
from pandas import Series



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

# ---------------- CONFIG ----------------

# !!! JALUR PENYIMPANAN ANDA !!!
BASE_PATH = "D:\SNEIJDERLINO\MY ASN\APLIKASI LAPORAN MY ASN\DATABASE_APLIKASI _LAPORAN_E-KINERJA"

EXCEL_FILE = os.path.join(BASE_PATH, "laporan_asn.xlsx")
IDENT_FILE = os.path.join(BASE_PATH, "identitas_asn.json")

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




def parse_date_flexible(date_str):
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
    d = parse_date_flexible(date_or_str)
    return d.strftime("%d-%m-%Y") if d else ""



def ensure_base_dir(file_path):
    """Memastikan direktori tempat file akan disimpan itu ada."""
    dir_name = os.path.dirname(file_path)
    if dir_name and not os.path.exists(dir_name):
        try:

            os.makedirs(dir_name)
        except Exception as e:

            messagebox.showerror("Error Folder", f"Gagal membuat folder tujuan: {dir_name}\n{e}")
            raise # Menghentikan eksekusi jika gagal membuat folder

# ----------------- Excel helpers ------------------------

def ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

def load_data():

    try:
        ensure_base_dir(EXCEL_FILE) 
    except Exception:
        return pd.DataFrame(columns=COLUMNS) 
        
    ensure_excel()
    try:
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

        for c in COLUMNS:
            if c not in df.columns:
                df[c] = ""

        return df.reindex(columns=COLUMNS).copy()
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membaca data: {e}")
        return pd.DataFrame(columns=COLUMNS)

def save_data(df):
    try:
        ensure_base_dir(EXCEL_FILE) 
        df = df.reindex(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan data: {e}")

def _insert_image_into_excel_last_row(foto_path):
    if not (PIL_AVAILABLE and OPENPYXL_AVAILABLE):
        return
    if not foto_path or not os.path.exists(foto_path):
        return
    try:

        ensure_base_dir(EXCEL_FILE)
        
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
        
        # Simpan file sementara di folder yang sama
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



def _write_report_to_pdf_page(pdf: "FPDF", row: pd.Series):
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
    if FPDF is None:
        messagebox.showwarning("Missing lib", "Install 'fpdf' (pip install fpdf) untuk export PDF.")
        return
    
    os.makedirs(out_dir, exist_ok=True)
    created = []
    
    for idx, row in df.iterrows():

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



def style_button_hover(btn, normal_bg=None, hover_bg=None, active_bg=None):
    try:
        if isinstance(btn, ttk.Button):
            return
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

# ----------------- identity persistence ------------------------

def load_identitas():

    try:
        ensure_base_dir(IDENT_FILE)
    except Exception:
        return {}
        
    if os.path.exists(IDENT_FILE):
        try:
            with open(IDENT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_identitas(data):
    try:
        ensure_base_dir(IDENT_FILE) 
        with open(IDENT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Gagal menyimpan identitas:", e)

def hapus_identitas_file():
    try:
        if os.path.exists(IDENT_FILE):
            os.remove(IDENT_FILE)
    except Exception as e:
        print("Gagal hapus identitas:", e)

# ----------------- Main App  ------------------------

class LaporanApp:
    def __init__(self, root):
        self.root = root
        root.title("Aplikasi Ekinerja — koordinasi dengan kepala regu terkait informasi kejadian kebakaran")
        root.geometry("1280x900")
        self.selected_foto_path = ""
        self.preview_img = None
        self.initial_preview_size = (380, 320)
        
        # --- TEMA/STYLING PROFESIONAL --------
        style = ttk.Style()
        style.theme_use('clam')
        
        # WARNA
        self.PRIMARY_COLOR = "#5D4AA0"
        self.ACCENT_COLOR = "#785BB8"
        self.LIGHT_BG = "#F4F4F9"
        self.DARK_TEXT = "#333333"
        self.WHITE_TEXT = "white"
        
        root.configure(bg=self.LIGHT_BG)

        # Style Header
        style.configure("Header.TFrame", background=self.PRIMARY_COLOR)
        style.configure("Header.TLabel", background=self.PRIMARY_COLOR, foreground=self.WHITE_TEXT, font=("Segoe UI", 16, "bold"))
        style.configure("SubHeader.TLabel", background=self.PRIMARY_COLOR, foreground=self.WHITE_TEXT, font=("Segoe UI", 12))
        
        # Style Frame Input
        style.configure("Form.TFrame", background=self.LIGHT_BG)
        style.configure("FormLabel.TLabel", background=self.LIGHT_BG, foreground=self.DARK_TEXT)
        
        # Style Tombol
        style.configure("TButton", font=("Segoe UI", 10), padding=6, background="#E0E0E0", foreground=self.DARK_TEXT)
        style.map("TButton", background=[('active', '#C0C0C0')])
        
        # Style Tombol Aksi Utama (CRUD & Export)
        style.configure("Add.TButton", background="#4CAF50", foreground=self.WHITE_TEXT)
        style.map("Add.TButton", background=[('active', '#66BB6A')])
        style.configure("Edit.TButton", background="#2196F3", foreground=self.WHITE_TEXT)
        style.map("Edit.TButton", background=[('active', '#42A5F5')])
        style.configure("Delete.TButton", background="#F44336", foreground=self.WHITE_TEXT)
        style.map("Delete.TButton", background=[('active', '#E57373')])
        style.configure("Reset.TButton", background="#FF9800", foreground=self.WHITE_TEXT)
        style.map("Reset.TButton", background=[('active', '#FFB74D')])
        style.configure("Upload.TButton", background=self.ACCENT_COLOR, foreground=self.WHITE_TEXT)
        style.map("Upload.TButton", background=[('active', self.PRIMARY_COLOR)])
        
        # Style Treeview (Tabel)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=self.ACCENT_COLOR, foreground=self.WHITE_TEXT)
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        style.map("Treeview", background=[('selected', self.ACCENT_COLOR)], foreground=[('selected', self.WHITE_TEXT)])

        # Header
        header = ttk.Frame(root, style="Header.TFrame", height=56)
        header.pack(fill="x")
        ttk.Label(header, text="APLIKASI BY SNEIJDERLINO ", style="Header.TLabel").pack(side="left", padx=12, pady=8)
        ttk.Label(header, text="Laporan koordinasi dengan kepala regu terkait informasi kejadian kebakaran", style="SubHeader.TLabel").pack(side="left", padx=12)
        btn_logout = tk.Button(header, text="Log Out", bg="#DCD7FF", relief="flat", command=self.on_logout)
        btn_logout.pack(side="right", padx=12, pady=8)
        style_button_hover(btn_logout, normal_bg="#DCD7FF", hover_bg="#E9E6FF", active_bg="#C0B5FF")

        # Form top
        form_container = ttk.Frame(root, style="Form.TFrame", padding="10 8 10 8")
        form_container.pack(fill="x")
        form_frame = ttk.Frame(form_container, style="Form.TFrame")
        form_frame.pack(fill="x")

        # Left basic fields
        left = ttk.Frame(form_frame, style="Form.TFrame", padding="6 4 12 4")
        left.pack(side="left")
        
        self.entries = {}
        labels = ["Nama", "NIP", "Jabatan", "Unit Kerja", "Tanggal", "Waktu"]
        for i, field in enumerate(labels):
            lbl = ttk.Label(left, text=field + ":", style="FormLabel.TLabel", anchor="w")
            lbl.grid(row=i, column=0, sticky="w", pady=4, padx=(0, 8))
            ent = ttk.Entry(left, width=35)
            ent.grid(row=i, column=1, pady=4)
            self.entries[field] = ent
        
        self.entries["Tanggal"].insert(0, datetime.now().strftime("%d-%m-%Y"))
        self.entries["Waktu"].insert(0, datetime.now().strftime("%H:%M:%S"))
        
        # --- Modifikasi/Panggil untuk Identitas ---
        self._load_and_lock_identitas()
        self.entries["Nama"].bind("<Double-1>", self._on_nama_double_click)
        self._create_identitas_context_menu()
        
        # Middle multiline
        mid = ttk.Frame(form_frame, style="Form.TFrame", padding="6 4")
        mid.pack(side="left", padx=15)
        self.texts = {}
        
        # KOLOM MULTILINE YANG BARU:
        multiline = ["Uraian Kejadian", "Waktu Kebakaran", "Kerusakan", "Tindakan"]
        for i, field in enumerate(multiline):
            lbl = ttk.Label(mid, text=field + ":", style="FormLabel.TLabel", anchor="w")
            lbl.grid(row=i*2, column=0, sticky="w", pady=(4,0))
            
            txt_frame = ttk.Frame(mid)
            txt_frame.grid(row=i*2+1, column=0, pady=(2, 10))
            
            txt = tk.Text(txt_frame, width=60, height=4, wrap="word", relief="flat", borderwidth=1, highlightthickness=1, highlightcolor=self.PRIMARY_COLOR, highlightbackground="#CCCCCC")
            txt.pack(side="left", fill="y", expand=True)
            
            vscroll = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
            vscroll.pack(side="right", fill="y")
            txt.config(yscrollcommand=vscroll.set)
            self.texts[field] = txt

        # Right preview
        right = ttk.Frame(form_frame, style="Form.TFrame", padding="8 4")
        right.pack(side="left", fill="y", padx=(20, 0))
        
        preview_container = ttk.LabelFrame(right, text="Preview Foto", padding="10 10 10 10")
        preview_container.pack(padx=6, pady=4, fill="both", expand=True)
        
        self.foto_label = tk.Label(preview_container, text="[Area Preview Foto]", width=48, height=18, relief="flat", bg="#EEEEEE", fg="#888888", anchor="center", font=("Segoe UI", 10))
        self.foto_label.pack(fill="both", expand=True)
        self.foto_label.bind("<Configure>", self._on_preview_resize)
        
        btn_frame_foto = ttk.Frame(right, padding="0 8 0 0")
        btn_frame_foto.pack(fill="x")
        
        btn_upload = ttk.Button(btn_frame_foto, text="📸 Upload Foto", command=self.pilih_foto, style="Upload.TButton")
        btn_upload.pack(side="left", padx=(0, 10))
        btn_hapus_foto = ttk.Button(btn_frame_foto, text="🗑 Hapus Foto", command=self.reset_foto)
        btn_hapus_foto.pack(side="left")

        # Actions row (CRUD) - TOMBOL EXPORT
        actions = ttk.Frame(root, padding="12 6")
        actions.pack(fill="x")

        # LOGIKA DI TAMBAH/EDIT LAPORAN
        btn_tambah = ttk.Button(actions, text="Tambah Laporan", width=18, command=self.tambah_laporan, style="Add.TButton")
        btn_tambah.pack(side="left", padx=6)
        btn_edit = ttk.Button(actions, text="Update Laporan", width=18, command=self.edit_laporan, style="Edit.TButton")
        btn_edit.pack(side="left", padx=6)
        btn_hapus = ttk.Button(actions, text="Hapus Laporan", width=18, command=self.hapus_laporan, style="Delete.TButton")
        btn_hapus.pack(side="left", padx=6)
        btn_reset = ttk.Button(actions, text="Reset Form", width=18, command=self.reset_form, style="Reset.TButton")
        btn_reset.pack(side="left", padx=12)

        # ------------------- EXPORT BUTTONS (DIPINDAH) -------------------
        export_frame = ttk.LabelFrame(actions, text="Export Data", padding="8 6 8 6")
        export_frame.pack(side="left", padx=12)
        
        btn_export_sel = ttk.Button(export_frame, text="Export Selected → PDF", command=self.export_selected_pdf, style="Edit.TButton")
        btn_export_sel.pack(side="left", padx=4)
        btn_export_all_pdf = ttk.Button(export_frame, text="Export Semua → 1 PDF", command=self.export_all_single_pdf, style="Add.TButton")
        btn_export_all_pdf.pack(side="left", padx=4)
        btn_export_many = ttk.Button(export_frame, text="Export Semua → Banyak PDFs", command=self.export_all_multiple_pdfs, style="Reset.TButton")
        btn_export_many.pack(side="left", padx=4)
        btn_export_excel = ttk.Button(export_frame, text="Export Semua → Excel", command=self.export_all_excel, style="Edit.TButton")
        btn_export_excel.pack(side="left", padx=4)
        # -----------------------------------------------------------------

        # Filter box
        filter_frame = ttk.LabelFrame(root, text="Filter / Pencarian", padding="8 6 8 6")
        filter_frame.pack(fill="x", padx=12, pady=6)
        
        ttk.Label(filter_frame, text="Kolom:", style="FormLabel.TLabel").grid(row=0, column=0, padx=4, sticky="w")
        self.filter_col = ttk.Combobox(filter_frame, values=["Nama", "NIP", "Jabatan", "Unit Kerja", "Uraian Kejadian", "Waktu Kebakaran", "Kerusakan", "Tindakan"], width=18, state="readonly")
        self.filter_col.set("Nama")
        self.filter_col.grid(row=0, column=1, padx=4, pady=2)
        
        ttk.Label(filter_frame, text="Keyword:", style="FormLabel.TLabel").grid(row=0, column=2, padx=4, sticky="w")
        self.filter_kw = ttk.Entry(filter_frame, width=20)
        self.filter_kw.grid(row=0, column=3, padx=4, pady=2)
        
        ttk.Label(filter_frame, text="Tanggal (DD-MM-YYYY) From:", style="FormLabel.TLabel").grid(row=0, column=4, padx=8, sticky="w")
        self.date_from = ttk.Entry(filter_frame, width=12)
        self.date_from.grid(row=0, column=5, padx=4, pady=4)
        
        ttk.Label(filter_frame, text="To:", style="FormLabel.TLabel").grid(row=0, column=6, padx=4, sticky="w")
        self.date_to = ttk.Entry(filter_frame, width=12)
        self.date_to.grid(row=0, column=7, padx=4, pady=4)
        
        btn_apply_filter = ttk.Button(filter_frame, text="🔍 Terapkan Filter", command=self.apply_filter, style="Edit.TButton")
        btn_apply_filter.grid(row=0, column=8, padx=8)
        
        btn_reset_filter = ttk.Button(filter_frame, text="🔄 Reset Filter", command=self.reset_filter, style="Reset.TButton")
        btn_reset_filter.grid(row=0, column=9, padx=6)
        
        filter_frame.grid_columnconfigure(10, weight=1) # Agar tombol terdorong ke kiri

        # Table area
        table_frame = ttk.Frame(root, padding="12 8 12 8")
        table_frame.pack(fill="both", expand=True)
        
        cols = COLUMNS.copy()
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        
        for c in cols:
            self.tree.heading(c, text=c)
            if c in ("Uraian Kejadian","Waktu Kebakaran","Kerusakan","Tindakan"):
                self.tree.column(c, width=180, anchor="w")
            elif c=="Foto":
                self.tree.column(c, width=100, anchor="center")
            elif c=="Generated At":
                self.tree.column(c, width=140, anchor="center")
            else:
                self.tree.column(c, width=100, anchor="w")
                
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", self.on_row_double_click)
        self.tampilkan_data()

    # ---------------- identity functions ----------------
    
    def _load_and_lock_identitas(self):
        data = load_identitas()
        for k in ["Nama","NIP","Jabatan","Unit Kerja"]:
            ent = self.entries.get(k)
            if ent:

                try:
                    ent.config(state="normal")
                    ent.delete(0, "end")
                except Exception:
                    pass
                

                if data and k in data:
                    ent.insert(0, data.get(k, ""))
                    try:
                        ent.config(state="readonly")
                    except Exception:
                        pass

    def _create_identitas_context_menu(self):
        self.ident_menu = tk.Menu(self.root, tearoff=0)
        self.ident_menu.add_command(label="Hapus identitas tersimpan", command=self._on_hapus_identitas)
        self.entries["Nama"].bind("<Button-3>", lambda e: self.ident_menu.post(e.x_root, e.y_root))

    def _on_nama_double_click(self, event=None):
        ent = self.entries["Nama"]
        state = ent.cget("state")
        
        if state == "readonly":
            # Buka kunci untuk edit
            for k in ["Nama","NIP","Jabatan","Unit Kerja"]:
                try:
                    self.entries[k].config(state="normal")
                except Exception:
                    pass
            messagebox.showinfo("Edit Identitas", "Field identitas dibuka. Ubah data lalu double-click lagi di Nama untuk menyimpan.")
            return
        else:
            # Simpan dan kunci
            data = {}
            for k in ["Nama","NIP","Jabatan","Unit Kerja"]:
                data[k] = self.entries[k].get().strip()
            
            if not data.get("Nama"):
                messagebox.showwarning("Validasi", "Nama tidak boleh kosong saat menyimpan identitas.")
                return
            
            nip_val = data.get("NIP","")
            if nip_val and not self._validate_nip(nip_val):
                messagebox.showwarning("Validasi", "NIP harus berupa angka dan minimal 5 digit.")
                return
            
            save_identitas(data)
            
            for k in ["Nama","NIP","Jabatan","Unit Kerja"]:
                try:
                    self.entries[k].config(state="readonly")
                except Exception:
                    pass
                    
            messagebox.showinfo("Sukses", "Identitas ASN disimpan dan dikunci. (Double-click Nama untuk ubah lagi)")

    def _on_hapus_identitas(self):
        if not os.path.exists(IDENT_FILE):
            messagebox.showinfo("Info", "Tidak ada identitas tersimpan.")
            return
            
        if not messagebox.askyesno("Konfirmasi", "Yakin ingin menghapus identitas ASN yang tersimpan?"):
            return
            
        hapus_identitas_file()
        
        # Buka kunci field dan kosongkan isinya
        for k in ["Nama","NIP","Jabatan","Unit Kerja"]:
            try:
                self.entries[k].config(state="normal")
                self.entries[k].delete(0, "end")
            except Exception:
                pass
                
        messagebox.showinfo("Sukses", "Identitas ASN berhasil dihapus.")
        
    # menyimpan identitas secara otomatis saat pertama kali tambah laporan
    def _save_new_identitas_if_needed(self):

        if self.entries.get("Nama") and self.entries["Nama"].cget("state") == "readonly":
            return 
            
        data = {}
        for k in ["Nama", "NIP", "Jabatan", "Unit Kerja"]:
            data[k] = self.entries[k].get().strip()
        

        if data.get("Nama") and data.get("NIP") and self._validate_nip(data.get("NIP")):
            if messagebox.askyesno("Simpan Identitas?", 
                                   "Apakah Anda ingin **menyimpan** Nama, NIP, Jabatan, dan Unit Kerja ini agar terisi otomatis dan terkunci di laporan berikutnya?"):
                save_identitas(data)

                for k in ["Nama", "NIP", "Jabatan", "Unit Kerja"]:
                    try:
                        self.entries[k].config(state="readonly")
                    except Exception:
                        pass


    # ---------------- photo handlers ----------------
    
    def pilih_foto(self):
        if not PIL_AVAILABLE:
            messagebox.showwarning("Missing lib", "Install 'Pillow' (pip install pillow) untuk upload & preview foto.")
            return
            
        path = filedialog.askopenfilename(title="Pilih foto", filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp")])
        if path:
            self.selected_foto_path = path
            self._render_preview_image(path)

    def _on_preview_resize(self, event):
        if self.selected_foto_path:
            self._render_preview_image(self.selected_foto_path)

    def _render_preview_image(self, path):
        if not PIL_AVAILABLE:
            return
        
        try:
            img = Image.open(path)
            lbl = self.foto_label
            
            w = lbl.winfo_width()
            h = lbl.winfo_height()
            if w <= 1 or h <= 1:
                w, h = self.initial_preview_size
                
            iw, ih = img.size
            ratio = min(w / iw, h / ih)
            if ratio <= 0:
                ratio = 1.0
            
            new_size = (max(1, int(iw * ratio)), max(1, int(ih * ratio)))
            
            # Use Image.Resampling.LANCZOS or Image.LANCZOS depending on PIL version
            try:
                img_resized = img.resize(new_size, Image.Resampling.LANCZOS)
            except AttributeError:
                img_resized = img.resize(new_size, Image.LANCZOS)
                
            photo = ImageTk.PhotoImage(img_resized)
            lbl.configure(image=photo, text="")
            lbl.image = photo # Keep a reference
            self.selected_foto_path = path
            
        except Exception as e:
            lbl = self.foto_label
            lbl.configure(image="", text="[Preview gagal atau format tidak didukung]", fg="red")
            print("Preview error:", e)

    def reset_foto(self):
        self.selected_foto_path = ""
        self.foto_label.configure(image="", text="[Area Preview Foto]", bg="#EEEEEE", fg="#888888")
        try:
            del self.foto_label.image
        except Exception:
            pass

    # ---------------- form helpers ----------------
    
    def get_form_data(self):
        data = {}
        
        for k, ent in self.entries.items():
            val = ent.get().strip()
            
            if k == "Tanggal":
                if val:
                    parsed = parse_date_flexible(val)
                    if parsed:
                        val = parsed.strftime("%d-%m-%Y")
                    else:

                        val = datetime.now().strftime("%d-%m-%Y") 
                
            data[k] = val
            

        texts_map = {
            "Uraian Kejadian": self.texts["Uraian Kejadian"],
            "Waktu Kebakaran": self.texts["Waktu Kebakaran"],
            "Kerusakan": self.texts["Kerusakan"],
            "Tindakan": self.texts["Tindakan"],
        }
        
        for k, txt in texts_map.items():
            data[k] = txt.get("1.0", "end").strip()
            
        data["Foto"] = self.selected_foto_path or ""
        data["Generated At"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        
        return data

    def set_form_data_from_row(self, row):

        for k in ["Nama","NIP","Jabatan","Unit Kerja","Tanggal","Waktu"]:
            ent = self.entries.get(k)
            if not ent: continue
            
            val = row.get(k,"")
            if pd.isna(val): val = ""
                
            if k == "Tanggal" and val:
                parsed = parse_date_flexible(val)
                if parsed:
                    val = parsed.strftime("%d-%m-%Y")
            
            current_state = ent.cget("state")
            

            if current_state == "readonly":
                ent.config(state="normal")
                ent.delete(0, "end")
                ent.insert(0, str(val))
                ent.config(state="readonly")
            else:
                ent.delete(0, "end")
                ent.insert(0, str(val))


        texts_map = {
            "Uraian Kejadian": self.texts["Uraian Kejadian"],
            "Waktu Kebakaran": self.texts["Waktu Kebakaran"],
            "Kerusakan": self.texts["Kerusakan"],
            "Tindakan": self.texts["Tindakan"],
        }
        
        for k in texts_map.keys():
            txt = self.texts.get(k)
            if not txt: continue
            
            txt.delete("1.0", "end")
            v = row.get(k,"")
            if pd.isna(v): v = ""
            txt.insert("1.0", str(v))

        # 3. Handle Foto
        foto = row.get("Foto","")
        if pd.notna(foto) and foto and os.path.exists(foto):
            self.selected_foto_path = foto
            if PIL_AVAILABLE:
                self._render_preview_image(foto)
            else:
                self.foto_label.configure(text="[Foto ada, instal Pillow untuk preview]")
        else:
            self.reset_foto()

    def tampilkan_data(self, filtered_df=None):
        for r in self.tree.get_children():
            self.tree.delete(r)
            
        df = filtered_df if filtered_df is not None else load_data()
        
        for idx, row in df.reset_index(drop=True).iterrows():
            vals = []
            for c in COLUMNS:
                v = row.get(c,"")
                if pd.isna(v): v = ""
                
                if c == "Tanggal" and v:
                    parsed = parse_date_flexible(v)
                    if parsed:
                        v = parsed.strftime("%d-%m-%Y")
                        
                vals.append(str(v))
                
            self.tree.insert("", "end", iid=str(idx), values=vals)

    def _validate_date(self, s):
        return parse_date_flexible(s) is not None

    def _validate_time(self, s):
        if not s: 
            return False
        for fmt in ("%H:%M:%S","%H:%M"):
            try:
                datetime.strptime(s, fmt)
                return True
            except Exception:
                continue
        return False

    def _validate_nip(self, nip):
        n = str(nip).strip()

        return n.isdigit() and len(n) >= 5 

    # FUNGSI BARU UNTUK PREVIEW SESUAI GAMBAR
    def _show_data_preview_popup(self, data, action_text, confirm_callback):
        # Membuat jendela pop-up
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Konfirmasi Laporan")
        # Mengubah ukuran agar lebih pas dengan konten
        preview_window.geometry("850x550") 
        preview_window.transient(self.root) # Tetap di atas jendela utama
        preview_window.grab_set() # Fokus di jendela ini

        # Gaya untuk jendela pop-up
        style = ttk.Style()
        style.configure("Preview.TLabel", background=self.LIGHT_BG, foreground=self.DARK_TEXT)
        style.configure("PreviewHeader.TLabel", background=self.ACCENT_COLOR, foreground=self.WHITE_TEXT, font=("Segoe UI", 10, "bold"))
        
        main_frame = ttk.Frame(preview_window, padding="10", style="Form.TFrame")
        main_frame.pack(fill="both", expand=True)

        # Header preview
        ttk.Label(main_frame, text=f"Periksa sebelum **{action_text.upper()}**:", style="Preview.TLabel", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10), anchor="w")

        # --- Tampilan Data (Mirip Gambar) ---
        preview_cols = ["NIP", "Jabatan", "Unit Kerja", "Tanggal", "Waktu", "Uraian Kejadian"]
        preview_tree = ttk.Treeview(main_frame, columns=preview_cols, show="headings", selectmode="none", height=1) 
        
        # Styling Header Preview
        style.configure("PreviewTree.Heading", font=("Segoe UI", 10, "bold"), background="#333333", foreground="white") # DARK BG HEADERS
        preview_tree.tag_configure('oddrow', background='white')
        preview_tree.tag_configure('evenrow', background='#EFEFEF')
        
        for col in preview_cols:
            preview_tree.heading(col, text=col.replace(" ", "\n"))
            if col in ("Tanggal", "Waktu"):
                preview_tree.column(col, width=80, anchor="center")
            elif col == "NIP":
                preview_tree.column(col, width=120, anchor="w")
            elif col == "Uraian Kejadian":
                preview_tree.column(col, width=250, anchor="w")
            else:
                preview_tree.column(col, width=120, anchor="w")
                
        preview_tree.pack(fill="x", padx=5, pady=5)

        # Data untuk preview baris
        preview_values = []
        for col in preview_cols:
            val = data.get(col, "")
            if col == "Tanggal":
                val = to_ddmmyyyy(val)

            if col == "Uraian Kejadian":
                val = (val[:80] + '...') if len(val) > 80 else val
            preview_values.append(str(val))
            
        preview_tree.insert("", "end", values=preview_values, tags=('oddrow',))


        detail_frame = ttk.Frame(main_frame, style="Form.TFrame", padding="5")
        detail_frame.pack(fill="both", expand=True, pady=(15, 0))


        detail_frame.grid_columnconfigure(0, weight=0) 
        detail_frame.grid_columnconfigure(1, weight=1)


        ttk.Label(detail_frame, text="Kerusakan:", style="Preview.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="nw", padx=(0, 10), pady=(5, 0))
        kerusakan_text = tk.Text(detail_frame, width=50, height=4, wrap="word", relief="flat", borderwidth=1, state="disabled")
        kerusakan_text.grid(row=0, column=1, sticky="ew", padx=(0, 15), pady=(5, 10))
        kerusakan_text.config(state="normal")
        kerusakan_text.insert("1.0", data.get("Kerusakan", "Tidak ada kerusakan tercatat."))
        kerusakan_text.config(state="disabled")


        ttk.Label(detail_frame, text="Tindakan:", style="Preview.TLabel", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="nw", padx=(0, 10), pady=(5, 0))
        tindakan_text = tk.Text(detail_frame, width=50, height=4, wrap="word", relief="flat", borderwidth=1, state="disabled")
        tindakan_text.grid(row=1, column=1, sticky="ew", padx=(0, 15), pady=(5, 10))
        tindakan_text.config(state="normal")
        tindakan_text.insert("1.0", data.get("Tindakan", "Tidak ada tindakan tercatat."))
        tindakan_text.config(state="disabled")

        ttk.Label(detail_frame, text="Foto Bukti:", style="Preview.TLabel", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="nw", pady=(10, 0))
        
        foto_path = data.get("Foto", "")
        if PIL_AVAILABLE and foto_path and os.path.exists(foto_path):
            img = Image.open(foto_path)

            w, h = img.size
            max_w = 200
            ratio = max_w / float(w)
            
            try:
                img_resized = img.resize((max_w, int(h * ratio)), Image.Resampling.LANCZOS)
            except AttributeError:
                img_resized = img.resize((max_w, int(h * ratio)), Image.LANCZOS)
                
            photo = ImageTk.PhotoImage(img_resized)
            foto_label = tk.Label(detail_frame, image=photo, relief="solid", borderwidth=1)
            foto_label.image = photo
            foto_label.grid(row=2, column=1, sticky="w", pady=(5, 0))
        else:
            ttk.Label(detail_frame, text="[Tidak ada foto / Gagal load]", style="Preview.TLabel", foreground="red").grid(row=2, column=1, sticky="w", pady=(5, 0))

        # --- Tombol Aksi ---
        btn_frame = ttk.Frame(main_frame, style="Form.TFrame")
        btn_frame.pack(fill="x", pady=15)
        
        # Tombol Batal
        btn_batal = tk.Button(btn_frame, text="X Batal", command=preview_window.destroy, bg="#F44336", fg="white", font=("Segoe UI", 10, "bold"), padx=10, pady=5, relief="flat")
        btn_batal.pack(side="right", padx=10)
        style_button_hover(btn_batal, normal_bg="#F44336", hover_bg="#E57373")
        
        # Tombol Konfirmasi
        btn_konfirmasi = tk.Button(btn_frame, text="✔ Konfirmasi", command=lambda: [preview_window.destroy(), confirm_callback()], bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"), padx=10, pady=5, relief="flat")
        btn_konfirmasi.pack(side="right")
        style_button_hover(btn_konfirmasi, normal_bg="#4CAF50", hover_bg="#66BB6A")
        
        preview_window.wait_window(preview_window)

    # ---------------- CRUD operations ----------------
    
    def tambah_laporan(self):
        
    
        self._save_new_identitas_if_needed()
        
        data = self.get_form_data()

        # --- Validasi ---
        if not data.get("Nama"):
            messagebox.showwarning("Validasi", "Nama harus diisi.")
            return
        if not self._validate_nip(data.get("NIP","")):
            messagebox.showwarning("Validasi", "NIP harus berupa angka dan minimal 5 digit.")
            return
        if not self._validate_date(data.get("Tanggal","")):
            messagebox.showwarning("Format tanggal", "Format 'Tanggal' tidak dikenali. Gunakan DD-MM-YYYY.")
            return
        if not self._validate_time(data.get("Waktu","")):
            messagebox.showwarning("Format waktu", "Format 'Waktu' tidak dikenali. Gunakan HH:MM:SS.")
            return

        # --- Panggil Preview Pop-up ---
        self._show_data_preview_popup(data, "Simpan", lambda: self._perform_tambah_laporan(data))

    def _perform_tambah_laporan(self, data):
        # --- Proses Simpan ---
        df = load_data()
        row_to_add = {c: data.get(c, "") for c in COLUMNS}
        
        # Pastikan format data tanggal sudah sesuai dengan output dari get_form_data
        
        df = pd.concat([df, pd.DataFrame([row_to_add], columns=COLUMNS)], ignore_index=True)
        save_data(df)
        
        # Hanya sisipkan gambar ke Excel jika tersedia
        if data.get("Foto") and OPENPYXL_AVAILABLE:
            _insert_image_into_excel_last_row(data.get("Foto"))
            
        self.tampilkan_data()
        messagebox.showinfo("Sukses", "Laporan berhasil ditambahkan.")
        self.reset_form()

    def hapus_laporan(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Peringatan", "Pilih baris yang akan dihapus.")
            return

        if not messagebox.askyesno("Konfirmasi", "Yakin menghapus laporan?"):
            return

        idx = int(sel[0])
        df = load_data()
        
        if idx < 0 or idx >= len(df):
            messagebox.showerror("Error", "Indeks tidak valid.")
            return

        df = df.drop(index=idx).reset_index(drop=True)
        save_data(df)
        
        self.tampilkan_data()
        messagebox.showinfo("Sukses", "Laporan dihapus.")

    def edit_laporan(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Peringatan", "Pilih baris yang akan diedit (double-click untuk load).")
            return
            
        idx = int(sel[0])
        df = load_data()
        
        if idx < 0 or idx >= len(df):
            messagebox.showerror("Error", "Indeks tidak valid.")
            return


        new = self.get_form_data()

        new["Generated At"] = df.at[idx, "Generated At"] if "Generated At" in df.columns and pd.notna(df.at[idx, "Generated At"]) else datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        # --- Validasi ---
        if not new.get("Nama"):
            messagebox.showwarning("Validasi", "Nama harus diisi.")
            return
        if not self._validate_nip(new.get("NIP","")):
            messagebox.showwarning("Validasi", "NIP harus berupa angka dan minimal 5 digit.")
            return
        if not self._validate_date(new.get("Tanggal","")):
            messagebox.showwarning("Format tanggal", "Format 'Tanggal' tidak dikenali. Gunakan DD-MM-YYYY.")
            return
        if not self._validate_time(new.get("Waktu","")):
            messagebox.showwarning("Format waktu", "Format 'Waktu' tidak dikenali. Gunakan HH:MM:SS.")
            return

        # --- Panggil Preview Pop-up ---
        self._show_data_preview_popup(new, "Update", lambda: self._perform_edit_laporan(idx, new))

    def _perform_edit_laporan(self, idx, new_data):
        # --- Proses Update ---
        df = load_data()
        
        # Lakukan pembaruan per kolom
        for k, v in new_data.items():
            if k not in df.columns:
                df[k] = ""
            df.at[idx, k] = v
            
        save_data(df)
        
        self.tampilkan_data()
        messagebox.showinfo("Sukses", "Laporan diperbarui.")

    def reset_form(self):

        for ent in self.entries.values():
            try:

                if ent.cget("state") != "readonly":
                    ent.delete(0, "end")
            except Exception:
                pass 

        texts_map = {
            "Uraian Kejadian": self.texts["Uraian Kejadian"],
            "Waktu Kebakaran": self.texts["Waktu Kebakaran"],
            "Kerusakan": self.texts["Kerusakan"],
            "Tindakan": self.texts["Tindakan"],
        }
        for txt in texts_map.values():
            txt.delete("1.0", "end")
            

        self.entries["Tanggal"].delete(0, "end")
        self.entries["Tanggal"].insert(0, datetime.now().strftime("%d-%m-%Y"))
        self.entries["Waktu"].delete(0, "end")
        self.entries["Waktu"].insert(0, datetime.now().strftime("%H:%M:%S"))
        

        self.reset_foto()

    def on_row_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
            
        idx = int(sel[0])
        df = load_data()
        
        if idx < 0 or idx >= len(df):
            return
            
        row = df.iloc[idx]
        self.set_form_data_from_row(row)

    # ---------------- Filter / Search ----------------
    
    def apply_filter(self):
        df = load_data()
        col = self.filter_col.get().strip()
        kw = self.filter_kw.get().strip().lower()
        df_filtered = df.copy()

        if kw:
            if col and col in df.columns:
                df_filtered = df_filtered[df_filtered[col].astype(str).str.lower().str.contains(kw, na=False)]
            else:
                mask = False
                for c in ["Nama","NIP","Jabatan","Unit Kerja","Uraian Kejadian","Waktu Kebakaran","Kerusakan","Tindakan"]:
                    if c in df_filtered.columns:
                        mask = mask | df_filtered[c].astype(str).str.lower().str.contains(kw, na=False)
                df_filtered = df_filtered[mask]
                
        df_filtered2 = df_filtered.copy()

        dfrom = self.date_from.get().strip()
        dto = self.date_to.get().strip()
        dfrom_dt = None
        dto_dt = None

        if dfrom:
            try:
                dfrom_dt = parse_date_flexible(dfrom)
                if not dfrom_dt: raise ValueError("Tanggal From tidak valid")
            except Exception:
                messagebox.showwarning("Format tanggal", "Format 'Tanggal From' harus DD-MM-YYYY atau YYYY-MM-DD")
                return
                
        if dto:
            try:
                dto_dt = parse_date_flexible(dto)
                if not dto_dt: raise ValueError("Tanggal To tidak valid")
            except Exception:
                messagebox.showwarning("Format tanggal", "Format 'Tanggal To' harus DD-MM-YYYY atau YYYY-MM-DD")
                return

        if dfrom_dt or dto_dt:
            def _date_in_range_filter(date_str):
                parsed = parse_date_flexible(date_str)
                if not parsed: return False
                
                in_range = True
                if dfrom_dt:
                    in_range = in_range and (parsed >= dfrom_dt)
                if dto_dt:
                    in_range = in_range and (parsed <= dto_dt)
                return in_range
                
            df_filtered2 = df_filtered2[df_filtered2["Tanggal"].astype(str).apply(_date_in_range_filter)]

        self.tampilkan_data(df_filtered2.reset_index(drop=True))

    def reset_filter(self):
        self.filter_kw.delete(0, "end")
        self.date_from.delete(0, "end")
        self.date_to.delete(0, "end")
        self.filter_col.set("Nama")
        self.tampilkan_data()

    # ---------------- Exports ----------------

    def export_selected_pdf(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Peringatan", "Pilih baris yang akan di-export.")
            return

        df = load_data().reset_index(drop=True)
        idx = int(sel[0])
        
        if idx < 0 or idx >= len(df):
            messagebox.showerror("Error", "Indeks tidak valid.")
            return
            
        row = df.iloc[idx]
        
        nama_safe = "".join(c for c in str(row.get("Nama","")) if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_") or "laporan_selected"
        
        out_path = filedialog.asksaveasfilename(title="Simpan PDF", defaultextension=".pdf", filetypes=[("PDF","*.pdf")], initialfile=f"laporan_{nama_safe}.pdf")
        
        if out_path:
            if FPDF is None:
                messagebox.showwarning("Missing lib", "Install 'fpdf' (pip install fpdf) untuk export PDF.")
                return
                
            pdf = FPDF()
            pdf.add_page()
            _write_report_to_pdf_page(pdf, row)
            
            try:
                pdf.output(out_path)
                messagebox.showinfo("Sukses", f"PDF dibuat: {out_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan PDF: {e}")

    def export_all_single_pdf(self):
        df = load_data()
        if df.empty:
            messagebox.showwarning("Kosong", "Tidak ada data untuk di-export.")
            return
            
        out_path = filedialog.asksaveasfilename(title="Simpan semua laporan ke 1 PDF", defaultextension=".pdf", filetypes=[("PDF","*.pdf")], initialfile=f"laporan_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        
        if not out_path: 
            return
            
        export_all_to_single_pdf(df, out_path)

    def export_all_multiple_pdfs(self):
        df = load_data()
        if df.empty:
            messagebox.showwarning("Kosong", "Tidak ada data untuk di-export.")
            return
            
        out_dir = filedialog.askdirectory(title="Pilih folder tujuan (per-file PDFs)")
        
        if not out_dir: 
            return
            
        export_each_row_to_pdf_files(df, out_dir)

    def export_all_excel(self):
        df = load_data()
        if df.empty:
            messagebox.showwarning("Kosong", "Tidak ada data untuk di-export.")
            return
            
        out_path = filedialog.asksaveasfilename(title="Simpan semua data ke Excel", defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")], initialfile=f"laporan_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        if not out_path: 
            return
            
        df_export = df.copy()
        
        if "Tanggal" in df_export.columns:

            df_export["Tanggal"] = df_export["Tanggal"].apply(lambda x: to_ddmmyyyy(x) if x else "")
            
        try:
            df_export.to_excel(out_path, index=False, engine="openpyxl")
            messagebox.showinfo("Sukses", f"Excel dibuat: {out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan Excel: {e}\nPastikan library 'openpyxl' terinstal.")

    # ---------------- Logout ----------------

    def on_logout(self):
        if messagebox.askyesno("Logout", "Yakin ingin logout?"):
            try:
                self.root.destroy()
            except Exception:
                os._exit(0) 

# ---------------- main ----------------

def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        style.theme_use('clam')
    except Exception:
        pass
        
    app = LaporanApp(root)
    root.mainloop() 
    
if __name__ == "__main__":
    main()