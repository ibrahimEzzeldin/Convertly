import tkinter as tk
from tkinter import filedialog, ttk
import threading
import os

# ── Conversion Functions ─────────────────────────────────────────────────────

def pdf_to_word(pdf_path, out_path):
    from pdf2docx import Converter
    cv = Converter(pdf_path)
    cv.convert(out_path)
    cv.close()

def pdf_to_excel(pdf_path, out_path):
    import pdfplumber, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        ws.append([c if c else "" for c in row])
            else:
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        ws.append([line])
    wb.save(out_path)

def word_to_pdf(docx_path, out_path):
    from docx2pdf import convert
    convert(docx_path, out_path)

def excel_to_pdf(xlsx_path, out_path):
    import subprocess
    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", os.path.dirname(out_path), xlsx_path
    ])

# ── Palette ───────────────────────────────────────────────────────────────────
BG        = "#f7f8fc"
WHITE     = "#ffffff"
BORDER    = "#e8eaf0"
BORDER2   = "#d0d4e0"
TEXT      = "#1a1d2e"
MUTED     = "#8b90a7"
MUTED2    = "#b0b5c8"

MODES = [
    {
        "label":   "PDF → Word",
        "desc":    "Convert PDF to editable Word doc",
        "icon":    "📄",
        "color":   "#4f6ef7",
        "light":   "#eef1ff",
        "ft":      [("PDF Files", "*.pdf")],
        "ext":     "_converted.docx",
        "fn":      pdf_to_word,
    },
    {
        "label":   "PDF → Excel",
        "desc":    "Extract tables into a spreadsheet",
        "icon":    "📊",
        "color":   "#10b981",
        "light":   "#ecfdf5",
        "ft":      [("PDF Files", "*.pdf")],
        "ext":     "_converted.xlsx",
        "fn":      pdf_to_excel,
    },
    {
        "label":   "Word → PDF",
        "desc":    "Convert Word document to PDF",
        "icon":    "📝",
        "color":   "#f43f5e",
        "light":   "#fff1f3",
        "ft":      [("Word Files", "*.docx")],
        "ext":     "_converted.pdf",
        "fn":      word_to_pdf,
    },
    {
        "label":   "Excel → PDF",
        "desc":    "Convert Excel sheet to PDF",
        "icon":    "📈",
        "color":   "#f59e0b",
        "light":   "#fffbeb",
        "ft":      [("Excel Files", "*.xlsx")],
        "ext":     "_converted.pdf",
        "fn":      excel_to_pdf,
    },
]

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertly")
        self.root.geometry("560x740")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.file_path   = tk.StringVar()
        self.active_mode = MODES[0]
        self.card_refs   = []
        self.convert_btn = None
        self.progress    = None
        self.status_var  = tk.StringVar(value="Select a format and choose your file to begin.")

        self._build()

    # ─────────────────────────────────────────────────────────────────────────
    def _build(self):

        # ── Header ──
        hdr = tk.Frame(self.root, bg=WHITE,
                       highlightthickness=1,
                       highlightbackground=BORDER)
        hdr.pack(fill="x")

        hdr_inner = tk.Frame(hdr, bg=WHITE)
        hdr_inner.pack(fill="x", padx=28, pady=16)

        # Logo
        logo = tk.Frame(hdr_inner, bg=WHITE)
        logo.pack(side="left")

        badge = tk.Frame(logo, bg="#4f6ef7", width=32, height=32)
        badge.pack(side="left")
        badge.pack_propagate(False)
        tk.Label(badge, text="C", font=("Segoe UI", 14, "bold"),
                 bg="#4f6ef7", fg=WHITE).pack(expand=True)

        tk.Label(logo, text="  Convertly",
                 font=("Segoe UI", 14, "bold"),
                 bg=WHITE, fg=TEXT).pack(side="left")

        tk.Label(hdr_inner, text="File Format Converter",
                 font=("Segoe UI", 9),
                 bg=WHITE, fg=MUTED).pack(side="right", anchor="center")

        # ── Body ──
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=20)

        # Section title
        self._section(body, "Choose conversion type")

        # ── Cards grid ──
        grid = tk.Frame(body, bg=BG)
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        for i, mode in enumerate(MODES):
            self._make_card(grid, mode, i // 2, i % 2)

        # ── File picker section ──
        self._section(body, "Select your file", top=20)

        file_box = tk.Frame(body, bg=WHITE,
                            highlightthickness=1,
                            highlightbackground=BORDER)
        file_box.pack(fill="x")

        file_inner = tk.Frame(file_box, bg=WHITE)
        file_inner.pack(fill="x", padx=16, pady=14)

        self.file_icon_lbl = tk.Label(file_inner, text="📂",
                                      font=("Segoe UI", 20),
                                      bg=WHITE)
        self.file_icon_lbl.pack(side="left", padx=(0, 12))

        txt_col = tk.Frame(file_inner, bg=WHITE)
        txt_col.pack(side="left", fill="both", expand=True)

        self.fname_lbl = tk.Label(txt_col,
                                  text="No file chosen",
                                  font=("Segoe UI", 10, "bold"),
                                  bg=WHITE, fg=MUTED, anchor="w")
        self.fname_lbl.pack(fill="x")

        self.fpath_lbl = tk.Label(txt_col,
                                  text="Supports PDF, DOCX, XLSX",
                                  font=("Segoe UI", 8),
                                  bg=WHITE, fg=MUTED2, anchor="w")
        self.fpath_lbl.pack(fill="x", pady=(2, 0))

        self.browse_btn = tk.Button(file_inner,
                                    text="Browse",
                                    command=self.browse,
                                    bg="#4f6ef7", fg=WHITE,
                                    font=("Segoe UI", 10, "bold"),
                                    relief="flat",
                                    padx=18, pady=8,
                                    cursor="hand2",
                                    activebackground="#3a58e0",
                                    activeforeground=WHITE, bd=0)
        self.browse_btn.pack(side="right")

        # ── Convert button ──
        self.convert_btn = tk.Button(body,
                                     text="⚡  Convert Now",
                                     command=self.start_conversion,
                                     bg="#4f6ef7", fg=WHITE,
                                     font=("Segoe UI", 12, "bold"),
                                     relief="flat",
                                     pady=15,
                                     cursor="hand2",
                                     activebackground="#3a58e0",
                                     activeforeground=WHITE, bd=0)
        self.convert_btn.pack(fill="x", pady=(20, 0))

        # ── Progress ──
        style = ttk.Style()
        style.theme_use("default")
        style.configure("App.Horizontal.TProgressbar",
                        troughcolor=BORDER,
                        background="#4f6ef7",
                        thickness=4)

        self.progress = ttk.Progressbar(body,
                                        mode="indeterminate",
                                        style="App.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(10, 0))

        # Status
        tk.Label(body,
                 textvariable=self.status_var,
                 font=("Segoe UI", 9),
                 bg=BG, fg=MUTED, anchor="w").pack(fill="x", pady=(6, 0))

        # ── Footer ──
        footer = tk.Frame(self.root, bg=WHITE,
                          highlightthickness=1,
                          highlightbackground=BORDER)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer,
                 text="🔒  Your files are converted locally — nothing is uploaded to the internet.",
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=MUTED).pack(pady=10)

        # Activate default card
        self._activate_card(0)

    # ─────────────────────────────────────────────────────────────────────────
    def _section(self, parent, text, top=0):
        tk.Label(parent, text=text,
                 font=("Segoe UI", 11, "bold"),
                 bg=BG, fg=TEXT, anchor="w").pack(fill="x", pady=(top, 8))

    def _make_card(self, parent, mode, row, col):
        pad = (0, 6) if col == 0 else (6, 0)

        outer = tk.Frame(parent, bg=BORDER, cursor="hand2")
        outer.grid(row=row, column=col, sticky="ew", padx=pad, pady=(0, 10))

        card = tk.Frame(outer, bg=WHITE, cursor="hand2", height=100)
        card.pack(fill="both", padx=1, pady=1)
        card.pack_propagate(False)

        top_row = tk.Frame(card, bg=WHITE)
        top_row.pack(fill="x", padx=14, pady=(14, 6))

        # Icon badge
        icon_bg = tk.Frame(top_row, bg=mode["light"],
                           width=36, height=36)
        icon_bg.pack(side="left")
        icon_bg.pack_propagate(False)
        tk.Label(icon_bg, text=mode["icon"],
                 font=("Segoe UI", 16),
                 bg=mode["light"]).pack(expand=True)

        check = tk.Label(top_row, text="",
                         font=("Segoe UI", 14),
                         bg=WHITE, fg=mode["color"])
        check.pack(side="right", padx=4)

        tk.Label(card, text=mode["label"],
                 font=("Segoe UI", 10, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(fill="x", padx=14)

        tk.Label(card, text=mode["desc"],
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=MUTED, anchor="w").pack(fill="x", padx=14)

        idx = len(self.card_refs)
        self.card_refs.append({
            "outer": outer, "card": card,
            "check": check, "mode": mode
        })

        for w in outer.winfo_children() + [outer]:
            pass
        for widget in [outer, card, top_row, icon_bg, check] + \
                       list(card.winfo_children()) + \
                       list(top_row.winfo_children()) + \
                       list(icon_bg.winfo_children()):
            widget.bind("<Button-1>", lambda e, i=idx: self._activate_card(i))

    def _activate_card(self, idx):
        for i, ref in enumerate(self.card_refs):
            ref["outer"].configure(bg=BORDER)
            ref["card"].configure(bg=WHITE)
            ref["check"].configure(text="", bg=WHITE)
            for child in ref["card"].winfo_children():
                try:
                    child.configure(bg=WHITE)
                except Exception:
                    pass

        ref = self.card_refs[idx]
        mode = ref["mode"]
        ref["outer"].configure(bg=mode["color"])
        ref["card"].configure(bg=mode["light"])
        ref["check"].configure(text="✓", bg=mode["light"])
        for child in ref["card"].winfo_children():
            try:
                child.configure(bg=mode["light"])
            except Exception:
                pass

        self.active_mode = mode

        if self.convert_btn:
            self.convert_btn.configure(bg=mode["color"],
                                       activebackground=mode["color"])
        if self.browse_btn:
            self.browse_btn.configure(bg=mode["color"],
                                      activebackground=mode["color"])
        if self.progress:
            style = ttk.Style()
            style.configure("App.Horizontal.TProgressbar",
                            background=mode["color"])

    # ─────────────────────────────────────────────────────────────────────────
    def browse(self):
        path = filedialog.askopenfilename(filetypes=self.active_mode["ft"])
        if path:
            self.file_path.set(path)
            name = os.path.basename(path)
            self.fname_lbl.config(
                text=name if len(name) <= 40 else name[:37] + "...",
                fg=TEXT)
            self.fpath_lbl.config(
                text=path if len(path) <= 55 else "..." + path[-52:],
                fg=MUTED)
            self.file_icon_lbl.config(text="📄")
            self.status_var.set("File selected — ready to convert.")

    def start_conversion(self):
        if not self.file_path.get():
            self._error_popup("No file selected. Please click Browse first.")
            return
        self.convert_btn.config(state="disabled", text="Converting...  ⏳")
        threading.Thread(target=self._do_convert, daemon=True).start()

    def _do_convert(self):
        self.progress.start(8)
        self.status_var.set("Converting — please wait...")
        try:
            src  = self.file_path.get()
            base = os.path.splitext(src)[0]
            out  = base + self.active_mode["ext"]
            self.active_mode["fn"](src, out)

            self.progress.stop()
            self.status_var.set(f"Done! Saved as  {os.path.basename(out)}")
            self.convert_btn.config(state="normal", text="⚡  Convert Now")
            self.root.after(0, lambda: self._success_popup(out))

        except Exception as e:
            self.progress.stop()
            self.status_var.set("An error occurred.")
            self.convert_btn.config(state="normal", text="⚡  Convert Now")
            self.root.after(0, lambda: self._error_popup(str(e)))

    # ── Popups ────────────────────────────────────────────────────────────────
    def _center_popup(self, popup, w, h):
        popup.update_idletasks()
        x = self.root.winfo_x() + self.root.winfo_width()  // 2 - w // 2
        y = self.root.winfo_y() + self.root.winfo_height() // 2 - h // 2
        popup.geometry(f"{w}x{h}+{x}+{y}")

    def _success_popup(self, out_path):
        p = tk.Toplevel(self.root)
        p.title("Done!")
        p.resizable(False, False)
        p.configure(bg=WHITE)
        p.grab_set()
        self._center_popup(p, 460, 270)

        # Top color strip
        tk.Frame(p, bg=self.active_mode["color"], height=5).pack(fill="x")

        body = tk.Frame(p, bg=WHITE)
        body.pack(fill="both", expand=True, padx=28, pady=22)

        # Check circle
        circle = tk.Frame(body, bg=self.active_mode["light"],
                          width=48, height=48)
        circle.pack(anchor="w")
        circle.pack_propagate(False)
        tk.Label(circle, text="✓",
                 font=("Segoe UI", 18, "bold"),
                 bg=self.active_mode["light"],
                 fg=self.active_mode["color"]).pack(expand=True)

        tk.Label(body, text="Conversion complete!",
                 font=("Segoe UI", 14, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(fill="x", pady=(10, 2))

        tk.Label(body,
                 text=f"📄  {os.path.basename(out_path)}",
                 font=("Segoe UI", 10),
                 bg=WHITE, fg=MUTED, anchor="w").pack(fill="x")

        tk.Label(body, text=out_path,
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=MUTED2, anchor="w",
                 wraplength=400).pack(fill="x", pady=(2, 16))

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x")

        btns = tk.Frame(body, bg=WHITE)
        btns.pack(fill="x", pady=(14, 0))

        tk.Button(btns, text="Open File",
                  command=lambda: (os.startfile(out_path), p.destroy()),
                  bg=self.active_mode["color"], fg=WHITE,
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=18, pady=8,
                  cursor="hand2").pack(side="left", padx=(0, 8))

        tk.Button(btns, text="Open Folder",
                  command=lambda: (os.startfile(os.path.dirname(out_path)), p.destroy()),
                  bg=BG, fg=TEXT,
                  font=("Segoe UI", 10),
                  relief="flat", padx=18, pady=8,
                  cursor="hand2").pack(side="left")

        tk.Button(btns, text="Close",
                  command=p.destroy,
                  bg=WHITE, fg=MUTED,
                  font=("Segoe UI", 10),
                  relief="flat", padx=18, pady=8,
                  cursor="hand2").pack(side="right")

    def _error_popup(self, message):
        p = tk.Toplevel(self.root)
        p.title("Error")
        p.resizable(False, False)
        p.configure(bg=WHITE)
        p.grab_set()
        self._center_popup(p, 420, 210)

        tk.Frame(p, bg="#f43f5e", height=5).pack(fill="x")

        body = tk.Frame(p, bg=WHITE)
        body.pack(fill="both", expand=True, padx=28, pady=22)

        tk.Label(body, text="Something went wrong",
                 font=("Segoe UI", 13, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(fill="x")

        tk.Label(body, text=message,
                 font=("Segoe UI", 9),
                 bg=WHITE, fg=MUTED, anchor="w",
                 wraplength=360, justify="left").pack(fill="x", pady=(8, 20))

        tk.Button(body, text="Close",
                  command=p.destroy,
                  bg="#f43f5e", fg=WHITE,
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=8,
                  cursor="hand2").pack(anchor="w")

# ── Run ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()