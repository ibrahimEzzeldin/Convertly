"""
Convertly — File Converter  v2.0
Developer: Ibrahim Ezzeldin Mirghani
"""
import tkinter as tk
from tkinter import filedialog, ttk
import threading
import os
import subprocess
import platform
from datetime import datetime


def _open_path(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)

# ── Conversion Functions ─────────────────────────────────────────────────────

def pdf_to_word(pdf_path, out_path, stop_event=None):
    from pdf2docx import Converter
    cv = Converter(pdf_path)
    cv.convert(out_path)
    cv.close()

def pdf_to_excel(pdf_path, out_path, stop_event=None):
    import pdfplumber, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if stop_event and stop_event.is_set():
                raise InterruptedError("Cancelled by user.")
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

def word_to_pdf(docx_path, out_path, stop_event=None):
    """
    Converts DOCX to PDF using python-docx + reportlab.
    No Microsoft Word required — pure Python.
    """
    from docx import Document
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable)
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

    doc  = Document(docx_path)
    pdf  = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=2.5*cm,  bottomMargin=2.5*cm,
    )

    base_styles = getSampleStyleSheet()
    story = []

    style_normal = ParagraphStyle(
        "Normal2", parent=base_styles["Normal"],
        fontSize=11, leading=16, spaceAfter=4,
    )
    style_h1 = ParagraphStyle(
        "H1", parent=base_styles["Heading1"],
        fontSize=18, leading=22, spaceBefore=12, spaceAfter=6,
        textColor=colors.HexColor("#0F1117"),
    )
    style_h2 = ParagraphStyle(
        "H2", parent=base_styles["Heading2"],
        fontSize=14, leading=18, spaceBefore=10, spaceAfter=4,
        textColor=colors.HexColor("#1a1d2e"),
    )
    style_h3 = ParagraphStyle(
        "H3", parent=base_styles["Heading3"],
        fontSize=12, leading=16, spaceBefore=8, spaceAfter=3,
        textColor=colors.HexColor("#2d3250"),
    )

    heading_map = {
        "Heading 1": style_h1,
        "Heading 2": style_h2,
        "Heading 3": style_h3,
    }

    for para in doc.paragraphs:
        if stop_event and stop_event.is_set():
            raise InterruptedError("Cancelled by user.")

        text = para.text.strip()
        if not text:
            story.append(Spacer(1, 6))
            continue

        # Escape XML special chars
        text = (text.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;"))

        style_name = para.style.name if para.style else "Normal"
        p_style = heading_map.get(style_name, style_normal)

        # Bold / italic from runs
        parts = []
        for run in para.runs:
            rt = (run.text
                  .replace("&", "&amp;")
                  .replace("<", "&lt;")
                  .replace(">", "&gt;"))
            if run.bold and run.italic:
                parts.append(f"<b><i>{rt}</i></b>")
            elif run.bold:
                parts.append(f"<b>{rt}</b>")
            elif run.italic:
                parts.append(f"<i>{rt}</i>")
            else:
                parts.append(rt)

        rich_text = "".join(parts) if parts else text
        story.append(Paragraph(rich_text, p_style))

    # Tables
    for table in doc.tables:
        if stop_event and stop_event.is_set():
            raise InterruptedError("Cancelled by user.")
        data = []
        for row in table.rows:
            data.append([cell.text for cell in row.cells])
        if data:
            num_cols = max(len(r) for r in data)
            col_w = (A4[0] - 5*cm) / max(num_cols, 1)
            tbl = Table(data, colWidths=[col_w]*num_cols)
            tbl.setStyle(TableStyle([
                ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#4361EE")),
                ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
                ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",    (0,0), (-1,-1), 9),
                ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#DDDFE8")),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#F5F7FF")]),
                ("TOPPADDING",  (0,0), (-1,-1), 4),
                ("BOTTOMPADDING",(0,0),(-1,-1), 4),
                ("LEFTPADDING", (0,0), (-1,-1), 6),
            ]))
            story.append(Spacer(1, 6))
            story.append(tbl)
            story.append(Spacer(1, 6))

    if not story:
        story.append(Paragraph("(Empty document)", style_normal))

    pdf.build(story)

def excel_to_pdf(xlsx_path, out_path, stop_event=None):
    """
    Uses openpyxl + reportlab to convert Excel → PDF.
    No LibreOffice required — pure Python.
    """
    import openpyxl
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    story = []
    styles = getSampleStyleSheet()

    for sheet_name in wb.sheetnames:
        if stop_event and stop_event.is_set():
            raise InterruptedError("Cancelled by user.")

        ws = wb[sheet_name]

        # Sheet title
        story.append(Paragraph(f"<b>{sheet_name}</b>", styles["Heading2"]))
        story.append(Spacer(1, 0.3 * cm))

        # Collect data
        data = []
        col_widths = []
        for row in ws.iter_rows(values_only=True):
            row_data = [str(cell) if cell is not None else "" for cell in row]
            data.append(row_data)

        if not data:
            continue

        # Auto column widths
        num_cols = max(len(r) for r in data)
        page_w = landscape(A4)[0] - 2 * cm
        col_w  = page_w / max(num_cols, 1)
        col_widths = [col_w] * num_cols

        # Pad rows
        data = [r + [""] * (num_cols - len(r)) for r in data]

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0),  colors.HexColor("#4361EE")),
            ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.white),
            ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 8),
            ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FF")]),
            ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#DDDFE8")),
            ("TOPPADDING",  (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 0.5 * cm))

    doc = SimpleDocTemplate(
        out_path,
        pagesize=landscape(A4),
        leftMargin=1*cm, rightMargin=1*cm,
        topMargin=1*cm,  bottomMargin=1*cm
    )
    doc.build(story)


# ── Design Tokens ─────────────────────────────────────────────────────────────
BG          = "#F8F9FB"
WHITE       = "#FFFFFF"
BORDER      = "#EAEDF3"
BORDER_SOFT = "#F0F2F7"
TEXT        = "#0F1117"
TEXT_SEC    = "#6B7280"
TEXT_MUTED  = "#9CA3AF"
ACCENT      = "#4361EE"
ACCENT_DARK = "#3451D1"
ACCENT_SOFT = "#EEF1FF"

MODES = [
    {
        "label": "PDF → Word",
        "desc":  "Editable Word document",
        "icon":  "📄",
        "color": "#4361EE",
        "light": "#EEF1FF",
        "ft":    [("PDF Files", "*.pdf")],
        "ext":   "_converted.docx",
        "fn":    pdf_to_word,
    },
    {
        "label": "PDF → Excel",
        "desc":  "Extract tables & data",
        "icon":  "📊",
        "color": "#10B981",
        "light": "#ECFDF5",
        "ft":    [("PDF Files", "*.pdf")],
        "ext":   "_converted.xlsx",
        "fn":    pdf_to_excel,
    },
    {
        "label": "Word → PDF",
        "desc":  "Professional PDF output",
        "icon":  "📝",
        "color": "#EF4444",
        "light": "#FFF1F2",
        "ft":    [("Word Files", "*.docx")],
        "ext":   "_converted.pdf",
        "fn":    word_to_pdf,
    },
    {
        "label": "Excel → PDF",
        "desc":  "Spreadsheet to PDF",
        "icon":  "📈",
        "color": "#F59E0B",
        "light": "#FFFBEB",
        "ft":    [("Excel Files", "*.xlsx")],
        "ext":   "_converted.pdf",
        "fn":    excel_to_pdf,
    },
]


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertly — File Converter")
        self.root.geometry("620x900")
        self.root.minsize(540, 820)
        self.root.resizable(True, True)
        self.root.configure(bg=BG)

        # Center on screen
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"620x900+{(sw-620)//2}+{(sh-900)//2}")

        self.file_path   = tk.StringVar()
        self.active_mode = MODES[0]
        self.card_refs   = []
        self.status_var  = tk.StringVar(value="Choose a format, then select your file.")
        self._stop_event = threading.Event()
        self._converting = False

        self._build()

    # ─────────────────────────────────────────────────────────────────────────
    def _build(self):
        self._build_header()
        self._build_body()
        self._build_footer()
        self._activate_card(0)

    # ── Header ───────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=WHITE)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=BORDER, height=1).pack(side="bottom", fill="x")

        inner = tk.Frame(hdr, bg=WHITE)
        inner.pack(fill="x", padx=32, pady=18)

        logo = tk.Frame(inner, bg=WHITE)
        logo.pack(side="left")

        badge = tk.Frame(logo, bg=ACCENT, width=40, height=40)
        badge.pack(side="left")
        badge.pack_propagate(False)
        tk.Label(badge, text="CV", font=("Segoe UI", 12, "bold"),
                 bg=ACCENT, fg=WHITE).pack(expand=True)

        name_col = tk.Frame(logo, bg=WHITE)
        name_col.pack(side="left", padx=(12, 0))
        tk.Label(name_col, text="Convertly",
                 font=("Segoe UI", 15, "bold"),
                 bg=WHITE, fg=TEXT).pack(anchor="w")
        tk.Label(name_col, text="File Format Converter",
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=TEXT_MUTED).pack(anchor="w")

        pill = tk.Frame(inner, bg=ACCENT_SOFT, padx=12, pady=5)
        pill.pack(side="right", anchor="center")
        tk.Label(pill, text="v 2.0", font=("Segoe UI", 8, "bold"),
                 bg=ACCENT_SOFT, fg=ACCENT).pack()

    # ── Body ─────────────────────────────────────────────────────────────────
    def _build_body(self):
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=28, pady=16)
        body.columnconfigure(0, weight=1)

        self._section(body, "Conversion Type")

        grid = tk.Frame(body, bg=BG)
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        for i, mode in enumerate(MODES):
            self._make_card(grid, mode, i // 2, i % 2)

        self._section(body, "Select File", top=16)

        fzone = tk.Frame(body, bg=WHITE,
                         highlightthickness=1,
                         highlightbackground=BORDER)
        fzone.pack(fill="x")

        fz_in = tk.Frame(fzone, bg=WHITE)
        fz_in.pack(fill="x", padx=18, pady=14)

        self.file_icon_lbl = tk.Label(fz_in, text="📂",
                                      font=("Segoe UI", 22), bg=WHITE)
        self.file_icon_lbl.pack(side="left", padx=(0, 14))

        txt = tk.Frame(fz_in, bg=WHITE)
        txt.pack(side="left", fill="both", expand=True)

        self.fname_lbl = tk.Label(txt, text="No file selected",
                                  font=("Segoe UI", 10, "bold"),
                                  bg=WHITE, fg=TEXT_SEC, anchor="w")
        self.fname_lbl.pack(fill="x")

        self.fpath_lbl = tk.Label(txt, text="Supports PDF · DOCX · XLSX",
                                  font=("Segoe UI", 8),
                                  bg=WHITE, fg=TEXT_MUTED, anchor="w")
        self.fpath_lbl.pack(fill="x", pady=(2, 0))

        self.browse_btn = tk.Button(fz_in, text="Browse",
                                    command=self.browse,
                                    bg=ACCENT, fg=WHITE,
                                    font=("Segoe UI", 9, "bold"),
                                    relief="flat", padx=18, pady=8,
                                    cursor="hand2", bd=0,
                                    activebackground=ACCENT_DARK,
                                    activeforeground=WHITE)
        self.browse_btn.pack(side="right")
        self._btn_hover(self.browse_btn, ACCENT, ACCENT_DARK)

        # ── Convert + Stop buttons row ──
        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(fill="x", pady=(20, 0))
        btn_row.columnconfigure(0, weight=1)

        self.convert_btn = tk.Button(btn_row, text="  Convert Now  →",
                                     command=self.start_conversion,
                                     bg=ACCENT, fg=WHITE,
                                     font=("Segoe UI", 12, "bold"),
                                     relief="flat", pady=14,
                                     cursor="hand2", bd=0,
                                     activebackground=ACCENT_DARK,
                                     activeforeground=WHITE)
        self.convert_btn.pack(side="left", fill="x", expand=True)
        self._btn_hover(self.convert_btn, ACCENT, ACCENT_DARK)

        # Stop button — hidden until conversion starts
        self.stop_btn = tk.Button(btn_row, text="✕  Stop",
                                  command=self.stop_conversion,
                                  bg="#EF4444", fg=WHITE,
                                  font=("Segoe UI", 11, "bold"),
                                  relief="flat", pady=14, padx=20,
                                  cursor="hand2", bd=0,
                                  activebackground="#C0392B",
                                  activeforeground=WHITE)
        self._btn_hover(self.stop_btn, "#EF4444", "#C0392B")

        # Progress
        style = ttk.Style()
        style.theme_use("default")
        style.configure("App.Horizontal.TProgressbar",
                        troughcolor=BORDER_SOFT,
                        background=ACCENT,
                        thickness=3)

        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                        style="App.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(10, 0))

        tk.Label(body, textvariable=self.status_var,
                 font=("Segoe UI", 9),
                 bg=BG, fg=TEXT_MUTED, anchor="w").pack(fill="x", pady=(6, 0))

    # ── Footer ───────────────────────────────────────────────────────────────
    def _build_footer(self):
        year = datetime.now().year

        footer = tk.Frame(self.root, bg=WHITE)
        footer.pack(fill="x", side="bottom")
        tk.Frame(footer, bg=BORDER, height=1).pack(fill="x")

        f_in = tk.Frame(footer, bg=WHITE)
        f_in.pack(fill="x", padx=28, pady=10)

        tk.Label(f_in,
                 text="🔒  Converted locally — your files stay private.",
                 font=("Segoe UI", 8), bg=WHITE, fg=TEXT_MUTED,
                 anchor="w").pack(side="left")

        dev = tk.Frame(f_in, bg=ACCENT_SOFT,
                       highlightthickness=1, highlightbackground="#D8DEFF")
        dev.pack(side="right")
        dev_in = tk.Frame(dev, bg=ACCENT_SOFT)
        dev_in.pack(padx=12, pady=6)
        tk.Label(dev_in, text="Ibrahim Ezzeldin Mirghani",
                 font=("Segoe UI", 8, "bold"),
                 bg=ACCENT_SOFT, fg=TEXT).pack(anchor="e")
        tk.Label(dev_in, text="Application's Developer",
                 font=("Consolas", 7),
                 bg=ACCENT_SOFT, fg=ACCENT).pack(anchor="e")

        copy_bar = tk.Frame(self.root, bg=BORDER_SOFT)
        copy_bar.pack(fill="x", side="bottom")
        tk.Label(copy_bar,
                 text=f"© {year}  Convertly — File Converter  •  All rights reserved.",
                 font=("Segoe UI", 8),
                 bg=BORDER_SOFT, fg=TEXT_MUTED).pack(pady=5)

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _section(self, parent, text, top=0):
        tk.Label(parent, text=text,
                 font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT, anchor="w").pack(fill="x", pady=(top, 8))

    def _btn_hover(self, btn, normal, hover):
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
        btn.bind("<Leave>", lambda e: btn.configure(bg=normal))

    def _make_card(self, parent, mode, row, col):
        pad = (0, 6) if col == 0 else (6, 0)

        outer = tk.Frame(parent, bg=BORDER, cursor="hand2")
        outer.grid(row=row, column=col, sticky="ew", padx=pad, pady=(0, 10))

        card = tk.Frame(outer, bg=WHITE, cursor="hand2", height=110)
        card.pack(fill="both", padx=1, pady=1)
        card.pack_propagate(False)

        top_row = tk.Frame(card, bg=WHITE)
        top_row.pack(fill="x", padx=14, pady=(13, 4))

        icon_frame = tk.Frame(top_row, bg=mode["light"], width=34, height=34)
        icon_frame.pack(side="left")
        icon_frame.pack_propagate(False)
        tk.Label(icon_frame, text=mode["icon"],
                 font=("Segoe UI", 15), bg=mode["light"]).pack(expand=True)

        check = tk.Label(top_row, text="",
                         font=("Segoe UI", 13),
                         bg=WHITE, fg=mode["color"])
        check.pack(side="right")

        dot = tk.Frame(top_row, bg=mode["color"], width=7, height=7)
        dot.pack(side="right", padx=(0, 4))

        tk.Label(card, text=mode["label"],
                 font=("Segoe UI", 10, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(fill="x", padx=14)

        tk.Label(card, text=mode["desc"],
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=TEXT_MUTED, anchor="w").pack(fill="x", padx=14)

        idx = len(self.card_refs)
        self.card_refs.append({
            "outer": outer, "card": card,
            "check": check, "top_row": top_row,
            "icon_frame": icon_frame, "dot": dot,
            "mode": mode
        })

        all_widgets = [outer, card, top_row, icon_frame, check, dot] + \
                      list(card.winfo_children()) + \
                      list(top_row.winfo_children()) + \
                      list(icon_frame.winfo_children())

        for w in all_widgets:
            w.bind("<Button-1>", lambda e, i=idx: self._activate_card(i))
            w.bind("<Enter>",    lambda e, i=idx: self._card_hover(i, True))
            w.bind("<Leave>",    lambda e, i=idx: self._card_hover(i, False))

    def _card_hover(self, idx, entering):
        ref = self.card_refs[idx]
        if self.active_mode != ref["mode"]:
            ref["outer"].configure(bg=BORDER_SOFT if entering else BORDER)

    def _activate_card(self, idx):
        if self._converting:
            return  # don't allow switching while converting

        for ref in self.card_refs:
            ref["outer"].configure(bg=BORDER)
            ref["card"].configure(bg=WHITE)
            ref["check"].configure(text="", bg=WHITE)
            ref["top_row"].configure(bg=WHITE)
            ref["icon_frame"].configure(bg=ref["mode"]["light"])
            for child in list(ref["card"].winfo_children()) + \
                         list(ref["top_row"].winfo_children()) + \
                         list(ref["icon_frame"].winfo_children()):
                try:    child.configure(bg=WHITE)
                except: pass

        ref  = self.card_refs[idx]
        mode = ref["mode"]
        ref["outer"].configure(bg=mode["color"])
        ref["card"].configure(bg=mode["light"])
        ref["check"].configure(text="✓", bg=mode["light"])
        ref["top_row"].configure(bg=mode["light"])
        ref["icon_frame"].configure(bg=mode["light"])
        for child in list(ref["card"].winfo_children()) + \
                     list(ref["top_row"].winfo_children()) + \
                     list(ref["icon_frame"].winfo_children()):
            try:    child.configure(bg=mode["light"])
            except: pass

        self.active_mode = mode

        for btn in [self.convert_btn, self.browse_btn]:
            if btn:
                btn.configure(bg=mode["color"], activebackground=mode["color"])
                self._btn_hover(btn, mode["color"], self._darken(mode["color"]))

        style = ttk.Style()
        style.configure("App.Horizontal.TProgressbar",
                        background=mode["color"])

    def _darken(self, hex_color, factor=0.85):
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        return "#{:02x}{:02x}{:02x}".format(
            int(r * factor), int(g * factor), int(b * factor))

    # ── File & Conversion ────────────────────────────────────────────────────
    def browse(self):
        path = filedialog.askopenfilename(filetypes=self.active_mode["ft"])
        if path:
            self.file_path.set(path)
            name = os.path.basename(path)
            self.fname_lbl.config(
                text=name if len(name) <= 42 else name[:39] + "…",
                fg=TEXT)
            self.fpath_lbl.config(
                text=path if len(path) <= 58 else "…" + path[-55:],
                fg=TEXT_MUTED)
            self.file_icon_lbl.config(text="📄")
            self.status_var.set("File ready — click Convert Now to proceed.")

    def start_conversion(self):
        if not self.file_path.get():
            self._error_popup("No file selected. Please click Browse first.")
            return

        self._stop_event.clear()
        self._converting = True

        # Show stop button
        self.convert_btn.config(state="disabled", text="  Converting…  ⏳")
        self.stop_btn.pack(side="left", padx=(8, 0))

        threading.Thread(target=self._do_convert, daemon=True).start()

    def stop_conversion(self):
        self._stop_event.set()
        self.status_var.set("⚠  Stopping — please wait…")
        self.stop_btn.config(state="disabled", text="Stopping…")

    def _do_convert(self):
        self.progress.start(8)
        self.status_var.set("Converting — please wait…")
        try:
            src  = self.file_path.get()
            base = os.path.splitext(src)[0]
            out  = base + self.active_mode["ext"]
            self.active_mode["fn"](src, out, self._stop_event)

            if self._stop_event.is_set():
                # Clean up partial output
                if os.path.exists(out):
                    try: os.remove(out)
                    except: pass
                self._finish_ui("⚠  Conversion stopped.")
            else:
                self._finish_ui(f"✓  Done!  Saved as {os.path.basename(out)}", out)

        except InterruptedError:
            self._finish_ui("⚠  Conversion cancelled.")
        except Exception as e:
            self._finish_ui("An error occurred during conversion.", error=str(e))

    def _finish_ui(self, status_msg, out_path=None, error=None):
        self.progress.stop()
        self._converting = False
        self._stop_event.clear()

        self.status_var.set(status_msg)
        self.convert_btn.config(state="normal", text="  Convert Now  →")
        self.stop_btn.config(state="normal", text="✕  Stop")

        # Hide stop button
        self.stop_btn.pack_forget()

        if out_path:
            captured = out_path
            self.root.after(0, lambda p=captured: self._success_popup(p))
        elif error:
            captured = error
            self.root.after(0, lambda e=captured: self._error_popup(e))

    # ── Popups ───────────────────────────────────────────────────────────────
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
        self._center_popup(p, 480, 340)

        tk.Frame(p, bg=self.active_mode["color"], height=4).pack(fill="x")

        body = tk.Frame(p, bg=WHITE)
        body.pack(fill="both", expand=True, padx=30, pady=20)

        # Top: badge + title side by side
        top = tk.Frame(body, bg=WHITE)
        top.pack(fill="x", pady=(0, 12))

        badge = tk.Frame(top, bg=self.active_mode["light"], width=46, height=46)
        badge.pack(side="left")
        badge.pack_propagate(False)
        tk.Label(badge, text="✓",
                 font=("Segoe UI", 17, "bold"),
                 bg=self.active_mode["light"],
                 fg=self.active_mode["color"]).pack(expand=True)

        title_col = tk.Frame(top, bg=WHITE)
        title_col.pack(side="left", padx=(14, 0))
        tk.Label(title_col, text="Conversion complete!",
                 font=("Segoe UI", 13, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(anchor="w")
        tk.Label(title_col, text="Your file has been saved successfully.",
                 font=("Segoe UI", 8),
                 bg=WHITE, fg=TEXT_MUTED, anchor="w").pack(anchor="w")

        # File info box
        info_box = tk.Frame(body, bg=BORDER_SOFT,
                            highlightthickness=1,
                            highlightbackground=BORDER)
        info_box.pack(fill="x", pady=(0, 16))
        info_in = tk.Frame(info_box, bg=BORDER_SOFT)
        info_in.pack(fill="x", padx=12, pady=10)

        tk.Label(info_in, text="📄  " + os.path.basename(out_path),
                 font=("Segoe UI", 10, "bold"),
                 bg=BORDER_SOFT, fg=TEXT, anchor="w").pack(fill="x")
        tk.Label(info_in, text=out_path,
                 font=("Segoe UI", 8),
                 bg=BORDER_SOFT, fg=TEXT_MUTED, anchor="w",
                 wraplength=400).pack(fill="x", pady=(3, 0))

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x")

        # Buttons — always visible, full width
        btns = tk.Frame(body, bg=WHITE)
        btns.pack(fill="x", pady=(14, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)

        c = self.active_mode["color"]
        tk.Button(btns, text="📂  Open File",
                  command=lambda: (_open_path(out_path), p.destroy()),
                  bg=c, fg=WHITE, font=("Segoe UI", 10, "bold"),
                  relief="flat", pady=10,
                  cursor="hand2", bd=0).grid(row=0, column=0, sticky="ew", padx=(0, 4))

        tk.Button(btns, text="📁  Open Folder",
                  command=lambda: (_open_path(os.path.dirname(out_path)), p.destroy()),
                  bg=BORDER_SOFT, fg=TEXT, font=("Segoe UI", 10),
                  relief="flat", pady=10,
                  cursor="hand2", bd=0).grid(row=0, column=1, sticky="ew", padx=4)

        tk.Button(btns, text="✕  Close", command=p.destroy,
                  bg=WHITE, fg=TEXT_MUTED, font=("Segoe UI", 10),
                  relief="flat", pady=10,
                  cursor="hand2", bd=0,
                  highlightthickness=1,
                  highlightbackground=BORDER).grid(row=0, column=2, sticky="ew", padx=(4, 0))

    def _error_popup(self, message):
        p = tk.Toplevel(self.root)
        p.title("Error")
        p.resizable(False, False)
        p.configure(bg=WHITE)
        p.grab_set()
        self._center_popup(p, 420, 220)

        tk.Frame(p, bg="#EF4444", height=4).pack(fill="x")

        body = tk.Frame(p, bg=WHITE)
        body.pack(fill="both", expand=True, padx=30, pady=24)

        tk.Label(body, text="Something went wrong",
                 font=("Segoe UI", 13, "bold"),
                 bg=WHITE, fg=TEXT, anchor="w").pack(fill="x")

        tk.Label(body, text=message,
                 font=("Segoe UI", 9),
                 bg=WHITE, fg=TEXT_SEC, anchor="w",
                 wraplength=360, justify="left").pack(fill="x", pady=(8, 20))

        tk.Button(body, text="Close", command=p.destroy,
                  bg="#EF4444", fg=WHITE,
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=8,
                  cursor="hand2", bd=0).pack(anchor="w")


# ── Run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    app  = ConverterApp(root)
    root.mainloop()