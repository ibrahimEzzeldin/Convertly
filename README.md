# Convertly — File Format Converter

**v2.0 · Built with Python · No Microsoft Office or LibreOffice required**

A clean, lightweight desktop application that converts documents locally — your files never leave your machine.

---

## Features

| Conversion | Description |
|---|---|
| PDF → Word | Turns any PDF into an editable `.docx` document |
| PDF → Excel | Extracts tables and text from PDFs into `.xlsx` |
| Word → PDF | Converts `.docx` files to professional PDF output |
| Excel → PDF | Exports spreadsheets to landscape PDF with styled tables |

- **100% local** — no internet connection, no cloud upload, full privacy
- **No Microsoft Office or LibreOffice needed** — pure Python implementation
- **Cancel mid-conversion** — stop button available while converting
- **Modern UI** — clean card-based interface built with Tkinter

---

## Tech Stack

- **Python 3** with `tkinter` for the GUI
- `pdf2docx` — PDF to Word conversion
- `pdfplumber` — PDF text and table extraction
- `openpyxl` — Excel reading and writing
- `python-docx` — Word document parsing
- `reportlab` — PDF generation for Word→PDF and Excel→PDF

---

## Getting Started

### Prerequisites

```bash
pip install pdf2docx pdfplumber openpyxl python-docx reportlab
```

### Run

```bash
python converter.py
```

---

## How to Use

1. Select a **conversion type** from the four cards
2. Click **Browse** to pick your source file
3. Click **Convert Now** — the output file is saved in the same folder as your input
4. Open the result directly from the success popup

---

## Developer

**Ibrahim Ezzeldin Mirghani**
Application Developer

---

© 2025 Convertly — File Converter · All rights reserved.
