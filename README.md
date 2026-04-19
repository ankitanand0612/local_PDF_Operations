# PDF Tools — Local PDF Operations

A lightweight, **100% local** web app for common PDF tasks. No uploads to any server — your files stay on your machine.

Built with Python (Flask + pikepdf + pdfplumber) and a clean dark UI served from a single file.

---

## Features

| Tool | What it does |
|------|-------------|
| 🔓 **Unlock** | Remove password protection from a PDF |
| 🔒 **Lock** | Add password protection to a PDF |
| 🗂 **Merge** | Combine multiple PDFs into one file |
| ✂️ **Split** | Split a PDF into individual pages (downloads as ZIP) |
| 📊 **PDF to XLSX** | Extract tables & text from a PDF into an Excel file |
| 🗜 **Compress** | Reduce PDF file size by compressing streams |

---

## Getting Started

### Prerequisites

- Python 3.8+
- pip

### Installation

```bash
# Clone the repo
git clone https://github.com/ankitanand0612/local_PDF_Operations.git
cd local_PDF_Operations

# Create and activate a virtual environment
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

# Install dependencies
pip install flask pikepdf pdfplumber openpyxl
```

### Run

```bash
python app.py
```

The app starts a local server and automatically opens your browser at `http://127.0.0.1:5000`.

To stop the server, press `Ctrl+C` in the terminal.

---

## How It Works

- The entire UI is a single-page app served by Flask as an inline HTML string.
- Each tool posts your file to a Flask route (`/unlock`, `/lock`, `/merge`, `/split`, `/to_xlsx`, `/compress`).
- The processed file is returned directly as a download — nothing is written to disk on the server side.
- All processing happens in memory using `pikepdf` (PDF manipulation) and `pdfplumber` (table/text extraction).

---

## Tech Stack

| Layer | Library |
|-------|---------|
| Web server | [Flask](https://flask.palletsprojects.com/) |
| PDF engine | [pikepdf](https://pikepdf.readthedocs.io/) |
| Table extraction | [pdfplumber](https://github.com/jsvine/pdfplumber) |
| Excel output | [openpyxl](https://openpyxl.readthedocs.io/) |

---

## Privacy

All operations run locally on your machine. No file is ever sent to an external server.
