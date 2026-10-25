# PPT Unify - Batch Presentation Merger & PDF Exporter

**PPT Unify** is a fast, error-free program designed to upload up to 100 PowerPoint presentations (`.pptx`), organize and reorder them line-by-line, merge them into a single unified presentation (with sequential or interleaved line-by-line ordering), and export the result as both **PPTX** and **PDF**.

---

## ✨ Features

- **Upload up to 100 PPT/PPTX Files Error-Free**:
  - Drag-and-drop or batch file selection.
  - Chunked, error-tolerant upload pipeline.
  - Automated per-file validation (slide count, file integrity, aspect ratio, corrupted file detection).
- **Merge All Presentations by Line**:
  - **Sequential Order (Default - By Table Line)**: Merges all slides from Deck #1, then Deck #2, Deck #3... matching the table order.
  - **Interleaved Order (Line-by-Line Collation)**: Collates slide-by-slide across decks (Slide 1 of each deck, then Slide 2 of each deck, etc.).
  - Intuitive line management: Move Up, Move Down, drag-and-drop reordering, Natural Sort (Deck 1, 2, ... 10, 100), A-Z, Reverse.
- **Dual Export Options**:
  - **Export as PPTX**: High-fidelity PowerPoint presentation.
  - **Export as PDF**: Vector-sharp, font-perfect PDF (powered by native macOS Keynote automation or headless conversion).
  - **Export Both (ZIP)**: One-click bundle containing both PPTX and PDF.
- **Modern User Interface & CLI**:
  - Responsive web dashboard with live progress bar and status streaming.
  - Direct command-line interface (`cli.py`) for automated workflows.

---

## 🚀 Quick Start

### 1. Launch the Web Interface

Double-click `start.sh` or run:

```bash
./start.sh
```

Or run directly with Python:

```bash
python3 run.py
```

This will automatically launch the application server and open `http://127.0.0.1:5005` in your default browser.

---

### 2. Command-Line Usage (CLI)

You can also merge presentations directly from your terminal:

```bash
# Merge specific files and export both PPTX and PDF
.venv/bin/python cli.py presentation1.pptx presentation2.pptx -o unified.pptx --pdf

# Merge an entire directory of PPTX files with interleaved collation
.venv/bin/python cli.py --dir /path/to/presentations -o unified.pptx --mode interleaved --pdf
```

---

## 🛠️ Architecture & Requirements

- Python 3.10+
- Installed libraries (configured in `.venv`):
  - `python-pptx` (PowerPoint manipulation and XML cloning)
  - `reportlab` & `pillow` (Python PDF rendering engine)
  - `pypdf` (PDF processing)
- macOS Native Integration:
  - Apple Keynote AppleScript automation for vector-quality PDF rendering.

---

## 📁 Project Structure

```
├── app.py              # Standalone HTTP server & REST API
├── merger_engine.py    # Core slide cloner, merger & PDF converter
├── cli.py              # Command-line interface
├── run.py              # One-click launcher with auto-browser launch
├── start.sh            # Shell executable script
├── static/
│   ├── index.html      # Modern dashboard UI
│   ├── style.css       # Responsive dark theme styling
│   └── app.js          # Client-side validation, drag-drop & progress tracker
├── uploads/            # Temporary upload storage
└── exports/            # Output merged PPTX and PDF files
```
