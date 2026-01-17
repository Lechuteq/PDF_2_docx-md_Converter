# PDF-2-DOCX & MD Converter
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)

A professional, high-fidelity PDF conversion toolkit designed to transform static documents into editable formats. Whether you need clean Markdown for documentation or perfectly structured Word documents for further editing, this tool leverages advanced layout analysis and OCR to deliver superior results.

---

## 🚀 Overview

This project was developed by **[Lechuteq](https://github.com/Lechuteq/)** to bridge the gap between static PDF files and flexible digital formats. It is particularly effective for:
-   **Legal & Official Documents**: Preserving signatures, stamps, and logos in Word.
-   **Documentation Pipelines**: Converting PDFs to Markdown with intelligent OCR for scanned pages.
-   **Bulk Processing**: Handling large folders of documents with a single command.

## ✨ Key Features

-   **Intelligent Markdown Extraction (`-md`)**: 
    -   Automatically detects if a PDF is text-based or a scan.
    -   Integrated **Tesseract OCR** with Polish language support for scanned documents.
-   **Full-Structure Word Conversion (`-docx`)**: 
    -   Reconstructs the document into editable Word blocks (text, tables, images).
    -   Maintains original document flow for easy editing.
-   **Visual Preservation Word Conversion (`-docx_image`)**: 
    -   Renders each page as a high-resolution image.
    -   Guarantees 100% visual fidelity for complex layouts, stamps, and signatures.
-   **Modern Stack**: Built with `uv` for lightning-fast performance and reproducible environments.

## 🛠️ Technology Stack

-   **Logic**: Python 3.12+
-   **Environment**: [uv](https://astral.sh/uv)
-   **Libraries**:
    -   `pdf2docx`: Advanced PDF layout analysis and rebuilding.
    -   `pytesseract`: Optical Character Recognition (OCR).
    -   `pdf2image`: High-resolution PDF rendering via Poppler.
    -   `pypdf`: Direct text layer extraction.

---

## 📦 Installation & Setup

### 1. Prerequisites
-   **Tesseract OCR**: Install [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) (default path: `C:\Program Files\Tesseract-OCR`).
-   **Poppler**: Download [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases) (default path: `C:\Program Files\poppler-25.07.0`).

### 2. Environment Initialization
Clone the repository and run:
```bash
uv sync
```

---

## 📖 Usage Guide

Prepare your PDF files by placing them in the `source_pdf/` directory, then execute the converter with your desired flags:

| Flag | Mode | Best For |
| :--- | :--- | :--- |
| `-md` | Markdown | Technical documentation, clean text extraction, searches. |
| `-docx` | Editable Word | General editing, modifying text, updating tables. |
| `-docx_image` | Static Word | Hard-to-convert layouts, preserving official stamps/signatures. |

### Example Commands:
```powershell
# Convert all files to Markdown and Editable Word
uv run main.py -md -docx

# Convert all files to all supported formats
uv run main.py -md -docx -docx_image
```

---

## 🏗️ Development & Configuration

The project is designed to be easily configurable via the header of `main.py`:
-   `TESSDATA_LANG`: Set to `pol` for Polish; change to `eng` for English.
-   `POPPLER_PATH` / `TESSERACT_PATH`: System-specific binary locations.

### 👨‍💻 Author

Lesław Nowakowski  
AI Enthusiast | Python Developer | BI & KYC Analyst | DPO | Automatization Specialist  
🔗 www.linkedin.com/in/leslaw-nowakowski


## ⚖️ License

Distributed under the **MIT License**. See `LICENSE` for more information.

