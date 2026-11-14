# local-doc-consolidator - Copilot & RAG Document Consolidator v6.0 🚀📁

[![Platform: Windows 11 & macOS](https://img.shields.io/badge/Platform-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacy: 100% Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Deep-brightgreen.svg)](.)
[![Compliance: PII Redacted](https://img.shields.io/badge/Privacy-PII%20Redaction%20ON-blue.svg)](.)
[![Performance: Multicore Parallel](https://img.shields.io/badge/Performance-Multicore%20Parallel-orange.svg)](.)
[![Export: Markdown + JSONL RAG](https://img.shields.io/badge/Export-Markdown%20%2B%20JSONL%20RAG-purple.svg)](.)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An enterprise-grade, **100% local, offline, and secure** tool to index, convert, and consolidate massive volumes of confidential documents (PDFs, Word documents with embedded images, Excel sheets with embedded images, PowerPoint slides, and standalone scanned images) into clean, continuous Markdown optimized for **Microsoft 365 Copilot** and structured JSONL for **local RAG pipelines (Chroma, Qdrant, LangChain)**.

Process **hundreds to thousands of files** across multiple departmental folders with zero data leakage to external cloud services and without requiring Administrator privileges on Windows or macOS (no external system dependencies like Poppler or Tesseract).

---

> 🇧🇷 *Para a versão deste documento em **Português do Brasil**, abra [LEIAME.md](LEIAME.md).*

---

## 🌟 Key Features (v6.0)

* **🧠 Deep Offline OCR (RapidOCR + ONNX Runtime):**
  * **Scanned & Hybrid PDFs:** Reads native digital text and automatically executes OCR on fully scanned pages or embedded images within mixed pages.
  * **Word (`.docx`) with Images:** Extracts embedded screenshots, invoices, or signed slips (`word/media/`) in context and transcribes them via OCR.
  * **Excel (`.xlsx` & `.xls`) with Images:** Detects and transcribes receipts, tables, and images pasted into spreadsheet sheets (`xl/media/`).
  * **PowerPoint (`.pptx`) with Images:** Transcribes diagrams and charts embedded across presentation slides.
  * **Standalone Images (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`):** Directly parses receipts and document photos into Markdown with resolution metadata.
* **🔤 Automatic Mojibake & Encoding Repair:**
  * Auto-detects and repairs corrupted UTF-8 encodings (e.g. UTF-8 misinterpreted as CP1252/Latin-1, converting `ObservaÃ§Ã£o` -> `Observação`, `MÃ©todo` -> `Método`, `nÃ£o` -> `não`).
* **✨ OCR Phonotactic Word De-breaking:**
  * Eliminates OCR segmentation space artifacts (e.g., `Aute nticacao` -> `Autenticacao`, `comprova nte` -> `comprovante`, `docu mento` -> `documento`, `estabeleci mento` -> `estabelecimento`).
* **🔍 Near-Duplicate Content Detection:**
  * Computes lexical overlap (Jaccard / Overlap >= 80%) across distinct documents (revisions, draft templates) and includes disambiguation notes in `00_INDICE_MESTRE.md` and audit reports.
* **🛡️ Privacy & PII Redaction (Safe-by-Default):**
  * Masks CPFs, CNPJs, email addresses, and phone numbers before ingestion into Microsoft 365 Copilot cloud (disable with `--no-pii` when desired).
* **📦 Smart Output Partitioning Without Redundancy:**
  * **Continuous Clean Markdown (`corpus_completo.md`):** Fluid text without disruptive page separators.
  * **No Redundant Single Parts:** If the whole corpus fits inside the chunk limit (e.g. < 8 MB), suppresses duplicate `corpus_parte_01.md`, saving 50% disk and ZIP size.
  * **Structured Vector RAG Dataset (`corpus_rag.jsonl`):** Page-by-page records with metadata for vector databases.
* **📊 Dedicated Native Bilingual Reports:**
  * Generates `relatorio_auditoria.md` (100% Brazilian Portuguese) and `audit_report.md` (100% English).
  * Explicit heuristic token estimation labeling (4 characters/token).
* **⚡ Multicore Safety & Cryptographic Cache:**
  * Caps worker processes (`min(4, os.cpu_count())`) to prevent RAM exhaustion.
  * Parametric SHA-256 cache taking DPI, OCR thresholds, and PII masking into account.

---

## 🛠️ Quick Start

### English Interface (Windows)
* **Command Prompt (CMD - Type directly):**
  ```cmd
  run.bat
  :: or
  start.bat
  ```
* **PowerShell / Windows Terminal (Current Window):**
  ```powershell
  .\start.ps1
  ```

### Portuguese Interface (Windows)
* **Prompt de Comando / Windows Explorer:**
  ```cmd
  iniciar_windows.bat
  ```
* **PowerShell / Windows Terminal:**
  ```powershell
  .\iniciar.ps1
  ```

### macOS Launcher
- English: `./copilot_consolidator/start_macos.command`
- Portuguese: `./copilot_consolidator/iniciar_macos.command`

### Command Line (Automated / Headless)
```powershell
# English execution (PII redaction ON by default)
python copilot_consolidator/app.py --auto --lang en

# Portuguese execution
python copilot_consolidator/app.py --auto --lang pt

# Disable PII redaction if processing public open datasets
python copilot_consolidator/app.py --auto --lang en --no-pii
```

---

## 🧪 Automated Unit Tests

To run the full unit test suite:
```powershell
python -m unittest discover -s copilot_consolidator/tests -v
```

---

## 📄 License
Distributed under the **MIT License**. See [LICENSE](LICENSE) for details.
