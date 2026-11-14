# local-doc-consolidator - Copilot & RAG Document Consolidator 🚀📁

[![Platform: Windows 11 & macOS](https://img.shields.io/badge/Platform-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacy: 100% Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Deep-brightgreen.svg)](.)
[![Performance: Multicore Parallel](https://img.shields.io/badge/Performance-Multicore%20Parallel-orange.svg)](.)
[![Export: Markdown + JSONL RAG](https://img.shields.io/badge/Export-Markdown%20%2B%20JSONL%20RAG-purple.svg)](.)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An enterprise-grade, **100% local, offline, and secure** tool to index, convert, and consolidate massive volumes of confidential documents (PDFs, Word documents with embedded images, Excel sheets with embedded images, PowerPoint slides, and standalone scanned images) into clean, continuous Markdown optimized for **Microsoft 365 Copilot** and structured JSONL for **local RAG pipelines (Chroma, Qdrant, LangChain)**.

Process **hundreds to thousands of files** across multiple departmental folders with zero data leakage to external cloud services and without requiring Administrator privileges on Windows or macOS (no external system dependencies like Poppler or Tesseract).

---

> 🇧🇷 *Para a versão deste documento em **Português do Brasil**, [role para a seção abaixo](#-versão-em-português-do-brasil) ou abra [LEIAME.md](copilot_consolidator/LEIAME.md).*

---

## 🌟 Key Features

* **🧠 Deep Offline OCR (RapidOCR + ONNX Runtime):**
  * **Scanned & Hybrid PDFs:** Reads native digital text and automatically executes OCR on fully scanned pages or embedded images within mixed pages.
  * **Word (`.docx`) with Images:** Automatically extracts embedded screenshots, invoices, or signed slips (`word/media/`) and transcribes them via OCR.
  * **Excel (`.xlsx` & `.xls`) with Images:** Detects and transcribes receipts, tables, and images pasted into spreadsheet sheets (`xl/media/`).
  * **PowerPoint (`.pptx`) with Images:** Transcribes diagrams and charts embedded across presentation slides.
  * **Standalone Images (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`):** Directly parses receipts and document photos into Markdown with resolution metadata.
* **🛡️ Smart Filtering & Noise Reduction:**
  * **OCR Confidence Filter (`score >= 0.50`):** Discards low-confidence OCR artifacts and noise.
  * **Image Hash Deduplication:** Computes SHA-256 on original raw image bytes. If a company logo or header repeats on 100 pages, OCR runs only once.
  * **Safe Aspect Ratio Filter (`> 20:1`):** Ignores divider lines and rules without dropping wide flowcharts, timelines, or financial tables.
  * **Text Normalization:** Reconstructs hyphenated line-endings (`develop-` \n `ment` -> `development`) and un-breaks hard line wraps.
* **📦 Flexible Markdown Splitting & RAG Export:**
  * **Continuous Clean Markdown (`corpus_completo.md`):** Fluid, uninterrupted text without intrusive middle-of-page split markers.
  * **Structured Vector RAG Dataset (`corpus_rag.jsonl`):** Dual export retaining exact page numbers, metadata, and relative paths for vector search engines.
  * **Auto-Chunking by Size (MB):** Automatically splits output when reaching **X MB** (customizable: 8 MB, 50 MB, 100 MB, up to **200 MB**).
  * **Batching by Document Count:** Group files into batches of *N* documents (e.g., 25, 50, or 100 docs per file).
  * **1:1 Individual Markdown:** Mirror documents individually into `saida_copilot/individual/`.
  * **Per-Department Folder:** Consolidated `.md` files for each monitored source folder in `saida_copilot/por_pasta/`.
* **📅 Timestamped Runs & Automatic Verified .ZIP Archiving:**
  * Every execution writes to a dedicated timestamped directory (`processado_YYYYMMDD_HHMMSS/`).
  * Automatically creates an optimized `.zip` archive verified with `zipfile.testzip()`.
  * Maintains a `ultimo_processamento/` shortcut folder for immediate access to the latest run.
* **⚡ Multicore Parallel Processing (CPU):**
  * Leverages all available CPU cores using `concurrent.futures.ProcessPoolExecutor` for high-throughput batch conversion.
* **🔄 Deterministic SHA-256 Parametric Cache:**
  * Caches OCR results factoring in DPI resolution, minimum dimensions, and confidence scores. Subsequent runs complete in seconds.
* **🤖 Dual Anti-Hallucination Prompt Templates:**
  * Generates both `COPILOT_PROMPT.txt` (English) and `PROMPT_PARA_COPILOT.txt` (Portuguese) enforcing strict source citations (`[001]`, `[002]`).

---

## 🛠️ Quick Start

### English Launcher (Windows)
Double-click **`start.bat`** (in root or in `copilot_consolidator/`).  
On first run, it automatically creates a virtual environment (`.venv`), installs local libraries, and launches the English menu.

### Portuguese Launcher (Windows)
Double-click **`iniciar_windows.bat`** (in root or in `copilot_consolidator/`).  
Launches the Brazilian Portuguese interface.

### macOS Launcher
- English: `./copilot_consolidator/start_macos.command`
- Portuguese: `./copilot_consolidator/iniciar_macos.command`

### Command Line (Headless / Automated)
```powershell
# English execution
python copilot_consolidator/app.py --auto --lang en

# Portuguese execution
python copilot_consolidator/app.py --auto --lang pt
```

---

## 🖥️ Interactive Terminal Menu (English)

```text
============================================================================
   LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (v4)
   Environment: Windows 11 | 100% Offline | 12 Concurrent CPU Cores
============================================================================
 Current Status:
  • Monitored folders        : 3
  • Located files            : 15 (8.62 MB)
  • Multicore Parallel       : ENABLED (12 cores)
  • OCR in PDFs and Images   : ENABLED
  • OCR Embedded Images      : ENABLED (Word/Excel/PPTX)
  • File Splitting           : Single | 8.0MB | Folder | RAG JSONL
  • Organization / Timestamp : Timestamped Folder + Compress .ZIP
  • Base destination         : ./saida_copilot
----------------------------------------------------------------------------
 [1] EXECUTE CONSOLIDATION AND GENERATE COPILOT FILES
 [2] Add folders (Windows Explorer, type path, or Parent Folder)
 [3] List / Check / Remove configured folders (3)
 [4] Reload folder list from 'pastas.txt'
 [5] Configure output options (Sizes/MB, Batches, ZIP, Timestamp, OCR, RAG JSONL)
 [6] Clear SHA-256 extraction cache (force full re-extraction)
 [L] Language / Idioma (Current: English)
 [0] Exit
============================================================================
```

---

## 📁 Supported Formats

| Category | File Extensions | Extraction & Processing Strategy |
|---|---|---|
| **PDF** | `.pdf` | Native text & tables + hybrid RapidOCR ONNX on scanned pages and embedded images |
| **Word** | `.docx` | Native paragraphs, styles, headings, tables + OCR on embedded images (`word/media/`) |
| **Excel** | `.xlsx`, `.xls` | Sheet-by-sheet Markdown tables + OCR on embedded receipts and figures (`xl/media/`) |
| **PowerPoint**| `.pptx` | Slides, text blocks, speaker notes, tables + OCR on slide diagrams (`ppt/media/`) |
| **Images** | `.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp` | Direct RapidOCR transcription with resolution and confidence metadata |
| **Text** | `.txt`, `.md`, `.csv` | Encoded with auto-fallback (UTF-8, Latin-1, CP1252) + CSV delimiter sniffing |

---

## 📂 Output Folder Structure

```text
saida_copilot/
  ├── ultimo_processamento/              # Fast-access shortcut to the most recent run
  └── processado_20261114_125217/        # Timestamped execution run
        ├── corpus_completo.md           # Master single-file Markdown with [001]..[N] index
        ├── corpus_parte_01.md           # Chunked parts if MB size limit is enabled
        ├── corpus_rag.jsonl             # Structured page-level JSONL for RAG vector stores
        ├── por_pasta/                   # Departmental consolidations
        │     ├── 01_Financeiro.md
        │     └── 02_Juridico.md
        ├── COPILOT_PROMPT.txt           # Anti-hallucination prompt (English)
        ├── PROMPT_PARA_COPILOT.txt      # Anti-hallucination prompt (Portuguese)
        ├── audit_report.md              # Technical audit metrics (English)
        ├── relatorio_auditoria.md       # Technical audit metrics (Portuguese)
        └── consolidated_20261114_125217.zip  # Integrity-verified .ZIP package
```

---

<br>

# 🇧🇷 Versão em Português do Brasil

O **local-doc-consolidator** é uma solução corporativa, **100% local, offline e segura** em Python para indexar, converter e consolidar volumes massivos de documentos confidenciais corporativos (PDFs nativos e escaneados, Word com imagens, Excel com comprovantes, apresentações PowerPoint e fotos de documentos) em Markdown contínuo para o **Microsoft 365 Copilot** e base JSONL estruturada para **RAG local (Chroma, Qdrant, Milvus, LangChain)**.

### Destaques da Versão:
1. **OCR Local Profundo com RapidOCR + ONNX Runtime:** Sem dependência de serviços em nuvem ou executáveis externos pesados (sem Tesseract, sem Poppler).
2. **Deduplicação de Imagens por Hash:** Logotipos repetidos em centenas de páginas são processados no OCR apenas uma vez.
3. **Filtro de Confiança e Proporção Segura:** Descarta ruídos com score < 0.50 e réguas divisórias (> 20:1), mantendo fluxogramas e tabelas analíticas.
4. **Markdown Limpo e Contínuo:** Sem quebras ou marcadores de página intrusivos no meio do texto, permitindo leitura fluida pelo Copilot.
5. **Exportação Dual RAG (`corpus_rag.jsonl`):** Preserva a rastreabilidade página a página nos metadados para indexação vetorial.
6. **Bilingue Completo:** Menus, relatórios de auditoria e prompts gerados em Português e Inglês.
7. **Empacotamento ZIP com Verificação Automática:** Validação estrita via `zipfile.testzip()`.

### Como Executar:
- **Windows (Português):** Dê dois cliques em **`iniciar_windows.bat`**.
- **Windows (Inglês):** Dê dois cliques em **`start.bat`**.
- **macOS:** Execute `./copilot_consolidator/iniciar_macos.command` ou `./copilot_consolidator/start_macos.command`.

---

## 📄 License / Licença

Distributed under the **MIT License**. See [LICENSE](LICENSE) for details.  
Copyright (c) 2026 marceloasl99.
