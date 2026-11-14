# local-doc-consolidator - Copilot Document Consolidator 🚀📁

[![Platform: Windows 11](https://img.shields.io/badge/Platform-Windows%2011-lightgrey.svg)](.)
[![Privacy: 100% Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Deep-brightgreen.svg)](.)
[![Performance: Multicore Parallel](https://img.shields.io/badge/Performance-Multicore%20Parallel-orange.svg)](.)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An enterprise-grade, **100% local, offline, and secure** Python tool to index, convert, and consolidate massive volumes of confidential documents (PDFs, Word documents with embedded images, Excel sheets with embedded images, PowerPoint slides, and standalone scanned images) into structured Markdown ready for **Microsoft 365 Copilot**.

Process **hundreds to thousands of files** across multiple departmental folders with zero data leakage to external cloud services and without requiring Administrator privileges on Windows (no heavy external system dependencies like Poppler or Tesseract).

---

> 🇧🇷 *Para a versão deste documento em **Português do Brasil**, [role para a seção abaixo](#-versão-em-português-do-brasil).*

---

## 🌟 Key Features

* **🧠 Deep Offline OCR (RapidOCR + ONNX Runtime):**
  * **Scanned & Hybrid PDFs:** Reads native digital text and automatically executes OCR on fully scanned pages or embedded images within mixed pages.
  * **Word (`.docx`) with Images:** Automatically extracts embedded screenshots, invoices, or signed slips (`word/media/`) and transcribes them via OCR.
  * **Excel (`.xlsx`) with Images:** Detects and transcribes receipts, tables, and images pasted into spreadsheet sheets (`xl/media/`).
  * **PowerPoint (`.pptx`) with Images:** Transcribes diagrams and charts embedded across presentation slides.
  * **Standalone Images (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`):** Directly parses receipts and document photos into Markdown with resolution metadata.
* **📦 Flexible Markdown Splitting & Packaging:**
  * **Master Single File (`corpus_completo.md`):** Complete consolidated knowledge base with Master Index IDs `[001]` to `[N]`.
  * **Auto-Chunking by Size (MB):** Automatically split output when reaching **X MB** (customizable: 8 MB, 50 MB, 100 MB, up to **200 MB**).
  * **Batching by Document Count:** Group files into batches of *N* documents (e.g., 25, 50, or 100 docs per file).
  * **1:1 Individual Markdown:** Mirror documents individually into `saida_copilot/individual/`.
  * **Per-Department Folder:** Consolidated `.md` files for each monitored source folder in `saida_copilot/por_pasta/`.
* **📅 Timestamped Runs & Automatic .ZIP Archiving:**
  * Every execution writes to a dedicated timestamped directory (e.g., `processado_YYYYMMDD_HHMMSS/`).
  * Automatically creates an optimized `.zip` archive ready for upload.
  * Maintains a `ultimo_processamento/` shortcut folder for immediate access to the latest run.
* **⚡ Multicore Parallel Processing (CPU):**
  * Leverages all available CPU cores using `concurrent.futures.ProcessPoolExecutor` for high-throughput batch conversion.
* **🔄 Incremental SHA-256 Cryptographic Cache:**
  * Re-converts only newly added or modified documents on subsequent runs, generating full outputs in sub-second time.
* **🤖 Anti-Hallucination Prompt Template:**
  * Auto-generates `PROMPT_PARA_COPILOT.txt` to enforce strict source citations (`[001]`, `[002]`) in Copilot responses.

---

## 🛠️ Installation & Quick Start (Windows 11)

### Option 1: Double-Click Launcher (Recommended)
1. Double-click **`iniciar_windows.bat`**.
2. On first run, it automatically creates a sandboxed virtual environment (`.venv`), installs all local dependencies, and opens the interactive terminal menu.

### Option 2: Command Line (PowerShell / CMD)
```powershell
cd copilot_consolidator
iniciar_windows.bat
```
*(For headless/automated execution, run `python app.py --auto`).*

---

## 🖥️ Interactive Terminal Menu

```text
============================================================================
   LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (v4)
   Environment: Windows 11 | 100% Offline | 12 Concurrent CPU Cores
============================================================================
 Current Status:
  • Monitored Folders       : 3
  • Files Located           : 10 (488.1 KB)
  • Multicore Parallel      : ENABLED (12 cores)
  • OCR on PDFs & Images    : ENABLED
  • OCR Embedded Images     : ENABLED (Word/Excel/PPTX)
  • Output Splitting        : Single | Limit 8.0MB | Per Folder
  • Timestamp & Packaging   : Timestamped Folder + Zip Compression
  • Output Directory        : ./saida_copilot
----------------------------------------------------------------------------
 [1] RUN CONVERSION AND GENERATE FILES FOR COPILOT
 [2] Add folders (Windows Explorer picker, type paths, or Parent Folder)
 [3] List / Check / Remove configured folders
 [4] Reload folder list from 'pastas.txt'
 [5] Configure output options (Sizes/MB, Batches, ZIP, Timestamp, OCR)
 [6] Clear SHA-256 extraction cache (force fresh re-read)
 [0] Exit
============================================================================
```

---

## 🤖 How to Query with Microsoft 365 Copilot

1. Open **Microsoft 365 Copilot Chat** or **Copilot Notebooks**.
2. Attach the generated `corpus_completo.md` (or the split parts from `saida_copilot/ultimo_processamento/`).
3. Paste the system prompt from `PROMPT_PARA_COPILOT.txt`:

```text
Analyze the attached consolidated documentation. All documents are cataloged in the MASTER DOCUMENT INDEX at the top with unique identifiers in brackets, such as [001], [002], etc.

Mandatory rules for your responses:
1. For every piece of information, number, or rule cited, you must cite the document ID and original relative path (e.g., 'According to [001] 01_Finance/Report_2026.pdf...').
2. If there are conflicting rules or figures between documents, identify both IDs and highlight the discrepancy clearly.
3. Do not extrapolate or fabricate data not present in the attached file; if something is not mentioned, state it explicitly.
```

---

## 📋 Supported Formats

| Format | Extension | Capabilities |
| :--- | :--- | :--- |
| **PDF** | `.pdf` | Native text & tables + Offline OCR on scanned pages & embedded images |
| **Word** | `.docx` | Headings, styled paragraphs, Markdown tables + OCR on embedded images |
| **Excel** | `.xlsx`, `.xls` | Converts all sheets to Markdown tables + OCR on embedded image slips |
| **PowerPoint**| `.pptx` | Slide outlines, presenter notes, tables + OCR on slide diagrams |
| **Images** | `.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp` | Direct OCR extraction with resolution & format metadata |
| **Text/Data** | `.txt`, `.md`, `.csv` | Automatic encoding detection (UTF-8, Latin-1, CP1252) and CSV delimiter sniffing |

---

<br>

---

# 🇧🇷 Versão em Português do Brasil

## Consolidador Local de Documentos para Microsoft Copilot 🚀📁

Ferramenta corporativa **100% local, offline e segura** desenvolvida em Python para indexar, converter e consolidar grandes volumes de documentos confidenciais (PDFs com texto e escaneados, Word com imagens, Excel com imagens, apresentações, tabelas e imagens escaneadas) em bases Markdown perfeitamente estruturadas para o **Microsoft 365 Copilot**.

Processa centenas a milhares de arquivos espalhados em dezenas de pastas sem enviar nenhum dado para a nuvem e sem exigir privilégios de Administrador no Windows.

---

### 🌟 Principais Recursos

* **🧠 OCR Local Profundo (100% Offline - RapidOCR ONNX):**
  * **PDFs com Imagens / Escaneados:** Lê texto nativo e executa OCR em páginas escaneadas e em imagens de páginas mistas.
  * **Word (`.docx`) com Imagens:** Extrai prints de sistemas, comprovantes e fotos coladas no Word e aplica OCR.
  * **Excel (`.xlsx`) com Imagens:** Extrai comprovantes e recibos colados nas abas da planilha e transcreve via OCR.
  * **PowerPoint (`.pptx`) com Imagens:** Lê gráficos e diagramas presentes nos slides.
  * **Imagens Avulsas (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`):** Converte notas e recibos diretamente em Markdown com OCR.
* **📅 Organização por Timestamp e Pacote .ZIP:**
  * Cada execução cria uma pasta isolada identificada por data e hora (ex: `processado_YYYYMMDD_HHMMSS/`).
  * Gera automaticamente um arquivo compactado `.zip` pronto para uso.
  * Mantém a pasta `ultimo_processamento/` atualizada com os arquivos mais recentes.
* **📊 Contagem de Arquivos, Volume de Dados e Confirmação Prévia:**
  * Calcula em tempo real a quantidade de arquivos e o volume de dados em MB/GB.
  * Tela de confirmação antes de executar com detalhamento por extensão e ajuste rápido de limites de MB.
* **📁 Navegação e Confirmação de Pastas:**
  * Janela visual nativa do Windows Explorer para navegar e selecionar pastas.
  * Valida e exibe a quantidade e tamanho de arquivos encontrados antes de adicionar.
* **📦 Particionamento Flexível de Markdown:**
  * Arquivo único (`corpus_completo.md`), divisão por limite de tamanho em MB (até 200 MB), divisão por lotes de quantidade, 1:1 individual e consolidado por pasta departamental.
* **⚡ Multicore Paralelo e Cache SHA-256:**
  * Distribui a carga entre todos os núcleos da CPU e só reconverte arquivos novos ou alterados nas execuções seguintes.

---

### 🛠️ Como Executar no Windows 11

1. Dê **duplo clique** em **`iniciar_windows.bat`**.
2. Na primeira vez, o ambiente virtual isolado (`.venv`) será configurado automaticamente com todas as dependências.
3. O menu interativo abrirá diretamente no terminal pronto para uso.

---

### 📝 Licença
Distribuído sob a licença **MIT**.
