# Local Document Consolidator for Microsoft Copilot & RAG v6.1

[![Platform: Windows 11 & macOS](https://img.shields.io/badge/Platform-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacy: 100% Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local-brightgreen.svg)](.)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A **100% local, offline** tool that reads large volumes of mixed documents and produces **one organized Markdown knowledge base** for **Microsoft 365 Copilot** to answer from. Every document gets an ID (`[001]`, `[002]`...) so Copilot can cite exactly where each fact came from.

Supported formats: PDF (native or scanned), Word (`.docx`), Excel (`.xlsx`, `.xlsm`, `.xls`), PowerPoint (`.pptx`), CSV, TXT, Markdown, HTML, `.eml` e-mails, JSON, XML, logs and images (`.png`, `.jpg`, `.tif`, `.bmp`, `.webp`). Images inside PDF, Word, Excel and PowerPoint files go through local OCR.

> 🇧🇷 *Versão em português: [LEIAME.md](LEIAME.md).* · 📜 *Full history: [HISTORICO_E_ESTADO_ATUAL.md](HISTORICO_E_ESTADO_ATUAL.md).*

> 🧭 **Never used Python or a terminal?** Follow the step-by-step guide [HOW_TO_USE.md](HOW_TO_USE.md) (Windows and macOS).

---

## 🚀 How to use

1. **Start it** (the first run installs dependencies into `.venv` automatically):
   - PowerShell: `.\start.ps1` · CMD/Explorer: `run.bat` · macOS: `./copilot_consolidator/start_macos.command`
   - Portuguese: `.\iniciar.ps1`, `iniciar_windows.bat` or `iniciar_macos.command`
2. **[2] Input folders**: add folders (file explorer dialog, pasted path, or a parent folder with all its subfolders).
3. **[1] Build the consolidated base**: review the summary and press ENTER.
4. **Upload `corpus_completo.md` to Copilot** (or the `corpus_parte_XX.md` parts if the base was split) and paste `COPILOT_PROMPT.txt` into the chat.

### Main menu

```
 [1] BUILD THE CONSOLIDATED COPILOT BASE
 [2] Input folders (add, remove, view)
 [3] Settings (outputs, OCR, privacy)
 [4] Open the latest result folder
 [5] Clear extraction cache (re-read everything)
 [L] Idioma: Português
 [0] Exit
```

---

## 📦 Output

In `saida_copilot/processado_YYYYMMDD_HHMMSS/` (always mirrored to `saida_copilot/ultimo_processamento/`):

| File | Content |
| --- | --- |
| `corpus_completo.md` | **The single base**: Copilot instructions, document index grouped by folder, and every document |
| `corpus_parte_XX.md` | The same base split into parts of up to N MB (only when it does not fit in one), each with its own index |
| `00_INDICE_MESTRE.md` | The catalog only (ID, document, type, notes) |
| `por_pasta/*.md` | One file per folder/department, for topic-specific questions |
| `corpus_rag.jsonl` | One line per page with metadata, for vector databases |
| `audit_report.md` | Failures, documents without text, low-confidence OCR, duplicates, similar versions, skipped formats |
| `COPILOT_PROMPT.txt` | Answering rules to paste into the chat |

Quality safeguards: duplicate files are listed but their content is not repeated; near-identical versions are flagged; PDF tables become Markdown tables without duplicated text; repeated headers/footers and page numbers are removed; OCR keeps the line layout of receipts and forms; corrupted accents (`ObservaÃ§Ã£o`) are repaired; personal data is kept by default (enable masking with `--redact-pii`).

## ⌨️ Command line

```powershell
python copilot_consolidator/app.py --auto --lang en                     # uses config.json
python copilot_consolidator/app.py --auto --input "D:\Docs" --output "D:\Base"
python copilot_consolidator/app.py --auto --redact-pii --workers 2 --no-ocr
python copilot_consolidator/app.py --help
```

## 🧪 Tests

```powershell
python -m unittest discover -s copilot_consolidator/tests -v
```

## 📄 License
Distributed under the **MIT License**. See [LICENSE](LICENSE).
