# How to use: step by step (Windows and macOS)

The program merges all your documents (PDF, Word, Excel, PowerPoint, images...) into **one file** to upload to Copilot. Everything runs on your computer: nothing is sent to the internet. You only need to do **Part 1** once.

## 🪟 Windows
1. **Install Python (once):** go to **https://www.python.org/downloads/** → yellow **Download Python** button → open the file → ⚠️ **check "Add python.exe to PATH"** on the first screen → **Install Now** → **Close**.
2. **Download the program:** on GitHub, click **Code** → **Download ZIP**, then right-click the ZIP → **Extract All...**.
3. **Open it:** double-click **`start.bat`** (English) or **`iniciar_windows.bat`** (Portuguese). If Windows shows *"Windows protected your PC"*, click **More info** → **Run anyway**. The first run installs what it needs (a few minutes, internet required).

## 🍎 macOS
1. **Install Python (once):** go to **https://www.python.org/downloads/** → **Download Python** → open the `.pkg` → **Continue** until **Install**.
2. **Download the program:** on GitHub, **Code** → **Download ZIP**, then double-click the ZIP.
3. **Open it:** open **Terminal** (`⌘ Command + Space`, type *Terminal*, Enter), type `bash ` (with a space), drag **`start_macos.command`** into the Terminal window, press **Enter**. The first run installs what it needs (a few minutes).
   *Double-clicking also works; if macOS blocks it, right-click → **Open** → **Open**.*

## 📋 Using it (same on both)
Type the option number and press Enter:
1. **`2`** (Input folders) → **`A`** → pick the folder with your documents (`0` goes back).
2. **`1`** (Build base) → press **Enter** to start.
3. When it finishes, press **`A`** to open the result folder (`copilot_consolidator/saida_copilot/ultimo_processamento`).

**Upload `corpus_completo.md` to Copilot** (or `corpus_parte_XX.md` if the base was split) and paste the text of `COPILOT_PROMPT.txt` into the chat. Copilot will cite documents by ID, e.g. **[004] Contract.pdf**.

| Problem | Fix |
| --- | --- |
| "Python was not found" | Redo step 1 and check **"Add python.exe to PATH"** (Windows). |
| Window opens and closes, or nothing happens (Windows) | Double-click `start.bat`. In PowerShell type `.start.bat` (with the leading `.`). Typing just `start` runs a built-in Windows command. |
| Library install fails | Check your internet connection; corporate proxies may block it. |
| `.doc`, `.ppt`, `.msg` skipped | Re-save them as `.docx`, `.pptx` or PDF. |
