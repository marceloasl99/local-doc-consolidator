# Versão sem Python — problema, opções e recomendação

> Esta pasta (`novo/`) é só de planejamento. **Nenhum arquivo do projeto atual foi alterado.**
> Qualquer código novo que for criado deve ficar aqui dentro.

---

## 1. O problema

O projeto atual (`copilot_consolidator/`) é feito em **Python** e precisa de:

1. **Python 3.10+ instalado** no computador.
2. **Bibliotecas instaladas com `pip`** (pymupdf, pypdf, python-docx, python-pptx, openpyxl, xlrd, pandas, pillow, rapidocr-onnxruntime).
3. **Internet** na primeira execução (para baixar as bibliotecas e os modelos de OCR).

No computador da empresa (Windows 11, sem permissão de administrador):

| Requisito | Situação provável | Consequência |
| --- | --- | --- |
| Instalar Python | O instalador padrão exige admin ou é bloqueado | Não roda |
| `pip install` | Exige acesso ao PyPI, que pode estar bloqueado por proxy | Não roda |
| Programas "não aprovados" | Política de TI proíbe instalar executáveis novos | Não roda |
| Python portátil (zip) | Pode ser bloqueado por AppLocker/WDAC (executável em pasta do usuário) | Pode não rodar |

Ou seja: **o código funciona, mas o ambiente de execução não pode ser instalado.** A solução precisa ser executada com o que o Windows e o Office **já trazem**, ou com algo que a empresa já aprovou.

---

## 2. O que o programa precisa fazer (requisitos a manter)

Independentemente da tecnologia, a versão nova deve manter:

- Ler pastas inteiras: **PDF, DOCX, XLSX, PPTX, imagens (PNG/JPG/TIF…), TXT, MD, CSV, HTML, EML, JSON, XML, LOG**.
- **OCR** para PDFs escaneados e imagens (com português).
- Gerar a **mesma saída** que o Copilot já usa:
  - `corpus_completo.md` (arquivo que se envia ao Copilot), com IDs `[001]`, `[002]`…
  - `00_INDICE_MESTRE.md`, `por_pasta/*.md`, `COPILOT_PROMPT.txt`, `audit_report.md`.
- Relatório de auditoria (falhas, documentos sem texto, OCR de baixa confiança, duplicados, arquivos ignorados).
- Opção de **mascarar dados pessoais** (CPF, RG, e-mail, telefone).
- Cache de extração (para não reprocessar tudo).

Tudo isso é lógica de texto, XML e arquivos; **o que é específico do Python é só a camada de leitura de formatos e de OCR.**

---

## 3. Como cada dependência pode ser substituída

| Função | Hoje (Python) | Alternativa sem Python instalado |
| --- | --- | --- |
| Texto de PDF nativo | pymupdf / pypdf | Windows não tem leitor de texto de PDF nativo; ver opções A e B (OCR ou pdf.js) |
| Renderizar página PDF em imagem | pymupdf | **Windows.Data.Pdf** (WinRT, já no Windows 10/11) ou pdf.js (navegador) |
| OCR | RapidOCR (ONNX) | **Windows.Media.Ocr** (WinRT, já no Windows; pt-BR presente, ver seção 5) ou Tesseract.js (navegador) |
| DOCX | python-docx | Ler o ZIP e interpretar `word/document.xml` (`System.IO.Compression`, nativo) |
| PPTX | python-pptx | Ler `ppt/slides/*.xml`, notas e gráficos (mesmo método) |
| XLSX | openpyxl / pandas | Ler `xl/sharedStrings.xml` e `xl/worksheets/*.xml` (mesmo método) |
| XLS, DOC, PPT antigos | xlrd / não suportados | Só via Office instalado (automação COM) |
| Imagens (redimensionar, TIF multipágina) | pillow | System.Drawing (PowerShell) ou APIs WinRT de imagem |
| Escolher pasta | tkinter | `FolderBrowserDialog` (WinForms, nativo) ou `<input webkitdirectory>` (navegador) |
| Paralelismo | ProcessPoolExecutor | Runspaces / `Start-Job` (PowerShell), ou Web Workers (navegador) |
| CSV, JSON, hash, e-mail, HTML, cache | stdlib | Equivalentes nativos em PowerShell ou JavaScript |

**Conclusão:** nenhuma dependência é insubstituível. O principal custo é reescrever os leitores de DOCX/PPTX/XLSX (são XML dentro de ZIP, trabalho previsível) e escolher uma forma de OCR.

---

## 4. As opções

### Opção A — PowerShell 5.1 (Windows PowerShell), script `.ps1` + `.bat`

Usa o PowerShell que **já vem no Windows 10/11**, sem instalar nada.

- **Prós**
  - Zero instalação e zero admin.
  - OCR do Windows (pt-BR) e renderizador de PDF já disponíveis (testado nesta máquina).
  - Leitura de DOCX/PPTX/XLSX por XML, sem bibliotecas externas.
  - Gera exatamente os mesmos arquivos de saída.
- **Contras**
  - Não existe extrator de **texto nativo de PDF** no WinRT: PDFs viram imagem e passam por OCR (mais lento; boa qualidade com 200–300 dpi).
  - Tabelas de Word/Excel precisam ser convertidas manualmente a partir do XML (mais código).
  - `.doc`, `.ppt`, `.xls` antigos só com Office instalado (opção C).
  - **Risco de política:** se a empresa usa *Constrained Language Mode* ou GPO que bloqueia PowerShell, o script não roda (ver checklist, seção 6).
- **Esforço estimado:** médio. A lógica de saída (índice, auditoria, cache, PII) pode ser portada quase direto.

### Opção B — Página HTML offline (abre no navegador)

Um único arquivo `consolidador.html` (com JavaScript embutido) abre no Edge ou Chrome, sem servidor e sem instalação.

- **Prós**
  - Navegador é o programa mais liberado em qualquer empresa.
  - Nada sai do computador (processamento local).
  - Bibliotecas podem ser embutidas no próprio arquivo: JSZip (ZIP), pdf.js (PDF nativo **e** renderização) e Tesseract.js (OCR, com pacote `por`).
  - Funciona mesmo se PowerShell estiver bloqueado.
- **Contras**
  - Não escreve direto em pastas: gera os arquivos para **download** (`corpus_completo.md` etc.).
  - OCR em JavaScript é **bem mais lento** que o nativo e o modelo de idioma pesa dezenas de MB.
  - Arquivos muito grandes podem estourar a memória do navegador; é preciso processar em lotes.
  - Abrir arquivo local (`file://`) pode limitar Web Workers e carregamento de WASM; a solução é embutir tudo ou usar um servidor local mínimo.
- **Esforço estimado:** médio-alto (mais código de interface e de empacotamento, mas sem dependência de sistema).

### Opção C — Office instalado (automação COM)

Se o Microsoft 365 estiver instalado (nesta máquina está: `O365ProPlusRetail`), o Word, o Excel e o PowerPoint podem ser controlados pelo PowerShell.

- **Prós:** converte **`.doc`, `.xls`, `.ppt`, `.rtf`** (que hoje não são suportados), e o Word pode abrir PDF e converter para texto.
- **Contras:** a conversão de PDF pelo Word perde estrutura; o COM pode ser bloqueado por política; precisa de Office no PC da empresa (geralmente tem, mas confirmar).
- **Uso ideal:** complemento da Opção A, não substituto.

### Opção D — Python portátil (sem instalar, só descompactar)

Baixar o Python "embeddable" (zip) ou WinPython e rodar o projeto atual sem instalador.

- **Prós:** o código atual roda **sem nenhuma alteração**.
- **Contras:**
  - **Ainda é Python**, então não atende ao requisito principal.
  - `pip` continua precisando de internet/PyPI.
  - Executáveis em pasta do usuário costumam ser bloqueados (AppLocker/WDAC) e pode ser "software não aprovado".
- **Recomendação:** só como teste em casa; não é solução para o PC da empresa.

### Opção E — Pedir aprovação ou usar o que a empresa já libera

- **Portal da Empresa / Central de Software (Software Center):** pedir a instalação do Python, que muitas vezes é instalado **sem admin** por esse portal.
- **Microsoft Store:** o Python da Store instala **por usuário**, sem admin, se a loja estiver liberada.
- **Se aprovado, o projeto atual roda como está** e não precisa de reescrita. É o caminho de menor esforço.

### Opção F — Não processar localmente (verificar com TI)

Se a empresa usa **Microsoft 365**, o Copilot pode já acessar arquivos do **SharePoint/OneDrive** diretamente. Nesse caso, a consolidação local pode nem ser necessária.

- Vale perguntar à TI se o Copilot já indexa as pastas de interesse.
- Atenção: a ferramenta lida com **dados pessoais (CPF, RG, e-mail)**. Antes de usar qualquer solução no PC da empresa, confirmar com a TI e o encarregado de dados (LGPD).

---

## 5. Comparação

| Critério | A. PowerShell | B. HTML offline | C. Office COM | D. Python portátil | E. Python aprovado |
| --- | --- | --- | --- | --- | --- |
| Precisa de admin | Não | Não | Não | Não | Não (se via portal/Store) |
| Instala algo | Não | Não | Não | Descompacta | Sim (aprovado) |
| Depende de política liberada | PowerShell | Navegador (quase sempre liberado) | COM | Executável | Portal/Store |
| OCR | Nativo do Windows (pt-BR) | Tesseract.js (pesado) | — | RapidOCR | RapidOCR |
| PDF nativo (texto) | Não (vira OCR) | Sim (pdf.js) | Parcial (via Word) | Sim | Sim |
| `.doc`/`.xls`/`.ppt` antigos | Só com Office | Não | Sim | Não | Não |
| Velocidade | Média | Baixa | Média | Alta | Alta |
| Esforço de implementação | Médio | Médio-alto | Complementa A/B | Nenhum | Nenhum |
| Manter 100% da saída atual | Sim | Sim | Sim | Sim | Sim |

---

## 6. Checklist: rodar no computador da empresa (sem admin)

Os testes abaixo **não instalam nada**. Rodar no PowerShell comum. Os resultados decidem qual opção é viável.

**1. Modo de linguagem do PowerShell** (precisa ser `FullLanguage`):
```powershell
$ExecutionContext.SessionState.LanguageMode
```

**2. Política de execução** (se `Restricted` ou GPO travada, scripts `.ps1` não rodam):
```powershell
Get-ExecutionPolicy -List
```

**3. OCR do Windows (idiomas disponíveis):**
```powershell
powershell -NoProfile -Command "[void][Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]; [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object { $_.LanguageTag }"
```
Esperado: aparece `pt-BR`. Se não aparecer, o OCR nativo não serve para português.

**4. Renderizador de PDF do Windows:**
```powershell
powershell -NoProfile -Command "[void][Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]; 'PDF renderer OK'"
```

**5. Office instalado** (para a opção C):
```powershell
Test-Path "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"
```

**6. Navegador** (para a opção B): abrir o Edge e confirmar que abre página local.

**7. Python liberado?** Ver se a Central de Software ou a Microsoft Store deixa instalar Python (opção E).

Se o item **1** ou **2** falhar, a **Opção A** fica bloqueada e a **Opção B** é a alternativa mais segura.

---

## 7. Recomendação

1. **Primeiro, perguntar à TI** se o Python pode ser instalado pela Central de Software ou pela Microsoft Store (opção E). Se sim, o projeto atual roda sem reescrita.
2. **Se não for liberado, rodar o checklist** (seção 6) no PC da empresa.
3. **Escolha conforme o resultado:**
   - PowerShell em `FullLanguage` + OCR pt-BR presente → **Opção A** (sem instalar nada, OCR nativo, mais rápida de desenvolver).
   - PowerShell bloqueado, mas navegador liberado → **Opção B** (HTML offline).
   - Ambas possíveis → começar pela **A** e usar a **C** para `.doc`/`.xls`/`.ppt`.
4. **Em qualquer caso**, validar com dados sintéticos e comparar a saída com a do Python (rodado em casa, onde o Python existe). Os arquivos `corpus_completo.md` e `00_INDICE_MESTRE.md` devem ter a mesma estrutura.

---

## 8. Limitações que ficam em qualquer opção

- **Qualidade do OCR** é menor que a do RapidOCR em alguns casos (documento torto, tabela, fonte pequena). Scans bons funcionam bem.
- **Tabelas de PDF** perdem formatação na extração sem pdf.js (OCR devolve texto, não tabela).
- **Desempenho:** OCR de muitas páginas leva tempo; a versão nova deve ter cache e progresso.
- **Sem testes automatizados** no código atual para o novo caminho: é preciso criar um conjunto de testes próprio (a pasta `exemplos_teste/` do projeto serve como base).
- **Dados pessoais:** continuam sendo processados localmente; a regra de mascarar PII deve ser mantida.

---

## 9. Decisões pendentes (para você escolher)

1. **Caminho preferido:** Opção A (PowerShell), Opção B (HTML), ou esperar a resposta da TI sobre Python (opção E)?
2. **Rodar o checklist** do item 6 no PC da empresa e me enviar o resultado (sem dados pessoais).
3. **Confirmar com a TI** se scripts ou páginas locais que leem arquivos de pasta são permitidos.
4. **Escopo do primeiro passo:** só os formatos Office + texto (sem OCR) para validar rápido, ou tudo de uma vez?

Depois dessas respostas, o próximo passo é criar o código dentro desta pasta `novo/`.
