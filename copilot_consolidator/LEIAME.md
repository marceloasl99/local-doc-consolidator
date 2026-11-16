# Consolidador Local de Documentos para Microsoft Copilot v7.0

[![Plataforma: Windows 11 & macOS](https://img.shields.io/badge/Plataforma-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacidade: 100% Offline](https://img.shields.io/badge/Privacidade-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local-brightgreen.svg)](.)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)

Ferramenta **100% local e offline** que lê grandes volumes de documentos variados e gera **uma base Markdown estruturada e totalmente referenciada** para o **Microsoft 365 Copilot** usar como fonte das respostas. Cada documento recebe um ID (`[001]`, `[002]`...) em ordem numérica natural para que o Copilot cite exatamente de onde tirou cada informação.

Formatos lidos: PDF (com texto ou digitalizado), Word (`.docx`), Excel (`.xlsx`, `.xlsm`, `.xls`), PowerPoint (`.pptx`), CSV, TXT, Markdown, HTML, e-mails `.eml` (incluindo anexos suportados), JSON, XML, logs e imagens (`.png`, `.jpg`, `.tif`, `.bmp`, `.webp`). Imagens dentro de PDFs, Word, Excel e PowerPoint passam por OCR local em passe único.

> 🇺🇸 *English version: [README.md](README.md).* · 📜 *Histórico completo: [HISTORICO_E_ESTADO_ATUAL.md](HISTORICO_E_ESTADO_ATUAL.md).*

> 🧭 **Nunca usou Python ou terminal?** Siga o passo a passo [COMO_USAR.md](../COMO_USAR.md) (Windows e macOS).

---

## 🚀 Como usar

1. **Inicie** (a primeira execução instala as dependências em `.venv` automaticamente):
   - Windows: dois cliques em `iniciar_windows.bat` (no PowerShell digite `.\iniciar_windows.bat`, com o `.\` na frente) · macOS: `bash iniciar_macos.command`
   - Em inglês: `start.bat` / `.\start.bat` / `bash start_macos.command`. ⚠️ Digitar só `start` não funciona: é um comando do próprio Windows.
2. **[2] Pastas de entrada** → adicione as pastas (janela do Explorer, colando o caminho ou uma pasta mãe com todas as subpastas).
3. **[1] Gerar base consolidada** → confira o resumo e pressione ENTER.
4. **Envie `corpus_completo.md` ao Copilot** (ou as partes `corpus_parte_XX.md`, se a base tiver sido dividida) e cole o conteúdo de `PROMPT_PARA_COPILOT.txt` no chat.

### Menu principal

```
 [1] GERAR BASE CONSOLIDADA PARA O COPILOT
 [2] Pastas de entrada (adicionar, remover, ver)
 [3] Configurações (saídas, OCR, privacidade)
 [4] Abrir a pasta do último resultado
 [5] Limpar cache de extração (força reler tudo)
 [L] Language: English
 [0] Sair
```

O menu mostra quantos arquivos foram encontrados por tipo e avisa sobre pastas inexistentes. A tela de confirmação permite ajustar na hora o tamanho das partes, os processos paralelos, o OCR, a anonimização de PII e o `.zip`.

---

## 📦 O que é gerado

Em `saida_copilot/processado_AAAAMMDD_HHMMSS/` (cópia sempre atualizada em `saida_copilot/ultimo_processamento/`), com foco padrão em **Markdown puro e referenciado**:

| Arquivo | Conteúdo |
| --- | --- |
| `corpus_completo.md` | **A base única**: instruções para o Copilot, índice em ordem natural, catálogo de diagramas e todos os documentos |
| `corpus_parte_XX.md` | A mesma base dividida em partes de até N MB (só quando não cabe em uma), cada uma com índice e catálogo de diagramas |
| `00_INDICE_MESTRE.md` | Só o catálogo (ID, documento, tipo, observações e tabela de diagramas) |
| `por_pasta/*.md` | Um arquivo por pasta/departamento (com índice e catálogo de diagramas filtrados), para perguntas sobre um tema específico |
| `imagens/` | Pasta ordenada (`docXXX_pYY_figZZ.png`) contendo **apenas** diagramas e esquemas visuais não traduzíveis em puro texto, referenciados no Markdown com transcrição OCR completa |
| `relatorio_auditoria.md` | Falhas, documentos sem texto, OCR com baixa confiança, duplicados, versões semelhantes e formatos ignorados |
| `PROMPT_PARA_COPILOT.txt` | Regras de resposta para colar no chat |
| `corpus_rag.jsonl` / `.zip` | Saídas opcionais (desativadas por padrão; ativáveis no menu `[3] → [8]` e `[9]`) |

### Estrutura de cada documento na base

```markdown
## [004] Relatorio_Financeiro_2026.pdf

- **Origem:** `Empresa/Financeiro/Relatorio_Financeiro_2026.pdf`
- **Tipo:** PDF · **Páginas:** 12 · **OCR:** 3 trecho(s), confiança média 91%
- **Metadados:** Título: ... · Autor: ... · Criado em: 2026-03-01

(texto, títulos, listas e tabelas em Markdown, na ordem visual do documento original)

*[Fim do documento [004]]*
```

### Cuidados com a qualidade da base

- **Ordenação natural e Unicode NFC:** arquivos numerados (`1_Aula`, `2_Aula`, ..., `10_Aula`) ficam em ordem correta e acentos decompostos (NFD) são normalizados.
- **Duplicados e grandes volumes:** arquivos idênticos aparecem como `duplicado de [ID]` sem repetir conteúdo; detecção escalável de versões semelhantes (≥ 80% de vocabulário em comum) para até 20.000 documentos.
- **PDF estruturado:** ordem visual inteligente (rótulos laterais saem antes dos parágrafos), filtro anti-falsas tabelas em PDFs de slides/Canva (`|Col1|Col2|`), remoção de divisórias e cabeçalhos/rodapés multi-linha (até 6 linhas) e junção de frases quebradas.
- **OCR em passe único (Texto Puro vs. Diagrama):** imagens de texto puro são transcritas integralmente no Markdown; apenas diagramas/esquemas visuais são salvos em `imagens/` preservando todo o texto OCR detectado dentro do card; repetições quase idênticas no mesmo documento são evitadas e palavras quebradas (`Aute nticacao`) são unidas sem colar siglas (`LTDA`).
- **Acentuação corrompida** (`ObservaÃ§Ã£o`) é corrigida automaticamente.
- **Dados pessoais** são preservados por padrão; ative a anonimização em `[3] → [21]`, com `[P]` na confirmação ou com `--redact-pii`.

---

## ⌨️ Linha de comando

```powershell
python copilot_consolidator/app.py --auto --lang pt                     # usa config.json
python copilot_consolidator/app.py --auto --input "D:\Docs" --output "D:\Base"
python copilot_consolidator/app.py --auto --redact-pii --workers 2 --no-ocr
python copilot_consolidator/app.py --help
```

## 🧪 Testes

```powershell
python -m unittest discover -s copilot_consolidator/tests -v
```

## 📄 Licença
Distribuído sob a licença **MIT**. Consulte [LICENSE](LICENSE).
