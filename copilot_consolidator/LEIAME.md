# Consolidador Local de Documentos para Microsoft Copilot & RAG v6.1

[![Plataforma: Windows 11 & macOS](https://img.shields.io/badge/Plataforma-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacidade: 100% Offline](https://img.shields.io/badge/Privacidade-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local-brightgreen.svg)](.)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)

Ferramenta **100% local e offline** que lê grandes volumes de documentos variados e gera **uma única base Markdown organizada** para o **Microsoft 365 Copilot** usar como fonte das respostas. Cada documento recebe um ID (`[001]`, `[002]`...) para que o Copilot cite exatamente de onde tirou cada informação.

Formatos lidos: PDF (com texto ou digitalizado), Word (`.docx`), Excel (`.xlsx`, `.xlsm`, `.xls`), PowerPoint (`.pptx`), CSV, TXT, Markdown, HTML, e-mails `.eml`, JSON, XML, logs e imagens (`.png`, `.jpg`, `.tif`, `.bmp`, `.webp`). Imagens dentro de PDFs, Word, Excel e PowerPoint passam por OCR local.

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

Em `saida_copilot/processado_AAAAMMDD_HHMMSS/` (cópia sempre atualizada em `saida_copilot/ultimo_processamento/`):

| Arquivo | Conteúdo |
| --- | --- |
| `corpus_completo.md` | **A base única**: instruções para o Copilot, índice de documentos agrupado por pasta e todos os documentos |
| `corpus_parte_XX.md` | A mesma base dividida em partes de até N MB (só quando não cabe em uma), cada uma com seu índice |
| `00_INDICE_MESTRE.md` | Só o catálogo (ID, documento, tipo, observações) |
| `por_pasta/*.md` | Um arquivo por pasta/departamento, para perguntas sobre um tema específico |
| `corpus_rag.jsonl` | Uma linha por página, com metadados, para bancos vetoriais (Chroma, Qdrant, LangChain) |
| `relatorio_auditoria.md` | Falhas, documentos sem texto, OCR com baixa confiança, duplicados, versões semelhantes e formatos ignorados |
| `PROMPT_PARA_COPILOT.txt` | Regras de resposta para colar no chat |

### Estrutura de cada documento na base

```markdown
## [004] Relatorio_Financeiro_2026.pdf

- **Origem:** `Empresa/Financeiro/Relatorio_Financeiro_2026.pdf`
- **Tipo:** PDF · **Páginas:** 12 · **OCR:** 3 trecho(s), confiança média 91%
- **Metadados:** Título: ... · Autor: ... · Criado em: 2026-03-01

(texto, títulos, listas e tabelas em Markdown, na ordem do documento original)

*[Fim do documento [004]]*
```

### Cuidados com a qualidade da base

- **Duplicados:** arquivos idênticos (ou com o mesmo texto) aparecem no índice como `duplicado de [ID]` e o conteúdo não é repetido.
- **Versões semelhantes:** documentos com ≥ 80% do vocabulário em comum recebem um aviso para o Copilot não misturar versões.
- **PDF:** tabelas viram tabelas Markdown sem duplicar o texto, cabeçalhos/rodapés repetidos e números de página são removidos (a 1ª página fica intacta), e frases quebradas em linhas são reunidas.
- **OCR:** linhas e colunas de recibos e formulários são preservadas; logos repetidos são lidos uma vez; palavras quebradas (`Aute nticacao`) são unidas sem colar siglas (`LTDA`) nem palavras reais.
- **Acentuação corrompida** (`ObservaÃ§Ã£o`) é corrigida automaticamente.
- **Dados pessoais** são preservados por padrão; ative a anonimização em `[3] → [20]`, com `[P]` na confirmação ou com `--redact-pii`.

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
