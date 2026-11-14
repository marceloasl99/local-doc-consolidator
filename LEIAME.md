# Consolidador Local de Documentos para Microsoft Copilot & RAG v6.0 🚀📁

[![Plataforma: Windows 11 & macOS](https://img.shields.io/badge/Plataforma-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacidade: 100% Offline](https://img.shields.io/badge/Privacidade-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Profundo-brightgreen.svg)](.)
[![Privacidade: Dados Originais Preservados](https://img.shields.io/badge/Privacidade-Dados%20Originais%20Preservados-blue.svg)](.)
[![Performance: Multicore Paralelo](https://img.shields.io/badge/Performance-Multicore%20Paralelo-orange.svg)](.)
[![Exportação: Markdown + JSONL RAG](https://img.shields.io/badge/Exporta%C3%A7%C3%A3o-Markdown%20%2B%20JSONL%20RAG-purple.svg)](.)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)

Ferramenta corporativa em Python, **100% local, privada e offline**, projetada para indexar, converter e consolidar volumes massivos de documentos corporativos confidenciais (PDFs nativos e escaneados, Word com capturas de tela, Excel com comprovantes embutidos, apresentações PowerPoint e fotos avulsas de documentos) em Markdown contínuo para o **Microsoft 365 Copilot** e base estruturada JSONL para **RAG local (Chroma, Qdrant, Milvus, LangChain)**.

Processe **centenas a milhares de arquivos** espalhados em pastas departamentais sem nenhum vazamento para a nuvem e sem exigir privilégios de administrador no Windows ou dependências externas como Tesseract ou Poppler.

---

## 🌟 Principais Recursos (v6.0)

1. **🧠 OCR Local Profundo com RapidOCR (ONNX Runtime):**
   - Transcreve páginas escaneadas e imagens embutidas em Word, Excel, PowerPoint e PDFs.
   - 100% offline, executando na CPU ou GPU local sem nenhuma requisição de rede.
2. **🔤 Correção Automática de Mojibake & Encodings:**
   - Detecta e repara corrupções de codificação em arquivos de entrada (ex: UTF-8 interpretado como CP1252/Latin-1 e re-salvo, convertendo `ObservaÃ§Ã£o` -> `Observação`, `MÃ©todo` -> `Método`, `nÃ£o` -> `não`).
3. **✨ Recombinação Fonotática de Palavras Quebradas por OCR:**
   - Desfaz erros de segmentação espacial e kerning do OCR (ex: `Aute nticacao` -> `Autenticacao`, `comprova nte` -> `comprovante`, `docu mento` -> `documento`, `estabeleci mento` -> `estabelecimento`).
4. **🔍 Detecção de Conteúdo Quase-Duplicado (Near-Duplicates):**
   - Analisa sobreposição lexical (Jaccard / Overlap >= 80%) entre documentos distintos (versões revisadas, templates com nomes alterados) e registra avisos explícitos no `00_INDICE_MESTRE.md` e nos relatórios de auditoria para evitar alucinações no Copilot.
5. **🛡️ Anonimização Opcional de PII (Dados Originais Preservados por Padrão):**
   - Por padrão, preserva todos os dados na íntegra para máxima fidelidade analítica (mantém nomes, CPFs, RGs, endereços, e-mails e telefones).
   - Anonimização ativável a qualquer momento via opção `[5] -> [9]` do menu, tecla `[P]` na tela de confirmação prévia, ou flag `--redact-pii` para ocultar registros sensíveis.
6. **📦 Particionamento Inteligente Sem Redundância:**
   - **Markdown Contínuo (`corpus_completo.md`):** Texto limpo e fluido sem divisores intrusivos de página no meio do arquivo.
   - **Eliminação de Partes Redundantes:** Quando o corpus é menor que o limite de partição (ex: < 8 MB), suprime `corpus_parte_01.md` solteiro, economizando 50% de espaço em disco e ZIP.
   - **Base Estruturada para RAG (`corpus_rag.jsonl`):** Registro página a página com metadados para indexadores vetoriais.
7. **📊 Relatórios Nativos Bilíngues e Estimativa Transparente de Tokens:**
   - Gera `relatorio_auditoria.md` (100% em português) e `audit_report.md` (100% em inglês).
   - Rótulo explícito de tokens estimados via heurística (4 caracteres/token).
8. **⚡ Multicore Seguro e Cache Criptográfico:**
   - Limite de concorrência seguro (`min(4, os.cpu_count())`) para proteção de memória RAM.
   - Cache SHA-256 atômico parametrizado com DPI, limiares de OCR e flags de PII.

---

## 🚀 Como Executar

### Windows (Português)
* **PowerShell / Windows Terminal (Na mesma janela):**
  ```powershell
  .\iniciar.ps1
  ```
* **Prompt de Comando / Windows Explorer:**
  ```cmd
  iniciar_windows.bat
  ```

### Windows (Inglês)
* **PowerShell / Windows Terminal (In-Window):**
  ```powershell
  .\start.ps1
  ```
* **Prompt de Comando / Windows Explorer:**
  ```cmd
  start.bat
  ```

### macOS
- Português: `./copilot_consolidator/iniciar_macos.command`
- Inglês: `./copilot_consolidator/start_macos.command`

### Linha de Comando (Modo Automatizado)
```powershell
# Execução padrão (Dados originais preservados na íntegra, Português)
python copilot_consolidator/app.py --auto --lang pt

# Execução em Inglês
python copilot_consolidator/app.py --auto --lang en

# Opcional: Ativar anonimização de PII (mascarar nomes, CPFs, RGs, endereços, telefones)
python copilot_consolidator/app.py --auto --lang pt --redact-pii
```

---

## 🧪 Testes Automatizados

Para executar a suíte de testes unitários:
```powershell
python -m unittest discover -s copilot_consolidator/tests -v
```

---

## 📄 Licença
Distribuído sob a licença **MIT**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
