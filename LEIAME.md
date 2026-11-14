# Consolidador Local de Documentos para Microsoft Copilot & RAG 🚀📁

[![Plataforma: Windows 11 & macOS](https://img.shields.io/badge/Plataforma-Windows%2011%20%7C%20macOS-lightgrey.svg)](.)
[![Privacidade: 100% Offline](https://img.shields.io/badge/Privacidade-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Profundo-brightgreen.svg)](.)
[![Performance: Multicore Paralelo](https://img.shields.io/badge/Performance-Multicore%20Paralelo-orange.svg)](.)
[![Exportação: Markdown + JSONL RAG](https://img.shields.io/badge/Exporta%C3%A7%C3%A3o-Markdown%20%2B%20JSONL%20RAG-purple.svg)](.)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)

Ferramenta corporativa em Python, **100% local, privada e offline**, projetada para indexar, converter e consolidar volumes massivos de documentos corporativos confidenciais (PDFs nativos e escaneados, Word com capturas de tela, Excel com comprovantes embutidos, apresentações PowerPoint e fotos avulsas de documentos) em Markdown contínuo para o **Microsoft 365 Copilot** e base estruturada JSONL para **RAG local (Chroma, Qdrant, Milvus, LangChain)**.

Processe **centenas a milhares de arquivos** espalhados em pastas departamentais sem nenhum vazamento para a nuvem e sem exigir privilégios de administrador no Windows ou dependências externas como Tesseract ou Poppler.

---

## 🌟 Principais Recursos

1. **🧠 OCR Local Profundo com RapidOCR (ONNX Runtime):**
   - Transcreve páginas escaneadas e imagens embutidas em Word, Excel, PowerPoint e PDFs.
   - 100% offline, executando na CPU ou GPU local sem nenhuma requisição de rede.
2. **🛡️ Filtragem Inteligente e Deduplicação:**
   - **Filtro de Confiança (`score >= 0.50`):** Remove ruídos e artefatos de compressão visual.
   - **Deduplicação de Imagens por Hash SHA-256:** Logotipos e cabeçalhos repetidos em dezenas de páginas são processados no OCR apenas uma vez.
   - **Filtro de Proporção Seguro (`> 20:1`):** Descarta linhas divisórias sem perder fluxogramas, tabelas largas ou cronogramas.
   - **Normalização Textual:** Reconecta palavras hifenizadas e quebras duras de linha.
3. **📦 Particionamento Flexível e Exportação Dual:**
   - **Markdown Contínuo (`corpus_completo.md`):** Texto limpo e fluido sem divisores intrusivos de página no meio do arquivo.
   - **Base Estruturada para RAG (`corpus_rag.jsonl`):** Registro página a página com metadados para indexadores vetoriais.
   - **Divisão por Tamanho (MB):** Particiona quando atinge o limite configurado (8 MB, 50 MB, 100 MB até 200 MB).
   - **Divisão por Lote e por Pasta:** Saídas organizadas por departamento em `saida_copilot/por_pasta/`.
4. **📅 Organização com Timestamp e Validação de ZIP:**
   - Cada execução grava em pasta própria (`processado_YYYYMMDD_HHMMSS/`).
   - Compactação automática em `.zip` testada com `zipfile.testzip()`.
5. **⚡ Multicore Paralelo e Cache Criptográfico:**
   - Uso de todos os núcleos da CPU para máxima velocidade.
   - Cache SHA-256 parametrizado com DPI e limiares de OCR.
6. **🤖 Prompts Anti-Alucinação em Português e Inglês:**
   - Gera `PROMPT_PARA_COPILOT.txt` e `COPILOT_PROMPT.txt` para respostas estritamente fundamentadas.

---

## 🚀 Como Executar

### Windows (Português)
Dê dois cliques no arquivo **`iniciar_windows.bat`**.

### Windows (Inglês)
Dê dois cliques no arquivo **`start.bat`**.

### macOS
- Português: `./copilot_consolidator/iniciar_macos.command`
- Inglês: `./copilot_consolidator/start_macos.command`

### Linha de Comando (Modo Automatizado)
```powershell
python copilot_consolidator/app.py --auto --lang pt
```

---

## 📄 Licença
Distribuído sob a licença **MIT**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
