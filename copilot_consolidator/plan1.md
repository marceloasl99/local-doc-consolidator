# Plano de Implementação e Otimização: local-doc-consolidator (Revisão 2 - Aprovada)
**Documento de Planejamento Técnico (`plan1.md`)**  
**Projeto:** `marceloasl99/local-doc-consolidator`  
**Data:** 14 de Novembro de 2026  
**Status:** Aprovado para Execução Imediata  
**Diretriz Chave Alinhada:** Sem marcadores de paginação no meio do Markdown (texto contínuo e limpo para o Copilot). Rastreamento de páginas restrito aos metadados do cabeçalho e ao `corpus_rag.jsonl`.

---

## 1. Decisões Técnicas Alinhadas & Veredito Final

1. **Supressão de Paginação Interna no Markdown (P1-02):**
   - Atendendo estritamente à regra de negócio (*"sem paginação no meio do arquivo"*), o Markdown consolidado não terá quebras ou cabeçalhos repetidos como `### Página X` ou `--- [Página X de Y] ---`. O texto correrá unificado e contínuo.
   - O número total de páginas constará no cabeçalho mestre de metadados (`- **Dimensão Documental:** 119 páginas`).
   - A segmentação por página será preservada exclusivamente no arquivo estruturado `corpus_rag.jsonl` (Fase P3) para indexadores vetoriais.

2. **Filtro de Proporção de Aspecto Seguro (P0-03):**
   - O filtro descartará apenas proporções extremas `> 20:1` (linhas divisórias, réguas e bordas de artefatos visuais), preservando gráficos horizontais, cronogramas, tabelas largas e fluxogramas analíticos.
   - Dimensão mínima para OCR: maior lado `>= 150px` e menor lado `>= 25px`.

3. **Detecção de Página Escaneada Híbrida (P0-04):**
   - Página escaneada acionada se: `len(texto_nativo.strip()) < 80 and len(imagens) > 0` OU qualquer imagem embutida cobrir mais de `60%` da área física da página.

4. **Reordenação da Fase P1:**
   - 1º Filtro de confiança (< 0.50) e dimensão;
   - 2º Deduplicação por hash SHA-256 dos bytes originais da imagem;
   - 3º Normalização de texto OCR (hifenização e quebras duras) restrita ao bloco de imagem;
   - 4º Limpeza de `<br>` em tabelas;
   - 5º Metadados estruturados no cabeçalho;
   - 6º Texto contínuo sem divisores de página intermediários no Markdown.

5. **Deduplicação de Imagens por Hash de Bytes Brutos (P1-04):**
   - O hash SHA-256 é calculado nos bytes originais (`doc.extract_image(xref)["image"]` ou `z.read(nome)`) antes de qualquer conversão para evitar divergências de compressão PNG/pixmap.

6. **100% Offline Garantido (P2-01):**
   - Remoção de `markitdown` de `requirements.txt`. Zero chamadas externas de rede.

---

## 2. Ordem de Implementação

```
[Etapa 1: Fase P0 - Correções Críticas & Limiares]
   ├── [P0-01] Flag de cache com DPI, dimensões e limiares (flag_ocr)
   ├── [P0-02] Constante declarativa do Prompt do Copilot
   ├── [P0-03] Limiar 150px + Filtro de aspecto seguro (> 20:1)
   └── [P0-04] Detecção proporcional/híbrida de página escaneada

[Etapa 2: Fase P1 - Qualidade de Extração & RAG/Copilot]
   ├── [P1-06] Filtro de score/confiança RapidOCR (< 0.50)
   ├── [P1-04] Deduplicação por hash SHA-256 de bytes originais de imagens
   ├── [P1-03] Normalização de texto OCR (hífen e quebra) sem quebrar delimitadores
   ├── [P1-05] Limpeza de quebras <br> em tabelas Markdown
   ├── [P1-01] Metadados estruturados no cabeçalho do documento (autor, páginas, OCR)
   └── [P1-02] Remoção de paginação interna no Markdown (texto contínuo e limpo)

[Etapa 3: Fase P2 - Robustez, 100% Offline & Limpeza]
   ├── [P2-01] Remoção de markitdown de requirements.txt
   ├── [P2-02] Logging explícito no terminal para falhas de I/O e ZIP
   ├── [P2-03] Filtro seletivo de Warnings
   ├── [P2-04] Tratamento claro para PDFs com senha / corrompidos
   └── [P2-05] Arquivamento de pacotes ZIP sem poluir a raiz

[Etapa 4: Fase P3 - UX, RAG JSONL & Finalização]
   ├── [P3-01] Barra de progresso com contador e ETA no terminal
   ├── [P3-02] Geração simultânea de corpus_rag.jsonl com metadados por página
   ├── [P3-03] Validação segura de config.json com notificação visual
   ├── [P3-04] Verificação automática de integridade zipfile.testzip()
   └── [P3-05] Criação do arquivo LICENSE (MIT)
```

---

## 3. Validação e Homologação

- Validação em amostra rápida (`teste_rapido/`).
- Validação completa dos 15 arquivos de `5555.zip`.
- Auditoria do arquivo `corpus_completo.md` (conferir ausência de quebras no meio do texto e limpeza de ruídos).
- Auditoria de `corpus_rag.jsonl` (conferir metadados estruturados).
- Commit e push no GitHub (`origin/main`).
