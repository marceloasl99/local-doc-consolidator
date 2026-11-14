# Plano de Implementação: local-doc-consolidator (v5.0 Enterprise)
**Documento de Planejamento Técnico & Arquitetura v5.0 (`plan1.md`)**  
**Projeto:** `marceloasl99/local-doc-consolidator`  
**Data:** 14 de Novembro de 2026  
**Status:** Revisão Técnica Aprovada — Implementação da Versão 5.0  
**Base:** Auditoria Técnica (Augusto MT - 7.5/10), testes de produção com 15 arquivos reais e especificações para ambientes corporativos de alta confidencialidade.

---

## 1. Resumo Executivo da Evolução v5.0

A versão 4.0 comprovou a viabilidade do processamento 100% offline com RapidOCR ONNX, cache SHA-256 e particionamento inteligente. A versão 5.0 eleva a solução para o padrão definitivo de produção corporativa, focando em:

1. **Fidelidade Contextual Completa nos Documentos Office:**
   - **Word (`.docx`):** Fim do isolamento de tabelas e imagens no final do arquivo. Extração ordenada iterando sobre `doc.element.body.iterchildren()`, posicionando tabelas e blocos de OCR de imagens exatamente abaixo do parágrafo de referência.
   - **PowerPoint (`.pptx`):** Extração de imagens e diagramas com OCR vinculada diretamente a cada slide (`### Slide N`), incluindo tabelas e dados de gráficos (`shape.has_chart`).
   - **Excel (`.xlsx`):** Ancoragem de imagens de comprovantes à aba correspondente (`ws._images`) e preservação de fórmulas calculadas.
   - **CSV Estruturado:** Detecção automática de delimitadores (`;`, `,`, `\t`, `|`) via `csv.Sniffer()` e conversão em tabelas Markdown alinhadas.

2. **Diagnóstico Explícito de 3 Estados no OCR (Fim de Falhas Silenciosas):**
   - Estado 1: `ok` (texto extraído acima do limiar de confiança).
   - Estado 2: `empty_or_low_conf` (OCR executado com sucesso, sem texto ou abaixo de 0.50).
   - Estado 3: `error` (falha técnica na inicialização ou decodificação — registrado no relatório e no Markdown com aviso explícito).

3. **Particionamento Estrito para Janela de Contexto Copilot:**
   - Criação de `00_INDICE_MESTRE.md` como catálogo global independente.
   - Cada parte (`corpus_parte_01.md`, etc.) inclui seu índice local resumido.
   - Fatiamento de documentos gigantes que excedam o limite de MB em trechos lógicos `[004] (Trecho 1/2)`.
   - Estimador de tokens de contexto com alerta de saturação da janela do Copilot (> 128k tokens).

4. **Blindagem e Segurança Corporativa (LGPD / Privacidade):**
   - Gravação atômica de cache (`.tmp` + `os.replace`).
   - Redação opcional de PII (`--redact-pii` / `"mascarar_pii": true`): mascaramento automático por Regex de CPFs, CNPJs, e-mails e telefones.
   - Controle seguro de concorrência: `max_workers` padrão em `min(4, os.cpu_count())` para evitar pico de memória RAM em notebooks.
   - Higiene de repositório: proteção de caminhos internos da empresa via `.gitignore`.

---

## 2. Matriz de Tarefas v5.0

```
[Módulo P0: OCR Modernizado & Diagnóstico de 3 Estados]
   ├── [P0-01] Diagnóstico de 3 estados (ok, empty, error) sem falhas silenciosas
   ├── [P0-02] Configuração explícita de idioma de OCR (--lang pt: LATIN / --lang en: EN)
   ├── [P0-03] Controle seguro de memória multicore (padrão min(4, cpus), configurável)
   └── [P0-04] Redação opcional de PII (Regex de CPF, CNPJ, e-mail, telefone)

[Módulo P1: Fidelidade Contextual de Documentos Office & CSV]
   ├── [P1-01] DOCX: Iteração corporal ordenada (Parágrafos -> Tabelas -> Imagens inline)
   ├── [P1-02] PPTX: Extração inline de imagens e gráficos por Slide
   ├── [P1-03] XLSX: Ancoragem de imagens à planilha/aba e preservação de fórmulas
   ├── [P1-04] CSV: Auto-detecção de delimitadores (csv.Sniffer) e formatação tabular
   └── [P1-05] Deduplicação dupla: hash binário e equivalência textual normalizada

[Módulo P2: Particionamento Estrito & Rastreabilidade Copilot]
   ├── [P2-01] Geração de 00_INDICE_MESTRE.md e índices locais em cada parte
   ├── [P2-02] Subdivisão lógica de documentos individuais > limite MB ([ID] (Trecho X/Y))
   ├── [P2-03] Estimativa de tokens e alerta de saturação de contexto Copilot
   └── [P2-04] Relatório HTML de auditoria opcional e exportação corpus_rag.jsonl

[Módulo P3: Higiene de Repositório & Segurança]
   ├── [P3-01] Gravação atômica de cache (.tmp + os.replace)
   ├── [P3-02] Templates de configuração (config.example.json, pastas.example.txt)
   └── [P3-03] Proteção de dados locais no .gitignore
```

---

## 3. Validação e Homologação v5.0

1. **Teste Unitário e de Extração:**
   - Verificação da ordem de elementos no Word (tabela antes ou depois de parágrafo específico).
   - Verificação de imagens inline nos slides do PowerPoint.
   - Verificação de delimitação de CSV `;` vs `,`.
2. **Teste de Carga nos 15 Arquivos Reais:**
   - Execução com 4 workers controlados.
   - Verificação do tempo de execução e consumo de memória.
3. **Commit e Push para origin/main.**
