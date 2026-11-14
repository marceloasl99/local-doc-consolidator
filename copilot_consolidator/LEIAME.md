# Consolidador Local de Documentos para Microsoft Copilot (v4) 🚀📁

[![Platform: Windows 11](https://img.shields.io/badge/Platform-Windows%2011-lightgrey.svg)](.)
[![Privacy: 100% Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success.svg)](.)
[![OCR: RapidOCR ONNX](https://img.shields.io/badge/OCR-Local%20%26%20Deep-brightgreen.svg)](.)
[![Performance: Multicore Parallel](https://img.shields.io/badge/Performance-Multicore%20Parallel-orange.svg)](.)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Ferramenta corporativa **100% local, offline e segura** desenvolvida em Python para indexar, converter e consolidar grandes volumes de documentos confidenciais (PDFs com texto e escaneados, Word com imagens, Excel com imagens, apresentações, tabelas e imagens escaneadas) em bases Markdown perfeitamente estruturadas para o **Microsoft 365 Copilot**.

Processa **centenas a milhares de arquivos** espalhados em dezenas de pastas sem enviar nenhum dado para a nuvem e sem exigir privilégios de Administrador no Windows (sem dependência de instaladores externos pesados como Poppler ou Tesseract).

---

## 🌟 Principais Recursos e Diferenciais

* **🧠 OCR Local Profundo (100% Offline - RapidOCR ONNX):**
  * **PDFs com Imagens / Escaneados:** Reconhece páginas escaneadas completas e extrai textos de imagens embutidas em páginas mistas.
  * **Word (`.docx`) com Imagens:** Extrai automaticamente prints de sistemas, fotos de contratos e comprovantes colados no documento e aplica OCR.
  * **Excel (`.xlsx`) com Imagens:** Extrai comprovantes fiscais, recibos e tabelas coladas como imagem nas abas da planilha e transcreve via OCR.
  * **PowerPoint (`.pptx`) com Imagens:** Lê gráficos e diagramas com texto presentes nos slides.
  * **Imagens Avulsas (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`):** Converte notas, recibos e documentos escaneados diretamente em blocos Markdown com OCR.
* **📅 Organização por Timestamp e Pacote Compactado (.ZIP):**
  * Cada execução cria uma pasta isolada identificada por data e hora (ex: `processado_20261114_112841/`).
  * Gera opcionalmente um arquivo compactado `.zip` (ex: `consolidado_20261114_112841.zip`) pronto para envio.
  * Mantém uma pasta espelho `ultimo_processamento/` para acesso rápido aos arquivos mais recentes.
* **📊 Contagem de Arquivos, Volume de Dados e Confirmação Prévia:**
  * Calcula em tempo real o total de arquivos e volume de dados em MB/GB.
  * Tela de confirmação antes de executar com detalhamento por extensão (PDF, Word, Excel, Imagens, etc.) e ajuste rápido de limites.
* **📁 Navegação e Confirmação de Pastas:**
  * Opção de abrir a janela visual nativa do Windows Explorer para navegar e selecionar pastas.
  * Ao selecionar uma pasta, mostra a quantidade de arquivos e tamanho total antes de confirmar a adição.
* **📦 Estratégias Flexíveis de Separação e Particionamento:**
  * **Arquivo Único Geral (`corpus_completo.md`):** Toda a base consolidada com o Índice Mestre `[001]`..`[N]`.
  * **Agrupamento por Tamanho Máximo em MB:** Divida automaticamente a base quando atingir **X MB** (ajustável livremente para 8 MB, 50 MB, 100 MB ou **200 MB**).
  * **Agrupamento por Quantidade de Documentos:** Separe em lotes de **N documentos** (ex: 25, 50 ou 100 docs por arquivo).
  * **1 Arquivo por Documento (Individual 1:1):** Gera versões Markdown individuais para cada documento em `saida_copilot/individual/`.
  * **Agrupamento por Pasta Departamental:** Gera 1 consolidado para cada subpasta monitorada em `saida_copilot/por_pasta/`.
* **⚡ Processamento Multicore Paralelo (CPU):**
  * Utiliza todos os núcleos lógicos do processador (`ProcessPoolExecutor`) no Windows, processando dezenas de arquivos pesados simultaneamente.
* **🔄 Cache Criptográfico Incremental (SHA-256):**
  * Processa tudo na primeira vez. Nas execuções seguintes, valida o hash criptográfico e **só reconverte os arquivos adicionados ou modificados**, gerando o consolidado em frações de segundo.

---

## 🛠️ Como Executar no Windows 11

### Método 1: Duplo Clique (Recomendado)
1. Dê **duplo clique** no arquivo **`iniciar_windows.bat`**.
2. Na primeira execução, ele criará automaticamente o ambiente virtual isolado (`.venv`), instalará as bibliotecas necessárias e abrirá o menu interativo no terminal.

### Método 2: Via PowerShell / Prompt de Comando
```powershell
cd C:\55\art\copilot_consolidator
iniciar_windows.bat
```
*(Para execução autônoma sem pausas de menu, use `python app.py --auto`).*

---

## 🖥️ Menu Interativo do Terminal

Ao abrir o aplicativo, você terá controle total pelo menu:

```text
============================================================================
   CONSOLIDADOR LOCAL DE DOCUMENTOS PARA MICROSOFT COPILOT (v4)
   Ambiente: Windows 11 | 100% Offline | 12 Núcleos CPU Concorrentes
============================================================================
 Status Atual:
  • Pastas monitoradas      : 3
  • Arquivos localizados    : 10 (488.1 KB)
  • Multicore Paralelo      : ATIVADO (12 cores)
  • OCR em PDFs e Imagens   : ATIVADO
  • OCR Imagens Embutidas   : ATIVADO (Word/Excel/PPTX)
  • Divisão de Arquivo      : Único | Limite 8.0MB | Por Pasta
  • Organização / Timestamp : Pasta com Data/Hora + Compactar .ZIP
  • Destino base            : C:\55\art\copilot_consolidator\saida_copilot
----------------------------------------------------------------------------
 [1] EXECUTAR CONVERSÃO E GERAR ARQUIVOS PARA O COPILOT
 [2] Adicionar pastas (Windows Explorer, digitar caminhos ou Pasta Mãe)
 [3] Listar / Verificar / Remover pastas configuradas (3)
 [4] Recarregar lista de pastas do arquivo 'pastas.txt'
 [5] Configurar opções de saída (Tamanhos/MB, Lotes, ZIP, Timestamp, OCR)
 [6] Limpar cache de extração SHA-256 (forçar releitura completa)
 [0] Sair
============================================================================
```

---

## 🤖 Como Usar no Microsoft 365 Copilot

1. Acesse o **Microsoft 365 Copilot Chat** ou **Copilot Notebooks**.
2. Anexe o arquivo `corpus_completo.md` (ou os arquivos menores `corpus_parte_01.md`, lotes ou pastas da pasta `saida_copilot/ultimo_processamento/`).
3. Cole o seguinte prompt inicial presente em `PROMPT_PARA_COPILOT.txt`:

```text
Analise a base documental corporativa em anexo. Todos os documentos estão catalogados no ÍNDICE MESTRE DE DOCUMENTOS no início do arquivo, identificados por IDs únicos entre colchetes como [001], [002], etc.

Regras obrigatórias para suas respostas:
1. Para cada informação, valor, número ou regra citada, cite obrigatoriamente o ID e o caminho de origem (ex: 'Conforme [001] 01_Financeiro/Relatorio_2026.pdf...').
2. Se existirem dados divergentes entre dois ou mais documentos, aponte ambos os IDs e destaque a diferença com precisão.
3. Não invente nenhuma informação ausente nos documentos anexados; caso algo não conste na base, afirme expressamente.
```

---

## 📋 Formatos Suportados

* **PDF (`.pdf`)**: Leitura digital nativa de texto e tabelas + OCR offline condicional em páginas escaneadas e em imagens embutidas.
* **Microsoft Word (`.docx`)**: Extração de títulos, parágrafos, tabelas formatadas em Markdown e OCR em capturas de tela/recibos embutidos.
* **Microsoft Excel (`.xlsx`, `.xls`)**: Conversão de todas as abas em tabelas Markdown legíveis por IA e OCR em comprovantes colados na planilha.
* **Microsoft PowerPoint (`.pptx`)**: Extração de slides, anotações de rodapé, tabelas e OCR em imagens dos slides.
* **Arquivos de Imagem (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`)**: OCR direto com extração de texto e resolução.
* **Texto e Dados (`.txt`, `.md`, `.csv`)**: Detecção automática de encoding (`UTF-8`, `Latin-1`, `CP1252`) e dialetos delimitadores.

---

## 📝 Licença
Distribuído sob a licença **MIT**.
