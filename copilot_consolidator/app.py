#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidador Local de Documentos para Microsoft Copilot (v4 - Enterprise Edition)
Local Document Consolidator for Microsoft Copilot (v4 - Enterprise Edition)
Bilingual: English (start.bat) & Português (iniciar_windows.bat)
100% Offline | Multicore Parallel | Private Local Processing
"""

import csv
import hashlib
import io
import json
import os
import platform
import re
import shutil
import sys
import time
import warnings
import zipfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

# Suprime avisos seletivos de layout do PyMuPDF e avisos de bibliotecas conhecidas
os.environ["PYMUPDF_SUGGEST_LAYOUT_ANALYZER"] = "0"
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*fitz.*")
warnings.filterwarnings("ignore", message=".*pymupdf_layout.*")

# Habilita cores ANSI no terminal do Windows
if platform.system() == "Windows":
    os.system("")


class Cor:
    TITULO = "\033[96m\033[1m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    VERMELHO = "\033[91m"
    AZUL = "\033[94m"
    MAGENTA = "\033[95m"
    NEGRITO = "\033[1m"
    RESET = "\033[0m"


BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
PASTAS_TXT_FILE = BASE_DIR / "pastas.txt"

EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp"}
EXTENSOES_OFFICE = {".docx", ".pptx", ".xlsx", ".xls"}
EXTENSOES_TEXTO = {".txt", ".md", ".csv"}
EXTENSOES_PDF = {".pdf"}

EXTENSOES_PADRAO = sorted(list(EXTENSOES_PDF | EXTENSOES_OFFICE | EXTENSOES_TEXTO | EXTENSOES_IMAGEM))

PROMPT_SISTEMA_COPILOT_PT = (
    "Analise a base documental corporativa indexada em anexo. Todos os documentos estão formalmente catalogados "
    "no ÍNDICE MESTRE DE DOCUMENTOS no início do arquivo, identificados por IDs únicos entre colchetes como [001], [002], etc.\n\n"
    "Diretrizes mandatórias e inegociáveis para suas respostas:\n"
    "1. Rastreabilidade Obrigatória de Fontes: Para cada dado, valor, diretriz, data ou fato mencionado, cite expressamente "
    "o ID do documento e seu caminho relativo de origem (exemplo: 'Conforme [001] Relatorio_Financeiro.pdf...').\n"
    "2. Identificação de Divergências: Se houver informações conflitantes entre documentos distintos, aponte os IDs de todos "
    "os documentos envolvidos e descreva precisamente a discrepância constatada.\n"
    "3. Postura Anti-Alucinação Estrita: Responda única e exclusivamente com base nos fatos contidos nos documentos anexados. "
    "Jamais deduza, extrapole ou invente informações ausentes. Caso um dado solicitado não conste na base, afirme expressamente: "
    "'Esta informação não consta nos documentos fornecidos.'\n"
    "4. Integridade de Citações: Transcreva termos técnicos, códigos e nomes exatamente como grafados no texto original."
)

PROMPT_SISTEMA_COPILOT_EN = (
    "Analyze the attached indexed corporate knowledge base. All documents are formally cataloged in the "
    "MASTER DOCUMENT INDEX at the beginning of the file, identified by unique IDs in square brackets such as [001], [002], etc.\n\n"
    "Mandatory and non-negotiable guidelines for your answers:\n"
    "1. Mandatory Source Traceability: For every data point, figure, rule, date, or fact cited, explicitly reference "
    "the document ID and relative source path (e.g., 'According to [001] Finance_Report.pdf...').\n"
    "2. Conflict Resolution: If there is conflicting information between different documents, specify the IDs of all "
    "documents involved and detail the discrepancy accurately.\n"
    "3. Strict Anti-Hallucination Stance: Answer solely and exclusively based on facts contained in the attached documents. "
    "Never assume, extrapolate, or invent missing information. If requested information is not in the knowledge base, "
    "state explicitly: 'This information is not present in the provided documents.'\n"
    "4. Citation Integrity: Retain technical terms, identifiers, codes, and names exactly as written in the source text."
)

I18N = {
    "pt": {
        "titulo_app": "CONSOLIDADOR LOCAL DE DOCUMENTOS PARA MICROSOFT COPILOT (v4)",
        "subtitulo_app": "Ambiente: Windows 11 | 100% Offline | {cpus} Núcleos CPU Concorrentes",
        "status_atual": "Status Atual:",
        "pastas_monitoradas": "Pastas monitoradas",
        "arquivos_localizados": "Arquivos localizados",
        "multicore_paralelo": "Multicore Paralelo",
        "ocr_pdf_img": "OCR em PDFs e Imagens",
        "ocr_embutidas": "OCR Imagens Embutidas",
        "divisao_arq": "Divisão de Arquivo",
        "org_timestamp": "Organização / Timestamp",
        "destino_base": "Destino base",
        "ativado": "ATIVADO",
        "desativado": "DESATIVADO",
        "nenhuma": "Nenhuma",
        "pasta_datahora": "Pasta com Data/Hora",
        "direta": "Direta",
        "compactar_zip": "Compactar .ZIP",
        "op_executar": "[1] EXECUTAR CONVERSÃO E GERAR ARQUIVOS PARA O COPILOT",
        "op_adicionar": "[2] Adicionar pastas (Windows Explorer, digitar caminhos ou Pasta Mãe)",
        "op_listar": "[3] Listar / Verificar / Remover pastas configuradas ({qtd})",
        "op_recarregar": "[4] Recarregar lista de pastas do arquivo 'pastas.txt'",
        "op_config": "[5] Configurar opções de saída (Tamanhos/MB, Lotes, ZIP, Timestamp, OCR, RAG JSONL)",
        "op_limpar_cache": "[6] Limpar cache de extração SHA-256 (forçar releitura completa)",
        "op_idioma": "[L] Idioma / Language (Atual: Português)",
        "op_sair": "[0] Sair",
        "prompt_opcao": "Digite o número da opção desejada: ",
        "sem_pastas": "[!] Nenhuma pasta configurada! Adicione pastas na opção [2] ou no arquivo pastas.txt.",
        "sem_arquivos": "[!] Nenhum arquivo compatível foi encontrado nas pastas informadas.",
        "iniciando_conversao": "INICIANDO CONVERSÃO LOCAL PARA MICROSOFT COPILOT",
        "processando_paralelo": "Processando documentos em paralelo...",
        "gerando_zip": "Gerando arquivo compactado .ZIP e verificando integridade...",
        "concluido_sucesso": "CONSOLIDAÇÃO CONCLUÍDA COM SUCESSO EM {tempo:.2f}s!",
        "pasta_execucao": "Pasta da execução",
        "sucesso_total": "Sucesso / Total",
        "volume_proc": "Volume processado",
        "disponiveis_copilot": "Arquivos disponíveis para o Microsoft Copilot / RAG:",
        "pacote_zip": "Pacote Compactado (.ZIP)",
        "voltar_menu": "Pressione ENTER para voltar ao menu principal...",
        "base_copilot_titulo": "# BASE DOCUMENTAL CORPORATIVA CONSOLIDADA PARA COPILOT",
        "data_consolida": "Data de consolidação",
        "total_docs_index": "Total de documentos indexados",
        "vol_original": "Volume original de dados",
        "pastas_origem": "Pastas de origem",
        "indice_mestre": "ÍNDICE MESTRE DE DOCUMENTOS (ID -> ORIGEM)",
        "id_ref": "ID de Referência",
        "pasta_dep": "Pasta Departamental",
        "cam_rel": "Caminho Relativo",
        "metadados": "Metadados",
        "dim_doc": "Dimensão Documental",
        "qualidade_ocr": "Qualidade OCR",
        "conf_media": "Confiança média estimada",
        "sem_texto": "*(Documento sem texto extraível)*",
        "img_sem_texto": "*(Imagem sem texto legível detectado)*",
        "texto_extraido_ocr": "Texto Extraído via OCR Local",
        "imagem": "Imagem",
        "resolucao": "Resolução",
        "formato": "Formato",
        "pagina_escaneada": "*(Página {p} escaneada - extraída via OCR local)*",
        "pagina_escaneada_sem_texto": "*(Página {p} escaneada sem texto legível)*",
        "img_embutida": "OCR Imagem Embutida #{idx} ({w}x{h} px)",
        "doc_protegido_senha": "*(Documento PDF protegido por senha - extração restrita sem credenciais)*",
        "relatorio_auditoria_titulo": "# RELATÓRIO DE AUDITORIA E INDEXAÇÃO LOCAL (MULTICORE)",
        "tempo_total": "Tempo total de execução",
        "total_encontrados": "Total de documentos encontrados",
        "indexados_sucesso": "Indexados com sucesso",
        "recup_cache": "Recuperados instantaneamente do cache SHA-256",
        "docs_ocr_aplicado": "Documentos/Páginas/Imagens com OCR aplicado",
        "docs_identicos": "Documentos com conteúdo idêntico (duplicados)",
        "falhas_leitura": "Falhas de leitura",
        "estrutura_saida": "Estrutura de Saída Gerada",
        "docs_identicos_detectados": "Documentos Idênticos Detectados",
        "falhas_extracao": "Falhas de Extração",
    },
    "en": {
        "titulo_app": "LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (v4)",
        "subtitulo_app": "Environment: Windows 11 | 100% Offline | {cpus} Concurrent CPU Cores",
        "status_atual": "Current Status:",
        "pastas_monitoradas": "Monitored folders",
        "arquivos_localizados": "Located files",
        "multicore_paralelo": "Multicore Parallel",
        "ocr_pdf_img": "OCR in PDFs and Images",
        "ocr_embutidas": "OCR Embedded Images",
        "divisao_arq": "File Splitting",
        "org_timestamp": "Organization / Timestamp",
        "destino_base": "Base destination",
        "ativado": "ENABLED",
        "desativado": "DISABLED",
        "nenhuma": "None",
        "pasta_datahora": "Timestamped Folder",
        "direta": "Direct",
        "compactar_zip": "Compress .ZIP",
        "op_executar": "[1] EXECUTE CONSOLIDATION AND GENERATE COPILOT FILES",
        "op_adicionar": "[2] Add folders (Windows Explorer, type path, or Parent Folder)",
        "op_listar": "[3] List / Check / Remove configured folders ({qtd})",
        "op_recarregar": "[4] Reload folder list from 'pastas.txt'",
        "op_config": "[5] Configure output options (Sizes/MB, Batches, ZIP, Timestamp, OCR, RAG JSONL)",
        "op_limpar_cache": "[6] Clear SHA-256 extraction cache (force full re-extraction)",
        "op_idioma": "[L] Language / Idioma (Current: English)",
        "op_sair": "[0] Exit",
        "prompt_opcao": "Enter desired option: ",
        "sem_pastas": "[!] No folder configured! Add folders in option [2] or in pastas.txt.",
        "sem_arquivos": "[!] No compatible files found in configured folders.",
        "iniciando_conversao": "STARTING LOCAL CONSOLIDATION FOR MICROSOFT COPILOT",
        "processando_paralelo": "Processing documents in parallel...",
        "gerando_zip": "Creating .ZIP archive and verifying integrity...",
        "concluido_sucesso": "CONSOLIDATION COMPLETED SUCCESSFULLY IN {tempo:.2f}s!",
        "pasta_execucao": "Execution folder",
        "sucesso_total": "Success / Total",
        "volume_proc": "Processed volume",
        "disponiveis_copilot": "Files available for Microsoft Copilot / RAG:",
        "pacote_zip": "Compressed Package (.ZIP)",
        "voltar_menu": "Press ENTER to return to main menu...",
        "base_copilot_titulo": "# CONSOLIDATED CORPORATE DOCUMENT BASE FOR COPILOT",
        "data_consolida": "Consolidation Date",
        "total_docs_index": "Total indexed documents",
        "vol_original": "Original data volume",
        "pastas_origem": "Source folders",
        "indice_mestre": "MASTER DOCUMENT INDEX (ID -> SOURCE)",
        "id_ref": "Reference ID",
        "pasta_dep": "Department Folder",
        "cam_rel": "Relative Path",
        "metadados": "Metadata",
        "dim_doc": "Document Dimensions",
        "qualidade_ocr": "OCR Quality",
        "conf_media": "Estimated average confidence",
        "sem_texto": "*(Document without extractable text)*",
        "img_sem_texto": "*(Image without readable text detected)*",
        "texto_extraido_ocr": "Text Extracted via Local OCR",
        "imagem": "Image",
        "resolucao": "Resolution",
        "formato": "Format",
        "pagina_escaneada": "*(Page {p} scanned - extracted via local OCR)*",
        "pagina_escaneada_sem_texto": "*(Page {p} scanned without readable text)*",
        "img_embutida": "OCR Embedded Image #{idx} ({w}x{h} px)",
        "doc_protegido_senha": "*(PDF document password protected - restricted extraction)*",
        "relatorio_auditoria_titulo": "# LOCAL AUDIT AND INDEXING REPORT (MULTICORE)",
        "tempo_total": "Total execution time",
        "total_encontrados": "Total documents found",
        "indexados_sucesso": "Indexed successfully",
        "recup_cache": "Retrieved instantly from SHA-256 cache",
        "docs_ocr_aplicado": "Documents/Pages/Images with OCR applied",
        "docs_identicos": "Documents with identical content (duplicates)",
        "falhas_leitura": "Extraction failures",
        "estrutura_saida": "Generated Output Structure",
        "docs_identicos_detectados": "Identical Documents Detected",
        "falhas_extracao": "Extraction Failures",
    }
}

CONFIG_PADRAO = {
    "idioma": "pt",
    "pastas_entrada": [],
    "pasta_saida": str(BASE_DIR / "saida_copilot"),
    # Estratégias de particionamento e saída Markdown / RAG:
    "gerar_arquivo_unico": True,
    "dividir_por_tamanho": True,
    "limite_mb_por_parte": 8.0,
    "dividir_por_quantidade": False,
    "qtd_docs_por_parte": 50,
    "gerar_individual_por_arquivo": False,
    "gerar_por_pasta_raiz": True,
    "gerar_jsonl": True,
    # Compactação e Timestamp:
    "usar_pasta_timestamp": True,
    "gerar_arquivo_zip": True,
    # Configurações de OCR:
    "usar_ocr_em_pdf": True,
    "usar_ocr_em_imagens_embutidas": True,
    "dpi_ocr_pdf": 200,
    "min_dimensao_imagem_ocr": 150,
    "limiar_confianca_ocr": 0.50,
    "proporcao_max_aspecto_ocr": 20.0,
    # Configurações de hardware / multicore:
    "multicore": True,
    "extensoes": EXTENSOES_PADRAO,
}

_OCR_ENGINE = None


def obter_idioma(cfg: dict) -> str:
    """Detecta idioma a partir de argumentos CLI (--lang en|pt) ou configuração."""
    if "--lang" in sys.argv:
        try:
            idx = sys.argv.index("--lang")
            if idx + 1 < len(sys.argv):
                l = sys.argv[idx + 1].lower()
                if l in ("en", "pt"):
                    return l
        except Exception:
            pass
    if "--english" in sys.argv or "--en" in sys.argv:
        return "en"
    if "--portugues" in sys.argv or "--pt" in sys.argv:
        return "pt"
    return cfg.get("idioma", "pt")


def t(chave: str, cfg: dict, **kwargs) -> str:
    """Recupera texto traduzido conforme o idioma ativo."""
    lang = obter_idioma(cfg)
    dicionario = I18N.get(lang, I18N["pt"])
    texto = dicionario.get(chave, I18N["pt"].get(chave, chave))
    if kwargs:
        try:
            return texto.format(**kwargs)
        except Exception:
            return texto
    return texto


def obter_ocr_worker():
    """Inicializa ou recupera o motor RapidOCR do processo worker atual."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _OCR_ENGINE = RapidOCR()
        except Exception:
            _OCR_ENGINE = False
    return _OCR_ENGINE if _OCR_ENGINE is not False else None


def limpar_tela():
    if "--auto" not in sys.argv:
        os.system("cls" if platform.system() == "Windows" else "clear")


def resolver_caminho(p_str: str) -> Path:
    p = Path(p_str).expanduser()
    if not p.is_absolute():
        p = (BASE_DIR / p).resolve()
    return p


def formatar_tamanho_bytes(tam_bytes: int) -> str:
    if tam_bytes < 1024:
        return f"{tam_bytes} B"
    elif tam_bytes < 1024 * 1024:
        return f"{tam_bytes / 1024:.1f} KB"
    elif tam_bytes < 1024 * 1024 * 1024:
        return f"{tam_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{tam_bytes / (1024 * 1024 * 1024):.2f} GB"


def obter_estatisticas_arquivos(arquivos: list[tuple[Path, Path]]) -> dict:
    total_arquivos = len(arquivos)
    tamanho_total = sum(arq.stat().st_size for _, arq in arquivos if arq.exists())
    por_extensao: dict[str, dict] = {}

    for _, arq in arquivos:
        ext = arq.suffix.lower() or ".sem_extensao"
        if ext not in por_extensao:
            por_extensao[ext] = {"qtd": 0, "bytes": 0}
        por_extensao[ext]["qtd"] += 1
        try:
            por_extensao[ext]["bytes"] += arq.stat().st_size
        except Exception:
            pass

    return {
        "total_arquivos": total_arquivos,
        "tamanho_total_bytes": tamanho_total,
        "tamanho_formatado": formatar_tamanho_bytes(tamanho_total),
        "por_extensao": por_extensao,
    }


def ler_texto_com_fallback(caminho: Path) -> str:
    """Lê arquivos de texto plano testando encodings comuns."""
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1"):
        try:
            return caminho.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return caminho.read_bytes().decode("utf-8", errors="replace")


def tabela_para_markdown(linhas: list[list[str]]) -> str:
    """Converte matriz de células em tabela Markdown bem formatada."""
    if not linhas or not any(linhas):
        return ""
    colunas_max = max(len(l) for l in linhas)
    if colunas_max == 0:
        return ""

    linhas_padronizadas = []
    for l in linhas:
        l_str = [str(c or "").replace("\r", " ").replace("\n", " ").strip() for c in l]
        if len(l_str) < colunas_max:
            l_str.extend([""] * (colunas_max - len(l_str)))
        linhas_padronizadas.append(l_str)

    if not any(any(c for c in l) for l in linhas_padronizadas):
        return ""

    cabecalho = linhas_padronizadas[0]
    separador = ["---"] * colunas_max
    saida = [
        "| " + " | ".join(cabecalho) + " |",
        "| " + " | ".join(separador) + " |",
    ]
    for row in linhas_padronizadas[1:]:
        saida.append("| " + " | ".join(row) + " |")
    return "\n".join(saida)


def normalizar_texto_ocr(texto: str) -> str:
    """
    Normaliza o texto extraído via OCR:
    - Recombina palavras quebradas por hifenização no final da linha (ex: desen-\nvolvimento -> desenvolvimento)
    - Converte quebras de linha duras simples em espaços, preservando quebras de parágrafo duplas
    - Não corrompe tabelas, listas numeradas ou citações em bloco
    """
    if not texto:
        return ""

    texto = re.sub(r'(\b\w+)-\s*\n\s*(\w+\b)', r'\1\2', texto)
    paragrafos = re.split(r'\n\s*\n', texto)
    limpos = []

    for p in paragrafos:
        linhas = [l.strip() for l in p.splitlines() if l.strip()]
        if not linhas:
            continue
        if any(re.match(r'^\s*([*+-]|\d+\.|>|\|)\s+', l) for l in linhas):
            limpos.append("\n".join(linhas))
        else:
            limpos.append(" ".join(linhas))

    return "\n\n".join(limpos)


def executar_ocr_em_bytes(
    img_bytes: bytes,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
) -> tuple[str, float]:
    """
    Executa OCR em bytes de imagem (PNG, JPEG, etc.), filtrando ruídos, proporções extremas (> 20:1)
    e detecções com baixa confiança (< 0.50). Retorna (texto_extraido, confianca_media).
    """
    if not img_bytes:
        return "", 0.0
    engine = obter_ocr_worker()
    if not engine:
        return "", 0.0
    try:
        from PIL import Image
        with Image.open(io.BytesIO(img_bytes)) as im:
            w, h = im.size
            if max(w, h) < min_dim or min(w, h) < 25:
                return "", 0.0

            ratio_aspecto = max(w, h) / max(min(w, h), 1)
            if ratio_aspecto > proporcao_max_aspecto:
                return "", 0.0

            max_edge = max(w, h)
            if max_edge > 2400:
                ratio = 2400 / max_edge
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                img_bytes = buf.getvalue()

        resultado, _ = engine(img_bytes)
        if not resultado:
            return "", 0.0

        linhas_validas = []
        scores = []
        for item in resultado:
            if len(item) >= 3:
                txt = str(item[1]).strip()
                try:
                    sc = float(item[2])
                except (ValueError, TypeError):
                    sc = 1.0
                if txt and sc >= limiar_confianca:
                    linhas_validas.append(txt)
                    scores.append(sc)
            elif len(item) >= 2:
                txt = str(item[1]).strip()
                if txt:
                    linhas_validas.append(txt)
                    scores.append(1.0)

        if not linhas_validas:
            return "", 0.0

        conf_media = (sum(scores) / len(scores)) if scores else 0.0
        texto_bruto = "\n".join(linhas_validas)
        texto_limpo = normalizar_texto_ocr(texto_bruto)
        return texto_limpo, conf_media
    except Exception:
        return "", 0.0


def extrair_imagens_zip_office(
    caminho: Path,
    prefixo: str,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    cache_dedup: dict | None = None,
) -> list[tuple[str, str, float]]:
    """Extrai imagens de pacotes OpenXML aplicando deduplicação por hash SHA-256 de bytes originais."""
    if cache_dedup is None:
        cache_dedup = {}
    resultados = []
    try:
        with zipfile.ZipFile(caminho, "r") as z:
            media_files = [n for n in z.namelist() if n.startswith(prefixo)]
            media_files.sort()
            for nome in media_files:
                ext = Path(nome).suffix.lower()
                if ext in {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tiff", ".tif"}:
                    dados = z.read(nome)
                    h_img = hashlib.sha256(dados).hexdigest()
                    if h_img in cache_dedup:
                        txt_ocr, conf = cache_dedup[h_img]
                    else:
                        txt_ocr, conf = executar_ocr_em_bytes(
                            dados,
                            min_dim=min_dim,
                            limiar_confianca=limiar_confianca,
                            proporcao_max_aspecto=proporcao_max_aspecto,
                        )
                        cache_dedup[h_img] = (txt_ocr, conf)

                    if txt_ocr:
                        nome_base = Path(nome).name
                        resultados.append((nome_base, txt_ocr, conf))
    except Exception:
        pass
    return resultados


def extrair_imagem_avulsa(
    caminho: Path,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float]:
    """Processa arquivos de imagem avulsos diretamente via OCR."""
    try:
        from PIL import Image
        dados = caminho.read_bytes()
        with Image.open(io.BytesIO(dados)) as im:
            w, h = im.size
            formato = im.format or caminho.suffix.replace(".", "").upper()
    except Exception as e:
        return f"*(Erro ao abrir imagem: {e})*", 0, 1, {}, [], 0.0

    dim_avulsa = min(min_dim, 40)
    texto, conf = executar_ocr_em_bytes(
        dados,
        min_dim=dim_avulsa,
        limiar_confianca=limiar_confianca,
        proporcao_max_aspecto=proporcao_max_aspecto,
    )
    meta_info = {"formato": formato, "resolucao": f"{w}x{h} px"}

    label_img = "Image" if lang == "en" else "Imagem"
    label_res = "Resolution" if lang == "en" else "Resolução"
    label_fmt = "Format" if lang == "en" else "Formato"
    label_ocr = "Text Extracted via Local OCR" if lang == "en" else "Texto Extraído via OCR Local"
    label_sem_txt = "*(Image without readable text detected)*" if lang == "en" else "*(Imagem sem texto legível detectado)*"

    if texto:
        bloco = (
            f"**[{label_img}: {caminho.name}]** ({label_res}: {w}x{h} px | {label_fmt}: {formato})\n\n"
            f"> **{label_ocr}:**\n"
            + "\n".join(f"> {linha}" for linha in texto.splitlines())
        )
        paginas_rag = [{"pagina": 1, "texto": texto}]
        return bloco, 1, 1, meta_info, paginas_rag, conf
    else:
        bloco = f"**[{label_img}: {caminho.name}]** ({label_res}: {w}x{h} px | {label_fmt}: {formato})\n{label_sem_txt}"
        return bloco, 0, 1, meta_info, [], 0.0


def extrair_pdf_hibrido(
    caminho_pdf: Path,
    usar_ocr: bool = True,
    dpi: int = 200,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float]:
    """Lê PDFs nativamente e aciona OCR local mantendo o fluxo textual contínuo."""
    cont_ocr = 0
    blocos_conteudo = []
    paginas_rag = []
    scores_ocr = []
    meta_info = {}
    cache_dedup_doc = {}

    try:
        import pymupdf
        pymupdf.no_recommend_layout()
        if hasattr(pymupdf, "TOOLS"):
            pymupdf.TOOLS.mupdf_display_warnings(False)
            pymupdf.TOOLS.mupdf_display_errors(False)
        doc = pymupdf.open(caminho_pdf)
        total_paginas = len(doc)

        if hasattr(doc, "metadata") and doc.metadata:
            for campo in ("title", "author", "creationDate", "modDate", "keywords"):
                val = doc.metadata.get(campo)
                if val and isinstance(val, str) and val.strip():
                    meta_info[campo] = val.strip()

        if doc.is_encrypted:
            doc.close()
            msg_p = "*(PDF document password protected - restricted extraction)*" if lang == "en" else "*(Documento PDF protegido por senha - extração restrita sem credenciais)*"
            return msg_p, 0, total_paginas, meta_info, [], 0.0

        for num_pag, pagina in enumerate(doc, start=1):
            texto_nativo = pagina.get_text("text").strip()
            tabelas_md = ""

            try:
                tabs = pagina.find_tables()
                if tabs and tabs.tables:
                    tabelas_limpas = []
                    for t_item in tabs.tables:
                        md_t = t_item.to_markdown()
                        md_t = re.sub(r'<br\s*/?>', ' ', md_t)
                        tabelas_limpas.append(md_t)
                    if tabelas_limpas:
                        tabelas_md = "\n\n" + "\n\n".join(tabelas_limpas)
            except Exception:
                pass

            imagens_pagina = pagina.get_images()
            texto_ocr_extra = []

            area_pagina = abs(pagina.rect.width * pagina.rect.height) if hasattr(pagina, "rect") else 1.0
            cobertura_grande = False
            if imagens_pagina:
                for img_info in imagens_pagina:
                    try:
                        xref = img_info[0]
                        for r in pagina.get_image_rects(xref):
                            if (r.width * r.height) / max(area_pagina, 1.0) >= 0.60:
                                cobertura_grande = True
                                break
                    except Exception:
                        pass

            eh_escaneada = usar_ocr and (
                (len(texto_nativo.strip()) < 80 and len(imagens_pagina) > 0)
                or cobertura_grande
            )

            # 1. Caso A: Página escaneada
            if eh_escaneada:
                engine = obter_ocr_worker()
                if engine:
                    try:
                        pix = pagina.get_pixmap(dpi=dpi)
                        img_bytes = pix.tobytes("png")
                        txt_ocr, conf_p = executar_ocr_em_bytes(
                            img_bytes,
                            min_dim=min_dim,
                            limiar_confianca=limiar_confianca,
                            proporcao_max_aspecto=proporcao_max_aspecto,
                        )
                        if txt_ocr:
                            tag_p = f"*(Page {num_pag} scanned - extracted via local OCR)*" if lang == "en" else f"*(Página {num_pag} escaneada - extraída via OCR local)*"
                            texto_pagina = f"{tag_p}\n{txt_ocr}"
                            cont_ocr += 1
                            scores_ocr.append(conf_p)
                        else:
                            texto_pagina = f"*(Page {num_pag} scanned without readable text)*" if lang == "en" else f"*(Página {num_pag} escaneada sem texto legível)*"
                    except Exception:
                        texto_pagina = texto_nativo or (f"*(Page {num_pag} scanned)*" if lang == "en" else f"*(Página {num_pag} escaneada)*")
                else:
                    texto_pagina = texto_nativo or (f"*(Page {num_pag} scanned)*" if lang == "en" else f"*(Página {num_pag} escaneada)*")

            # 2. Caso B: Página mista (texto nativo + imagens embutidas)
            else:
                texto_pagina = (texto_nativo + tabelas_md).strip()
                if usar_ocr and len(imagens_pagina) > 0:
                    for img_idx, img_info in enumerate(imagens_pagina, start=1):
                        xref = img_info[0]
                        try:
                            img_dict = doc.extract_image(xref)
                            if img_dict:
                                w = img_dict.get("width", 0)
                                h = img_dict.get("height", 0)
                                img_bytes = img_dict.get("image", b"")
                                if img_bytes:
                                    h_img = hashlib.sha256(img_bytes).hexdigest()
                                    if h_img in cache_dedup_doc:
                                        txt_img, conf_img = cache_dedup_doc[h_img]
                                    else:
                                        txt_img, conf_img = executar_ocr_em_bytes(
                                            img_bytes,
                                            min_dim=min_dim,
                                            limiar_confianca=limiar_confianca,
                                            proporcao_max_aspecto=proporcao_max_aspecto,
                                        )
                                        cache_dedup_doc[h_img] = (txt_img, conf_img)

                                    if txt_img:
                                        lbl_emb = f"OCR Embedded Image #{img_idx}" if lang == "en" else f"OCR Imagem Embutida #{img_idx}"
                                        texto_ocr_extra.append(
                                            f"> **[{lbl_emb} ({w}x{h} px)]:**\n"
                                            + "\n".join(f"> {l}" for l in txt_img.splitlines())
                                        )
                                        cont_ocr += 1
                                        scores_ocr.append(conf_img)
                        except Exception:
                            continue

            bloco_pagina = texto_pagina
            if texto_ocr_extra:
                bloco_pagina += "\n\n" + "\n\n".join(texto_ocr_extra)

            bloco_pagina = bloco_pagina.strip()
            if bloco_pagina:
                blocos_conteudo.append(bloco_pagina)
                paginas_rag.append({"pagina": num_pag, "texto": bloco_pagina})

        doc.close()
        conf_media = (sum(scores_ocr) / len(scores_ocr)) if scores_ocr else 0.0
        return "\n\n".join(blocos_conteudo), cont_ocr, total_paginas, meta_info, paginas_rag, conf_media
    except ImportError:
        pass

    # Fallback com pypdf
    from pypdf import PdfReader
    reader = PdfReader(str(caminho_pdf))
    total_paginas = len(reader.pages)
    for num_pag, pagina in enumerate(reader.pages, start=1):
        txt = (pagina.extract_text() or "").strip()
        if txt:
            blocos_conteudo.append(txt)
            paginas_rag.append({"pagina": num_pag, "texto": txt})
    return "\n\n".join(blocos_conteudo), cont_ocr, total_paginas, meta_info, paginas_rag, 0.0


def extrair_docx_nativo(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float]:
    """Extrai parágrafos, tabelas e imagens com OCR em documentos Word."""
    import docx
    doc = docx.Document(str(caminho))
    partes = []
    meta_info = {}

    if hasattr(doc, "core_properties"):
        cp = doc.core_properties
        if cp.title:
            meta_info["title"] = cp.title
        if cp.author:
            meta_info["author"] = cp.author
        if cp.created:
            meta_info["creationDate"] = cp.created.strftime("%Y-%m-%d")

    for p in doc.paragraphs:
        t_p = p.text.strip()
        if not t_p:
            continue
        estilo = (p.style.name or "").lower() if p.style else ""
        if "heading 1" in estilo or "título 1" in estilo:
            partes.append(f"### {t_p}")
        elif "heading 2" in estilo or "título 2" in estilo:
            partes.append(f"#### {t_p}")
        else:
            partes.append(t_p)

    lbl_tab = "Table" if lang == "en" else "Tabela"
    for idx_t, tab in enumerate(doc.tables, start=1):
        linhas = [[cel.text.strip() for cel in row.cells] for row in tab.rows]
        md_tab = tabela_para_markdown(linhas)
        if md_tab:
            partes.append(f"\n**{lbl_tab} {idx_t}:**\n{md_tab}")

    cont_ocr = 0
    scores = []
    cache_dedup = {}
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(
            caminho,
            "word/media/",
            min_dim=min_dim,
            limiar_confianca=limiar_confianca,
            proporcao_max_aspecto=proporcao_max_aspecto,
            cache_dedup=cache_dedup,
        )
        if imagens_ocr:
            lbl_bloco = "\n### Embedded Images in Document (OCR)" if lang == "en" else "\n### Imagens Embutidas no Documento (OCR)"
            lbl_tag = "OCR Image" if lang == "en" else "OCR Imagem"
            blocos_img = [lbl_bloco]
            for nome_img, txt, conf in imagens_ocr:
                cont_ocr += 1
                scores.append(conf)
                blocos_img.append(
                    f"> **[{lbl_tag}: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    paginas_rag = [{"pagina": 1, "texto": texto_final}]
    return texto_final, cont_ocr, 1, meta_info, paginas_rag, conf_media


def extrair_xlsx_nativo(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float]:
    """Converte planilhas Excel em tabelas Markdown e extrai imagens com OCR."""
    import openpyxl
    wb = openpyxl.load_workbook(str(caminho), data_only=True, read_only=True)
    partes = []
    meta_info = {}
    paginas_rag = []
    lbl_plan = "Sheet" if lang == "en" else "Planilha"

    for idx_aba, nome_aba in enumerate(wb.sheetnames, start=1):
        ws = wb[nome_aba]
        linhas = []
        for row in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() != "" for c in row):
                linhas.append(["" if c is None else str(c) for c in row])
        md_tab = tabela_para_markdown(linhas)
        if md_tab:
            texto_aba = f"### {lbl_plan}: {nome_aba}\n\n{md_tab}"
            partes.append(texto_aba)
            paginas_rag.append({"pagina": idx_aba, "texto": texto_aba})
    wb.close()

    cont_ocr = 0
    scores = []
    cache_dedup = {}
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(
            caminho,
            "xl/media/",
            min_dim=min_dim,
            limiar_confianca=limiar_confianca,
            proporcao_max_aspecto=proporcao_max_aspecto,
            cache_dedup=cache_dedup,
        )
        if imagens_ocr:
            lbl_bloco = "\n### Embedded Images / Receipts in Spreadsheet (OCR)" if lang == "en" else "\n### Imagens / Comprovantes Embutidos na Planilha (OCR)"
            lbl_tag = "OCR Spreadsheet Image" if lang == "en" else "OCR Imagem Planilha"
            blocos_img = [lbl_bloco]
            for nome_img, txt, conf in imagens_ocr:
                cont_ocr += 1
                scores.append(conf)
                blocos_img.append(
                    f"> **[{lbl_tag}: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    return texto_final, cont_ocr, len(paginas_rag) or 1, meta_info, paginas_rag, conf_media


def extrair_pptx_nativo(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float]:
    """Extrai slides, anotações, tabelas e imagens com OCR de apresentações PowerPoint."""
    from pptx import Presentation
    prs = Presentation(str(caminho))
    partes = []
    meta_info = {}
    paginas_rag = []
    lbl_notas = "Presenter Notes" if lang == "en" else "Notas do Apresentador"

    for idx_s, slide in enumerate(prs.slides, start=1):
        textos_slide = [f"### Slide {idx_s}"]
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for p in shape.text_frame.paragraphs:
                    t_txt = p.text.strip()
                    if t_txt:
                        textos_slide.append(f"- {t_txt}")
            if getattr(shape, "has_table", False):
                linhas = [[cel.text.strip() for cel in row.cells] for row in shape.table.rows]
                md_t = tabela_para_markdown(linhas)
                if md_t:
                    textos_slide.append(f"\n{md_t}")
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                textos_slide.append(f"\n> **{lbl_notas}:** {notas}")

        bloco_slide = "\n".join(textos_slide)
        partes.append(bloco_slide)
        paginas_rag.append({"pagina": idx_s, "texto": bloco_slide})

    cont_ocr = 0
    scores = []
    cache_dedup = {}
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(
            caminho,
            "ppt/media/",
            min_dim=min_dim,
            limiar_confianca=limiar_confianca,
            proporcao_max_aspecto=proporcao_max_aspecto,
            cache_dedup=cache_dedup,
        )
        if imagens_ocr:
            lbl_bloco = "\n### Embedded Images / Diagrams in Slides (OCR)" if lang == "en" else "\n### Imagens / Diagramas Embutidos nos Slides (OCR)"
            lbl_tag = "OCR Slide Image" if lang == "en" else "OCR Imagem Slide"
            blocos_img = [lbl_bloco]
            for nome_img, txt, conf in imagens_ocr:
                cont_ocr += 1
                scores.append(conf)
                blocos_img.append(
                    f"> **[{lbl_tag}: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    return texto_final, cont_ocr, len(prs.slides), meta_info, paginas_rag, conf_media


def calcular_hash_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with caminho.open("rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()[:20]


def worker_processar_arquivo(args: tuple) -> tuple[str, str, str, dict, str | None]:
    """Função concorrente executada nos núcleos da CPU (ProcessPoolExecutor)."""
    raiz_str, arq_str, pasta_cache_str, cfg = args
    raiz = Path(raiz_str)
    arq = Path(arq_str)
    pasta_cache = Path(pasta_cache_str)
    lang = cfg.get("idioma", "pt")

    try:
        caminho_rel = f"{raiz.name}/{arq.relative_to(raiz)}".replace("\\", "/")
    except Exception:
        caminho_rel = arq.name

    try:
        pasta_cache.mkdir(parents=True, exist_ok=True)
        file_hash = calcular_hash_arquivo(arq)

        min_dim = int(cfg.get("min_dimensao_imagem_ocr", 150))
        limiar_conf = float(cfg.get("limiar_confianca_ocr", 0.50))
        dpi_val = int(cfg.get("dpi_ocr_pdf", 200))
        usar_pdf = int(cfg.get("usar_ocr_em_pdf", True))
        usar_img = int(cfg.get("usar_ocr_em_imagens_embutidas", True))
        prop_max = float(cfg.get("proporcao_max_aspecto_ocr", 20.0))

        flag_ocr = f"ocr_{usar_pdf}_{usar_img}_{dpi_val}_{min_dim}_{int(limiar_conf * 100)}"
        arq_cache = pasta_cache / f"{file_hash}_{flag_ocr}.json"

        if arq_cache.exists():
            try:
                dados = json.loads(arq_cache.read_text(encoding="utf-8"))
                return arq.name, caminho_rel, raiz.name, {
                    "cache": True,
                    "ocr_pages": dados.get("ocr_pages", 0),
                    "total_paginas": dados.get("total_paginas", 1),
                    "meta_info": dados.get("meta_info", {}),
                    "confianca_ocr": dados.get("confianca_ocr", 0.0),
                    "paginas_rag": dados.get("paginas_rag", []),
                    "hash": file_hash,
                    "conteudo": dados["conteudo"],
                }, None
            except Exception:
                pass

        ext = arq.suffix.lower()
        ocr_count = 0
        total_paginas = 1
        meta_info = {}
        paginas_rag = []
        confianca_ocr = 0.0
        conteudo = ""

        if ext == ".pdf":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr = extrair_pdf_hibrido(
                arq,
                usar_ocr=cfg.get("usar_ocr_em_pdf", True),
                dpi=dpi_val,
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext in EXTENSOES_IMAGEM:
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr = extrair_imagem_avulsa(
                arq,
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".docx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr = extrair_docx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".xlsx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr = extrair_xlsx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".pptx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr = extrair_pptx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext in EXTENSOES_TEXTO:
            conteudo = ler_texto_com_fallback(arq)
            paginas_rag = [{"pagina": 1, "texto": conteudo}]
        elif ext == ".xls":
            try:
                import pandas as pd
                xl = pd.ExcelFile(str(arq))
                blocos = []
                lbl_pl = "Sheet" if lang == "en" else "Planilha"
                for idx_s, nome in enumerate(xl.sheet_names, start=1):
                    df = xl.parse(nome)
                    linhas = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
                    txt_aba = f"### {lbl_pl}: {nome}\n\n{tabela_para_markdown(linhas)}"
                    blocos.append(txt_aba)
                    paginas_rag.append({"pagina": idx_s, "texto": txt_aba})
                conteudo = "\n\n".join(blocos)
                total_paginas = len(xl.sheet_names)
            except Exception:
                conteudo = ler_texto_com_fallback(arq)
                paginas_rag = [{"pagina": 1, "texto": conteudo}]

        conteudo = conteudo.strip()
        arq_cache.write_text(
            json.dumps(
                {
                    "arquivo": arq.name,
                    "conteudo": conteudo,
                    "ocr_pages": ocr_count,
                    "total_paginas": total_paginas,
                    "meta_info": meta_info,
                    "confianca_ocr": confianca_ocr,
                    "paginas_rag": paginas_rag,
                    "data_cache": datetime.now().isoformat(),
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return arq.name, caminho_rel, raiz.name, {
            "cache": False,
            "ocr_pages": ocr_count,
            "total_paginas": total_paginas,
            "meta_info": meta_info,
            "confianca_ocr": confianca_ocr,
            "paginas_rag": paginas_rag,
            "hash": file_hash,
            "conteudo": conteudo,
        }, None
    except Exception as e:
        return arq.name, caminho_rel, raiz.name, {}, str(e)


def listar_arquivos_elegiveis(pastas: list[str], extensoes: list[str]) -> list[tuple[Path, Path]]:
    exts = {e.lower() for e in extensoes}
    arquivos = []
    vistos_abs = set()

    for p_str in pastas:
        p_raiz = resolver_caminho(p_str)
        if not p_raiz.exists() or not p_raiz.is_dir():
            continue
        try:
            for item in p_raiz.rglob("*"):
                if item.is_file() and item.suffix.lower() in exts:
                    if item.name.startswith("~$") or item.name.startswith("."):
                        continue
                    abs_path = item.resolve()
                    if abs_path not in vistos_abs:
                        vistos_abs.add(abs_path)
                        arquivos.append((p_raiz, item))
        except Exception:
            continue

    arquivos.sort(key=lambda x: str(x[1]).lower())
    return arquivos


def compactar_pasta(origem: Path, destino_zip: Path) -> bool:
    """Gera pacote compactado .ZIP e valida sua integridade com testzip()."""
    try:
        destino_zip.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(destino_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for raiz, _, arqs in os.walk(origem):
                for a in arqs:
                    caminho_arq = Path(raiz) / a
                    if caminho_arq.suffix.lower() == ".zip":
                        continue
                    arcname = caminho_arq.relative_to(origem)
                    z.write(caminho_arq, arcname)

        with zipfile.ZipFile(destino_zip, "r") as z_check:
            falha = z_check.testzip()
            if falha is not None:
                print(f"{Cor.VERMELHO}[AVISO / WARNING] Zip integrity warning on file: {falha}{Cor.RESET}")
                return False
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}[ERRO / ERROR] Zip creation failed: {e}{Cor.RESET}")
        return False


def sanitizar_nome_arquivo(nome: str) -> str:
    invalido = '<>:"/\\|?*'
    limpo = "".join("_" if c in invalido else c for c in nome)
    return limpo.strip() or "documento"


def executar_pipeline(cfg: dict, pausar_no_final: bool = True):
    limpar_tela()
    lang = obter_idioma(cfg)
    cfg["idioma"] = lang

    print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
    print(f"{Cor.TITULO}   {t('titulo_app', cfg)}{Cor.RESET}")
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")

    pastas = cfg.get("pastas_entrada", [])
    if not pastas:
        print(f"{Cor.VERMELHO}{t('sem_pastas', cfg)}{Cor.RESET}")
        if pausar_no_final:
            input(f"\n{t('voltar_menu', cfg)}")
        return

    arquivos = listar_arquivos_elegiveis(pastas, cfg.get("extensoes", EXTENSOES_PADRAO))
    total = len(arquivos)

    if total == 0:
        print(f"{Cor.VERMELHO}{t('sem_arquivos', cfg)}{Cor.RESET}")
        if pausar_no_final:
            input(f"\n{t('voltar_menu', cfg)}")
        return

    stats = obter_estatisticas_arquivos(arquivos)
    pasta_saida_base = resolver_caminho(cfg["pasta_saida"])
    pasta_cache = pasta_saida_base / ".cache_extracao"
    pasta_saida_base.mkdir(parents=True, exist_ok=True)

    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    if cfg.get("usar_pasta_timestamp", True):
        pasta_destino = pasta_saida_base / f"processado_{timestamp_str}"
    else:
        pasta_destino = pasta_saida_base
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # TELA DE CONFIRMAÇÃO PRÉVIA (interativo)
    if "--auto" not in sys.argv:
        while True:
            limpar_tela()
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
            titulo_conf = "CONFIRMATION OF LOCAL PROCESSING" if lang == "en" else "CONFIRMAÇÃO DO PROCESSAMENTO LOCAL"
            print(f"{Cor.TITULO}   {titulo_conf}{Cor.RESET}")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")
            
            lbl_resumo = "Corpus Summary to Process:" if lang == "en" else "Resumo do Corpus a Processar:"
            print(f" {Cor.NEGRITO}{lbl_resumo}{Cor.RESET}")
            print(f"  • {t('pastas_monitoradas', cfg):<25}: {len(pastas)}")
            print(f"  • {t('arquivos_localizados', cfg):<25}: {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET}")
            print(f"  • {t('volume_proc', cfg):<25}: {Cor.AMARELO}{Cor.NEGRITO}{stats['tamanho_formatado']}{Cor.RESET}")
            print(f"  • Format breakdown:")
            for ext, info in sorted(stats["por_extensao"].items(), key=lambda x: x[1]["qtd"], reverse=True):
                print(f"     - {ext.upper():<6} : {info['qtd']:>3} files ({formatar_tamanho_bytes(info['bytes'])})")

            lbl_cfg = "Current Execution Settings:" if lang == "en" else "Configurações desta Execução:"
            print(f"\n {Cor.NEGRITO}{lbl_cfg}{Cor.RESET}")
            print(f"  • Destination folder      : {pasta_destino}")
            print(f"  • Split by MB limit       : {'YES (' + str(cfg.get('limite_mb_por_parte', 8.0)) + ' MB)' if cfg.get('dividir_por_tamanho', True) else 'NO'}")
            print(f"  • Single file output      : {'YES (corpus_completo.md)' if cfg.get('gerar_arquivo_unico', True) else 'NO'}")
            print(f"  • RAG dataset (.JSONL)    : {'YES (corpus_rag.jsonl)' if cfg.get('gerar_jsonl', True) else 'NO'}")
            print(f"  • Compress to .ZIP        : {'YES' if cfg.get('gerar_arquivo_zip', True) else 'NO'}")
            print(f"  • OCR in PDF and Images   : {'ENABLED (DPI ' + str(cfg.get('dpi_ocr_pdf', 200)) + ' | Min ' + str(cfg.get('min_dimensao_imagem_ocr', 150)) + 'px)' if cfg.get('usar_ocr_em_pdf', True) else 'DISABLED'}")
            print(f"{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
            if lang == "en":
                print(" [ENTER] or [C] Confirm and Start Processing")
                print(" [M] Change MB splitting limit (e.g., 8, 50, 100, 200 MB)")
                print(" [Z] Toggle .ZIP archive generation")
                print(" [V] Cancel and return to main menu")
            else:
                print(" [ENTER] ou [C] Confirmar e Iniciar o Processamento")
                print(" [M] Alterar limite de separação em MB (ex: 8, 50, 100, 200 MB)")
                print(" [Z] Alternar geração de arquivo compactado .ZIP")
                print(" [V] Cancelar e voltar ao menu")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

            escolha_conf = input("\nAction [ENTER to confirm]: ").strip().upper()
            if escolha_conf in ("", "C", "S", "SIM", "Y", "YES"):
                break
            elif escolha_conf in ("V", "0"):
                return
            elif escolha_conf == "M":
                novo_mb = input("\nLimit in MB (e.g., 8, 50, 200): ").strip().replace(",", ".")
                try:
                    v_mb = float(novo_mb)
                    if v_mb >= 0.5:
                        cfg["limite_mb_por_parte"] = v_mb
                        cfg["dividir_por_tamanho"] = True
                        salvar_config(cfg)
                except ValueError:
                    pass
            elif escolha_conf == "Z":
                cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", True)
                salvar_config(cfg)

    num_cpus = os.cpu_count() or 4
    usar_multicore = cfg.get("multicore", True)
    max_workers = num_cpus if usar_multicore else 1

    inicio = time.time()
    limpar_tela()
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
    print(f"{Cor.TITULO}  {t('iniciando_conversao', cfg)}{Cor.RESET}")
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")
    print(f"{t('pastas_monitoradas', cfg):<22}: {len(pastas)}")
    print(f"{t('arquivos_localizados', cfg):<22}: {Cor.NEGRITO}{total}{Cor.RESET} ({stats['tamanho_formatado']})")
    print(f"{t('multicore_paralelo', cfg):<22}: {Cor.VERDE}{t('ativado', cfg)} ({max_workers} cores){Cor.RESET}")
    print(f"{t('pasta_execucao', cfg):<22}: {pasta_destino}\n")

    tarefas = [
        (str(raiz), str(arq), str(pasta_cache), cfg)
        for raiz, arq in arquivos
    ]

    resultados_processados = []
    erros = []

    print(f"{t('processando_paralelo', cfg)}")
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futuros = {executor.submit(worker_processar_arquivo, tarefa): tarefa for tarefa in tarefas}
        docs_concluidos = 0
        for futuro in as_completed(futuros):
            docs_concluidos += 1
            nome_arq, caminho_rel, raiz_nome, meta_res, erro_str = futuro.result()

            tempo_passado = time.time() - inicio
            taxa_docs_seg = docs_concluidos / max(tempo_passado, 0.001)
            docs_restantes = total - docs_concluidos
            seg_restantes = int(docs_restantes / max(taxa_docs_seg, 0.001))
            eta_str = f"{seg_restantes // 60:02d}:{seg_restantes % 60:02d}"

            if erro_str:
                err_label = "[ERROR]" if lang == "en" else "[ERRO]"
                print(f"({docs_concluidos:03d}/{total:03d}) {Cor.VERMELHO}{err_label}{Cor.RESET} {caminho_rel} -> {erro_str}")
                erros.append(f"- **[{docs_concluidos:03d}]** `{caminho_rel}` -> `{erro_str}`")
            else:
                resultados_processados.append((docs_concluidos, nome_arq, caminho_rel, raiz_nome, meta_res))
                status_str = f"{Cor.AZUL}[CACHE]{Cor.RESET}" if meta_res["cache"] else f"{Cor.VERDE}[OK]{Cor.RESET}"
                if meta_res.get("ocr_pages", 0) > 0:
                    status_str += f" {Cor.AMARELO}(+{meta_res['ocr_pages']} OCR){Cor.RESET}"
                print(f"({docs_concluidos:03d}/{total:03d}) {status_str} {caminho_rel} | ETA: {eta_str}")

    resultados_processados.sort(key=lambda x: str(x[2]).lower())

    indice_geral = [
        t("base_copilot_titulo", cfg),
        f"- **{t('data_consolida', cfg)}:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"- **{t('total_docs_index', cfg)}:** {len(resultados_processados)}",
        f"- **{t('vol_original', cfg)}:** {stats['tamanho_formatado']}",
        f"- **{t('pastas_origem', cfg)}:** {len(pastas)}",
        f"\n## {t('indice_mestre', cfg)}\n",
    ]

    blocos_gerais = []
    blocos_por_pasta: dict[str, list[str]] = {}
    docs_individuais = []
    hashes_vistos: dict[str, str] = {}
    duplicados = []
    cont_cache = 0
    cont_ocr_itens = 0

    for novo_idx, (_, nome_arq, caminho_rel, raiz_nome, meta) in enumerate(resultados_processados, start=1):
        doc_id = f"[{novo_idx:03d}]"
        if meta.get("cache"):
            cont_cache += 1
        if meta.get("ocr_pages", 0) > 0:
            cont_ocr_itens += meta["ocr_pages"]

        f_hash = meta.get("hash", "")
        aviso_dup = ""
        if f_hash and f_hash in hashes_vistos:
            orig_id = hashes_vistos[f_hash]
            msg_d = f"has identical content to {orig_id}" if lang == "en" else f"possui conteúdo idêntico a {orig_id}"
            duplicados.append(f"- `{caminho_rel}` ({doc_id}) {msg_d}")
            aviso_dup = f" *(Identical to {orig_id})*" if lang == "en" else f" *(Idêntico a {orig_id})*"
        elif f_hash:
            hashes_vistos[f_hash] = f"{doc_id} (`{caminho_rel}`)"

        tag_ocr = f" *(OCR: {meta['ocr_pages']})*" if meta.get("ocr_pages", 0) > 0 else ""
        indice_geral.append(f"- **{doc_id}** `{caminho_rel}`{tag_ocr}{aviso_dup}")

        meta_linhas = []
        meta_info = meta.get("meta_info", {})
        if meta_info:
            detalhes_meta = [f"{k.capitalize()}: {v}" for k, v in meta_info.items()]
            meta_linhas.append(f"- **{t('metadados', cfg)}:** {' | '.join(detalhes_meta)}")

        total_pags = meta.get("total_paginas", 1)
        pags_str = f"{total_pags} page{'s' if total_pags > 1 else ''}" if lang == "en" else f"{total_pags} página{'s' if total_pags > 1 else ''}"
        ocr_str = f"{meta.get('ocr_pages', 0)} element(s) processed via OCR" if lang == "en" else f"{meta.get('ocr_pages', 0)} elemento(s) processado(s) via OCR"
        meta_linhas.append(f"- **{t('dim_doc', cfg)}:** {pags_str} | {ocr_str}")

        conf_ocr = meta.get("confianca_ocr", 0.0)
        if meta.get("ocr_pages", 0) > 0 and conf_ocr > 0:
            meta_linhas.append(f"- **{t('qualidade_ocr', cfg)}:** {t('conf_media', cfg)}: {conf_ocr * 100:.1f}%")

        bloco_meta_formatado = "\n".join(meta_linhas)
        if bloco_meta_formatado:
            bloco_meta_formatado = "\n" + bloco_meta_formatado

        bloco_md = (
            f"\n\n{'=' * 80}\n"
            f"## {doc_id} {nome_arq}\n"
            f"- **{t('id_ref', cfg)}:** {doc_id}\n"
            f"- **{t('pasta_dep', cfg)}:** `{raiz_nome}`\n"
            f"- **{t('cam_rel', cfg)}:** `{caminho_rel}`"
            f"{bloco_meta_formatado}\n"
            f"{'=' * 80}\n\n"
            f"{meta['conteudo'] if meta.get('conteudo') else t('sem_texto', cfg)}\n"
        )
        blocos_gerais.append(bloco_md)
        blocos_por_pasta.setdefault(raiz_nome, []).append(bloco_md)
        docs_individuais.append((doc_id, nome_arq, caminho_rel, bloco_md))

    cabecalho_str = "\n".join(indice_geral) + "\n"
    arquivos_gerados = []

    # 1. Estratégia A: Arquivo Único Consolidado
    if cfg.get("gerar_arquivo_unico", True):
        arq_unico = pasta_destino / "corpus_completo.md"
        with arq_unico.open("w", encoding="utf-8") as f:
            f.write(cabecalho_str)
            for b in blocos_gerais:
                f.write(b)
        tam_mb = arq_unico.stat().st_size / (1024 * 1024)
        arquivos_gerados.append(f"corpus_completo.md ({tam_mb:.2f} MB)")

    # 2. Estratégia B: Divisão por Tamanho em MB
    if cfg.get("dividir_por_tamanho", True) and blocos_gerais:
        limite_mb = float(cfg.get("limite_mb_por_parte", 8.0))
        limite_bytes = int(limite_mb * 1024 * 1024)

        for antigo in pasta_destino.glob("corpus_parte_*.md"):
            antigo.unlink(missing_ok=True)

        parte_num = 1
        buffer_atual = [cabecalho_str]
        bytes_atuais = len(cabecalho_str.encode("utf-8"))

        for b in blocos_gerais:
            tam_b = len(b.encode("utf-8"))
            if bytes_atuais + tam_b > limite_bytes and len(buffer_atual) > 1:
                nome_p = f"corpus_parte_{parte_num:02d}.md"
                caminho_p = pasta_destino / nome_p
                caminho_p.write_text("".join(buffer_atual), encoding="utf-8")
                arquivos_gerados.append(f"{nome_p} ({caminho_p.stat().st_size / (1024*1024):.2f} MB)")
                parte_num += 1
                cab_cont = (
                    f"# BASE DOCUMENTAL - PARTE {parte_num:02d}\n"
                    f"*(See Master Index in Part 01 / Consulte o Índice Mestre na Parte 01)*\n\n"
                )
                buffer_atual = [cab_cont]
                bytes_atuais = len(cab_cont.encode("utf-8"))

            buffer_atual.append(b)
            bytes_atuais += tam_b

        if len(buffer_atual) > 1:
            nome_p = f"corpus_parte_{parte_num:02d}.md"
            caminho_p = pasta_destino / nome_p
            caminho_p.write_text("".join(buffer_atual), encoding="utf-8")
            arquivos_gerados.append(f"{nome_p} ({caminho_p.stat().st_size / (1024*1024):.2f} MB)")

    # 3. Estratégia C: Divisão por Quantidade de Documentos
    if cfg.get("dividir_por_quantidade", False) and blocos_gerais:
        qtd_por_lote = max(1, int(cfg.get("qtd_docs_por_parte", 50)))
        dir_lotes = pasta_destino / "por_lotes"
        dir_lotes.mkdir(parents=True, exist_ok=True)

        for antigo in dir_lotes.glob("lote_*.md"):
            antigo.unlink(missing_ok=True)

        for lote_idx, i in enumerate(range(0, len(blocos_gerais), qtd_por_lote), start=1):
            fatia = blocos_gerais[i:i + qtd_por_lote]
            nome_lote = f"lote_{lote_idx:02d}_docs_{i+1:03d}_a_{min(i+qtd_por_lote, len(blocos_gerais)):03d}.md"
            caminho_lote = dir_lotes / nome_lote
            topo_lote = (
                f"# BATCH / LOTE {lote_idx:02d} ({len(fatia)} docs)\n"
                f"- **IDs:** [{i+1:03d}] .. [{min(i+qtd_por_lote, len(blocos_gerais)):03d}]\n\n"
            )
            caminho_lote.write_text(topo_lote + "".join(fatia), encoding="utf-8")
        arquivos_gerados.append(f"por_lotes/*.md ({len(range(0, len(blocos_gerais), qtd_por_lote))} batches)")

    # 4. Estratégia D: 1 Arquivo Markdown por Arquivo Original (Individual 1:1)
    if cfg.get("gerar_individual_por_arquivo", False) and docs_individuais:
        dir_indiv = pasta_destino / "individual"
        dir_indiv.mkdir(parents=True, exist_ok=True)
        for antigo in dir_indiv.glob("*.md"):
            antigo.unlink(missing_ok=True)
        for doc_id, nome_arq, caminho_rel, bloco_md in docs_individuais:
            nome_limpo = sanitizar_nome_arquivo(f"{doc_id}_{nome_arq}.md")
            (dir_indiv / nome_limpo).write_text(bloco_md.strip(), encoding="utf-8")
        arquivos_gerados.append(f"individual/*.md ({len(docs_individuais)} individual 1:1 docs)")

    # 5. Estratégia E: Consolidado por Pasta Raiz / Departamento
    if cfg.get("gerar_por_pasta_raiz", True) and blocos_por_pasta:
        dir_por_pasta = pasta_destino / "por_pasta"
        dir_por_pasta.mkdir(parents=True, exist_ok=True)
        for antigo in dir_por_pasta.glob("*.md"):
            antigo.unlink(missing_ok=True)
        for idx_p, (nome_pasta, lista_blocos) in enumerate(blocos_por_pasta.items(), start=1):
            nome_limpo = sanitizar_nome_arquivo(nome_pasta)
            arq_p = dir_por_pasta / f"{idx_p:02d}_{nome_limpo}.md"
            topo = (
                f"# CONSOLIDATED / CONSOLIDADO: {nome_pasta}\n"
                f"- **Total indexed documents:** {len(lista_blocos)}\n\n"
            )
            arq_p.write_text(topo + "".join(lista_blocos), encoding="utf-8")
        arquivos_gerados.append(f"por_pasta/*.md ({len(blocos_por_pasta)} folders)")

    # 6. Estratégia F: Base Estruturada JSONL para RAG
    if cfg.get("gerar_jsonl", True):
        arq_jsonl = pasta_destino / "corpus_rag.jsonl"
        with arq_jsonl.open("w", encoding="utf-8") as f_jsonl:
            for novo_idx, (_, nome_arq, caminho_rel, raiz_nome, meta) in enumerate(resultados_processados, start=1):
                doc_id = f"[{novo_idx:03d}]"
                pags = meta.get("paginas_rag", [])
                if pags:
                    for p in pags:
                        registro = {
                            "id": doc_id,
                            "arquivo": nome_arq,
                            "caminho_relativo": caminho_rel,
                            "departamento": raiz_nome,
                            "pagina": p.get("pagina", 1),
                            "total_paginas": meta.get("total_paginas", 1),
                            "metadados": meta.get("meta_info", {}),
                            "texto": p.get("texto", ""),
                        }
                        f_jsonl.write(json.dumps(registro, ensure_ascii=False) + "\n")
                else:
                    registro = {
                        "id": doc_id,
                        "arquivo": nome_arq,
                        "caminho_relativo": caminho_rel,
                        "departamento": raiz_nome,
                        "pagina": 1,
                        "total_paginas": meta.get("total_paginas", 1),
                        "metadados": meta.get("meta_info", {}),
                        "texto": meta.get("conteudo", ""),
                    }
                    f_jsonl.write(json.dumps(registro, ensure_ascii=False) + "\n")
        tam_jsonl_mb = arq_jsonl.stat().st_size / (1024 * 1024)
        arquivos_gerados.append(f"corpus_rag.jsonl ({tam_jsonl_mb:.2f} MB)")

    duracao = time.time() - inicio
    relatorio = [
        t("relatorio_auditoria_titulo", cfg),
        f"- **Date/Time / Data:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"- **{t('tempo_total', cfg)}:** {duracao:.2f} s ({max_workers} active cores)",
        f"- **{t('total_encontrados', cfg)}:** {total}",
        f"- **{t('vol_original', cfg)}:** {stats['tamanho_formatado']}",
        f"- **{t('indexados_sucesso', cfg)}:** {len(resultados_processados)}",
        f"- **{t('recup_cache', cfg)}:** {cont_cache}",
        f"- **{t('docs_ocr_aplicado', cfg)}:** {cont_ocr_itens}",
        f"- **{t('docs_identicos', cfg)}:** {len(duplicados)}",
        f"- **{t('falhas_leitura', cfg)}:** {len(erros)}\n",
        f"## {t('estrutura_saida', cfg)}:",
    ]
    for ag in arquivos_gerados:
        relatorio.append(f"- `{ag}`")

    if duplicados:
        relatorio.append(f"\n## {t('docs_identicos_detectados', cfg)}:")
        relatorio.extend(duplicados)

    if erros:
        relatorio.append(f"\n## {t('falhas_extracao', cfg)}:")
        relatorio.extend(erros)

    texto_relatorio = "\n".join(relatorio)
    (pasta_destino / "relatorio_auditoria.md").write_text(texto_relatorio, encoding="utf-8")
    (pasta_destino / "audit_report.md").write_text(texto_relatorio, encoding="utf-8")

    # Grava prompts nos dois idiomas para uso universal
    (pasta_destino / "PROMPT_PARA_COPILOT.txt").write_text(PROMPT_SISTEMA_COPILOT_PT, encoding="utf-8")
    (pasta_destino / "COPILOT_PROMPT.txt").write_text(PROMPT_SISTEMA_COPILOT_EN, encoding="utf-8")

    if cfg.get("usar_pasta_timestamp", True):
        link_ultimo = pasta_saida_base / "ultimo_processamento"
        try:
            if link_ultimo.exists():
                shutil.rmtree(link_ultimo, ignore_errors=True)
            shutil.copytree(pasta_destino, link_ultimo, ignore=shutil.ignore_patterns("*.zip"))
        except Exception:
            pass

    arquivo_zip_gerado = None
    if cfg.get("gerar_arquivo_zip", True):
        caminho_zip = pasta_destino / f"consolidated_{timestamp_str}.zip"
        print(f"\n{Cor.AZUL}{t('gerando_zip', cfg)}{Cor.RESET}")
        if compactar_pasta(pasta_destino, caminho_zip):
            tam_zip_mb = caminho_zip.stat().st_size / (1024 * 1024)
            arquivo_zip_gerado = f"{caminho_zip.name} ({tam_zip_mb:.2f} MB)"

    print(f"\n{Cor.VERDE}============================================================================{Cor.RESET}")
    print(f"{Cor.VERDE}  {t('concluido_sucesso', cfg, tempo=duracao)}{Cor.RESET}")
    print(f"{Cor.VERDE}============================================================================{Cor.RESET}")
    print(f"{t('pasta_execucao', cfg):<22}: {Cor.NEGRITO}{pasta_destino}{Cor.RESET}")
    print(f"{t('sucesso_total', cfg):<22}: {len(resultados_processados)} / {total} (Cache: {cont_cache} | OCR: {cont_ocr_itens})")
    print(f"{t('volume_proc', cfg):<22}: {stats['tamanho_formatado']}")

    print(f"\n{t('disponiveis_copilot', cfg)}")
    for ag in arquivos_gerados:
        print(f"  -> {Cor.TITULO}{ag}{Cor.RESET}")
    print(f"  -> {Cor.AMARELO}audit_report.md / relatorio_auditoria.md{Cor.RESET}")
    print(f"  -> {Cor.AMARELO}COPILOT_PROMPT.txt / PROMPT_PARA_COPILOT.txt{Cor.RESET}")

    if arquivo_zip_gerado:
        print(f"\n  -> {Cor.VERDE}{Cor.NEGRITO}{t('pacote_zip', cfg)}: {arquivo_zip_gerado}{Cor.RESET}")

    if pausar_no_final and "--auto" not in sys.argv:
        input(f"\n{t('voltar_menu', cfg)}")


def selecionar_pasta_gui() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta_sel = filedialog.askdirectory(title="Select Folder to Consolidate / Selecione a Pasta")
        root.destroy()
        if pasta_sel:
            return str(Path(pasta_sel).resolve())
    except Exception as e:
        print(f"{Cor.AMARELO}[Notice] Could not open GUI dialog: {e}{Cor.RESET}")
    return None


def ler_pastas_txt() -> list[str]:
    pastas = []
    if PASTAS_TXT_FILE.exists():
        for linha in PASTAS_TXT_FILE.read_text(encoding="utf-8").splitlines():
            l = linha.strip()
            if l and not l.startswith("#"):
                pastas.append(str(resolver_caminho(l)))
    return pastas


def salvar_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    linhas_txt = ["# Pastas monitoradas pelo Consolidador Local / Monitored Folders"]
    for p in cfg.get("pastas_entrada", []):
        linhas_txt.append(p)
    PASTAS_TXT_FILE.write_text("\n".join(linhas_txt) + "\n", encoding="utf-8")


def carregar_config() -> dict:
    cfg = CONFIG_PADRAO.copy()
    if CONFIG_FILE.exists():
        try:
            dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(dados, dict):
                cfg.update(dados)
        except json.JSONDecodeError as e:
            print(f"{Cor.AMARELO}[NOTICE] config.json has syntax error ({e}). Loaded defaults.{Cor.RESET}")
        except Exception as e:
            print(f"{Cor.AMARELO}[NOTICE] Error loading config.json ({e}). Loaded defaults.{Cor.RESET}")

    pastas_txt = ler_pastas_txt()
    for p in pastas_txt:
        if p not in cfg["pastas_entrada"]:
            cfg["pastas_entrada"].append(p)

    return cfg


def menu_adicionar_pastas(cfg: dict):
    lang = obter_idioma(cfg)
    while True:
        limpar_tela()
        titulo = "ADD INPUT FOLDERS" if lang == "en" else "ADICIONAR PASTAS DE ENTRADA"
        print(f"{Cor.TITULO}=== {titulo} ==={Cor.RESET}\n")
        lbl_pastas = "Currently monitored folders" if lang == "en" else "Pastas atualmente monitoradas"
        print(f"{lbl_pastas} ({len(cfg['pastas_entrada'])}):")
        for i, p in enumerate(cfg["pastas_entrada"], start=1):
            status = f"{Cor.VERDE}[OK]{Cor.RESET}" if Path(p).exists() else f"{Cor.VERMELHO}[MISSING]{Cor.RESET}"
            print(f"  [{i}] {status} {p}")

        print(f"\n{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
        if lang == "en":
            print(" [1] Open Windows Explorer Dialog (Browse and select folder)")
            print(" [2] Type or paste folder path manually")
            print(" [3] Add Parent Folder (automatically detect all subdirectories)")
            print(" [0] Return to main menu")
        else:
            print(" [1] Abrir Janela do Windows Explorer (Navegar e Selecionar)")
            print(" [2] Digitar ou colar caminho da pasta manualmente")
            print(" [3] Adicionar Pasta Mãe (detectar todas as subpastas automaticamente)")
            print(" [0] Voltar ao menu principal")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

        op = input("\nOption / Opção: ").strip()
        if op == "0" or not op:
            break
        elif op == "1":
            print("\nAbrindo janela de seleção / Opening browse dialog...")
            sel = selecionar_pasta_gui()
            if sel:
                if sel not in cfg["pastas_entrada"]:
                    cfg["pastas_entrada"].append(sel)
                    salvar_config(cfg)
                    print(f"{Cor.VERDE}[OK] Added: {sel}{Cor.RESET}")
                else:
                    print(f"{Cor.AMARELO}[Aviso] Folder already added.{Cor.RESET}")
                time.sleep(1.2)
        elif op == "2":
            prompt_c = "Enter or drop folder path: " if lang == "en" else "Digite ou arraste o caminho da pasta: "
            caminho = input(f"\n{prompt_c}").strip().strip('"')
            if caminho:
                p_res = str(resolver_caminho(caminho))
                if Path(p_res).exists():
                    if p_res not in cfg["pastas_entrada"]:
                        cfg["pastas_entrada"].append(p_res)
                        salvar_config(cfg)
                        print(f"{Cor.VERDE}[OK] Added: {p_res}{Cor.RESET}")
                    else:
                        print(f"{Cor.AMARELO}[Aviso] Already added.{Cor.RESET}")
                else:
                    print(f"{Cor.VERMELHO}[Erro] Path not found: {p_res}{Cor.RESET}")
                time.sleep(1.2)
        elif op == "3":
            prompt_m = "Enter parent folder path: " if lang == "en" else "Digite o caminho da Pasta Mãe: "
            caminho = input(f"\n{prompt_m}").strip().strip('"')
            if caminho:
                p_mae = resolver_caminho(caminho)
                if p_mae.exists() and p_mae.is_dir():
                    subpastas = [str(sp) for sp in p_mae.iterdir() if sp.is_dir()]
                    adicionadas = 0
                    for sp in subpastas:
                        if sp not in cfg["pastas_entrada"]:
                            cfg["pastas_entrada"].append(sp)
                            adicionadas += 1
                    salvar_config(cfg)
                    print(f"{Cor.VERDE}[OK] {adicionadas} subfolders added!{Cor.RESET}")
                else:
                    print(f"{Cor.VERMELHO}[Erro] Invalid parent directory.{Cor.RESET}")
                time.sleep(1.5)


def menu_listar_remover_pastas(cfg: dict):
    lang = obter_idioma(cfg)
    while True:
        limpar_tela()
        titulo = "MONITORED FOLDERS" if lang == "en" else "PASTAS MONITORADAS"
        print(f"{Cor.TITULO}=== {titulo} ==={Cor.RESET}\n")
        if not cfg["pastas_entrada"]:
            print(f"{Cor.AMARELO}(No folders currently configured / Nenhuma pasta configurada){Cor.RESET}")
            input(f"\n{t('voltar_menu', cfg)}")
            break

        for i, p in enumerate(cfg["pastas_entrada"], start=1):
            p_obj = Path(p)
            status = f"{Cor.VERDE}[OK]{Cor.RESET}" if p_obj.exists() else f"{Cor.VERMELHO}[MISSING]{Cor.RESET}"
            print(f" [{i:>2}] {status} {p}")

        if lang == "en":
            print("\nEnter folder number to REMOVE it")
            print("Enter [T] to REMOVE ALL folders")
            print("Enter [0] to Return")
        else:
            print("\nDigite o número da pasta para REMOVÊ-LA")
            print("Digite [T] para REMOVER TODAS as pastas")
            print("Digite [0] para Voltar")

        escolha = input("\nChoice / Escolha: ").strip().upper()
        if not escolha or escolha == "0":
            break
        elif escolha == "T":
            cfg["pastas_entrada"] = []
            salvar_config(cfg)
        elif escolha.isdigit():
            idx = int(escolha) - 1
            if 0 <= idx < len(cfg["pastas_entrada"]):
                cfg["pastas_entrada"].pop(idx)
                salvar_config(cfg)


def menu_opcoes_saida_e_ocr(cfg: dict):
    while True:
        limpar_tela()
        lang = obter_idioma(cfg)
        print(f"{Cor.TITULO}=== OUTPUT, SPLITTING, OCR & LANGUAGE SETTINGS ==={Cor.RESET}\n" if lang == "en" else f"{Cor.TITULO}=== CONFIGURAÇÃO DE SAÍDAS, PARTICIONAMENTO, OCR E IDIOMA ==={Cor.RESET}\n")

        if lang == "en":
            print(f"{Cor.NEGRITO}--- Markdown & RAG Output Options ---{Cor.RESET}")
            print(f"[1] Generate Single File (corpus_completo.md)       : {'YES' if cfg.get('gerar_arquivo_unico', True) else 'NO'}")
            print(f"[2] Split by Size Limit in MB                     : {'YES' if cfg.get('dividir_por_tamanho', True) else 'NO'}")
            print(f"[3]   ↳ Maximum Size per Part (MB)                : {cfg.get('limite_mb_por_parte', 8.0)} MB")
            print(f"[4] Split by Document Count per Batch             : {'YES' if cfg.get('dividir_por_quantidade', False) else 'NO'}")
            print(f"[5]   ↳ Documents per Batch                       : {cfg.get('qtd_docs_por_parte', 50)} docs")
            print(f"[6] Generate 1 Markdown per Document (1:1)        : {'YES' if cfg.get('gerar_individual_por_arquivo', False) else 'NO'}")
            print(f"[7] Generate Folder/Departmental Consolidated     : {'YES' if cfg.get('gerar_por_pasta_raiz', True) else 'NO'}")
            print(f"[8] Generate Structured RAG Dataset (corpus_rag)  : {'YES' if cfg.get('gerar_jsonl', True) else 'NO'}")
            print(f"\n{Cor.NEGRITO}--- Compression & Organization ---{Cor.RESET}")
            print(f"[9] Save in Timestamped Folder per Run            : {'YES' if cfg.get('usar_pasta_timestamp', True) else 'NO'}")
            print(f"[10] Generate Compressed Package (.ZIP)           : {'YES' if cfg.get('gerar_arquivo_zip', True) else 'NO'}")
            print(f"\n{Cor.NEGRITO}--- Local Optical Character Recognition (OCR) ---{Cor.RESET}")
            print(f"[11] OCR in PDFs (Scanned & Mixed Pages)          : {'ENABLED' if cfg.get('usar_ocr_em_pdf', True) else 'DISABLED'}")
            print(f"[12] OCR in Embedded Images (Word, Excel, PPTX)   : {'ENABLED' if cfg.get('usar_ocr_em_imagens_embutidas', True) else 'DISABLED'}")
            print(f"[13] DPI Resolution for PDF OCR                   : {cfg.get('dpi_ocr_pdf', 200)} DPI")
            print(f"[14] Minimum Image Dimension for OCR (px)         : {cfg.get('min_dimensao_imagem_ocr', 150)} px")
            print(f"[15] OCR Minimum Confidence Threshold (Score)     : {cfg.get('limiar_confianca_ocr', 0.50):.2f}")
            print(f"\n{Cor.NEGRITO}--- Performance & Destination ---{Cor.RESET}")
            print(f"[16] Multicore Concurrent CPU Processing          : {'ENABLED' if cfg.get('multicore', True) else 'DISABLED'}")
            print(f"[17] Change Output Folder (Current: {cfg.get('pasta_saida')})")
            print(f"[L] Switch Language / Idioma                      : English (Click to toggle)")
            print(f"\n{Cor.AMARELO}[R] Restore Factory Default Settings{Cor.RESET}")
            print("[0] Return to main menu")
        else:
            print(f"{Cor.NEGRITO}--- Modalidades de Saída Markdown & RAG ---{Cor.RESET}")
            print(f"[1] Gerar Arquivo Único (corpus_completo.md)       : {'SIM' if cfg.get('gerar_arquivo_unico', True) else 'NÃO'}")
            print(f"[2] Dividir por Tamanho Máximo em MB              : {'SIM' if cfg.get('dividir_por_tamanho', True) else 'NÃO'}")
            print(f"[3]   ↳ Limite de Tamanho por Parte (MB)          : {cfg.get('limite_mb_por_parte', 8.0)} MB")
            print(f"[4] Dividir por Quantidade de Documentos por Lote : {'SIM' if cfg.get('dividir_por_quantidade', False) else 'NÃO'}")
            print(f"[5]   ↳ Quantidade de Docs por Lote               : {cfg.get('qtd_docs_por_parte', 50)} docs")
            print(f"[6] Gerar 1 Arquivo Markdown por Documento (1:1)  : {'SIM' if cfg.get('gerar_individual_por_arquivo', False) else 'NÃO'}")
            print(f"[7] Gerar Consolidado por Pasta Departamental      : {'SIM' if cfg.get('gerar_por_pasta_raiz', True) else 'NÃO'}")
            print(f"[8] Gerar Base Estruturada RAG (corpus_rag.jsonl) : {'SIM' if cfg.get('gerar_jsonl', True) else 'NÃO'}")
            print(f"\n{Cor.NEGRITO}--- Compactação & Organização com Timestamp ---{Cor.RESET}")
            print(f"[9] Salvar em Pasta com Timestamp por Execução    : {'SIM' if cfg.get('usar_pasta_timestamp', True) else 'NÃO'}")
            print(f"[10] Gerar Pacote Compactado (.ZIP) ao final      : {'SIM' if cfg.get('gerar_arquivo_zip', True) else 'NÃO'}")
            print(f"\n{Cor.NEGRITO}--- Reconhecimento Óptico de Caracteres (OCR Local) ---{Cor.RESET}")
            print(f"[11] OCR em PDFs (Escaneados e Imagens Mistas)    : {'ATIVADO' if cfg.get('usar_ocr_em_pdf', True) else 'DESATIVADO'}")
            print(f"[12] OCR em Imagens Embutidas (Word, Excel, PPTX) : {'ATIVADO' if cfg.get('usar_ocr_em_imagens_embutidas', True) else 'DESATIVADO'}")
            print(f"[13] Resolução DPI para OCR em PDF               : {cfg.get('dpi_ocr_pdf', 200)} DPI")
            print(f"[14] Dimensão mínima de imagem para OCR (px)     : {cfg.get('min_dimensao_imagem_ocr', 150)} px")
            print(f"[15] Limiar de Confiança OCR (Score Mínimo)       : {cfg.get('limiar_confianca_ocr', 0.50):.2f}")
            print(f"\n{Cor.NEGRITO}--- Desempenho & Destino ---{Cor.RESET}")
            print(f"[16] Processamento Multicore Concorrente (CPU)   : {'ATIVADO' if cfg.get('multicore', True) else 'DESATIVADO'}")
            print(f"[17] Alterar Pasta de Saída (Atual: {cfg.get('pasta_saida')})")
            print(f"[L] Alternar Idioma / Switch Language             : Português (Clique para alternar)")
            print(f"\n{Cor.AMARELO}[R] Restaurar Configuração Padrão Original{Cor.RESET}")
            print("[0] Voltar ao menu principal")

        op = input("\nChoice / Opção: ").strip().upper()
        if op == "0" or not op:
            break
        elif op == "1":
            cfg["gerar_arquivo_unico"] = not cfg.get("gerar_arquivo_unico", True)
        elif op == "2":
            cfg["dividir_por_tamanho"] = not cfg.get("dividir_por_tamanho", True)
        elif op == "3":
            val = input("MB limit (e.g. 8, 50, 200): ").strip().replace(",", ".")
            try:
                mb = float(val)
                if mb >= 0.5:
                    cfg["limite_mb_por_parte"] = mb
            except ValueError:
                pass
        elif op == "4":
            cfg["dividir_por_quantidade"] = not cfg.get("dividir_por_quantidade", False)
        elif op == "5":
            val = input("Batch size (e.g. 25, 50, 100): ").strip()
            if val.isdigit() and int(val) > 0:
                cfg["qtd_docs_por_parte"] = int(val)
        elif op == "6":
            cfg["gerar_individual_por_arquivo"] = not cfg.get("gerar_individual_por_arquivo", False)
        elif op == "7":
            cfg["gerar_por_pasta_raiz"] = not cfg.get("gerar_por_pasta_raiz", True)
        elif op == "8":
            cfg["gerar_jsonl"] = not cfg.get("gerar_jsonl", True)
        elif op == "9":
            cfg["usar_pasta_timestamp"] = not cfg.get("usar_pasta_timestamp", True)
        elif op == "10":
            cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", True)
        elif op == "11":
            cfg["usar_ocr_em_pdf"] = not cfg.get("usar_ocr_em_pdf", True)
        elif op == "12":
            cfg["usar_ocr_em_imagens_embutidas"] = not cfg.get("usar_ocr_em_imagens_embutidas", True)
        elif op == "13":
            val = input("DPI (e.g. 150, 200, 300): ").strip()
            if val.isdigit() and int(val) >= 100:
                cfg["dpi_ocr_pdf"] = int(val)
        elif op == "14":
            val = input("Minimum dimension px (e.g. 100, 150, 200): ").strip()
            if val.isdigit() and int(val) >= 20:
                cfg["min_dimensao_imagem_ocr"] = int(val)
        elif op == "15":
            val = input("Minimum confidence score (0.10 to 0.90): ").strip().replace(",", ".")
            try:
                sc = float(val)
                if 0.05 <= sc <= 0.95:
                    cfg["limiar_confianca_ocr"] = sc
            except ValueError:
                pass
        elif op == "16":
            cfg["multicore"] = not cfg.get("multicore", True)
        elif op == "17":
            novo = input("Output folder: ").strip().strip('"')
            if novo:
                cfg["pasta_saida"] = str(resolver_caminho(novo))
        elif op == "L":
            cfg["idioma"] = "en" if cfg.get("idioma", "pt") == "pt" else "pt"
            print(f"\n{Cor.VERDE}Language set to: {'English' if cfg['idioma'] == 'en' else 'Português'}{Cor.RESET}")
            time.sleep(0.8)
        elif op == "R":
            pastas_atuais = cfg.get("pastas_entrada", [])
            lang_atual = cfg.get("idioma", "pt")
            cfg.clear()
            cfg.update(CONFIG_PADRAO.copy())
            cfg["pastas_entrada"] = pastas_atuais
            cfg["idioma"] = lang_atual
            print(f"\n{Cor.VERDE}[OK] Reset to default settings!{Cor.RESET}")
            time.sleep(1.0)

        salvar_config(cfg)


def menu_principal():
    cfg = carregar_config()
    while True:
        limpar_tela()
        lang = obter_idioma(cfg)
        cfg["idioma"] = lang

        qtd_pastas = len(cfg["pastas_entrada"])
        arquivos = listar_arquivos_elegiveis(cfg["pastas_entrada"], cfg["extensoes"]) if qtd_pastas else []
        stats = obter_estatisticas_arquivos(arquivos) if arquivos else {"total_arquivos": 0, "tamanho_formatado": "0 B"}
        num_cpus = os.cpu_count() or 4

        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f"{Cor.TITULO}   {t('titulo_app', cfg)}{Cor.RESET}")
        print(f"{Cor.TITULO}   {t('subtitulo_app', cfg, cpus=num_cpus)}{Cor.RESET}")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f" {t('status_atual', cfg)}")
        print(f"  • {t('pastas_monitoradas', cfg):<25}: {Cor.NEGRITO}{qtd_pastas}{Cor.RESET}")
        print(f"  • {t('arquivos_localizados', cfg):<25}: {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET} ({Cor.AMARELO}{stats['tamanho_formatado']}{Cor.RESET})")
        print(f"  • {t('multicore_paralelo', cfg):<25}: {Cor.VERDE + t('ativado', cfg) if cfg.get('multicore', True) else Cor.AMARELO + t('desativado', cfg)} ({num_cpus} cores){Cor.RESET}")
        print(f"  • {t('ocr_pdf_img', cfg):<25}: {Cor.VERDE + t('ativado', cfg) if cfg.get('usar_ocr_em_pdf', True) else Cor.AMARELO + t('desativado', cfg)}{Cor.RESET}")
        print(f"  • {t('ocr_embutidas', cfg):<25}: {Cor.VERDE + t('ativado', cfg) if cfg.get('usar_ocr_em_imagens_embutidas', True) else Cor.AMARELO + t('desativado', cfg)} (Word/Excel/PPTX){Cor.RESET}")

        partes_status = []
        if cfg.get("gerar_arquivo_unico", True):
            partes_status.append("Single" if lang == "en" else "Único")
        if cfg.get("dividir_por_tamanho", True):
            partes_status.append(f"{cfg.get('limite_mb_por_parte', 8.0)}MB")
        if cfg.get("dividir_por_quantidade", False):
            partes_status.append(f"{cfg.get('qtd_docs_por_parte', 50)} docs")
        if cfg.get("gerar_individual_por_arquivo", False):
            partes_status.append("1:1")
        if cfg.get("gerar_por_pasta_raiz", True):
            partes_status.append("Folder" if lang == "en" else "Por Pasta")
        if cfg.get("gerar_jsonl", True):
            partes_status.append("RAG JSONL")
        div_str = " | ".join(partes_status) if partes_status else t("nenhuma", cfg)

        print(f"  • {t('divisao_arq', cfg):<25}: {div_str}")
        print(f"  • {t('org_timestamp', cfg):<25}: {t('pasta_datahora', cfg) if cfg.get('usar_pasta_timestamp', True) else t('direta', cfg)} "
              f"{'+ ' + t('compactar_zip', cfg) if cfg.get('gerar_arquivo_zip', True) else ''}")
        print(f"  • {t('destino_base', cfg):<25}: {cfg.get('pasta_saida')}")
        print(f"{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
        print(f" {Cor.VERDE}{Cor.NEGRITO}{t('op_executar', cfg)}{Cor.RESET}")
        print(f" {t('op_adicionar', cfg)}")
        print(f" {t('op_listar', cfg, qtd=qtd_pastas)}")
        print(f" {t('op_recarregar', cfg)}")
        print(f" {t('op_config', cfg)}")
        print(f" {t('op_limpar_cache', cfg)}")
        print(f" {t('op_idioma', cfg)}")
        print(f" {t('op_sair', cfg)}")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

        opcao = input(f"\n{t('prompt_opcao', cfg)}").strip().upper()
        if opcao == "1":
            executar_pipeline(cfg)
        elif opcao == "2":
            menu_adicionar_pastas(cfg)
        elif opcao == "3":
            menu_listar_remover_pastas(cfg)
        elif opcao == "4":
            novas = ler_pastas_txt()
            for p in novas:
                if p not in cfg["pastas_entrada"]:
                    cfg["pastas_entrada"].append(p)
            salvar_config(cfg)
        elif opcao == "5":
            menu_opcoes_saida_e_ocr(cfg)
        elif opcao == "6":
            pasta_cache = resolver_caminho(cfg["pasta_saida"]) / ".cache_extracao"
            shutil.rmtree(pasta_cache, ignore_errors=True)
            msg_c = "[OK] Extraction cache cleared successfully!" if lang == "en" else "[OK] Cache SHA-256 limpo com sucesso!"
            print(f"\n{Cor.VERDE}{msg_c}{Cor.RESET}")
            time.sleep(1.0)
        elif opcao == "L":
            cfg["idioma"] = "en" if cfg.get("idioma", "pt") == "pt" else "pt"
            salvar_config(cfg)
        elif opcao == "0":
            break


if __name__ == "__main__":
    if "--auto" in sys.argv:
        cfg_auto = carregar_config()
        executar_pipeline(cfg_auto, pausar_no_final=False)
    else:
        try:
            menu_principal()
        except (KeyboardInterrupt, EOFError):
            print("\nClosed / Encerrado.")
