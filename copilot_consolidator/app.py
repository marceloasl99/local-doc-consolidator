#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidador Local de Documentos para Microsoft Copilot & RAG (v6.0 - Enterprise Edition)
Local Document Consolidator for Microsoft Copilot & RAG (v6.0 - Enterprise Edition)
Bilingual: English (start.bat / run.bat) & Português (iniciar_windows.bat)
100% Offline | Multicore Concurrency Control | Context-Aware Office Extraction
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

# Suprime avisos de bibliotecas e desativa sugestoes de layout do PyMuPDF
os.environ["PYMUPDF_SUGGEST_LAYOUT_ANALYZER"] = "0"
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*fitz.*")
warnings.filterwarnings("ignore", message=".*pymupdf_layout.*")

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
        "titulo_app": "CONSOLIDADOR LOCAL DE DOCUMENTOS PARA MICROSOFT COPILOT (v6.0)",
        "subtitulo_app": "Ambiente: Windows 11 | 100% Offline | {workers} Processos Concorrentes Seguros",
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
        "op_config": "[5] Configurar opções (Tamanhos/MB, Workers, PII, OCR, RAG JSONL)",
        "op_limpar_cache": "[6] Limpar cache de extração SHA-256 (forçar releitura completa)",
        "op_idioma": "[L] Idioma / Language (Atual: Português - Selecione para alternar)",
        "op_sair": "[0] Sair",
        "prompt_opcao": "Digite o número da opção desejada: ",
        "sem_pastas": "[!] Nenhuma pasta configurada! Adicione pastas na opção [2] ou no arquivo pastas.txt.",
        "sem_arquivos": "[!] Nenhum arquivo compatível foi encontrado nas pastas informadas.",
        "iniciando_conversao": "INICIANDO CONVERSÃO LOCAL PARA MICROSOFT COPILOT (v6.0)",
        "processando_paralelo": "Processando documentos com concorrência segura...",
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
        "titulo_app": "LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (v6.0)",
        "subtitulo_app": "Environment: Windows 11 | 100% Offline | {workers} Safe Concurrent Workers",
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
        "op_config": "[5] Configure options (Sizes/MB, Workers, PII, OCR, RAG JSONL)",
        "op_limpar_cache": "[6] Clear SHA-256 extraction cache (force full re-extraction)",
        "op_idioma": "[L] Language / Idioma (Current: English - Select to toggle)",
        "op_sair": "[0] Exit",
        "prompt_opcao": "Enter desired option: ",
        "sem_pastas": "[!] No folder configured! Add folders in option [2] or in pastas.txt.",
        "sem_arquivos": "[!] No compatible files found in configured folders.",
        "iniciando_conversao": "STARTING LOCAL CONSOLIDATION FOR MICROSOFT COPILOT (v6.0)",
        "processando_paralelo": "Processing documents with safe concurrency...",
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
    # Segurança, Concorrência e PII (Safe-by-default para Copilot / Nuvem):
    "usar_pasta_timestamp": True,
    "gerar_arquivo_zip": True,
    "mascarar_pii": False,
    "max_workers": min(4, os.cpu_count() or 4),
    # Configurações de OCR:
    "usar_ocr_em_pdf": True,
    "usar_ocr_em_imagens_embutidas": True,
    "dpi_ocr_pdf": 200,
    "min_dimensao_imagem_ocr": 150,
    "limiar_confianca_ocr": 0.50,
    "proporcao_max_aspecto_ocr": 20.0,
    "multicore": True,
    "extensoes": EXTENSOES_PADRAO,
}

_OCR_ENGINE = None

# Expressões regulares para redação de PII (LGPD / Privacidade - Opcional)
RE_CPF = re.compile(r'\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b')
RE_CNPJ = re.compile(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b')
RE_RG = re.compile(r'(?i)\b(RG|R\.G\.|Identidade|Cédula de Identidade)[\s:]+([0-9A-Za-z.-]{5,14})\b')
RE_RG_FORMATADO = re.compile(r'\b\d{1,2}\.\d{3}\.\d{3}-[\dX|x]\b')
RE_CEP = re.compile(r'(?i)\b(?:CEP[\s:]+)?\d{5}-\d{3}\b')
RE_EMAIL = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
RE_TELEFONE = re.compile(r'(?:\+?55\s?)?(?:\(\d{2}\)|\b\d{2}\b)[-.\s]?(?:9\d{4}|\d{4})[-.\s]?\d{4}\b')
RE_ENDERECO = re.compile(r'(?i)\b(?:Rua|Avenida|Av\.|Travessa|Alameda|Rodovia|Estrada|Praça|Praca)\s+[^,\n\r]{3,45},\s*(?:n[º°]?\s*)?\d+[A-Za-z0-9\s\-/]*')
RE_NOME_CONTEXTUAL = re.compile(r'(?i)\b(Nome|Titular|Cliente|Favorecido|Paciente|Colaborador|Funcion[aá]rio)[\s:]+([A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+(?:[ \t]+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+)+)')

# Padrões para correção de Mojibake (UTF-8 interpretado como CP1252/Latin-1)
PADRAO_MOJIBAKE = re.compile(r'(?:Ã[\x80-\xbf]|Â[\x80-\xbf]|â€[“”—‘’™\x80-\xbf]|[\w]Ã[\w])')

# Padrões para correção de palavras quebradas por OCR (erros de segmentação / kerning)
PADRAO_FONOTATICO_OCR = re.compile(
    r'\b([A-Za-zÀ-ÿ]{2,})\s+(nt|nd|mp|mb|rt|rd|lt|ld|lh|nh|rr|ss|ç)([a-zà-ÿ]+)\b',
    re.IGNORECASE,
)
PADRAO_SUFIXOS_OCR = re.compile(
    r'\b([A-Za-zÀ-ÿ]{3,})\s+(cao|ções|mentos?|idades?|mentais?|tivos?|tivas?|mente|ficados?|ficações|ficação)\b',
    re.IGNORECASE,
)
PADRAO_PREFIXOS_OCR = re.compile(
    r'\b(auto|tele|micro|macro|sub|sobre|contra|extra|multi|inter)\s+([a-zà-ÿ]{4,})\b',
    re.IGNORECASE,
)


def redigir_pii(texto: str) -> str:
    """Aplica anonimização local em dados sensíveis (CPF, RG, CNPJ, E-mail, Telefone, Endereços, Nomes)."""
    if not texto:
        return ""
    texto = RE_CPF.sub("[CPF REDIGIDO]", texto)
    texto = RE_CNPJ.sub("[CNPJ REDIGIDO]", texto)
    texto = RE_RG.sub(r'\1: [RG REDIGIDO]', texto)
    texto = RE_RG_FORMATADO.sub("[RG REDIGIDO]", texto)
    texto = RE_ENDERECO.sub("[ENDEREÇO REDIGIDO]", texto)
    texto = RE_CEP.sub("[CEP REDIGIDO]", texto)
    texto = RE_EMAIL.sub("[EMAIL REDIGIDO]", texto)
    texto = RE_TELEFONE.sub("[TELEFONE REDIGIDO]", texto)
    texto = RE_NOME_CONTEXTUAL.sub(r'\1: [NOME REDIGIDO]', texto)
    return texto


def reparar_mojibake(texto: str) -> tuple[str, int]:
    """
    Detecta e corrige corrupções clássicas de mojibake (ex: UTF-8 lido como CP1252/Latin-1 e re-salvo).
    Exemplos: 'ObservaÃ§Ã£o' -> 'Observação', 'MÃ©todo' -> 'Método', 'nÃ£o' -> 'não'.
    Retorna (texto_corrigido, total_correcoes).
    """
    if not texto:
        return texto, 0

    texto_trabalho = texto.lstrip("\ufeff")
    matches = PADRAO_MOJIBAKE.findall(texto_trabalho)
    if not matches:
        return texto_trabalho, 0

    total_correcoes = len(matches)

    # 1. Tentativa global: documento inteiro (mais rápido e consistente)
    try:
        candidato = texto_trabalho.encode("cp1252").decode("utf-8")
        if not PADRAO_MOJIBAKE.search(candidato):
            return candidato, total_correcoes
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass

    # 2. Tentativa linha por linha com fallback resiliente
    linhas = texto_trabalho.splitlines(keepends=True)
    linhas_corrigidas = []
    correcoes_feitas = 0

    for linha in linhas:
        if PADRAO_MOJIBAKE.search(linha):
            try:
                linha_recod = linha.encode("cp1252").decode("utf-8")
                linhas_corrigidas.append(linha_recod)
                correcoes_feitas += len(PADRAO_MOJIBAKE.findall(linha))
                continue
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass

            # 3. Tentativa token a token se a linha tiver caracteres fora de CP1252
            def substituir_token(match):
                nonlocal correcoes_feitas
                tok = match.group(0)
                try:
                    c = tok.encode("cp1252").decode("utf-8")
                    correcoes_feitas += 1
                    return c
                except Exception:
                    return tok

            token_regex = re.compile(r'\S*(?:Ã[\x80-\xbf]|Â[\x80-\xbf]|â€[“”—‘’™\x80-\xbf])\S*')
            linha_recod = token_regex.sub(substituir_token, linha)
            linhas_corrigidas.append(linha_recod)
        else:
            linhas_corrigidas.append(linha)

    return "".join(linhas_corrigidas), correcoes_feitas or total_correcoes


def corrigir_espacos_ocr(texto: str) -> tuple[str, int]:
    """
    Remove espaços espúrios inseridos no meio de palavras pelo OCR (erros de segmentação / kerning).
    Exemplos: 'Aute nticacao' -> 'Autenticacao', 'comprova nte' -> 'comprovante', 'docu mento' -> 'documento'.
    Retorna (texto_corrigido, total_correcoes).
    """
    if not texto:
        return texto, 0

    total_correcoes = 0

    def contar_sub(pat, t):
        nonlocal total_correcoes
        achados = len(pat.findall(t))
        total_correcoes += achados
        return pat.sub(r'\1\2\3' if pat == PADRAO_FONOTATICO_OCR else r'\1\2', t)

    t_atual = texto
    for _ in range(2):
        t_ant = t_atual
        t_atual = contar_sub(PADRAO_FONOTATICO_OCR, t_atual)
        t_atual = contar_sub(PADRAO_SUFIXOS_OCR, t_atual)
        t_atual = contar_sub(PADRAO_PREFIXOS_OCR, t_atual)
        if t_atual == t_ant:
            break

    return t_atual, total_correcoes


def obter_idioma(cfg: dict) -> str:
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
    lang = cfg.get("idioma") if (cfg and "idioma" in cfg) else obter_idioma(cfg)
    dicionario = I18N.get(lang, I18N["pt"])
    texto = dicionario.get(chave, I18N["pt"].get(chave, chave))
    if kwargs:
        try:
            return texto.format(**kwargs)
        except Exception:
            return texto
    return texto


def obter_ocr_worker():
    """Inicializa ou recupera o motor RapidOCR do processo worker atual com fallback automático."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            try:
                from rapidocr import RapidOCR
                _OCR_ENGINE = RapidOCR()
            except ImportError:
                from rapidocr_onnxruntime import RapidOCR
                _OCR_ENGINE = RapidOCR()
        except Exception:
            _OCR_ENGINE = False
    return _OCR_ENGINE if _OCR_ENGINE is not False else None


def limpar_tela():
    if "--auto" not in sys.argv:
        if platform.system() == "Windows":
            os.system("cls")
        else:
            try:
                sys.stdout.write("\033[2J\033[H")
                sys.stdout.flush()
            except Exception:
                os.system("clear")


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
    texto = ""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1", "iso-8859-1"):
        try:
            texto = caminho.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if not texto:
        try:
            texto = caminho.read_bytes().decode("utf-8", errors="replace")
        except Exception:
            texto = ""
    texto_corrigido, _ = reparar_mojibake(texto)
    return texto_corrigido


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


def extrair_csv_estruturado(caminho: Path) -> str:
    """Lê arquivos CSV identificando o delimitador e convertendo em tabela Markdown."""
    conteudo_bruto = ler_texto_com_fallback(caminho)
    if not conteudo_bruto.strip():
        return ""
    try:
        amostra = conteudo_bruto[:4096]
        dialeto = csv.Sniffer().sniff(amostra, delimiters=";,\t|")
        delimitador = dialeto.delimiter
    except Exception:
        primeira_linha = conteudo_bruto.splitlines()[0] if conteudo_bruto.splitlines() else ""
        delimitador = ";" if primeira_linha.count(";") > primeira_linha.count(",") else ","

    try:
        linhas = list(csv.reader(conteudo_bruto.splitlines(), delimiter=delimitador))
        if linhas and any(linhas):
            md_tab = tabela_para_markdown(linhas)
            if md_tab:
                return md_tab
    except Exception:
        pass
    return conteudo_bruto


def normalizar_texto_ocr(texto: str) -> str:
    """Recombina palavras hifenizadas e unifica quebras de linha duras de OCR."""
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
) -> tuple[str, float, str, str | None]:
    """
    Executa OCR em bytes de imagem retornando:
    (texto_extraido, confianca_media, status_estado, erro_str)
    status_estado pode ser: 'ok', 'empty_or_low_conf', 'error'
    """
    if not img_bytes:
        return "", 0.0, "empty_or_low_conf", None
    engine = obter_ocr_worker()
    if not engine:
        return "", 0.0, "error", "Motor OCR não pôde ser inicializado no worker"
    try:
        from PIL import Image
        with Image.open(io.BytesIO(img_bytes)) as im:
            w, h = im.size
            if max(w, h) < min_dim or min(w, h) < 25:
                return "", 0.0, "empty_or_low_conf", None

            ratio_aspecto = max(w, h) / max(min(w, h), 1)
            if ratio_aspecto > proporcao_max_aspecto:
                return "", 0.0, "empty_or_low_conf", None

            max_edge = max(w, h)
            if max_edge > 2400:
                ratio = 2400 / max_edge
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                img_bytes = buf.getvalue()

        resultado, _ = engine(img_bytes)
        if not resultado:
            return "", 0.0, "empty_or_low_conf", None

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
            return "", 0.0, "empty_or_low_conf", None

        conf_media = (sum(scores) / len(scores)) if scores else 0.0
        texto_bruto = "\n".join(linhas_validas)
        texto_limpo = normalizar_texto_ocr(texto_bruto)
        texto_limpo, _ = corrigir_espacos_ocr(texto_limpo)
        texto_limpo, _ = reparar_mojibake(texto_limpo)
        return texto_limpo, conf_media, "ok", None
    except Exception as e:
        return "", 0.0, "error", str(e)


def extrair_imagem_avulsa(
    caminho: Path,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float, dict]:
    """Processa arquivos de imagem avulsos diretamente via OCR."""
    diag = {"ok": 0, "empty": 0, "error": 0}
    try:
        from PIL import Image
        dados = caminho.read_bytes()
        with Image.open(io.BytesIO(dados)) as im:
            w, h = im.size
            formato = im.format or caminho.suffix.replace(".", "").upper()
    except Exception as e:
        diag["error"] += 1
        return f"*(Erro ao abrir imagem: {e})*", 0, 1, {}, [], 0.0, diag

    dim_avulsa = min(min_dim, 40)
    texto, conf, status, err = executar_ocr_em_bytes(
        dados,
        min_dim=dim_avulsa,
        limiar_confianca=limiar_confianca,
        proporcao_max_aspecto=proporcao_max_aspecto,
    )
    if status == "ok":
        diag["ok"] += 1
    elif status == "error":
        diag["error"] += 1
    else:
        diag["empty"] += 1

    meta_info = {"formato": formato, "resolucao": f"{w}x{h} px"}
    label_img = "Image" if lang == "en" else "Imagem"
    label_res = "Resolution" if lang == "en" else "Resolução"
    label_fmt = "Format" if lang == "en" else "Formato"
    label_ocr = "Text Extracted via Local OCR" if lang == "en" else "Texto Extraído via OCR Local"

    if status == "ok" and texto:
        bloco = (
            f"**[{label_img}: {caminho.name}]** ({label_res}: {w}x{h} px | {label_fmt}: {formato})\n\n"
            f"> **{label_ocr}:**\n"
            + "\n".join(f"> {linha}" for linha in texto.splitlines())
        )
        paginas_rag = [{"pagina": 1, "texto": texto}]
        return bloco, 1, 1, meta_info, paginas_rag, conf, diag
    elif status == "error":
        bloco = f"**[{label_img}: {caminho.name}]**\n> **[OCR Local - Falha Técnica]:** *({err})*"
        return bloco, 0, 1, meta_info, [], 0.0, diag
    else:
        label_sem_txt = "*(Image without readable text detected)*" if lang == "en" else "*(Imagem sem texto legível detectado)*"
        bloco = f"**[{label_img}: {caminho.name}]** ({label_res}: {w}x{h} px | {label_fmt}: {formato})\n{label_sem_txt}"
        return bloco, 0, 1, meta_info, [], 0.0, diag


def extrair_pdf_hibrido(
    caminho_pdf: Path,
    usar_ocr: bool = True,
    dpi: int = 200,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float, dict]:
    """Lê PDFs nativamente e aciona OCR com diagnóstico de 3 estados."""
    cont_ocr = 0
    blocos_conteudo = []
    paginas_rag = []
    scores_ocr = []
    meta_info = {}
    cache_dedup_doc = {}
    diag = {"ok": 0, "empty": 0, "error": 0}

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
            return msg_p, 0, total_paginas, meta_info, [], 0.0, diag

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
                try:
                    pix = pagina.get_pixmap(dpi=dpi)
                    img_bytes = pix.tobytes("png")
                    txt_ocr, conf_p, st, err = executar_ocr_em_bytes(
                        img_bytes,
                        min_dim=min_dim,
                        limiar_confianca=limiar_confianca,
                        proporcao_max_aspecto=proporcao_max_aspecto,
                    )
                    if st == "ok" and txt_ocr:
                        diag["ok"] += 1
                        tag_p = f"*(Page {num_pag} scanned - extracted via local OCR)*" if lang == "en" else f"*(Página {num_pag} escaneada - extraída via OCR local)*"
                        texto_pagina = f"{tag_p}\n{txt_ocr}"
                        cont_ocr += 1
                        scores_ocr.append(conf_p)
                    elif st == "error":
                        diag["error"] += 1
                        texto_pagina = f"*(Página {num_pag} escaneada - Erro de OCR: {err})*"
                    else:
                        diag["empty"] += 1
                        texto_pagina = f"*(Page {num_pag} scanned without readable text)*" if lang == "en" else f"*(Página {num_pag} escaneada sem texto legível)*"
                except Exception as e:
                    diag["error"] += 1
                    texto_pagina = texto_nativo or f"*(Página {num_pag} escaneada - Erro: {e})*"

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
                                        txt_img, conf_img, st, err = cache_dedup_doc[h_img]
                                    else:
                                        txt_img, conf_img, st, err = executar_ocr_em_bytes(
                                            img_bytes,
                                            min_dim=min_dim,
                                            limiar_confianca=limiar_confianca,
                                            proporcao_max_aspecto=proporcao_max_aspecto,
                                        )
                                        cache_dedup_doc[h_img] = (txt_img, conf_img, st, err)

                                    if st == "ok" and txt_img:
                                        diag["ok"] += 1
                                        lbl_emb = f"OCR Embedded Image #{img_idx}" if lang == "en" else f"OCR Imagem Embutida #{img_idx}"
                                        texto_ocr_extra.append(
                                            f"> **[{lbl_emb} ({w}x{h} px)]:**\n"
                                            + "\n".join(f"> {l}" for l in txt_img.splitlines())
                                        )
                                        cont_ocr += 1
                                        scores_ocr.append(conf_img)
                                    elif st == "error":
                                        diag["error"] += 1
                                        texto_ocr_extra.append(f"> **[OCR Imagem Embutida #{img_idx} - Falha Técnica]:** *({err})*")
                                    else:
                                        diag["empty"] += 1
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
        return "\n\n".join(blocos_conteudo), cont_ocr, total_paginas, meta_info, paginas_rag, conf_media, diag
    except ImportError:
        pass

    # Fallback pypdf
    from pypdf import PdfReader
    reader = PdfReader(str(caminho_pdf))
    total_paginas = len(reader.pages)
    for num_pag, pagina in enumerate(reader.pages, start=1):
        txt = (pagina.extract_text() or "").strip()
        if txt:
            blocos_conteudo.append(txt)
            paginas_rag.append({"pagina": num_pag, "texto": txt})
    return "\n\n".join(blocos_conteudo), cont_ocr, total_paginas, meta_info, paginas_rag, 0.0, diag


def extrair_docx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float, dict]:
    """Extrai parágrafos, tabelas e imagens em ordem contextual sequencial exata."""
    import docx
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    doc = docx.Document(str(caminho))
    partes = []
    meta_info = {}
    diag = {"ok": 0, "empty": 0, "error": 0}
    scores = []
    cont_ocr = 0
    cache_dedup = {}
    imagens_rids_vistos = set()

    if hasattr(doc, "core_properties"):
        cp = doc.core_properties
        if cp.title:
            meta_info["title"] = cp.title
        if cp.author:
            meta_info["author"] = cp.author
        if cp.created:
            meta_info["creationDate"] = cp.created.strftime("%Y-%m-%d")

    lbl_tab = "Table" if lang == "en" else "Tabela"
    lbl_tag = "OCR Image" if lang == "en" else "OCR Imagem"

    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            t_p = p.text.strip()
            if t_p:
                estilo = (p.style.name or "").lower() if p.style else ""
                if "heading 1" in estilo or "título 1" in estilo:
                    partes.append(f"### {t_p}")
                elif "heading 2" in estilo or "título 2" in estilo:
                    partes.append(f"#### {t_p}")
                else:
                    partes.append(t_p)

            # Imagens inline vinculadas a este parágrafo
            if usar_ocr_imagens:
                r_embeds = child.xpath('.//@r:embed')
                for r_id in r_embeds:
                    if r_id in doc.part.related_parts and r_id not in imagens_rids_vistos:
                        imagens_rids_vistos.add(r_id)
                        target = doc.part.related_parts[r_id]
                        if hasattr(target, "blob"):
                            img_bytes = target.blob
                            nome_img = getattr(target, "partname", f"image_{r_id}").split("/")[-1]
                            h_img = hashlib.sha256(img_bytes).hexdigest()
                            if h_img in cache_dedup:
                                txt_ocr, conf, st, err = cache_dedup[h_img]
                            else:
                                txt_ocr, conf, st, err = executar_ocr_em_bytes(
                                    img_bytes,
                                    min_dim=min_dim,
                                    limiar_confianca=limiar_confianca,
                                    proporcao_max_aspecto=proporcao_max_aspecto,
                                )
                                cache_dedup[h_img] = (txt_ocr, conf, st, err)

                            if st == "ok" and txt_ocr:
                                diag["ok"] += 1
                                cont_ocr += 1
                                scores.append(conf)
                                partes.append(
                                    f"> **[{lbl_tag}: `{nome_img}` (posicionada nesta seção)]:**\n"
                                    + "\n".join(f"> {l}" for l in txt_ocr.splitlines())
                                )
                            elif st == "error":
                                diag["error"] += 1
                                partes.append(f"> **[OCR Falha Técnica - `{nome_img}`]:** *({err})*")
                            else:
                                diag["empty"] += 1

        elif isinstance(child, CT_Tbl):
            t = Table(child, doc)
            linhas = [[cel.text.strip() for cel in row.cells] for row in t.rows]
            md_tab = tabela_para_markdown(linhas)
            if md_tab:
                partes.append(f"\n**{lbl_tab}:**\n{md_tab}")

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    paginas_rag = [{"pagina": 1, "texto": texto_final}]
    return texto_final, cont_ocr, 1, meta_info, paginas_rag, conf_media, diag


def extrair_xlsx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float, dict]:
    """Lê planilhas preservando fórmulas calculadas e vinculando imagens às respectivas abas."""
    import openpyxl
    partes = []
    meta_info = {}
    paginas_rag = []
    diag = {"ok": 0, "empty": 0, "error": 0}
    scores = []
    cont_ocr = 0
    cache_dedup = {}
    lbl_plan = "Sheet" if lang == "en" else "Planilha"

    wb_val = openpyxl.load_workbook(str(caminho), data_only=True, read_only=False)
    try:
        wb_form = openpyxl.load_workbook(str(caminho), data_only=False, read_only=False)
    except Exception:
        wb_form = None

    for idx_aba, nome_aba in enumerate(wb_val.sheetnames, start=1):
        ws_v = wb_val[nome_aba]
        ws_f = wb_form[nome_aba] if wb_form and nome_aba in wb_form.sheetnames else None

        linhas = []
        for r_idx, row in enumerate(ws_v.iter_rows(), start=1):
            linha_cols = []
            tem_dado = False
            for c_idx, cell in enumerate(row, start=1):
                val_v = cell.value
                val_str = "" if val_v is None else str(val_v).strip()
                if val_str:
                    tem_dado = True
                if ws_f:
                    try:
                        f_val = ws_f.cell(row=r_idx, column=c_idx).value
                        if isinstance(f_val, str) and f_val.startswith("="):
                            val_str += f" [Fórmula: {f_val}]"
                    except Exception:
                        pass
                linha_cols.append(val_str)
            if tem_dado:
                linhas.append(linha_cols)

        md_tab = tabela_para_markdown(linhas)
        texto_aba_blocos = []
        if md_tab:
            texto_aba_blocos.append(f"### {lbl_plan}: {nome_aba}\n\n{md_tab}")

        # Extrai imagens ancoradas nesta aba
        if usar_ocr_imagens and hasattr(ws_v, "_images") and ws_v._images:
            for im in ws_v._images:
                try:
                    img_data = im._data()
                    h_img = hashlib.sha256(img_data).hexdigest()
                    if h_img in cache_dedup:
                        txt_ocr, conf, st, err = cache_dedup[h_img]
                    else:
                        txt_ocr, conf, st, err = executar_ocr_em_bytes(
                            img_data,
                            min_dim=min_dim,
                            limiar_confianca=limiar_confianca,
                            proporcao_max_aspecto=proporcao_max_aspecto,
                        )
                        cache_dedup[h_img] = (txt_ocr, conf, st, err)

                    if st == "ok" and txt_ocr:
                        diag["ok"] += 1
                        cont_ocr += 1
                        scores.append(conf)
                        texto_aba_blocos.append(
                            f"\n> **[OCR Imagem na Planilha '{nome_aba}']:**\n"
                            + "\n".join(f"> {l}" for l in txt_ocr.splitlines())
                        )
                    elif st == "error":
                        diag["error"] += 1
                        texto_aba_blocos.append(f"\n> **[OCR Imagem Planilha '{nome_aba}' - Falha]:** *({err})*")
                    else:
                        diag["empty"] += 1
                except Exception:
                    pass

        if texto_aba_blocos:
            bloco_aba = "\n\n".join(texto_aba_blocos)
            partes.append(bloco_aba)
            paginas_rag.append({"pagina": idx_aba, "texto": bloco_aba})

    wb_val.close()
    if wb_form:
        wb_form.close()

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    return texto_final, cont_ocr, len(paginas_rag) or 1, meta_info, paginas_rag, conf_media, diag


def extrair_pptx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
) -> tuple[str, int, int, dict, list, float, dict]:
    """Extrai texto, tabelas, gráficos e imagens com OCR diretamente em cada slide."""
    from pptx import Presentation
    prs = Presentation(str(caminho))
    partes = []
    meta_info = {}
    paginas_rag = []
    diag = {"ok": 0, "empty": 0, "error": 0}
    scores = []
    cont_ocr = 0
    cache_dedup = {}
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
            if getattr(shape, "has_chart", False):
                try:
                    ch = shape.chart
                    tit = ch.chart_title.text_frame.text if ch.has_title else "Gráfico"
                    textos_slide.append(f"\n**[Gráfico: {tit}]**")
                except Exception:
                    pass
            if getattr(shape, "shape_type", None) == 13 or hasattr(shape, "image"):
                if usar_ocr_imagens:
                    try:
                        img_bytes = shape.image.blob
                        h_img = hashlib.sha256(img_bytes).hexdigest()
                        if h_img in cache_dedup:
                            txt_img, conf_img, st, err = cache_dedup[h_img]
                        else:
                            txt_img, conf_img, st, err = executar_ocr_em_bytes(
                                img_bytes,
                                min_dim=min_dim,
                                limiar_confianca=limiar_confianca,
                                proporcao_max_aspecto=proporcao_max_aspecto,
                            )
                            cache_dedup[h_img] = (txt_img, conf_img, st, err)

                        if st == "ok" and txt_img:
                            diag["ok"] += 1
                            cont_ocr += 1
                            scores.append(conf_img)
                            lbl_slide_img = "OCR Slide Image" if lang == "en" else "OCR Imagem do Slide"
                            textos_slide.append(
                                f"\n> **[{lbl_slide_img}]:**\n"
                                + "\n".join(f"> {l}" for l in txt_img.splitlines())
                            )
                        elif st == "error":
                            diag["error"] += 1
                            textos_slide.append(f"\n> **[OCR Slide {idx_s} - Falha]:** *({err})*")
                        else:
                            diag["empty"] += 1
                    except Exception:
                        pass

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                textos_slide.append(f"\n> **{lbl_notas}:** {notas}")

        bloco_slide = "\n".join(textos_slide)
        partes.append(bloco_slide)
        paginas_rag.append({"pagina": idx_s, "texto": bloco_slide})

    texto_final = "\n\n".join(partes)
    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    return texto_final, cont_ocr, len(prs.slides), meta_info, paginas_rag, conf_media, diag


def calcular_hash_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with caminho.open("rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def worker_processar_arquivo(args: tuple) -> tuple[str, str, str, dict, str | None]:
    """Worker multicore seguro que processa arquivos e grava cache atômico."""
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
        mascarar_pii = bool(cfg.get("mascarar_pii", False))

        flag_ocr = f"v5_{usar_pdf}_{usar_img}_{dpi_val}_{min_dim}_{int(limiar_conf * 100)}_{int(prop_max)}_{int(mascarar_pii)}"
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
                    "diag_ocr": dados.get("diag_ocr", {"ok": 0, "empty": 0, "error": 0}),
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
        diag_ocr = {"ok": 0, "empty": 0, "error": 0}

        if ext == ".pdf":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr = extrair_pdf_hibrido(
                arq,
                usar_ocr=cfg.get("usar_ocr_em_pdf", True),
                dpi=dpi_val,
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext in EXTENSOES_IMAGEM:
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr = extrair_imagem_avulsa(
                arq,
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".docx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr = extrair_docx_contextual(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".xlsx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr = extrair_xlsx_contextual(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".pptx":
            conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr = extrair_pptx_contextual(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
                limiar_confianca=limiar_conf,
                proporcao_max_aspecto=prop_max,
                lang=lang,
            )
        elif ext == ".csv":
            conteudo = extrair_csv_estruturado(arq)
            paginas_rag = [{"pagina": 1, "texto": conteudo}]
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

        conteudo, n_mojibake = reparar_mojibake(conteudo)
        for p in paginas_rag:
            p["texto"], _ = reparar_mojibake(p.get("texto", ""))

        if mascarar_pii:
            conteudo = redigir_pii(conteudo)
            for p in paginas_rag:
                p["texto"] = redigir_pii(p["texto"])

        conteudo = conteudo.strip()

        # Gravação atômica segura do cache
        tmp_cache = arq_cache.with_suffix(".tmp")
        tmp_cache.write_text(
            json.dumps(
                {
                    "arquivo": arq.name,
                    "conteudo": conteudo,
                    "ocr_pages": ocr_count,
                    "total_paginas": total_paginas,
                    "meta_info": meta_info,
                    "confianca_ocr": confianca_ocr,
                    "paginas_rag": paginas_rag,
                    "diag_ocr": diag_ocr,
                    "mojibake_repairs": n_mojibake,
                    "data_cache": datetime.now().isoformat(),
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        os.replace(tmp_cache, arq_cache)

        return arq.name, caminho_rel, raiz.name, {
            "cache": False,
            "ocr_pages": ocr_count,
            "total_paginas": total_paginas,
            "meta_info": meta_info,
            "confianca_ocr": confianca_ocr,
            "paginas_rag": paginas_rag,
            "diag_ocr": diag_ocr,
            "mojibake_repairs": n_mojibake,
            "hash": file_hash,
            "conteudo": conteudo,
        }, None
    except Exception as e:
        return arq.name, caminho_rel, raiz.name, {}, str(e)


def listar_arquivos_elegiveis(pastas: list[str], extensoes: list[str]) -> tuple[list[tuple[Path, Path]], list[str]]:
    exts = {e.lower() for e in extensoes}
    arquivos = []
    vistos_abs = set()
    pastas_sem_permissao = []

    for p_str in pastas:
        p_raiz = resolver_caminho(p_str)
        if not p_raiz.exists() or not p_raiz.is_dir():
            continue
        try:
            for item in p_raiz.rglob("*"):
                try:
                    if item.is_file() and item.suffix.lower() in exts:
                        if item.name.startswith("~$") or item.name.startswith("."):
                            continue
                        abs_path = item.resolve()
                        if abs_path not in vistos_abs:
                            vistos_abs.add(abs_path)
                            arquivos.append((p_raiz, item))
                except PermissionError:
                    pastas_sem_permissao.append(str(item))
                except Exception:
                    continue
        except PermissionError:
            pastas_sem_permissao.append(str(p_raiz))
        except Exception:
            continue

    arquivos.sort(key=lambda x: str(x[1]).lower())
    return arquivos, pastas_sem_permissao


def compactar_pasta(origem: Path, destino_zip: Path) -> bool:
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
                print(f"{Cor.VERMELHO}[AVISO / WARNING] Zip integrity notice on file: {falha}{Cor.RESET}")
                return False
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}[ERRO / ERROR] Zip compression failed: {e}{Cor.RESET}")
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

    arquivos, sem_permissao = listar_arquivos_elegiveis(pastas, cfg.get("extensoes", EXTENSOES_PADRAO))
    total = len(arquivos)

    if sem_permissao:
        msg_perm = f"[Notice] {len(sem_permissao)} files/folders ignored due to missing read permissions." if lang == "en" else f"[Aviso] {len(sem_permissao)} arquivos/pastas ignorados por falta de permissão de leitura."
        print(f"{Cor.AMARELO}{msg_perm}{Cor.RESET}")

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

    # Concorrência segura (padrão min(4, cpus) para evitar estouro de RAM com ONNX)
    num_cpus = os.cpu_count() or 4
    workers_config = int(cfg.get("max_workers", min(4, num_cpus)))
    usar_multicore = cfg.get("multicore", True)
    max_workers = workers_config if usar_multicore else 1

    # TELA DE CONFIRMAÇÃO PRÉVIA (interativo)
    if "--auto" not in sys.argv:
        while True:
            limpar_tela()
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
            titulo_conf = "CONFIRMATION OF LOCAL PROCESSING (v6.0)" if lang == "en" else "CONFIRMAÇÃO DO PROCESSAMENTO LOCAL (v6.0)"
            print(f"{Cor.TITULO}   {titulo_conf}{Cor.RESET}")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")

            lbl_resumo = "Corpus Summary to Process:" if lang == "en" else "Resumo do Corpus a Processar:"
            print(f" {Cor.NEGRITO}{lbl_resumo}{Cor.RESET}")
            print(f"  • {t('pastas_monitoradas', cfg):<25}: {len(pastas)}")
            print(f"  • {t('arquivos_localizados', cfg):<25}: {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET}")
            print(f"  • {t('volume_proc', cfg):<25}: {Cor.AMARELO}{Cor.NEGRITO}{stats['tamanho_formatado']}{Cor.RESET}")
            lbl_breakdown = "Format breakdown:" if lang == "en" else "Distribuição por formato:"
            lbl_files = "files" if lang == "en" else "arquivos"
            print(f"  • {lbl_breakdown}")
            for ext, info in sorted(stats["por_extensao"].items(), key=lambda x: x[1]["qtd"], reverse=True):
                print(f"     - {ext.upper():<6} : {info['qtd']:>3} {lbl_files} ({formatar_tamanho_bytes(info['bytes'])})")

            lbl_cfg = "Current Execution Settings:" if lang == "en" else "Configurações desta Execução:"
            print(f"\n {Cor.NEGRITO}{lbl_cfg}{Cor.RESET}")
            lbl_dest = "Destination folder" if lang == "en" else "Pasta de destino"
            lbl_cpu = "Safe CPU workers" if lang == "en" else "Processos concorrentes"
            lbl_cpu_desc = f"{max_workers} processes (RAM optimized)" if lang == "en" else f"{max_workers} processos seguros (RAM otimizada)"
            lbl_pii = "PII Redaction (Privacy)" if lang == "en" else "Privacidade PII (LGPD)"
            pii_val_conf = ("ENABLED" if cfg.get("mascarar_pii", False) else "DISABLED (Keep raw data)") if lang == "en" else ("ATIVADO" if cfg.get("mascarar_pii", False) else "DESATIVADO (Manter dados originais)")
            lbl_split_mb = "Split by MB limit" if lang == "en" else "Divisão por limite MB"
            split_mb_val = (f"YES ({cfg.get('limite_mb_por_parte', 8.0)} MB)" if cfg.get("dividir_por_tamanho", True) else "NO") if lang == "en" else (f"SIM ({cfg.get('limite_mb_por_parte', 8.0)} MB)" if cfg.get("dividir_por_tamanho", True) else "NÃO")
            lbl_single = "Single file output" if lang == "en" else "Gerar arquivo único"
            single_val = ("YES (corpus_completo.md)" if cfg.get("gerar_arquivo_unico", True) else "NO") if lang == "en" else ("SIM (corpus_completo.md)" if cfg.get("gerar_arquivo_unico", True) else "NÃO")
            lbl_rag = "RAG dataset (.JSONL)" if lang == "en" else "Base RAG (.JSONL)"
            rag_val = ("YES (corpus_rag.jsonl)" if cfg.get("gerar_jsonl", True) else "NO") if lang == "en" else ("SIM (corpus_rag.jsonl)" if cfg.get("gerar_jsonl", True) else "NÃO")
            lbl_zip = "Compress to .ZIP" if lang == "en" else "Compactar em .ZIP"
            zip_val = ("YES" if cfg.get("gerar_arquivo_zip", True) else "NO") if lang == "en" else ("SIM" if cfg.get("gerar_arquivo_zip", True) else "NÃO")
            lbl_ocr = "OCR in PDF & Images" if lang == "en" else "OCR em PDF e Imagens"
            ocr_val = (f"ENABLED (DPI {cfg.get('dpi_ocr_pdf', 200)} | Min {cfg.get('min_dimensao_imagem_ocr', 150)}px)" if cfg.get("usar_ocr_em_pdf", True) else "DISABLED") if lang == "en" else (f"ATIVADO (DPI {cfg.get('dpi_ocr_pdf', 200)} | Mín {cfg.get('min_dimensao_imagem_ocr', 150)}px)" if cfg.get("usar_ocr_em_pdf", True) else "DESATIVADO")

            print(f"  • {lbl_dest:<25}: {pasta_destino}")
            print(f"  • {lbl_cpu:<25}: {lbl_cpu_desc}")
            print(f"  • {lbl_pii:<25}: {pii_val_conf}")
            print(f"  • {lbl_split_mb:<25}: {split_mb_val}")
            print(f"  • {lbl_single:<25}: {single_val}")
            print(f"  • {lbl_rag:<25}: {rag_val}")
            print(f"  • {lbl_zip:<25}: {zip_val}")
            print(f"  • {lbl_ocr:<25}: {ocr_val}")
            print(f"{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
            if lang == "en":
                print(" [ENTER] or [C] Confirm and Start Processing")
                print(" [M] Change MB splitting limit (e.g., 8, 50, 100, 200 MB)")
                print(" [W] Adjust number of CPU workers (current: " + str(max_workers) + ")")
                print(" [P] Toggle PII Redaction (Names, CPF, RG, Addresses, Phone)")
                print(" [Z] Toggle .ZIP archive generation")
                print(" [V] Cancel and return to main menu")
            else:
                print(" [ENTER] ou [C] Confirmar e Iniciar o Processamento")
                print(" [M] Alterar limite de separação em MB (ex: 8, 50, 100, 200 MB)")
                print(" [W] Ajustar número de processos concorrentes (atual: " + str(max_workers) + ")")
                print(" [P] Alternar anonimização de PII (Excluir Nomes, CPF, RG, Endereços)")
                print(" [Z] Alternar geração de arquivo compactado .ZIP")
                print(" [V] Cancelar e voltar ao menu")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

            prompt_act = "Enter desired action: " if lang == "en" else "Digite a ação desejada: "
            escolha_conf = input(f"\n{prompt_act}").strip().upper()
            if escolha_conf in ("", "C", "S", "SIM", "Y", "YES"):
                break
            elif escolha_conf in ("V", "0"):
                return
            elif escolha_conf == "M":
                prompt_m_input = "Limit in MB (e.g., 8, 50, 200): " if lang == "en" else "Limite em MB (ex: 8, 50, 200): "
                novo_mb = input(f"\n{prompt_m_input}").strip().replace(",", ".")
                try:
                    v_mb = float(novo_mb)
                    if v_mb >= 0.5:
                        cfg["limite_mb_por_parte"] = v_mb
                        cfg["dividir_por_tamanho"] = True
                        salvar_config(cfg)
                except ValueError:
                    pass
            elif escolha_conf == "W":
                prompt_w_input = f"Workers count (1 to {num_cpus}): " if lang == "en" else f"Quantidade de workers (1 a {num_cpus}): "
                nw = input(f"\n{prompt_w_input}").strip()
                if nw.isdigit() and 1 <= int(nw) <= num_cpus:
                    cfg["max_workers"] = int(nw)
                    max_workers = int(nw)
                    salvar_config(cfg)
            elif escolha_conf == "P":
                cfg["mascarar_pii"] = not cfg.get("mascarar_pii", False)
                salvar_config(cfg)
            elif escolha_conf == "Z":
                cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", True)
                salvar_config(cfg)

    inicio = time.time()
    limpar_tela()
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
    print(f"{Cor.TITULO}  {t('iniciando_conversao', cfg)}{Cor.RESET}")
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")
    print(f"{t('pastas_monitoradas', cfg):<22}: {len(pastas)}")
    print(f"{t('arquivos_localizados', cfg):<22}: {Cor.NEGRITO}{total}{Cor.RESET} ({stats['tamanho_formatado']})")
    workers_label = f"{max_workers} processos seguros" if lang == "pt" else f"{max_workers} safe workers"
    print(f"{t('multicore_paralelo', cfg):<22}: {Cor.VERDE}{t('ativado', cfg)} ({workers_label}){Cor.RESET}")
    pii_ativo = cfg.get("mascarar_pii", False)
    pii_label = "Privacidade PII (LGPD)" if lang == "pt" else "PII Redaction (Privacy)"
    pii_desc = ("ATIVADO (Nomes, CPF, RG, Endereços, Tel mascarados)" if pii_ativo else "DESATIVADO (Dados originais preservados)") if lang == "pt" else ("ENABLED (Names, CPF, RG, Addresses, Phone masked)" if pii_ativo else "DISABLED (Raw data preserved)")
    print(f"{pii_label:<22}: {Cor.VERDE if pii_ativo else Cor.AMARELO}{pii_desc}{Cor.RESET}")
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
            eta_label = "Restante" if lang == "pt" else "ETA"

            if erro_str:
                err_label = "[ERROR]" if lang == "en" else "[ERRO]"
                print(f"({docs_concluidos:03d}/{total:03d}) {Cor.VERMELHO}{err_label}{Cor.RESET} {caminho_rel} -> {erro_str}")
                erros.append(f"- **[{docs_concluidos:03d}]** `{caminho_rel}` -> `{erro_str}`")
            else:
                resultados_processados.append((docs_concluidos, nome_arq, caminho_rel, raiz_nome, meta_res))
                status_str = f"{Cor.AZUL}[CACHE]{Cor.RESET}" if meta_res["cache"] else f"{Cor.VERDE}[OK]{Cor.RESET}"
                if meta_res.get("ocr_pages", 0) > 0:
                    status_str += f" {Cor.AMARELO}(+{meta_res['ocr_pages']} OCR){Cor.RESET}"
                print(f"({docs_concluidos:03d}/{total:03d}) {status_str} {caminho_rel} | {eta_label}: {eta_str}")

    resultados_processados.sort(key=lambda x: str(x[2]).lower())

    # Índice Mestre Global
    indice_geral = [
        t("base_copilot_titulo", cfg),
        f"- **{t('data_consolida', cfg)}:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"- **{t('total_docs_index', cfg)}:** {len(resultados_processados)}",
        f"- **{t('vol_original', cfg)}:** {stats['tamanho_formatado']}",
        f"- **{t('pastas_origem', cfg)}:** {len(pastas)}",
        f"\n## {t('indice_mestre', cfg)}\n",
    ]

    blocos_gerais = []
    blocos_info_docs = []
    blocos_por_pasta: dict[str, list[str]] = {}
    docs_individuais = []
    hashes_binarios: dict[str, str] = {}
    hashes_textuais: dict[str, str] = {}
    duplicados = []
    cont_cache = 0
    cont_ocr_itens = 0
    total_ocr_diag = {"ok": 0, "empty": 0, "error": 0}

    for novo_idx, (_, nome_arq, caminho_rel, raiz_nome, meta) in enumerate(resultados_processados, start=1):
        doc_id = f"[{novo_idx:03d}]"
        if meta.get("cache"):
            cont_cache += 1
        if meta.get("ocr_pages", 0) > 0:
            cont_ocr_itens += meta["ocr_pages"]

        # Acumula diagnóstico de 3 estados
        d_ocr = meta.get("diag_ocr", {})
        for k in total_ocr_diag:
            total_ocr_diag[k] += d_ocr.get(k, 0)

        # Deduplicação dupla (binária e textual)
        f_hash = meta.get("hash", "")
        conteudo_limpo = (meta.get("conteudo") or "").strip()
        h_textual = hashlib.sha256(conteudo_limpo.encode("utf-8")).hexdigest()[:20] if conteudo_limpo else ""
        aviso_dup = ""

        if f_hash and f_hash in hashes_binarios:
            orig_id = hashes_binarios[f_hash]
            msg_d = f"has identical binary file to {orig_id}" if lang == "en" else f"possui arquivo binário idêntico a {orig_id}"
            duplicados.append(f"- `{caminho_rel}` ({doc_id}) {msg_d}")
            aviso_dup = f" *(Binary identical to {orig_id})*" if lang == "en" else f" *(Binariamente idêntico a {orig_id})*"
        elif h_textual and h_textual in hashes_textuais:
            orig_id = hashes_textuais[h_textual]
            msg_d = f"has equivalent textual content to {orig_id}" if lang == "en" else f"possui conteúdo textual equivalente a {orig_id}"
            duplicados.append(f"- `{caminho_rel}` ({doc_id}) {msg_d}")
            aviso_dup = f" *(Equivalent text to {orig_id})*" if lang == "en" else f" *(Texto equivalente a {orig_id})*"
        else:
            if f_hash:
                hashes_binarios[f_hash] = f"{doc_id} (`{caminho_rel}`)"
            if h_textual:
                hashes_textuais[h_textual] = f"{doc_id} (`{caminho_rel}`)"

        tag_ocr = f" *(OCR: {meta['ocr_pages']})*" if meta.get("ocr_pages", 0) > 0 else ""
        linha_indice = f"- **{doc_id}** `{caminho_rel}`{tag_ocr}{aviso_dup}"
        indice_geral.append(linha_indice)

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
        blocos_info_docs.append((doc_id, caminho_rel, bloco_md))
        blocos_por_pasta.setdefault(raiz_nome, []).append(bloco_md)
        docs_individuais.append((doc_id, nome_arq, caminho_rel, bloco_md))

    cabecalho_str = "\n".join(indice_geral) + "\n"
    arquivos_gerados = []

    # 0. Geração de 00_INDICE_MESTRE.md isolado
    (pasta_destino / "00_INDICE_MESTRE.md").write_text(cabecalho_str, encoding="utf-8")
    arquivos_gerados.append("00_INDICE_MESTRE.md")

    # 1. Estratégia A: Arquivo Único Consolidado
    if cfg.get("gerar_arquivo_unico", True):
        arq_unico = pasta_destino / "corpus_completo.md"
        with arq_unico.open("w", encoding="utf-8") as f:
            f.write(cabecalho_str)
            for b in blocos_gerais:
                f.write(b)
        tam_mb = arq_unico.stat().st_size / (1024 * 1024)
        arquivos_gerados.append(f"corpus_completo.md ({tam_mb:.2f} MB)")

    # 2. Estratégia B: Particionamento Estrito por Limite MB com Índice Local
    if cfg.get("dividir_por_tamanho", True) and blocos_info_docs:
        limite_mb = float(cfg.get("limite_mb_por_parte", 8.0))
        limite_bytes = int(limite_mb * 1024 * 1024)

        for antigo in pasta_destino.glob("corpus_parte_*.md"):
            antigo.unlink(missing_ok=True)

        partes_para_gravar = []
        parte_num = 1
        buffer_docs = []
        buffer_bytes = 0
        docs_nesta_parte = []

        for doc_id, c_rel, b_texto in blocos_info_docs:
            tam_b = len(b_texto.encode("utf-8"))

            # Se um documento individual excede o limite sozinho, divide em trechos lógicos
            if tam_b > limite_bytes:
                paragrafos = b_texto.split("\n\n")
                trecho_buffer = []
                trecho_bytes = 0
                trecho_num = 1
                for par in paragrafos:
                    tam_par = len((par + "\n\n").encode("utf-8"))
                    if trecho_bytes + tam_par > limite_bytes and trecho_buffer:
                        # Fecha parte atual
                        nome_p = f"corpus_parte_{parte_num:02d}.md"
                        cab_local = (
                            f"# BASE DOCUMENTAL CORPORATIVA - PARTE {parte_num:02d}\n"
                            f"## ÍNDICE DESTA PARTE:\n- **{doc_id}** `{c_rel}` (Trecho {trecho_num})\n"
                            f"*(Consulte 00_INDICE_MESTRE.md para o catálogo global)*\n\n"
                        )
                        conteudo_parte = cab_local + "\n\n".join(trecho_buffer)
                        partes_para_gravar.append((nome_p, conteudo_parte))
                        parte_num += 1
                        trecho_num += 1
                        trecho_buffer = [par]
                        trecho_bytes = tam_par
                    else:
                        trecho_buffer.append(par)
                        trecho_bytes += tam_par
                if trecho_buffer:
                    nome_p = f"corpus_parte_{parte_num:02d}.md"
                    cab_local = (
                        f"# BASE DOCUMENTAL CORPORATIVA - PARTE {parte_num:02d}\n"
                        f"## ÍNDICE DESTA PARTE:\n- **{doc_id}** `{c_rel}` (Trecho {trecho_num})\n"
                        f"*(Consulte 00_INDICE_MESTRE.md para o catálogo global)*\n\n"
                    )
                    conteudo_parte = cab_local + "\n\n".join(trecho_buffer)
                    partes_para_gravar.append((nome_p, conteudo_parte))
                    parte_num += 1
                continue

            if buffer_bytes + tam_b > limite_bytes and buffer_docs:
                nome_p = f"corpus_parte_{parte_num:02d}.md"
                linhas_indice_local = [f"- **{did}** `{crl}`" for did, crl in docs_nesta_parte]
                cab_local = (
                    f"# BASE DOCUMENTAL CORPORATIVA - PARTE {parte_num:02d}\n"
                    f"## ÍNDICE DESTA PARTE ({len(docs_nesta_parte)} documentos):\n"
                    + "\n".join(linhas_indice_local)
                    + "\n\n*(Consulte 00_INDICE_MESTRE.md para o catálogo global de todos os documentos)*\n\n"
                )
                partes_para_gravar.append((nome_p, cab_local + "".join(buffer_docs)))
                parte_num += 1
                buffer_docs = [b_texto]
                buffer_bytes = tam_b
                docs_nesta_parte = [(doc_id, c_rel)]
            else:
                buffer_docs.append(b_texto)
                buffer_bytes += tam_b
                docs_nesta_parte.append((doc_id, c_rel))

        if buffer_docs:
            nome_p = f"corpus_parte_{parte_num:02d}.md"
            linhas_indice_local = [f"- **{did}** `{crl}`" for did, crl in docs_nesta_parte]
            cab_local = (
                f"# BASE DOCUMENTAL CORPORATIVA - PARTE {parte_num:02d}\n"
                f"## ÍNDICE DESTA PARTE ({len(docs_nesta_parte)} documentos):\n"
                + "\n".join(linhas_indice_local)
                + "\n\n*(Consulte 00_INDICE_MESTRE.md para o catálogo global de todos os documentos)*\n\n"
            )
            partes_para_gravar.append((nome_p, cab_local + "".join(buffer_docs)))

        # Evita redundância de saída: se o corpus couber em 1 parte e o arquivo único estiver ativo,
        # suprime a parte 01 solteira (corpus_completo.md já representa a base unificada).
        if len(partes_para_gravar) == 1 and cfg.get("gerar_arquivo_unico", True):
            pass
        else:
            for nome_p, conteudo_parte in partes_para_gravar:
                caminho_p = pasta_destino / nome_p
                caminho_p.write_text(conteudo_parte, encoding="utf-8")
                tam_p_mb = caminho_p.stat().st_size / (1024 * 1024)
                arquivos_gerados.append(f"{nome_p} ({tam_p_mb:.2f} MB)")

    # 3. Estratégia C: Lotes
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

    # 4. Estratégia D: 1:1 Individual
    if cfg.get("gerar_individual_por_arquivo", False) and docs_individuais:
        dir_indiv = pasta_destino / "individual"
        dir_indiv.mkdir(parents=True, exist_ok=True)
        for antigo in dir_indiv.glob("*.md"):
            antigo.unlink(missing_ok=True)
        for doc_id, nome_arq, caminho_rel, bloco_md in docs_individuais:
            nome_limpo = sanitizar_nome_arquivo(f"{doc_id}_{nome_arq}.md")
            (dir_indiv / nome_limpo).write_text(bloco_md.strip(), encoding="utf-8")
        arquivos_gerados.append(f"individual/*.md ({len(docs_individuais)} individual docs)")

    # 5. Estratégia E: Por Pasta
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

    # Detecção de Quase-Duplicados (Alta similaridade de vocabulário e extensão)
    doc_palavras = []
    for novo_idx, (_, nome_a, c_rel, r_nome, m) in enumerate(resultados_processados, start=1):
        doc_id = f"[{novo_idx:03d}]"
        txt_d = m.get("conteudo") or ""
        words = set(re.findall(r'\b[a-zA-ZÀ-ÿ]{4,}\b', txt_d.lower()))
        if len(words) >= 30:
            doc_palavras.append((doc_id, nome_a, c_rel, words, len(txt_d)))

    quase_duplicados_pt = []
    quase_duplicados_en = []
    for i in range(len(doc_palavras)):
        did1, n1, rel1, w1, l1 = doc_palavras[i]
        for j in range(i + 1, len(doc_palavras)):
            did2, n2, rel2, w2, l2 = doc_palavras[j]
            ratio_l = min(l1, l2) / max(l1, l2) if max(l1, l2) > 0 else 0
            if ratio_l < 0.60:
                continue
            inter = len(w1 & w2)
            menor_len = min(len(w1), len(w2))
            if menor_len == 0:
                continue
            overlap = inter / menor_len
            if overlap >= 0.80:
                pct = overlap * 100
                quase_duplicados_pt.append(f"- **{did1}** `{rel1}` e **{did2}** `{rel2}`: {pct:.1f}% de sobreposição textual (versão/modelo quase duplicado)")
                quase_duplicados_en.append(f"- **{did1}** `{rel1}` and **{did2}** `{rel2}`: {pct:.1f}% text overlap (near-duplicate template/version)")

    if quase_duplicados_pt:
        aviso_qd = (
            "\n## NOTAS DE SIMILARIDADE E QUASE-DUPLICATAS (ATENÇÃO COPILOT):\n"
            + "\n".join(quase_duplicados_pt) + "\n"
        )
        (pasta_destino / "00_INDICE_MESTRE.md").write_text(cabecalho_str + aviso_qd, encoding="utf-8")

    # Estimativa de tokens do corpus via heurística (4 caracteres por token)
    total_chars = sum(len(b) for b in blocos_gerais)
    tokens_estimados = total_chars // 4
    duracao = time.time() - inicio
    total_mojibake = sum(m.get("mojibake_repairs", 0) for _, _, _, _, m in resultados_processados)

    def construir_relatorio(idioma: str) -> str:
        c_dict = dict(cfg)
        c_dict["idioma"] = idioma
        d_tit = t("relatorio_auditoria_titulo", c_dict)
        dt_lbl = "Date/Time" if idioma == "en" else "Data/Hora"
        tok_lbl = f"(~{tokens_estimados:,} estimated heuristic tokens)" if idioma == "en" else f"(~{tokens_estimados:,} tokens estimados via heurística)"
        workers_lbl = f"{max_workers} safe workers" if idioma == "en" else f"{max_workers} processos seguros"
        pii_lbl = ("ENABLED (Names, CPF, RG, Addresses, Phone masked)" if cfg.get("mascarar_pii", False) else "DISABLED (Raw data preserved)") if idioma == "en" else ("ATIVADO (Nomes, CPF, RG, Endereços, Tel mascarados)" if cfg.get("mascarar_pii", False) else "DESATIVADO (Dados originais preservados)")
        mojibake_lbl = f"- **Corrupted encoding fixes (Mojibake auto-repair):** {total_mojibake} occurrences corrected" if idioma == "en" else f"- **Correções de codificação corrompida (Mojibake reparado):** {total_mojibake} ocorrências corrigidas"

        linhas_rel = [
            d_tit,
            f"- **{dt_lbl}:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            f"- **{t('tempo_total', c_dict)}:** {duracao:.2f} s ({workers_lbl})",
            f"- **{t('total_encontrados', c_dict)}:** {total}",
            f"- **{t('vol_original', c_dict)}:** {stats['tamanho_formatado']} {tok_lbl}",
            f"- **{t('indexados_sucesso', c_dict)}:** {len(resultados_processados)}",
            f"- **{t('recup_cache', c_dict)}:** {cont_cache}",
            f"- **{t('docs_ocr_aplicado', c_dict)}:** {cont_ocr_itens}",
            f"  - {'OCR Success' if idioma == 'en' else 'OCR Sucesso'} (Score >= {cfg.get('limiar_confianca_ocr', 0.50):.2f}): {total_ocr_diag['ok']}",
            f"  - {'Images Without Readable Text' if idioma == 'en' else 'Imagens Sem Texto Legível'}: {total_ocr_diag['empty']}",
            f"  - {'OCR Technical Failures' if idioma == 'en' else 'Falhas Técnicas de OCR'}: {total_ocr_diag['error']}",
            mojibake_lbl,
            f"- **{t('docs_identicos', c_dict)}:** {len(duplicados)}",
            f"- **{t('falhas_leitura', c_dict)}:** {len(erros)}",
            f"- **{'PII Redaction / Privacy' if idioma == 'en' else 'Filtro PII / LGPD'}:** {pii_lbl}\n",
            f"## {t('estrutura_saida', c_dict)}:",
        ]
        for ag in arquivos_gerados:
            linhas_rel.append(f"- `{ag}`")

        if tokens_estimados > 128_000:
            if idioma == "en":
                linhas_rel.append(
                    f"\n> ⚠️ **Context Window Advisory:** The corpus totals ~{tokens_estimados:,} estimated heuristic tokens. "
                    f"For Microsoft 365 Copilot queries, prefer attaching split parts (`corpus_parte_*.md`) or departmental files (`por_pasta/`) "
                    f"to prevent model context window saturation."
                )
            else:
                linhas_rel.append(
                    f"\n> ⚠️ **Aviso de Janela de Contexto:** O corpus totaliza ~{tokens_estimados:,} tokens estimados via heurística. "
                    f"Para consultas no Microsoft 365 Copilot, prefira anexar partes fatiadas (`corpus_parte_*.md`) ou os consolidados por departamento (`por_pasta/`) "
                    f"para evitar saturação da janela do modelo."
                )

        qd_list = quase_duplicados_en if idioma == "en" else quase_duplicados_pt
        if qd_list:
            tit_qd = "Near-Duplicate Documents Detected (High Textual Overlap)" if idioma == "en" else "Documentos Quase-Duplicados Detectados (Alta Sobreposição Textual)"
            linhas_rel.append(f"\n## {tit_qd}:")
            linhas_rel.extend(qd_list)

        if duplicados:
            linhas_rel.append(f"\n## {t('docs_identicos_detectados', c_dict)}:")
            linhas_rel.extend(duplicados)

        if erros:
            linhas_rel.append(f"\n## {t('falhas_extracao', c_dict)}:")
            linhas_rel.extend(erros)

        return "\n".join(linhas_rel)

    (pasta_destino / "relatorio_auditoria.md").write_text(construir_relatorio("pt"), encoding="utf-8")
    (pasta_destino / "audit_report.md").write_text(construir_relatorio("en"), encoding="utf-8")

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
    token_heuristica = f"(~{tokens_estimados:,} estimated heuristic tokens)" if lang == "en" else f"(~{tokens_estimados:,} tokens estimados via heurística)"
    print(f"{t('volume_proc', cfg):<22}: {stats['tamanho_formatado']} {token_heuristica}")
    if total_mojibake > 0:
        mojibake_term = f"Mojibake repaired     : {total_mojibake} occurrences" if lang == "en" else f"Mojibake corrigido    : {total_mojibake} ocorrências"
        print(f"{mojibake_term}")

    print(f"\n{t('disponiveis_copilot', cfg)}")
    for ag in arquivos_gerados:
        print(f"  -> {Cor.TITULO}{ag}{Cor.RESET}")
    if lang == "en":
        print(f"  -> {Cor.AMARELO}audit_report.md (detailed audit report){Cor.RESET}")
        print(f"  -> {Cor.AMARELO}COPILOT_PROMPT.txt (anti-hallucination system prompt){Cor.RESET}")
    else:
        print(f"  -> {Cor.AMARELO}relatorio_auditoria.md (relatório de auditoria detalhado){Cor.RESET}")
        print(f"  -> {Cor.AMARELO}PROMPT_PARA_COPILOT.txt (prompt anti-alucinação){Cor.RESET}")

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
        title_dlg = "Select Folder to Consolidate" if "--lang" in sys.argv or "--en" in sys.argv else "Selecione a Pasta para Consolidar"
        pasta_sel = filedialog.askdirectory(title=title_dlg)
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
    # Normaliza caminhos e elimina duplicatas de caminhos relativos/absolutos
    pastas_normalizadas = []
    vistos = set()
    for p in cfg.get("pastas_entrada", []) + pastas_txt:
        p_res = str(resolver_caminho(p))
        if p_res not in vistos:
            vistos.add(p_res)
            pastas_normalizadas.append(p_res)
    cfg["pastas_entrada"] = pastas_normalizadas

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
            lbl_miss = "[MISSING]" if lang == "en" else "[INEXISTENTE]"
            status = f"{Cor.VERDE}[OK]{Cor.RESET}" if Path(p).exists() else f"{Cor.VERMELHO}{lbl_miss}{Cor.RESET}"
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

        prompt_opt = "Enter option: " if lang == "en" else "Digite a opção: "
        op = input(f"\n{prompt_opt}").strip()
        if op == "0" or not op:
            break
        elif op == "1":
            msg_dlg = "Opening browse dialog..." if lang == "en" else "Abrindo janela de seleção do Windows..."
            print(f"\n{msg_dlg}")
            sel = selecionar_pasta_gui()
            if sel:
                if sel not in cfg["pastas_entrada"]:
                    cfg["pastas_entrada"].append(sel)
                    salvar_config(cfg)
                    msg_add = f"[OK] Added: {sel}" if lang == "en" else f"[OK] Adicionada: {sel}"
                    print(f"{Cor.VERDE}{msg_add}{Cor.RESET}")
                else:
                    msg_warn = "[Notice] Folder already added." if lang == "en" else "[Aviso] Pasta já adicionada."
                    print(f"{Cor.AMARELO}{msg_warn}{Cor.RESET}")
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
                        msg_add = f"[OK] Added: {p_res}" if lang == "en" else f"[OK] Adicionada: {p_res}"
                        print(f"{Cor.VERDE}{msg_add}{Cor.RESET}")
                    else:
                        msg_warn = "[Notice] Folder already added." if lang == "en" else "[Aviso] Pasta já adicionada."
                        print(f"{Cor.AMARELO}{msg_warn}{Cor.RESET}")
                else:
                    msg_err = f"[Error] Path not found: {p_res}" if lang == "en" else f"[Erro] Caminho não encontrado: {p_res}"
                    print(f"{Cor.VERMELHO}{msg_err}{Cor.RESET}")
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
                    msg_ok = f"[OK] {adicionadas} subfolders added!" if lang == "en" else f"[OK] {adicionadas} subpastas adicionadas!"
                    print(f"{Cor.VERDE}{msg_ok}{Cor.RESET}")
                else:
                    msg_err = "[Error] Invalid parent directory." if lang == "en" else "[Erro] Diretório mãe inválido."
                    print(f"{Cor.VERMELHO}{msg_err}{Cor.RESET}")
                time.sleep(1.5)


def menu_listar_remover_pastas(cfg: dict):
    lang = obter_idioma(cfg)
    while True:
        limpar_tela()
        titulo = "MONITORED FOLDERS" if lang == "en" else "PASTAS MONITORADAS"
        print(f"{Cor.TITULO}=== {titulo} ==={Cor.RESET}\n")
        if not cfg["pastas_entrada"]:
            msg_empty = "(No folders currently configured)" if lang == "en" else "(Nenhuma pasta configurada no momento)"
            print(f"{Cor.AMARELO}{msg_empty}{Cor.RESET}")
            input(f"\n{t('voltar_menu', cfg)}")
            break

        for i, p in enumerate(cfg["pastas_entrada"], start=1):
            p_obj = Path(p)
            lbl_miss = "[MISSING]" if lang == "en" else "[INEXISTENTE]"
            status = f"{Cor.VERDE}[OK]{Cor.RESET}" if p_obj.exists() else f"{Cor.VERMELHO}{lbl_miss}{Cor.RESET}"
            print(f" [{i:>2}] {status} {p}")

        if lang == "en":
            print("\nEnter folder number to REMOVE it")
            print("Enter [T] to REMOVE ALL folders")
            print("Enter [0] to Return to main menu")
        else:
            print("\nDigite o número da pasta para REMOVÊ-LA")
            print("Digite [T] para REMOVER TODAS as pastas")
            print("Digite [0] para Voltar ao menu principal")

        prompt_c = "Enter choice: " if lang == "en" else "Digite sua opção: "
        escolha = input(f"\n{prompt_c}").strip().upper()
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
        print(f"{Cor.TITULO}=== SETTINGS (v6.0) - OUTPUT, OCR, PRIVACY & CONCURRENCY ==={Cor.RESET}\n" if lang == "en" else f"{Cor.TITULO}=== CONFIGURAÇÕES (v6.0) - SAÍDAS, OCR, LGPD E CONCORRÊNCIA ==={Cor.RESET}\n")

        num_cpus = os.cpu_count() or 4
        workers_atual = int(cfg.get("max_workers", min(4, num_cpus)))

        if lang == "en":
            print(f"{Cor.NEGRITO}--- Markdown & RAG Output Options ---{Cor.RESET}")
            print(f"[1] Generate Single File (corpus_completo.md)       : {'YES' if cfg.get('gerar_arquivo_unico', True) else 'NO'}")
            print(f"[2] Split by Size Limit in MB (Strict Chunks)     : {'YES' if cfg.get('dividir_por_tamanho', True) else 'NO'}")
            print(f"[3]   ↳ Maximum Size per Part (MB)                : {cfg.get('limite_mb_por_parte', 8.0)} MB")
            print(f"[4] Split by Document Count per Batch             : {'YES' if cfg.get('dividir_por_quantidade', False) else 'NO'}")
            print(f"[5]   ↳ Documents per Batch                       : {cfg.get('qtd_docs_por_parte', 50)} docs")
            print(f"[6] Generate 1 Markdown per Document (1:1)        : {'YES' if cfg.get('gerar_individual_por_arquivo', False) else 'NO'}")
            print(f"[7] Generate Folder/Departmental Consolidated     : {'YES' if cfg.get('gerar_por_pasta_raiz', True) else 'NO'}")
            print(f"[8] Generate Structured RAG Dataset (corpus_rag)  : {'YES' if cfg.get('gerar_jsonl', True) else 'NO'}")
            print(f"\n{Cor.NEGRITO}--- Privacy, Security & Concurrency ---{Cor.RESET}")
            print(f"[9] PII Redaction (Mask Names, CPF, RG, Addresses, etc.) : {'ENABLED' if cfg.get('mascarar_pii', False) else 'DISABLED'}")
            print(f"[10] Safe Concurrent CPU Workers (RAM Protect)    : {workers_atual} workers (max: {num_cpus})")
            print(f"[11] Generate Compressed Package (.ZIP)           : {'YES' if cfg.get('gerar_arquivo_zip', True) else 'NO'}")
            print(f"\n{Cor.NEGRITO}--- Local Optical Character Recognition (OCR) ---{Cor.RESET}")
            print(f"[12] OCR in PDFs (Scanned & Mixed Pages)          : {'ENABLED' if cfg.get('usar_ocr_em_pdf', True) else 'DISABLED'}")
            print(f"[13] OCR in Embedded Images (Word, Excel, PPTX)   : {'ENABLED' if cfg.get('usar_ocr_em_imagens_embutidas', True) else 'DISABLED'}")
            print(f"[14] DPI Resolution for PDF OCR                   : {cfg.get('dpi_ocr_pdf', 200)} DPI")
            print(f"[15] Minimum Image Dimension for OCR (px)         : {cfg.get('min_dimensao_imagem_ocr', 150)} px")
            print(f"[16] OCR Minimum Confidence Threshold (Score)     : {cfg.get('limiar_confianca_ocr', 0.50):.2f}")
            print(f"[17] Change Output Folder (Current: {cfg.get('pasta_saida')})")
            print(f"[L] Switch Language                               : English (Select to toggle)")
            print(f"\n{Cor.AMARELO}[R] Restore Factory Default Settings{Cor.RESET}")
            print("[0] Return to main menu")
        else:
            print(f"{Cor.NEGRITO}--- Modalidades de Saída Markdown & RAG ---{Cor.RESET}")
            print(f"[1] Gerar Arquivo Único (corpus_completo.md)       : {'SIM' if cfg.get('gerar_arquivo_unico', True) else 'NÃO'}")
            print(f"[2] Dividir por Tamanho Máximo em MB (Estrito)    : {'SIM' if cfg.get('dividir_por_tamanho', True) else 'NÃO'}")
            print(f"[3]   ↳ Limite de Tamanho por Parte (MB)          : {cfg.get('limite_mb_por_parte', 8.0)} MB")
            print(f"[4] Dividir por Quantidade de Documentos por Lote : {'SIM' if cfg.get('dividir_por_quantidade', False) else 'NÃO'}")
            print(f"[5]   ↳ Quantidade de Docs por Lote               : {cfg.get('qtd_docs_por_parte', 50)} docs")
            print(f"[6] Gerar 1 Arquivo Markdown por Documento (1:1)  : {'SIM' if cfg.get('gerar_individual_por_arquivo', False) else 'NÃO'}")
            print(f"[7] Gerar Consolidado por Pasta Departamental      : {'SIM' if cfg.get('gerar_por_pasta_raiz', True) else 'NÃO'}")
            print(f"[8] Gerar Base Estruturada RAG (corpus_rag.jsonl) : {'SIM' if cfg.get('gerar_jsonl', True) else 'NÃO'}")
            print(f"\n{Cor.NEGRITO}--- Privacidade, Segurança e Concorrência ---{Cor.RESET}")
            print(f"[9] Anonimização de PII (Excluir Nomes, CPF, RG, Endereços) : {'ATIVADO' if cfg.get('mascarar_pii', False) else 'DESATIVADO'}")
            print(f"[10] Processos Concorrentes Seguros (Anti-Estouro): {workers_atual} workers (máx: {num_cpus})")
            print(f"[11] Gerar Pacote Compactado (.ZIP) ao final      : {'SIM' if cfg.get('gerar_arquivo_zip', True) else 'NÃO'}")
            print(f"\n{Cor.NEGRITO}--- Reconhecimento Óptico de Caracteres (OCR Local) ---{Cor.RESET}")
            print(f"[12] OCR em PDFs (Escaneados e Imagens Mistas)    : {'ATIVADO' if cfg.get('usar_ocr_em_pdf', True) else 'DESATIVADO'}")
            print(f"[13] OCR em Imagens Embutidas (Word, Excel, PPTX) : {'ATIVADO' if cfg.get('usar_ocr_em_imagens_embutidas', True) else 'DESATIVADO'}")
            print(f"[14] Resolução DPI para OCR em PDF               : {cfg.get('dpi_ocr_pdf', 200)} DPI")
            print(f"[15] Dimensão mínima de imagem para OCR (px)     : {cfg.get('min_dimensao_imagem_ocr', 150)} px")
            print(f"[16] Limiar de Confiança OCR (Score Mínimo)       : {cfg.get('limiar_confianca_ocr', 0.50):.2f}")
            print(f"[17] Alterar Pasta de Saída (Atual: {cfg.get('pasta_saida')})")
            print(f"[L] Alternar Idioma                               : Português (Selecione para alternar)")
            print(f"\n{Cor.AMARELO}[R] Restaurar Configuração Padrão Original{Cor.RESET}")
            print("[0] Voltar ao menu principal")

        prompt_c = "Enter option: " if lang == "en" else "Digite a opção: "
        op = input(f"\n{prompt_c}").strip().upper()
        if op == "0" or not op:
            break
        elif op == "1":
            cfg["gerar_arquivo_unico"] = not cfg.get("gerar_arquivo_unico", True)
        elif op == "2":
            cfg["dividir_por_tamanho"] = not cfg.get("dividir_por_tamanho", True)
        elif op == "3":
            prompt_mb = "MB limit (e.g. 8, 50, 200): " if lang == "en" else "Limite em MB (ex: 8, 50, 200): "
            val = input(f"\n{prompt_mb}").strip().replace(",", ".")
            try:
                mb = float(val)
                if mb >= 0.5:
                    cfg["limite_mb_por_parte"] = mb
            except ValueError:
                pass
        elif op == "4":
            cfg["dividir_por_quantidade"] = not cfg.get("dividir_por_quantidade", False)
        elif op == "5":
            prompt_b = "Batch size (e.g. 25, 50, 100): " if lang == "en" else "Tamanho do lote (ex: 25, 50, 100): "
            val = input(f"\n{prompt_b}").strip()
            if val.isdigit() and int(val) > 0:
                cfg["qtd_docs_por_parte"] = int(val)
        elif op == "6":
            cfg["gerar_individual_por_arquivo"] = not cfg.get("gerar_individual_por_arquivo", False)
        elif op == "7":
            cfg["gerar_por_pasta_raiz"] = not cfg.get("gerar_por_pasta_raiz", True)
        elif op == "8":
            cfg["gerar_jsonl"] = not cfg.get("gerar_jsonl", True)
        elif op == "9":
            cfg["mascarar_pii"] = not cfg.get("mascarar_pii", False)
        elif op == "10":
            prompt_w = f"Safe workers (1 to {num_cpus}): " if lang == "en" else f"Workers seguros (1 a {num_cpus}): "
            nw = input(f"\n{prompt_w}").strip()
            if nw.isdigit() and 1 <= int(nw) <= num_cpus:
                cfg["max_workers"] = int(nw)
        elif op == "11":
            cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", True)
        elif op == "12":
            cfg["usar_ocr_em_pdf"] = not cfg.get("usar_ocr_em_pdf", True)
        elif op == "13":
            cfg["usar_ocr_em_imagens_embutidas"] = not cfg.get("usar_ocr_em_imagens_embutidas", True)
        elif op == "14":
            prompt_dpi = "DPI (e.g. 150, 200, 300): " if lang == "en" else "DPI (ex: 150, 200, 300): "
            val = input(f"\n{prompt_dpi}").strip()
            if val.isdigit() and int(val) >= 100:
                cfg["dpi_ocr_pdf"] = int(val)
        elif op == "15":
            prompt_dim = "Minimum dimension px (e.g. 100, 150, 200): " if lang == "en" else "Dimensão mínima em px (ex: 100, 150, 200): "
            val = input(f"\n{prompt_dim}").strip()
            if val.isdigit() and int(val) >= 20:
                cfg["min_dimensao_imagem_ocr"] = int(val)
        elif op == "16":
            prompt_sc = "Minimum confidence score (0.10 to 0.90): " if lang == "en" else "Score mínimo de confiança (0.10 a 0.90): "
            val = input(f"\n{prompt_sc}").strip().replace(",", ".")
            try:
                sc = float(val)
                if 0.05 <= sc <= 0.95:
                    cfg["limiar_confianca_ocr"] = sc
            except ValueError:
                pass
        elif op == "17":
            prompt_out = "Output folder: " if lang == "en" else "Pasta de saída: "
            novo = input(f"\n{prompt_out}").strip().strip('"')
            if novo:
                cfg["pasta_saida"] = str(resolver_caminho(novo))
        elif op == "L":
            cfg["idioma"] = "en" if cfg.get("idioma", "pt") == "pt" else "pt"
            msg_lang = f"Language set to: {'English' if cfg['idioma'] == 'en' else 'Português'}" if cfg["idioma"] == "en" else f"Idioma definido para: {'English' if cfg['idioma'] == 'en' else 'Português'}"
            print(f"\n{Cor.VERDE}{msg_lang}{Cor.RESET}")
            time.sleep(0.8)
        elif op == "R":
            pastas_atuais = cfg.get("pastas_entrada", [])
            lang_atual = cfg.get("idioma", "pt")
            cfg.clear()
            cfg.update(CONFIG_PADRAO.copy())
            cfg["pastas_entrada"] = pastas_atuais
            cfg["idioma"] = lang_atual
            msg_rst = "[OK] Reset to factory default settings!" if lang == "en" else "[OK] Configurações restauradas para o padrão!"
            print(f"\n{Cor.VERDE}{msg_rst}{Cor.RESET}")
            time.sleep(1.0)

        salvar_config(cfg)


def menu_principal():
    cfg = carregar_config()
    lang_init = obter_idioma(cfg)
    cfg["idioma"] = lang_init
    while True:
        limpar_tela()
        lang = obter_idioma(cfg)
        cfg["idioma"] = lang

        qtd_pastas = len(cfg["pastas_entrada"])
        arquivos, _ = listar_arquivos_elegiveis(cfg["pastas_entrada"], cfg["extensoes"]) if qtd_pastas else ([], [])
        stats = obter_estatisticas_arquivos(arquivos) if arquivos else {"total_arquivos": 0, "tamanho_formatado": "0 B"}
        num_cpus = os.cpu_count() or 4
        workers_atual = int(cfg.get("max_workers", min(4, num_cpus)))

        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f"{Cor.TITULO}   {t('titulo_app', cfg)}{Cor.RESET}")
        print(f"{Cor.TITULO}   {t('subtitulo_app', cfg, workers=workers_atual)}{Cor.RESET}")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f" {t('status_atual', cfg)}")
        print(f"  • {t('pastas_monitoradas', cfg):<25}: {Cor.NEGRITO}{qtd_pastas}{Cor.RESET}")
        print(f"  • {t('arquivos_localizados', cfg):<25}: {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET} ({Cor.AMARELO}{stats['tamanho_formatado']}{Cor.RESET})")
        lbl_workers_unit = "safe workers" if lang == "en" else "workers seguros"
        print(f"  • {t('multicore_paralelo', cfg):<25}: {Cor.VERDE + t('ativado', cfg) if cfg.get('multicore', True) else Cor.AMARELO + t('desativado', cfg)} ({workers_atual} {lbl_workers_unit}){Cor.RESET}")
        lbl_pii = "PII Redaction (Privacy)" if lang == "en" else "Privacidade PII (LGPD)"
        val_pii = ("ENABLED" if cfg.get('mascarar_pii', False) else "DISABLED (Keep raw data)") if lang == "en" else ("ATIVADO" if cfg.get('mascarar_pii', False) else "DESATIVADO (Manter dados originais)")
        print(f"  • {lbl_pii:<25}: {Cor.VERDE if cfg.get('mascarar_pii', False) else Cor.AMARELO}{val_pii}{Cor.RESET}")
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
    if "--help" in sys.argv or "-h" in sys.argv:
        print("""
Local Document Consolidator for Microsoft Copilot (v6.0)
Windows 11 | 100% Offline | Multicore Parallel & Local RapidOCR

Usage:
  python app.py [options]

Options:
  --lang en|pt          Set interface language (English or Portuguese)
  --auto                Run non-interactively without prompting
  --no-pii              Disable PII masking (keep raw CPF/CNPJ/Email)
  --redact-pii          Force enable PII masking
  --workers N           Set number of concurrent worker processes
  --help, -h            Show this help message and exit
        """.strip())
        sys.exit(0)

    if "--auto" in sys.argv:
        cfg_auto = carregar_config()
        if "--redact-pii" in sys.argv:
            cfg_auto["mascarar_pii"] = True
        elif "--no-pii" in sys.argv:
            cfg_auto["mascarar_pii"] = False
        if "--workers" in sys.argv:
            try:
                w_idx = sys.argv.index("--workers")
                cfg_auto["max_workers"] = int(sys.argv[w_idx + 1])
            except Exception:
                pass
        executar_pipeline(cfg_auto, pausar_no_final=False)
    else:
        try:
            menu_principal()
        except (KeyboardInterrupt, EOFError):
            print("\nClosed / Encerrado.")
