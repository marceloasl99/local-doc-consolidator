#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidador Local de Documentos para Microsoft Copilot & RAG (v7.0)
Local Document Consolidator for Microsoft Copilot & RAG (v7.0)

Objetivo: ler grandes volumes de documentos heterogêneos (PDF, Word, Excel, PowerPoint,
CSV, texto, HTML, e-mails e imagens) e gerar UMA base Markdown única, organizada e
rastreável (IDs [001], [002]...) para servir de base de conhecimento ao Microsoft Copilot.

100% offline | OCR local (RapidOCR/ONNX) | Multiprocessado com recuperação de falhas
Bilíngue: Português (iniciar_windows.bat / iniciar.ps1) e English (run.bat / start.ps1)
"""

import csv
import email
import hashlib
import io
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import warnings
import zipfile
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool
from datetime import date, datetime
from email import policy as email_policy
from html.parser import HTMLParser
from pathlib import Path

# Suprime avisos de bibliotecas e desativa sugestões de layout do PyMuPDF
os.environ["PYMUPDF_SUGGEST_LAYOUT_ANALYZER"] = "0"
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*fitz.*")
warnings.filterwarnings("ignore", message=".*pymupdf_layout.*")

# Console robusto: nunca quebrar por caractere não suportado pelo terminal
for _fluxo in (sys.stdout, sys.stderr):
    try:
        if hasattr(_fluxo, "isatty") and _fluxo.isatty():
            _fluxo.reconfigure(errors="replace")
        else:
            _fluxo.reconfigure(encoding="utf-8", errors="replace")  # saída redirecionada para log
    except Exception:
        pass

if platform.system() == "Windows":
    os.system("")  # habilita sequências ANSI no console do Windows

VERSAO = "7.0"
# Incrementar sempre que a lógica de extração mudar: invalida o cache antigo automaticamente.
VERSAO_EXTRATOR = "7.0.0"


class Cor:
    TITULO = "\033[96m\033[1m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    VERMELHO = "\033[91m"
    AZUL = "\033[94m"
    MAGENTA = "\033[95m"
    NEGRITO = "\033[1m"
    CINZA = "\033[90m"
    RESET = "\033[0m"


if not (hasattr(sys.stdout, "isatty") and sys.stdout.isatty()):
    # Saída redirecionada (logs, --auto em agendador): sem códigos de cor
    for _attr in ("TITULO", "VERDE", "AMARELO", "VERMELHO", "AZUL", "MAGENTA", "NEGRITO", "CINZA", "RESET"):
        setattr(Cor, _attr, "")


BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
PASTAS_TXT_FILE = BASE_DIR / "pastas.txt"

EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp"}
EXTENSOES_OFFICE = {".docx", ".pptx", ".xlsx", ".xlsm", ".xls"}
EXTENSOES_TEXTO = {".txt", ".md", ".csv", ".log", ".json", ".xml", ".html", ".htm", ".eml"}
EXTENSOES_PDF = {".pdf"}

EXTENSOES_PADRAO = sorted(EXTENSOES_PDF | EXTENSOES_OFFICE | EXTENSOES_TEXTO | EXTENSOES_IMAGEM)

# Lista padrão das versões anteriores: configs que ainda a usam são atualizadas automaticamente.
_EXTENSOES_PADRAO_V6_0 = {
    ".bmp", ".csv", ".docx", ".jpeg", ".jpg", ".md", ".pdf", ".png",
    ".pptx", ".tif", ".tiff", ".txt", ".webp", ".xls", ".xlsx",
}

# Formatos comuns que NÃO são lidos, com dica de conversão (aparecem no relatório).
DICAS_FORMATOS_NAO_SUPORTADOS = {
    ".doc": ("Word 97-2003: salve como .docx", "Word 97-2003: save as .docx"),
    ".ppt": ("PowerPoint 97-2003: salve como .pptx", "PowerPoint 97-2003: save as .pptx"),
    ".msg": ("E-mail do Outlook: salve como .eml ou PDF", "Outlook e-mail: save as .eml or PDF"),
    ".rtf": ("RTF: salve como .docx ou PDF", "RTF: save as .docx or PDF"),
    ".odt": ("LibreOffice Writer: salve como .docx", "LibreOffice Writer: save as .docx"),
    ".ods": ("LibreOffice Calc: salve como .xlsx", "LibreOffice Calc: save as .xlsx"),
    ".odp": ("LibreOffice Impress: salve como .pptx", "LibreOffice Impress: save as .pptx"),
    ".heic": ("Foto do iPhone: converta para .jpg", "iPhone photo: convert to .jpg"),
    ".zip": ("Arquivo compactado: extraia antes", "Archive: extract it first"),
    ".rar": ("Arquivo compactado: extraia antes", "Archive: extract it first"),
    ".7z": ("Arquivo compactado: extraia antes", "Archive: extract it first"),
}

NOMES_TIPO = {
    ".pdf": ("PDF", "PDF"),
    ".docx": ("Word", "Word"),
    ".xlsx": ("Excel", "Excel"),
    ".xlsm": ("Excel", "Excel"),
    ".xls": ("Excel 97-2003", "Excel 97-2003"),
    ".pptx": ("PowerPoint", "PowerPoint"),
    ".csv": ("CSV", "CSV"),
    ".txt": ("Texto", "Text"),
    ".md": ("Markdown", "Markdown"),
    ".log": ("Log", "Log"),
    ".json": ("JSON", "JSON"),
    ".xml": ("XML", "XML"),
    ".html": ("HTML", "HTML"),
    ".htm": ("HTML", "HTML"),
    ".eml": ("E-mail", "Email"),
}

_PROMPT_LINHAS_PT = [
    "Esta base reúne documentos corporativos convertidos para texto. O ÍNDICE DE DOCUMENTOS no início lista "
    "todos os documentos agrupados por pasta, cada um com um ID único entre colchetes ([001], [002]...). "
    "Cada documento começa com o título \"## [ID] Nome do arquivo\", seguido da sua origem, e termina com a linha "
    "\"[Fim do documento [ID]]\".",
    "Regras obrigatórias para responder:",
    "1. Responda única e exclusivamente com base no conteúdo desta base. Nunca deduza, extrapole ou invente. "
    "Se a informação não estiver aqui, responda: \"Esta informação não consta nos documentos fornecidos.\"",
    "2. Cite sempre o ID e o nome do documento de onde tirou cada dado (exemplo: \"conforme [004] Relatorio_Financeiro_2026.pdf\").",
    "3. Se documentos diferentes trouxerem informações conflitantes, aponte os IDs de todos os envolvidos e descreva a divergência.",
    "4. Transcreva números, valores, datas, códigos e nomes exatamente como aparecem no texto original.",
    "5. Trechos indicados como OCR vieram de imagens ou digitalizações e podem conter erros de leitura; "
    "avise quando a resposta depender deles.",
    "6. Quando a resposta envolver arquiteturas, fluxogramas, esquemas ou diagramas técnicos, cite a figura correspondente "
    "(exemplo: \"conforme diagrama em imagens/doc001_p03_fig01.png\") e a página original para que o usuário possa consultar o desenho visual.",
]

_PROMPT_LINHAS_EN = [
    "This knowledge base contains corporate documents converted to text. The DOCUMENT INDEX at the beginning lists "
    "every document grouped by folder, each one with a unique ID in square brackets ([001], [002]...). "
    "Each document starts with the heading \"## [ID] File name\", followed by its source, and ends with the line "
    "\"[End of document [ID]]\".",
    "Mandatory rules for your answers:",
    "1. Answer solely and exclusively from the content of this knowledge base. Never assume, extrapolate or invent. "
    "If the information is not here, answer: \"This information is not present in the provided documents.\"",
    "2. Always cite the ID and name of the document each fact comes from (e.g. \"according to [004] Finance_Report_2026.pdf\").",
    "3. If different documents contain conflicting information, list the IDs of all documents involved and describe the discrepancy.",
    "4. Reproduce numbers, amounts, dates, codes and names exactly as written in the source text.",
    "5. Passages marked as OCR come from images or scans and may contain recognition errors; "
    "say so when your answer depends on them.",
    "6. When an answer involves architectures, flowcharts, schematics or technical diagrams, cite the corresponding file "
    "in the images folder (e.g. \"according to diagram in imagens/doc001_p03_fig01.png\") and the original page so the user can inspect the visual drawing.",
]

PROMPT_SISTEMA_COPILOT_PT = "\n\n".join(_PROMPT_LINHAS_PT)
PROMPT_SISTEMA_COPILOT_EN = "\n\n".join(_PROMPT_LINHAS_EN)

# Rótulos principais da interface (mesmas chaves nos dois idiomas)
I18N = {
    "pt": {
        "titulo_app": f"CONSOLIDADOR LOCAL DE DOCUMENTOS PARA MICROSOFT COPILOT (v{VERSAO})",
        "subtitulo_app": "100% Offline | OCR Local | {workers} processos paralelos",
        "pastas_monitoradas": "Pastas de entrada",
        "arquivos_localizados": "Arquivos encontrados",
        "ativado": "ATIVADO",
        "desativado": "DESATIVADO",
        "sim": "SIM",
        "nao": "NÃO",
        "voltar_menu": "Pressione ENTER para voltar ao menu...",
        "sem_pastas": "[!] Nenhuma pasta de entrada configurada. Use a opção [2] do menu principal.",
        "sem_arquivos": "[!] Nenhum arquivo compatível foi encontrado nas pastas configuradas.",
        "concluido_sucesso": "CONSOLIDAÇÃO CONCLUÍDA EM {tempo:.1f}s",
        "pasta_execucao": "Pasta do resultado",
        "titulo_base": "Base documental consolidada",
        "indice_docs": "ÍNDICE DE DOCUMENTOS",
        "relatorio_titulo": "Relatório de auditoria da consolidação",
        "prompt_opcao": "Escolha uma opção: ",
    },
    "en": {
        "titulo_app": f"LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (v{VERSAO})",
        "subtitulo_app": "100% Offline | Local OCR | {workers} parallel workers",
        "pastas_monitoradas": "Input folders",
        "arquivos_localizados": "Files found",
        "ativado": "ENABLED",
        "desativado": "DISABLED",
        "sim": "YES",
        "nao": "NO",
        "voltar_menu": "Press ENTER to return to the menu...",
        "sem_pastas": "[!] No input folder configured. Use option [2] in the main menu.",
        "sem_arquivos": "[!] No compatible files were found in the configured folders.",
        "concluido_sucesso": "CONSOLIDATION COMPLETED IN {tempo:.1f}s",
        "pasta_execucao": "Result folder",
        "titulo_base": "Consolidated document base",
        "indice_docs": "DOCUMENT INDEX",
        "relatorio_titulo": "Consolidation audit report",
        "prompt_opcao": "Choose an option: ",
    },
}

CONFIG_PADRAO = {
    "idioma": "pt",
    "pastas_entrada": [],
    "pasta_saida": "./saida_copilot",
    # Saídas
    "gerar_arquivo_unico": True,
    "dividir_por_tamanho": True,
    "limite_mb_por_parte": 8.0,
    "dividir_por_quantidade": False,
    "qtd_docs_por_parte": 50,
    "gerar_individual_por_arquivo": False,
    "gerar_por_pasta_raiz": True,
    "gerar_jsonl": False,
    "omitir_duplicados": True,
    "remover_cabecalhos_rodapes": True,
    "usar_pasta_timestamp": True,
    "gerar_arquivo_zip": False,
    # Privacidade e desempenho
    "mascarar_pii": False,
    "max_workers": min(4, os.cpu_count() or 4),
    "multicore": True,
    # OCR
    "usar_ocr_em_pdf": True,
    "usar_ocr_em_imagens_embutidas": True,
    "salvar_diagramas_em_pasta": True,
    "pasta_diagramas": "imagens",
    "min_dimensao_diagrama": 150,
    "dpi_ocr_pdf": 200,
    "min_dimensao_imagem_ocr": 150,
    "limiar_confianca_ocr": 0.50,
    "proporcao_max_aspecto_ocr": 20.0,
    "extensoes": EXTENSOES_PADRAO,
}

_OCR_ENGINE = None

# ---------------------------------------------------------------------------
# Expressões regulares: PII (opcional), mojibake e correções de OCR
# ---------------------------------------------------------------------------
RE_CPF = re.compile(r'\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b')
RE_CNPJ = re.compile(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b')
RE_RG = re.compile(r'(?i)\b(RG|R\.G\.|Identidade|Cédula de Identidade)[\s:]+([0-9A-Za-z.-]{5,14})\b')
RE_RG_FORMATADO = re.compile(r'\b\d{1,2}\.\d{3}\.\d{3}-[\dXx]\b')
RE_CEP = re.compile(r'(?i)\b(?:CEP[\s:]+)?\d{5}-\d{3}\b')
RE_EMAIL = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
RE_TELEFONE = re.compile(r'(?:\+?55\s?)?(?:\(\d{2}\)|\b\d{2}\b)[-.\s]?(?:9\d{4}|\d{4})[-.\s]?\d{4}\b')
RE_ENDERECO = re.compile(r'(?i)\b(?:Rua|Avenida|Av\.|Travessa|Alameda|Rodovia|Estrada|Praça|Praca)\s+[^,\n\r]{3,45},\s*(?:n[º°]?\s*)?\d+[A-Za-z0-9\s\-/]*')
RE_NOME_CONTEXTUAL = re.compile(r'(?i)\b(Nome|Titular|Cliente|Favorecido|Paciente|Colaborador|Funcion[aá]rio)[\s:]+([A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+(?:[ \t]+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+)+)')

# Mojibake: bytes UTF-8 lidos como CP1252. "Ã"/"Â" seguidos de um caractere que o CP1252 produz
# a partir de um byte de continuação UTF-8 (0x80-0xBF). Não casa com texto legítimo como "SÃO".
_CP1252_CONTINUACAO = (
    "\u0080-¿ŒœŠšŸŽžƒˆ˜"
    "–—‘-‚“-„†-•…‰‹›€™"
)
_NUCLEO_MOJIBAKE = f"(?:[ÂÃ][{_CP1252_CONTINUACAO}]|â€[{_CP1252_CONTINUACAO}])"
PADRAO_MOJIBAKE = re.compile(_NUCLEO_MOJIBAKE)
_TOKEN_MOJIBAKE = re.compile(r"\S*" + _NUCLEO_MOJIBAKE + r"\S*")

# Correção de palavras quebradas por OCR. Os fragmentos precisam começar em minúscula e a junção
# só ocorre com espaço simples na mesma linha, para não colar siglas ("Comércio LTDA") nem
# palavras reais ("que lhe", "sobre proteção", "nossa mente").
PADRAO_FONOTATICO_OCR = re.compile(r'\b([A-Za-zÀ-ÿ]{2,}) (nt|nd|mp|mb|rt|rd|lt|ld|nh|rr|ss|ç)([a-zà-ÿ]+)\b')
PADRAO_SUFIXOS_OCR = re.compile(r'\b([A-Za-zÀ-ÿ]{3,}) (cao|ção|ções|coes|mento|mentos|tivo|tivos|tiva|tivas|ficação|ficações)\b')
PADRAO_PREFIXOS_OCR = re.compile(r'\b(micro|macro|multi|tele) ([a-zà-ÿ]{4,})\b')


# ---------------------------------------------------------------------------
# Utilidades de idioma e formatação
# ---------------------------------------------------------------------------
def tx(lang: str, pt: str, en: str) -> str:
    """Seleciona o texto no idioma ativo."""
    return en if lang == "en" else pt


def idioma_linha_comando() -> str | None:
    """Idioma pedido na linha de comando (--lang en|pt, --english, --pt...), se houver."""
    argv = sys.argv
    if "--lang" in argv:
        idx = argv.index("--lang")
        if idx + 1 < len(argv) and argv[idx + 1].lower() in ("en", "pt"):
            return argv[idx + 1].lower()
    if "--english" in argv or "--en" in argv:
        return "en"
    if "--portugues" in argv or "--pt" in argv:
        return "pt"
    return None


def obter_idioma(cfg: dict) -> str:
    """Idioma ativo: o da configuração da sessão; se não houver, o da linha de comando; senão 'pt'.

    A linha de comando é aplicada uma única vez na inicialização (ver aplicar_argumentos),
    de modo que a troca de idioma pelo menu continue funcionando mesmo quando o launcher
    passa --lang.
    """
    if cfg and cfg.get("idioma") in ("pt", "en"):
        return cfg["idioma"]
    return idioma_linha_comando() or "pt"


def t(chave: str, cfg: dict, **kwargs) -> str:
    lang = obter_idioma(cfg)
    texto = I18N.get(lang, I18N["pt"]).get(chave, I18N["pt"].get(chave, chave))
    if kwargs:
        try:
            return texto.format(**kwargs)
        except Exception:
            return texto
    return texto


def fmt_int(n: int, lang: str) -> str:
    s = f"{int(n):,}"
    return s.replace(",", ".") if lang == "pt" else s


def agora_fmt(lang: str) -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M" if lang == "pt" else "%Y-%m-%d %H:%M")


def formatar_tamanho_bytes(tam_bytes: int) -> str:
    if tam_bytes < 1024:
        return f"{tam_bytes} B"
    if tam_bytes < 1024 * 1024:
        return f"{tam_bytes / 1024:.1f} KB"
    if tam_bytes < 1024 * 1024 * 1024:
        return f"{tam_bytes / (1024 * 1024):.2f} MB"
    return f"{tam_bytes / (1024 * 1024 * 1024):.2f} GB"


def nome_tipo(ext: str, lang: str) -> str:
    if ext in EXTENSOES_IMAGEM:
        return tx(lang, "Imagem", "Image")
    par = NOMES_TIPO.get(ext)
    if par:
        return par[1] if lang == "en" else par[0]
    return ext.lstrip(".").upper() or "?"


def limpar_tela():
    if "--auto" in sys.argv:
        return
    if platform.system() == "Windows":
        os.system("cls")
    else:
        try:
            sys.stdout.write("\033[2J\033[H")
            sys.stdout.flush()
        except Exception:
            os.system("clear")


def resolver_caminho(p_str: str) -> Path:
    p = Path(str(p_str).strip().strip('"')).expanduser()
    if not p.is_absolute():
        p = BASE_DIR / p
    try:
        return p.resolve()
    except Exception:
        return p.absolute()


def _chave_caminho(p: Path) -> str:
    return os.path.normcase(str(p))


def normalizar_unicode(texto: str) -> str:
    """Normaliza Unicode para NFC (evita acentos decompostos NFD de macOS/Canva)."""
    if not texto:
        return ""
    return unicodedata.normalize("NFC", texto)


def chave_ordenacao_natural(texto: str) -> list:
    """Chave de ordenação natural numérica (ex.: 1_Aula, 2_Aula, ..., 10_Aula)."""
    s = unicodedata.normalize("NFC", str(texto)).lower()
    return [(0, int(p)) if p.isdigit() else (1, p) for p in re.split(r"(\d+)", s) if p]


def linha(cor: str = Cor.TITULO, char: str = "=") -> str:
    return f"{cor}{char * 76}{Cor.RESET}"


def perguntar(msg: str) -> str:
    try:
        return input(msg)
    except EOFError:
        return "0"


def pausar(lang: str):
    perguntar(f"\n{tx(lang, 'Pressione ENTER para continuar...', 'Press ENTER to continue...')}")


# ---------------------------------------------------------------------------
# Limpeza de texto: PII, mojibake, OCR, Markdown
# ---------------------------------------------------------------------------
def redigir_pii(texto: str) -> str:
    """Anonimiza dados sensíveis (CPF, RG, CNPJ, e-mail, telefone, endereço, nomes com rótulo)."""
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
    Corrige mojibake clássico (UTF-8 lido como CP1252): 'ObservaÃ§Ã£o' -> 'Observação'.
    Tenta o texto inteiro, depois linha a linha e por fim palavra a palavra; trechos que não
    puderem ser recodificados com segurança ficam como estão.
    Retorna (texto_corrigido, total_correcoes).
    """
    if not texto:
        return texto, 0

    texto_trabalho = texto.lstrip("﻿")
    total = len(PADRAO_MOJIBAKE.findall(texto_trabalho))
    if not total:
        return texto_trabalho, 0

    try:
        candidato = texto_trabalho.encode("cp1252").decode("utf-8")
        if not PADRAO_MOJIBAKE.search(candidato):
            return candidato, total
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass

    correcoes = 0

    def substituir_token(match):
        nonlocal correcoes
        tok = match.group(0)
        try:
            novo = tok.encode("cp1252").decode("utf-8")
            correcoes += len(PADRAO_MOJIBAKE.findall(tok))
            return novo
        except (UnicodeEncodeError, UnicodeDecodeError):
            return tok

    saida = []
    for ln in texto_trabalho.splitlines(keepends=True):
        if not PADRAO_MOJIBAKE.search(ln):
            saida.append(ln)
            continue
        try:
            saida.append(ln.encode("cp1252").decode("utf-8"))
            correcoes += len(PADRAO_MOJIBAKE.findall(ln))
            continue
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
        saida.append(_TOKEN_MOJIBAKE.sub(substituir_token, ln))

    return "".join(saida), correcoes


def corrigir_espacos_ocr(texto: str) -> tuple[str, int]:
    """
    Remove espaços espúrios inseridos pelo OCR no meio de palavras.
    Exemplos: 'Aute nticacao' -> 'Autenticacao', 'docu mento' -> 'documento'.
    Retorna (texto_corrigido, total_correcoes).
    """
    if not texto:
        return texto, 0

    total = 0

    def aplicar(pat, t_in, repl):
        nonlocal total
        novo, n = pat.subn(repl, t_in)
        total += n
        return novo

    atual = texto
    for _ in range(2):
        anterior = atual
        atual = aplicar(PADRAO_FONOTATICO_OCR, atual, r'\1\2\3')
        atual = aplicar(PADRAO_SUFIXOS_OCR, atual, r'\1\2')
        atual = aplicar(PADRAO_PREFIXOS_OCR, atual, r'\1\2')
        if atual == anterior:
            break
    return atual, total


def escapar_celula(valor) -> str:
    s = "" if valor is None else str(valor)
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r" {2,}", " ", s).strip()
    return s.replace("|", "\\|")


def tabela_para_markdown(linhas: list[list]) -> str:
    """Converte uma matriz de células em tabela Markdown (remove linhas/colunas totalmente vazias)."""
    if not linhas:
        return ""
    matriz = [[escapar_celula(c) for c in ln] for ln in linhas if ln is not None]
    matriz = [ln for ln in matriz if any(ln)]
    if not matriz:
        return ""
    n_cols = max(len(ln) for ln in matriz)
    for ln in matriz:
        ln.extend([""] * (n_cols - len(ln)))
    cols_usadas = [c for c in range(n_cols) if any(ln[c] for ln in matriz)]
    if not cols_usadas:
        return ""
    matriz = [[ln[c] for c in cols_usadas] for ln in matriz]

    cabecalho = matriz[0]
    saida = [
        "| " + " | ".join(cabecalho) + " |",
        "| " + " | ".join(["---"] * len(cols_usadas)) + " |",
    ]
    for row in matriz[1:]:
        saida.append("| " + " | ".join(row) + " |")
    return "\n".join(saida)


def bloco_codigo(texto: str, linguagem: str = "") -> str:
    cerca = "```"
    while cerca in texto:
        cerca += "`"
    return f"{cerca}{linguagem}\n{texto.rstrip()}\n{cerca}"


def rebaixar_titulos_markdown(texto: str, niveis: int = 2) -> str:
    """Rebaixa títulos de arquivos .md para não competirem com os títulos da base (## [ID])."""
    saida = []
    em_codigo = False
    for ln in texto.splitlines():
        if ln.lstrip().startswith(("```", "~~~")):
            em_codigo = not em_codigo
        m = None if em_codigo else re.match(r"^(#{1,6})(\s+.*)$", ln)
        if m:
            nivel = min(len(m.group(1)) + niveis, 6)
            saida.append("#" * nivel + m.group(2))
        else:
            saida.append(ln)
    return "\n".join(saida)


def normalizar_espacos(texto: str) -> str:
    """Remove espaços no fim das linhas e excesso de linhas em branco, normalizando Unicode NFC."""
    if not texto:
        return ""
    texto = normalizar_unicode(texto).replace("\r\n", "\n").replace("\r", "\n").replace(" ", " ")
    texto = "\n".join(ln.rstrip() for ln in texto.split("\n"))
    return re.sub(r"\n{3,}", "\n\n", texto).strip()


def normalizar_texto_ocr(texto: str) -> str:
    """Une palavras hifenizadas no fim da linha e remove linhas vazias redundantes."""
    if not texto:
        return ""
    texto = re.sub(r"(\w)-\n(?=[a-zà-ÿ])", r"\1", texto)
    return normalizar_espacos(texto)


# ---------------------------------------------------------------------------
# Leitura de arquivos de texto
# ---------------------------------------------------------------------------
def ler_texto_com_fallback(caminho: Path) -> str:
    dados = caminho.read_bytes()
    texto = None
    for enc in ("utf-8-sig", "utf-16") if dados[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = dados.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        texto = dados.decode("utf-8", errors="replace")
    texto_corrigido, _ = reparar_mojibake(texto)
    return normalizar_unicode(texto_corrigido)


def extrair_csv_estruturado(caminho: Path) -> str:
    """Lê CSV detectando o delimitador e converte em tabela Markdown."""
    conteudo_bruto = ler_texto_com_fallback(caminho)
    if not conteudo_bruto.strip():
        return ""
    try:
        csv.field_size_limit(2**31 - 1)
    except Exception:
        pass
    try:
        dialeto = csv.Sniffer().sniff(conteudo_bruto[:8192], delimiters=";,\t|")
        delimitador = dialeto.delimiter
    except Exception:
        primeira = conteudo_bruto.splitlines()[0] if conteudo_bruto.splitlines() else ""
        contagens = {d: primeira.count(d) for d in (";", ",", "\t", "|")}
        delimitador = max(contagens, key=contagens.get) if any(contagens.values()) else ","
    try:
        linhas = list(csv.reader(conteudo_bruto.splitlines(), delimiter=delimitador))
        md_tab = tabela_para_markdown(linhas)
        if md_tab:
            return md_tab
    except Exception:
        pass
    return conteudo_bruto


class _ConversorHTML(HTMLParser):
    """Conversor HTML -> Markdown simples (títulos, parágrafos, listas e tabelas)."""

    IGNORAR = {"script", "style", "noscript", "template", "svg", "head"}
    BLOCOS = {"p", "div", "section", "article", "header", "footer", "main", "aside", "nav",
              "blockquote", "pre", "ul", "ol", "table", "form", "hr", "dl", "dt", "dd", "figure"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.partes: list[str] = []
        self.ignorar = 0
        self.titulo_pagina = ""
        self._em_title = False
        self.linha_tab: list[str] | None = None
        self.linhas_tabela = 0

    def handle_starttag(self, tag, attrs):
        if tag == "title":
            self._em_title = True
        if tag in self.IGNORAR:
            self.ignorar += 1
            return
        if self.ignorar:
            return
        if re.fullmatch(r"h[1-6]", tag):
            self.partes.append("\n\n" + "#" * min(int(tag[1]) + 2, 6) + " ")
        elif tag == "li":
            self.partes.append("\n- ")
        elif tag == "br":
            if self.linha_tab is not None and self.linha_tab:
                self.linha_tab[-1] += " "
            else:
                self.partes.append("\n")
        elif tag == "table":
            self.linhas_tabela = 0
            self.partes.append("\n\n")
        elif tag == "tr":
            self.linha_tab = []
        elif tag in ("td", "th"):
            if self.linha_tab is not None:
                self.linha_tab.append("")
        elif tag in self.BLOCOS:
            self.partes.append("\n\n")

    def handle_endtag(self, tag):
        if tag == "title":
            self._em_title = False
        if tag in self.IGNORAR:
            self.ignorar = max(0, self.ignorar - 1)
            return
        if self.ignorar:
            return
        if tag == "tr" and self.linha_tab is not None:
            celulas = [escapar_celula(c) for c in self.linha_tab]
            if any(celulas):
                self.partes.append("\n| " + " | ".join(celulas) + " |")
                if self.linhas_tabela == 0:
                    self.partes.append("\n| " + " | ".join(["---"] * len(celulas)) + " |")
                self.linhas_tabela += 1
            self.linha_tab = None
        elif re.fullmatch(r"h[1-6]", tag) or tag in self.BLOCOS:
            self.partes.append("\n\n")

    def handle_data(self, data):
        if self._em_title:
            self.titulo_pagina += data.strip()
        if self.ignorar:
            return
        texto = re.sub(r"\s+", " ", data)
        if self.linha_tab is not None and self.linha_tab:
            self.linha_tab[-1] += texto
        elif self.linha_tab is None:
            self.partes.append(texto)

    def resultado(self) -> str:
        bruto = "".join(self.partes)
        linhas = [ln.strip() for ln in bruto.split("\n")]
        return normalizar_espacos("\n".join(linhas))


def html_para_markdown(conteudo_html: str) -> tuple[str, str]:
    conv = _ConversorHTML()
    try:
        conv.feed(conteudo_html)
        conv.close()
    except Exception:
        pass
    return conv.resultado(), conv.titulo_pagina


def extrair_email(
    caminho: Path,
    lang: str,
    cfg: dict | None = None,
    pasta_cache: Path | None = None,
    permitir_anexos: bool = True,
) -> tuple[str, dict]:
    msg = email.message_from_bytes(caminho.read_bytes(), policy=email_policy.default)
    meta: dict = {}
    cab = []
    for campo, rot_pt, rot_en in (("From", "De", "From"), ("To", "Para", "To"), ("Cc", "Cc", "Cc"),
                                  ("Date", "Data", "Date"), ("Subject", "Assunto", "Subject")):
        valor = msg.get(campo)
        if valor:
            cab.append(f"- **{tx(lang, rot_pt, rot_en)}:** {escapar_celula(normalizar_unicode(str(valor)))}")
            if campo == "Subject":
                meta["title"] = normalizar_unicode(str(valor))
    corpo = ""
    try:
        parte = msg.get_body(preferencelist=("plain", "html"))
        if parte is not None:
            corpo = parte.get_content()
            if parte.get_content_type() == "text/html":
                corpo, _ = html_para_markdown(corpo)
    except Exception:
        corpo = ""

    partes = ["\n".join(cab)] if cab else []
    if corpo.strip():
        partes.append(normalizar_espacos(corpo))

    anexos_nao_extraidos = []
    diagramas_anexos: list[dict] = []
    ocr_total = 0
    diag_ocr_total = {"ok": 0, "empty": 0, "error": 0}
    scores_anexos: list[float] = []
    exts_anexo = set(EXTENSOES_PADRAO) - {".eml"}

    try:
        anexos_iter = list(msg.iter_attachments())
    except Exception:
        anexos_iter = []

    if anexos_iter:
        with tempfile.TemporaryDirectory() as tmp_dir:
            for idx, anexo in enumerate(anexos_iter, start=1):
                nome = normalizar_unicode(anexo.get_filename() or "")
                if not nome:
                    continue
                ext_a = Path(nome).suffix.lower()
                if not permitir_anexos or ext_a not in exts_anexo:
                    anexos_nao_extraidos.append(nome)
                    continue
                payload = None
                try:
                    payload = anexo.get_payload(decode=True)
                    if payload is None:
                        conteudo_a = anexo.get_content()
                        if isinstance(conteudo_a, bytes):
                            payload = conteudo_a
                        elif isinstance(conteudo_a, str):
                            payload = conteudo_a.encode("utf-8", errors="replace")
                except Exception:
                    payload = None
                if not payload:
                    anexos_nao_extraidos.append(nome)
                    continue
                arq_tmp = Path(tmp_dir) / f"{idx:02d}_{sanitizar_nome_arquivo(nome)}"
                if arq_tmp.suffix.lower() != ext_a:
                    arq_tmp = arq_tmp.with_suffix(ext_a)
                try:
                    arq_tmp.write_bytes(payload)
                    txt_a, ocr_a, _pags_a, meta_a, _rag_a, conf_a, diag_a = _despachar_extrator(
                        arq_tmp, cfg or CONFIG_PADRAO, pasta_cache, permitir_anexos_eml=False
                    )
                    ocr_total += ocr_a
                    for k in diag_ocr_total:
                        diag_ocr_total[k] += (diag_a or {}).get(k, 0)
                    if conf_a:
                        scores_anexos.append(conf_a)
                    for d_img in (meta_a or {}).get("diagramas", []):
                        diagramas_anexos.append(d_img)
                    if (txt_a or "").strip():
                        partes.append(f"### {tx(lang, 'Anexo', 'Attachment')}: {nome}\n\n{txt_a.strip()}")
                    else:
                        anexos_nao_extraidos.append(nome)
                except Exception:
                    anexos_nao_extraidos.append(nome)

    if anexos_nao_extraidos:
        partes.append(f"**{tx(lang, 'Anexos (não extraídos)', 'Attachments (not extracted)')}:** " + ", ".join(anexos_nao_extraidos))
    if diagramas_anexos:
        meta["diagramas"] = diagramas_anexos
    if ocr_total or any(diag_ocr_total.values()):
        meta["_ocr_count"] = ocr_total
        meta["_diag_ocr"] = diag_ocr_total
        meta["_confianca"] = (sum(scores_anexos) / len(scores_anexos)) if scores_anexos else 0.0
    return "\n\n".join(partes), meta


def extrair_texto_generico(
    caminho: Path,
    ext: str,
    lang: str,
    cfg: dict | None = None,
    pasta_cache: Path | None = None,
    permitir_anexos_eml: bool = True,
) -> tuple[str, dict]:
    """Lê .txt, .md, .log, .json, .xml, .html e .eml."""
    if ext == ".eml":
        return extrair_email(caminho, lang, cfg=cfg, pasta_cache=pasta_cache, permitir_anexos=permitir_anexos_eml)
    texto = ler_texto_com_fallback(caminho)
    meta = {}
    if ext in (".html", ".htm"):
        md, titulo = html_para_markdown(texto)
        if titulo:
            meta["title"] = titulo
        return md, meta
    if ext == ".md":
        return rebaixar_titulos_markdown(normalizar_espacos(texto)), meta
    if ext == ".json":
        try:
            texto = json.dumps(json.loads(texto), indent=2, ensure_ascii=False)
        except Exception:
            pass
        return bloco_codigo(texto, "json"), meta
    if ext == ".xml":
        return bloco_codigo(texto, "xml"), meta
    return normalizar_espacos(texto), meta


# ---------------------------------------------------------------------------
# OCR local (RapidOCR / ONNX)
# ---------------------------------------------------------------------------
def obter_ocr_worker():
    """Inicializa (uma vez por processo) o motor RapidOCR."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            try:
                from rapidocr_onnxruntime import RapidOCR
            except ImportError:
                from rapidocr import RapidOCR
            _OCR_ENGINE = RapidOCR()
        except Exception:
            _OCR_ENGINE = False
    return _OCR_ENGINE if _OCR_ENGINE is not False else None


def _itens_resultado_ocr(resultado) -> list[tuple]:
    """Normaliza a saída das duas APIs do RapidOCR para [(caixa, texto, score)]."""
    if resultado is None:
        return []
    # rapidocr_onnxruntime: (lista, tempo) | rapidocr >= 2: objeto com boxes/txts/scores
    if isinstance(resultado, tuple) and len(resultado) == 2 and not hasattr(resultado, "txts"):
        resultado = resultado[0]
    if hasattr(resultado, "txts"):
        txts = list(resultado.txts or [])
        boxes = list(resultado.boxes) if getattr(resultado, "boxes", None) is not None else [None] * len(txts)
        scores = list(resultado.scores) if getattr(resultado, "scores", None) is not None else [1.0] * len(txts)
        return list(zip(boxes, txts, scores))
    itens = []
    for item in resultado or []:
        try:
            if len(item) >= 3:
                itens.append((item[0], item[1], item[2]))
            elif len(item) == 2:
                itens.append((item[0], item[1], 1.0))
        except TypeError:
            continue
    return itens


def reconstruir_layout_ocr(itens: list[tuple]) -> str:
    """
    Reconstrói linhas e parágrafos a partir das caixas do OCR: caixas na mesma altura viram
    uma linha (ordenadas da esquerda para a direita) e espaços verticais grandes viram
    quebras de parágrafo. Preserva a estrutura de recibos, formulários e tabelas simples.
    """
    caixas = []
    for box, txt, _sc in itens:
        txt = str(txt).strip()
        if not txt:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
            caixas.append((min(xs), min(ys), max(xs), max(ys), txt))
        except Exception:
            return "\n".join(str(i[1]).strip() for i in itens if str(i[1]).strip())
    if not caixas:
        return ""

    alturas = sorted(c[3] - c[1] for c in caixas)
    h_med = max(alturas[len(alturas) // 2], 1.0)
    caixas.sort(key=lambda c: ((c[1] + c[3]) / 2, c[0]))

    linhas = []  # [y_centro, y0, y1, [caixas]]
    for c in caixas:
        yc = (c[1] + c[3]) / 2
        if linhas and abs(yc - linhas[-1][0]) <= 0.5 * h_med:
            atual = linhas[-1]
            atual[3].append(c)
            atual[0] = sum((k[1] + k[3]) / 2 for k in atual[3]) / len(atual[3])
            atual[1] = min(atual[1], c[1])
            atual[2] = max(atual[2], c[3])
        else:
            linhas.append([yc, c[1], c[3], [c]])

    saida = []
    y_ant = None
    for _yc, y0, y1, cx in linhas:
        cx.sort(key=lambda k: k[0])
        partes = [cx[0][4]]
        for anterior, atual in zip(cx, cx[1:]):
            distancia = atual[0] - anterior[2]
            partes.append(("   " if distancia > 2 * h_med else " ") + atual[4])
        if y_ant is not None and (y0 - y_ant) > 1.5 * h_med:
            saida.append("")
        saida.append("".join(partes))
        y_ant = y1
    return "\n".join(saida)


def _preparar_imagem_ocr(img_bytes: bytes, min_dim: int, proporcao_max_aspecto: float):
    """Abre, valida e normaliza a imagem para o OCR. Retorna (png_bytes | None, motivo)."""
    from PIL import Image, ImageOps, UnidentifiedImageError
    try:
        im = Image.open(io.BytesIO(img_bytes))
        im.load()
    except (UnidentifiedImageError, OSError, ValueError):
        return None, "unsupported"
    w, h = im.size
    if max(w, h) < min_dim or min(w, h) < 20:
        return None, "small"
    if max(w, h) / max(min(w, h), 1) > proporcao_max_aspecto:
        return None, "aspect"
    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    # Transparência sobre fundo branco (texto preto em PNG transparente não pode virar preto no preto)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        fundo = Image.new("RGB", im.size, (255, 255, 255))
        fundo.paste(im, mask=im.split()[-1])
        im = fundo
    elif im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    w, h = im.size
    if max(w, h) > 2400:
        r = 2400 / max(w, h)
        im = im.resize((max(1, int(w * r)), max(1, int(h * r))), Image.Resampling.LANCZOS)
    elif max(w, h) < 600:
        im = im.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue(), "ok"


def _medir_residuo_visual(png_bytes: bytes, boxes: list) -> tuple[float, float]:
    """Mede a densidade de bordas não-textuais após mascarar as caixas detectadas pelo OCR."""
    from PIL import Image, ImageDraw, ImageFilter, ImageStat

    im_ocr = Image.open(io.BytesIO(png_bytes)).convert("L")
    w_ocr, h_ocr = im_ocr.size
    fundo = int(ImageStat.Stat(im_ocr).median[0])
    im_sem_texto = im_ocr.copy()
    draw = ImageDraw.Draw(im_sem_texto)
    for box in boxes:
        if not box:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
            x0 = max(0, int(min(xs)) - 4)
            y0 = max(0, int(min(ys)) - 4)
            x1 = min(w_ocr, int(max(xs)) + 4)
            y1 = min(h_ocr, int(max(ys)) + 4)
            if x1 > x0 and y1 > y0:
                draw.rectangle([x0, y0, x1, y1], fill=fundo)
        except Exception:
            continue

    bordas = im_sem_texto.filter(ImageFilter.FIND_EDGES)
    if w_ocr > 24 and h_ocr > 24:
        bordas = bordas.crop((4, 4, w_ocr - 4, h_ocr - 4))
    stat = ImageStat.Stat(bordas)
    return float(stat.mean[0]), float(stat.stddev[0])


def _analisar_imagem_completa(
    img_bytes: bytes,
    min_dim_ocr: int = 150,
    min_dim_diag: int = 150,
    limiar_confianca: float = 0.50,
    prop_max_ocr: float = 20.0,
    prop_max_diag: float = 12.0,
    analisar_diagrama: bool = True,
) -> dict:
    """
    Executa preparação + RapidOCR uma ÚNICA vez sobre a imagem e devolve tanto o texto OCR
    quanto a classificação de diagrama vs. imagem de texto puro.
    """
    from PIL import Image, UnidentifiedImageError

    vazio = {
        "texto": "", "conf": 0.0, "estado": "empty_or_low_conf", "erro": None,
        "eh_diag": False, "motivo_diag": "invalid", "rotulos": [], "png_bytes": None,
        "dims": (0, 0), "res_std": 0.0, "cobertura_texto": 0.0,
    }
    if not img_bytes:
        return vazio

    try:
        im = Image.open(io.BytesIO(img_bytes))
        im.load()
        w, h = im.size
    except (UnidentifiedImageError, OSError, ValueError) as e:
        vazio["estado"] = "empty_or_low_conf"
        vazio["motivo_diag"] = "unsupported"
        return vazio

    razao_aspecto = max(w, h) / max(min(w, h), 1)
    elegivel_ocr = max(w, h) >= min_dim_ocr and min(w, h) >= 20 and razao_aspecto <= prop_max_ocr
    elegivel_diag = (
        analisar_diagrama
        and razao_aspecto <= prop_max_diag
        and max(w, h) >= min_dim_diag
        and min(w, h) >= 50
    )

    motivo_diag = "decoracao_ou_foto"
    if analisar_diagrama:
        if razao_aspecto > prop_max_diag:
            motivo_diag = "divider_or_banner"
        elif max(w, h) < min_dim_diag or min(w, h) < 50:
            motivo_diag = "small_icon"

    if not elegivel_ocr and not elegivel_diag:
        vazio["dims"] = (w, h)
        vazio["motivo_diag"] = motivo_diag
        return vazio

    min_prep = min(min_dim_ocr, min_dim_diag) if (elegivel_ocr and elegivel_diag) else (min_dim_ocr if elegivel_ocr else min_dim_diag)
    prop_prep = max(prop_max_ocr, prop_max_diag)
    try:
        png_bytes, _ = _preparar_imagem_ocr(img_bytes, min_dim=min_prep, proporcao_max_aspecto=prop_prep)
    except Exception as e:
        vazio["estado"] = "error"
        vazio["erro"] = str(e)
        vazio["dims"] = (w, h)
        return vazio

    if not png_bytes:
        vazio["dims"] = (w, h)
        vazio["motivo_diag"] = motivo_diag
        return vazio

    engine = obter_ocr_worker()
    if not engine:
        vazio["estado"] = "error"
        vazio["erro"] = "OCR engine (rapidocr-onnxruntime) unavailable"
        vazio["png_bytes"] = png_bytes
        vazio["dims"] = (w, h)
        return vazio

    try:
        itens = _itens_resultado_ocr(engine(png_bytes))
    except Exception as e:
        vazio["estado"] = "error"
        vazio["erro"] = str(e)
        vazio["png_bytes"] = png_bytes
        vazio["dims"] = (w, h)
        return vazio

    validos, scores, boxes_validas, rotulos = [], [], [], []
    area_texto = 0.0
    for box, txt, sc in itens:
        try:
            sc = float(sc)
        except (TypeError, ValueError):
            sc = 1.0
        txt_str = str(txt).strip()
        if txt_str and sc >= limiar_confianca:
            validos.append((box, txt, sc))
            scores.append(sc)
            rotulos.append(txt_str)
            boxes_validas.append(box)
            try:
                xs = [float(p[0]) for p in box]
                ys = [float(p[1]) for p in box]
                area_texto += max(0.0, max(xs) - min(xs)) * max(0.0, max(ys) - min(ys))
            except Exception:
                pass

    conf_media = (sum(scores) / len(scores)) if scores else 0.0
    texto = ""
    estado = "empty_or_low_conf"
    if validos and elegivel_ocr:
        texto = reconstruir_layout_ocr(validos)
        texto = normalizar_texto_ocr(texto)
        texto, _ = corrigir_espacos_ocr(texto)
        texto, _ = reparar_mojibake(texto)
        if texto.strip():
            estado = "ok"

    eh_diag = False
    res_mean = 0.0
    res_std = 0.0
    cobertura_texto = 0.0
    if elegivel_diag:
        try:
            with Image.open(io.BytesIO(png_bytes)) as im_png:
                w_ocr, h_ocr = im_png.size
            cobertura_texto = area_texto / max(w_ocr * h_ocr, 1)
            res_mean, res_std = _medir_residuo_visual(png_bytes, boxes_validas)
        except Exception:
            res_mean, res_std = 0.0, 0.0

        qtd_linhas = len(rotulos)
        if cobertura_texto >= 0.40 and qtd_linhas >= 16 and res_mean < 6.0:
            eh_diag = False
            motivo_diag = "scanned_text_page"
        elif qtd_linhas >= 1 and (
            (res_mean < 4.5 and res_std < 18.0 and conf_media >= 0.60)
            or (cobertura_texto >= 0.30 and qtd_linhas >= 10 and res_mean < 8.0)
        ):
            # Imagem traduzível em puro texto (comprovantes, recibos, prints limpos de parágrafos)
            eh_diag = False
            motivo_diag = "pure_text_image"
        elif (res_std >= 8.0 or res_mean >= 2.0) and min(w, h) >= 120:
            eh_diag = True
            motivo_diag = "diagrama_com_rotulos" if qtd_linhas >= 1 else "esquema_visual"
        else:
            eh_diag = False
            motivo_diag = "decoracao_ou_foto"

    return {
        "texto": texto,
        "conf": conf_media,
        "estado": estado,
        "erro": None,
        "eh_diag": eh_diag,
        "motivo_diag": motivo_diag,
        "rotulos": rotulos,
        "png_bytes": png_bytes,
        "dims": (w, h),
        "res_std": res_std,
        "cobertura_texto": cobertura_texto,
    }


def executar_ocr_em_bytes(
    img_bytes: bytes,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
) -> tuple[str, float, str, str | None]:
    """
    Executa OCR em bytes de imagem. Retorna (texto, confianca_media, estado, erro).
    estado: 'ok' | 'empty_or_low_conf' | 'error'.
    """
    info = _analisar_imagem_completa(
        img_bytes,
        min_dim_ocr=min_dim,
        limiar_confianca=limiar_confianca,
        prop_max_ocr=proporcao_max_aspecto,
        analisar_diagrama=False,
    )
    return info["texto"], info["conf"], info["estado"], info["erro"]


def classificar_e_extrair_diagrama(
    img_bytes: bytes,
    min_dim_diag: int = 150,
    prop_max: float = 12.0,
    limiar_conf: float = 0.40,
) -> tuple[bool, str, list[str], bytes | None, tuple[int, int]]:
    """
    Avalia se uma imagem é um diagrama/esquema/gráfico técnico (não traduzível em puro texto).
    Retorna: (eh_diagrama, motivo, rotulos_ocr, png_bytes, (largura, altura))
    """
    info = _analisar_imagem_completa(
        img_bytes,
        min_dim_ocr=min_dim_diag,
        min_dim_diag=min_dim_diag,
        limiar_confianca=limiar_conf,
        prop_max_ocr=prop_max,
        prop_max_diag=prop_max,
        analisar_diagrama=True,
    )
    return info["eh_diag"], info["motivo_diag"], info["rotulos"], info["png_bytes"], info["dims"]


class _ContextoOCR:
    """Acumula estatísticas de OCR de um documento, detecta diagramas (passe único) e evita repetições."""

    def __init__(
        self,
        min_dim: int,
        limiar: float,
        prop_max: float,
        salvar_diagramas: bool = False,
        min_dim_diagrama: int = 150,
        pasta_cache: Path | None = None,
    ):
        self.min_dim = min_dim
        self.limiar = limiar
        self.prop_max = prop_max
        self.salvar_diagramas = salvar_diagramas
        self.min_dim_diagrama = min_dim_diagrama
        self.pasta_cache = Path(pasta_cache) if pasta_cache else None
        self.diag = {"ok": 0, "empty": 0, "error": 0}
        self.scores: list[float] = []
        self.qtd = 0
        self._cache: dict[str, tuple] = {}
        self.erros: list[str] = []
        self.diagramas: list[dict] = []
        self.contador_figuras = 0
        self._hashes_vistos: set[str] = set()
        self._textos_ocr_vistos: list[set[str]] = []

    def _registrar_diagrama(self, h: str, info: dict, num_local: int, tipo_local: str, guardar_texto: bool) -> str:
        self.contador_figuras += 1
        png_bytes = info.get("png_bytes")
        if self.pasta_cache and png_bytes:
            try:
                pasta_blobs = self.pasta_cache / "blobs"
                pasta_blobs.mkdir(parents=True, exist_ok=True)
                (pasta_blobs / f"{h}.png").write_bytes(png_bytes)
            except Exception:
                pass
        marcador = f"<!--DIAGRAMA_REF:{h}:{tipo_local}:{num_local}:{self.contador_figuras}-->"
        dims = info.get("dims", (0, 0))
        info_diag = {
            "id_blob": h,
            "tipo_local": tipo_local,
            "num_local": num_local,
            "indice": self.contador_figuras,
            "largura": dims[0],
            "altura": dims[1],
            "rotulos": info.get("rotulos", []),
            "texto_ocr": info.get("texto", "") if guardar_texto else "",
            "marcador": marcador,
        }
        self.diagramas.append(info_diag)
        return marcador

    def ocr(
        self,
        img_bytes: bytes,
        num_local: int = 1,
        tipo_local: str = "p",
        pagina_inteira: bool = False,
    ) -> str | None:
        """
        Executa OCR em passe único. Retorna texto OCR, marcador de diagrama, ambos (se página inteira
        tiver texto + diagrama) ou None se vazio/repetido.
        """
        if not img_bytes:
            return None
        h = hashlib.sha256(img_bytes).hexdigest()
        if h in self._hashes_vistos:
            return None  # imagem repetida no mesmo documento (logo, carimbo): não duplica
        self._hashes_vistos.add(h)

        info = _analisar_imagem_completa(
            img_bytes,
            min_dim_ocr=self.min_dim,
            min_dim_diag=self.min_dim_diagrama,
            limiar_confianca=self.limiar,
            prop_max_ocr=self.prop_max,
            prop_max_diag=min(self.prop_max, 12.0),
            analisar_diagrama=self.salvar_diagramas,
        )
        txt, conf, st, err = info["texto"], info["conf"], info["estado"], info["erro"]
        self._cache[h] = (txt, conf, st, err)

        if pagina_inteira:
            tem_diag_pagina = (
                self.salvar_diagramas
                and info["eh_diag"]
                and info["res_std"] >= 14.0
                and info["cobertura_texto"] < 0.35
            )
            marcador = self._registrar_diagrama(h, info, num_local, tipo_local, guardar_texto=False) if tem_diag_pagina else ""
            if st == "ok" and txt:
                self.diag["ok"] += 1
                self.qtd += 1
                self.scores.append(conf)
                return f"{txt}\n\n{marcador}" if marcador else txt
            if marcador:
                self.diag["ok"] += 1
                self.qtd += 1
                return marcador
            if st == "error":
                self.diag["error"] += 1
                if err and len(self.erros) < 5:
                    self.erros.append(err)
            else:
                self.diag["empty"] += 1
            return None

        if self.salvar_diagramas and info["eh_diag"]:
            marcador = self._registrar_diagrama(h, info, num_local, tipo_local, guardar_texto=True)
            self.diag["ok"] += 1
            self.qtd += 1
            if conf:
                self.scores.append(conf)
            return marcador

        if st == "ok" and txt:
            palavras = set(re.findall(r"\b\w{3,}\b", txt.lower()))
            if len(palavras) >= 15:
                for visto in self._textos_ocr_vistos:
                    if len(palavras & visto) / max(min(len(palavras), len(visto)), 1) >= 0.90:
                        return None
                self._textos_ocr_vistos.append(palavras)
            self.diag["ok"] += 1
            self.qtd += 1
            self.scores.append(conf)
            return txt
        if st == "error":
            self.diag["error"] += 1
            if err and len(self.erros) < 5:
                self.erros.append(err)
        else:
            self.diag["empty"] += 1
        return None

    @property
    def confianca(self) -> float:
        return sum(self.scores) / len(self.scores) if self.scores else 0.0


def citar_ocr(rotulo: str, texto: str) -> str:
    return f"> **[{rotulo}]**\n" + "\n".join(f"> {ln}" if ln.strip() else ">" for ln in texto.splitlines())


# ---------------------------------------------------------------------------
# Extratores por formato. Todos retornam:
# (conteudo, qtd_ocr, total_paginas, meta_info, paginas_rag, confianca_ocr, diag_ocr)
# ---------------------------------------------------------------------------
def extrair_imagem_avulsa(
    caminho: Path,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
    salvar_diagramas: bool = True,
    min_dim_diagrama: int = 150,
    pasta_cache: Path | None = None,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Processa imagens avulsas via OCR e identifica diagramas/esquemas técnicos."""
    from PIL import Image
    diag = {"ok": 0, "empty": 0, "error": 0}
    dados = caminho.read_bytes()
    try:
        with Image.open(io.BytesIO(dados)) as im:
            w, h = im.size
            formato = im.format or caminho.suffix.lstrip(".").upper()
            n_frames = getattr(im, "n_frames", 1) or 1
            quadros = []
            if n_frames > 1:
                for i in range(min(n_frames, 200)):
                    im.seek(i)
                    buf = io.BytesIO()
                    im.convert("RGB").save(buf, format="PNG")
                    quadros.append(buf.getvalue())
            else:
                quadros = [dados]
    except Exception as e:
        diag["error"] += 1
        aviso = tx(lang, f"Não foi possível abrir a imagem: {e}", f"Could not open the image: {e}")
        return "", 0, 1, {"aviso": aviso}, [], 0.0, diag

    meta_info = {"formato": formato, "resolucao": f"{w}x{h} px"}
    ctx = _ContextoOCR(
        min(min_dim, 40), limiar_confianca, proporcao_max_aspecto,
        salvar_diagramas=salvar_diagramas, min_dim_diagrama=min_dim_diagrama, pasta_cache=pasta_cache
    )
    blocos, paginas_rag = [], []
    for idx, quadro in enumerate(quadros, start=1):
        txt = ctx.ocr(quadro, num_local=idx, tipo_local="p")
        if txt:
            blocos.append(txt)
            paginas_rag.append({"pagina": idx, "texto": txt})
    meta_info["diagramas"] = ctx.diagramas
    return "\n\n".join(blocos), ctx.qtd, len(quadros), meta_info, paginas_rag, ctx.confianca, ctx.diag


def _limpar_data_pdf(valor: str) -> str:
    m = re.match(r"^D?:?(\d{4})(\d{2})?(\d{2})?", valor or "")
    if not m:
        return valor
    ano, mes, dia = m.group(1), m.group(2) or "01", m.group(3) or "01"
    return f"{ano}-{mes}-{dia}"


_VALORES_META_INUTEIS = {
    "untitled", "anonymous", "unknown", "unspecified", "sem título", "sem titulo", "none", "null", "-",
    "user", "usuario", "usuário", "admin", "administrator", "administrador", "windows user",
    "usuário do windows", "microsoft office user", "openpyxl", "python-docx", "python-pptx", "reportlab",
}


def limpar_meta_info(meta: dict) -> dict:
    limpo = {}
    for k, v in (meta or {}).items():
        if v is None:
            continue
        v = normalizar_unicode(str(v).strip())
        if k in ("creationDate", "modDate"):
            v = _limpar_data_pdf(v)
        v = re.sub(r"^Microsoft (Word|PowerPoint|Excel) - ", "", v)
        if not v or v.lower() in _VALORES_META_INUTEIS:
            continue
        limpo[k] = v
    if limpo.get("modDate") and limpo.get("modDate") == limpo.get("creationDate"):
        limpo.pop("modDate")
    return limpo


def _limpar_bloco_pdf(txt: str) -> str:
    txt = txt.replace("­", "")
    txt = re.sub(r"(\w)-\n(?=[a-zà-ÿ])", r"\1", txt)
    linhas = [re.sub(r"[ \t]+", " ", ln).strip() for ln in txt.splitlines()]
    linhas = [ln for ln in linhas if ln and not re.match(r"^[_=\-~*•·]{8,}$", ln)]
    return "\n".join(linhas)


def remover_cabecalhos_rodapes(paginas: list[str]) -> tuple[list[str], int]:
    """
    Remove cabeçalhos/rodapés repetidos (até as 6 primeiras e 6 últimas linhas de cada página
    que se repetem em >= 60% das páginas, incluindo números de página). A 1ª página fica intacta.
    """
    n = len(paginas)
    if n < 3:
        return paginas, 0

    # Só números de página têm os dígitos ignorados; qualquer outra linha precisa ser idêntica,
    # para nunca apagar dados que mudam de página para página (ex.: "Valor total: R$ 1.234,00").
    padrao_num_pagina = re.compile(r"^[-–—\s]*(p[áa]g(ina)?\.?|page|p\.|folha|fl\.)?\s*\d{1,4}(\s*(de|of|/)\s*\d{1,4})?[-–—\s]*$", re.IGNORECASE)

    def norm(ln: str) -> str:
        s = ln.strip().lower()
        return "<num-pagina>" if padrao_num_pagina.match(s) else s

    def candidatas(p: str) -> list[int]:
        idx = [i for i, ln in enumerate(p.split("\n")) if ln.strip()]
        sel = idx[:6] + idx[-6:]
        return sorted(set(sel))

    contagem = Counter()
    for p in paginas:
        lns = p.split("\n")
        chaves = {norm(lns[i]) for i in candidatas(p)
                  if len(lns[i].strip()) <= 120 and not lns[i].lstrip().startswith("|")}
        contagem.update(chaves)
    minimo = max(3, int(0.6 * n))
    repetidas = {k for k, c in contagem.items() if c >= minimo and k}
    if not repetidas:
        return paginas, 0

    saida, removidas = [paginas[0]], 0
    for p in paginas[1:]:
        lns = p.split("\n")
        remover = {i for i in candidatas(p) if norm(lns[i]) in repetidas and not lns[i].lstrip().startswith("|")}
        removidas += len(remover)
        saida.append(normalizar_espacos("\n".join(ln for i, ln in enumerate(lns) if i not in remover)))
    return saida, removidas


def _limpar_matriz_tabela_pdf(raw_rows: list[list]) -> str:
    """Valida e converte matriz bruta de tabela PDF em Markdown, descartando falsas tabelas."""
    linhas_limpas = []
    for row in raw_rows or []:
        if not row:
            continue
        cels = []
        ant = None
        for c in row:
            val = re.sub(r"\s+", " ", _limpar_bloco_pdf(str(c or ""))).strip()
            if val and val == ant:
                cels.append("")
            else:
                cels.append(val)
                if val:
                    ant = val
        if any(cels) and (not linhas_limpas or cels != linhas_limpas[-1]):
            linhas_limpas.append(cels)

    if len(linhas_limpas) < 2:
        return ""
    n_cols = max(len(ln) for ln in linhas_limpas)
    cols_ativas = sum(1 for c in range(n_cols) if any(c < len(ln) and ln[c] for ln in linhas_limpas))
    cels_unicas = {c for ln in linhas_limpas for c in ln if c}
    tem_linha_multi = any(sum(1 for c in ln if c) >= 2 for ln in linhas_limpas)
    if cols_ativas < 2 or len(cels_unicas) < 2 or not tem_linha_multi:
        return ""
    if len(cels_unicas) <= 2 and any(len(c) > 250 for c in cels_unicas):
        return ""
    return tabela_para_markdown(linhas_limpas)


def _ordenar_blocos_visualmente(elementos: list[tuple]) -> list[tuple[float, str, bool]]:
    """Ordena blocos de página PDF por faixas horizontais, priorizando rótulos curtos e ordem esquerda->direita."""
    if not elementos:
        return []
    ordenados = sorted(elementos, key=lambda it: (round(it[0].y0, 1), round(it[0].x0, 1)))
    faixas: list[dict] = []
    for r, txt, eh_tab in ordenados:
        alocado = False
        for f in reversed(faixas[-3:]):
            sobrep = min(r.y1, f["y1"]) - max(r.y0, f["y0"])
            alt_min = max(min(r.height, f["y1"] - f["y0"]), 1.0)
            mesma_faixa = sobrep > 0.35 * alt_min or (
                not eh_tab and len(txt) <= 35 and r.y0 >= f["y0"] - 8 and r.y1 <= f["y1"] + 8
            )
            if mesma_faixa:
                f["itens"].append((r, txt, eh_tab))
                f["y0"] = min(f["y0"], r.y0)
                f["y1"] = max(f["y1"], r.y1)
                alocado = True
                break
        if not alocado:
            faixas.append({"y0": r.y0, "y1": r.y1, "itens": [(r, txt, eh_tab)]})

    itens = []
    for f in faixas:
        f["itens"].sort(
            key=lambda it: (0 if (not it[2] and len(it[1]) <= 35) else 1, round(it[0].x0, 1), round(it[0].y0, 1))
        )
        for r, txt, eh_tab in f["itens"]:
            itens.append((r.y0, txt, eh_tab))
    return itens


def _texto_pagina_pdf(pymupdf, pagina) -> str:
    """Texto nativo da página em ordem visual, com tabelas reais em Markdown e sem duplicar texto."""
    blocos = [b for b in pagina.get_text("blocks") if len(b) > 6 and b[6] == 0]
    tabelas = []
    try:
        achadas = pagina.find_tables()
        for tb in getattr(achadas, "tables", []) or []:
            if getattr(tb, "row_count", 0) >= 2 and getattr(tb, "col_count", 0) >= 2:
                try:
                    raw_rows = tb.extract() or []
                except Exception:
                    raw_rows = []
                md = _limpar_matriz_tabela_pdf(raw_rows)
                if md:
                    tabelas.append((pymupdf.Rect(tb.bbox), md))
    except Exception:
        tabelas = []

    def dentro_de_tabela(r) -> bool:
        area = max(r.get_area(), 1e-6)
        for tr, _ in tabelas:
            inter = pymupdf.Rect(r)
            inter.intersect(tr)
            if not inter.is_empty and inter.get_area() / area >= 0.6:
                return True
        return False

    elementos = []
    for b in blocos:
        r = pymupdf.Rect(b[:4])
        if tabelas and dentro_de_tabela(r):
            continue
        txt = _limpar_bloco_pdf(b[4])
        if txt:
            elementos.append((r, txt, False))
    for tr, md in tabelas:
        elementos.append((tr, md, True))

    itens = _ordenar_blocos_visualmente(elementos)

    # Muitos PDFs gravam cada linha como um bloco: junta blocos que continuam a mesma frase
    # (bloco anterior sem pontuação final e o seguinte começando em minúscula/número).
    saida: list[str] = []
    for _, txt, _eh_tab in itens:
        if (saida and not txt.startswith("|") and not saida[-1].startswith("|")
                and saida[-1][-1:] not in ".:;!?)]" and re.match(r"[a-zà-ÿ0-9(\"“'«]", txt)):
            saida[-1] = saida[-1] + "\n" + txt
        else:
            saida.append(txt)
    return "\n\n".join(saida)


def extrair_pdf_hibrido(
    caminho_pdf: Path,
    usar_ocr: bool = True,
    dpi: int = 200,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
    remover_repetidos: bool = True,
    salvar_diagramas: bool = True,
    min_dim_diagrama: int = 150,
    pasta_cache: Path | None = None,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Lê o texto nativo do PDF e aplica OCR em páginas escaneadas e imagens embutidas."""
    try:
        import pymupdf
    except ImportError:
        pymupdf = None

    if pymupdf is None:
        from pypdf import PdfReader
        reader = PdfReader(str(caminho_pdf))
        paginas = [_limpar_bloco_pdf(p.extract_text() or "") for p in reader.pages]
        if remover_repetidos:
            paginas, _ = remover_cabecalhos_rodapes(paginas)
        rag = [{"pagina": i, "texto": p} for i, p in enumerate(paginas, start=1) if p.strip()]
        return "\n\n".join(p for p in paginas if p.strip()), 0, len(reader.pages), {}, rag, 0.0, {"ok": 0, "empty": 0, "error": 0}

    try:
        pymupdf.no_recommend_layout()
    except Exception:
        pass
    try:
        pymupdf.TOOLS.mupdf_display_warnings(False)
        pymupdf.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass

    ctx = _ContextoOCR(min_dim, limiar_confianca, proporcao_max_aspecto, salvar_diagramas=salvar_diagramas, min_dim_diagrama=min_dim_diagrama, pasta_cache=pasta_cache)
    doc = pymupdf.open(caminho_pdf)
    try:
        total_paginas = len(doc)
        meta_info = {}
        for campo in ("title", "author", "subject", "keywords", "creationDate", "modDate"):
            val = (doc.metadata or {}).get(campo)
            if isinstance(val, str) and val.strip():
                meta_info[campo] = val.strip()
        meta_info = limpar_meta_info(meta_info)

        if doc.needs_pass and not doc.authenticate(""):
            meta_info["aviso"] = tx(lang, "PDF protegido por senha: conteúdo não extraído.",
                                    "Password-protected PDF: content not extracted.")
            return "", 0, total_paginas, meta_info, [], 0.0, ctx.diag

        textos_paginas: list[str] = []
        paginas_ocr = 0
        for pagina in doc:
            texto_nativo = _texto_pagina_pdf(pymupdf, pagina)
            n_nativo = len(texto_nativo.strip())

            xrefs = []
            for img in pagina.get_images(full=True):
                if img[0] not in xrefs:
                    xrefs.append(img[0])

            area_pagina = max(abs(pagina.rect.width * pagina.rect.height), 1.0)
            cobertura_grande = False
            for xref in xrefs:
                try:
                    for r in pagina.get_image_rects(xref):
                        if (r.width * r.height) / area_pagina >= 0.60:
                            cobertura_grande = True
                            break
                except Exception:
                    pass
                if cobertura_grande:
                    break

            texto_vetorial = False
            if usar_ocr and n_nativo == 0 and not xrefs:
                try:
                    texto_vetorial = len(pagina.get_drawings()) >= 30  # texto convertido em curvas
                except Exception:
                    texto_vetorial = False

            escaneada = usar_ocr and (
                (n_nativo < 50 and (xrefs or texto_vetorial))
                or (cobertura_grande and n_nativo < 200)
            )

            if escaneada:
                texto_ocr = None
                try:
                    pix = pagina.get_pixmap(dpi=dpi)
                    texto_ocr = ctx.ocr(pix.tobytes("png"), num_local=pagina.number + 1, tipo_local="p", pagina_inteira=True)
                except Exception as e:
                    ctx.diag["error"] += 1
                    if len(ctx.erros) < 5:
                        ctx.erros.append(str(e))
                if texto_ocr:
                    paginas_ocr += 1
                    texto_pagina = texto_ocr
                else:
                    texto_pagina = texto_nativo
            else:
                partes = [texto_nativo] if texto_nativo else []
                # Imagens embutidas em página com texto (prints, gráficos, comprovantes colados).
                # Em páginas com texto e imagem cobrindo a página (PDF pesquisável), o texto nativo já basta.
                if usar_ocr and xrefs and not cobertura_grande:
                    for xref in xrefs:
                        try:
                            info = doc.extract_image(xref)
                        except Exception:
                            continue
                        if not info or not info.get("image"):
                            continue
                        txt_img = ctx.ocr(info["image"], num_local=pagina.number + 1, tipo_local="p")
                        if txt_img:
                            if "<!--DIAGRAMA_REF:" in txt_img:
                                partes.append(txt_img)
                            else:
                                rotulo = tx(lang, "Imagem embutida — texto via OCR", "Embedded image — OCR text")
                                partes.append(citar_ocr(rotulo, txt_img))
                texto_pagina = "\n\n".join(partes)

            textos_paginas.append(normalizar_espacos(texto_pagina))

        if remover_repetidos:
            textos_paginas, _ = remover_cabecalhos_rodapes(textos_paginas)

        paginas_rag = [{"pagina": i, "texto": p} for i, p in enumerate(textos_paginas, start=1) if p.strip()]
        conteudo = "\n\n".join(p for p in textos_paginas if p.strip())

        meta_info["diagramas"] = ctx.diagramas
        if paginas_ocr:
            meta_info["paginas_ocr"] = f"{paginas_ocr}/{total_paginas}"
        if not conteudo.strip():
            if not usar_ocr:
                meta_info["aviso"] = tx(lang, "PDF sem texto nativo e OCR desativado.", "PDF has no native text and OCR is disabled.")
            elif ctx.diag["error"]:
                meta_info["aviso"] = tx(lang, "Falha técnica no OCR: ", "OCR technical failure: ") + "; ".join(ctx.erros[:2])
            else:
                meta_info["aviso"] = tx(lang, "Nenhum texto legível encontrado.", "No readable text found.")
        return conteudo, ctx.qtd, total_paginas, meta_info, paginas_rag, ctx.confianca, ctx.diag
    finally:
        doc.close()


def _nivel_titulo_docx(nome_estilo: str) -> int | None:
    nome = (nome_estilo or "").strip().lower()
    if nome in ("title", "título", "titulo", "titre"):
        return 3
    m = re.match(r"^(heading|título|titulo|titre|überschrift|encabezado)\s*(\d)", nome)
    if m:
        return min(int(m.group(2)) + 2, 6)
    return None


def extrair_docx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
    salvar_diagramas: bool = True,
    min_dim_diagrama: int = 150,
    pasta_cache: Path | None = None,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Extrai parágrafos, títulos, listas, tabelas e imagens (OCR) na ordem em que aparecem."""
    import docx
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    documento = docx.Document(str(caminho))
    partes: list[str] = []
    ctx = _ContextoOCR(min_dim, limiar_confianca, proporcao_max_aspecto, salvar_diagramas=salvar_diagramas, min_dim_diagrama=min_dim_diagrama, pasta_cache=pasta_cache)
    meta_info = {}
    try:
        cp = documento.core_properties
        if cp.title:
            meta_info["title"] = cp.title
        if cp.author:
            meta_info["author"] = cp.author
        if cp.created:
            meta_info["creationDate"] = cp.created.strftime("%Y-%m-%d")
        if cp.modified:
            meta_info["modDate"] = cp.modified.strftime("%Y-%m-%d")
    except Exception:
        pass
    meta_info = limpar_meta_info(meta_info)

    rotulo_img = tx(lang, "Imagem no documento — texto via OCR", "Image in document — OCR text")
    rids_vistos: set[str] = set()

    def ocr_imagens_do_elemento(elemento):
        if not usar_ocr_imagens:
            return
        for r_id in elemento.xpath(".//@r:embed"):
            if r_id in rids_vistos:
                continue
            rids_vistos.add(r_id)
            alvo = documento.part.related_parts.get(r_id)
            blob = getattr(alvo, "blob", None)
            if blob:
                txt = ctx.ocr(blob, num_local=1, tipo_local="p")
                if txt:
                    if "<!--DIAGRAMA_REF:" in txt:
                        partes.append(txt)
                    else:
                        partes.append(citar_ocr(rotulo_img, txt))

    for filho in documento.element.body.iterchildren():
        if isinstance(filho, CT_P):
            p = Paragraph(filho, documento)
            texto = p.text.strip()
            # caixas de texto (w:txbxContent) não fazem parte de p.text
            caixas = [
                "".join(n.text or "" for n in tp.xpath(".//w:t")).strip()
                for tp in filho.xpath(".//w:txbxContent//w:p")
            ]
            if texto:
                try:
                    estilo = p.style.name if p.style is not None else ""
                except Exception:
                    estilo = ""
                nivel = _nivel_titulo_docx(estilo)
                eh_lista = filho.pPr is not None and filho.pPr.numPr is not None or "list" in (estilo or "").lower()
                if nivel:
                    partes.append("#" * nivel + " " + texto.replace("\n", " "))
                elif eh_lista:
                    partes.append("- " + texto)
                else:
                    partes.append(texto)
            for cx in caixas:
                if cx and cx != texto:
                    partes.append(cx)
            ocr_imagens_do_elemento(filho)
        elif isinstance(filho, CT_Tbl):
            tabela = Table(filho, documento)
            linhas = []
            for row in tabela.rows:
                celulas, ultimo = [], None
                for cel in row.cells:
                    if cel._tc is ultimo:  # células mescladas horizontalmente se repetem
                        continue
                    ultimo = cel._tc
                    celulas.append(cel.text.strip())
                linhas.append(celulas)
            md = tabela_para_markdown(linhas)
            if md:
                partes.append(md)
            ocr_imagens_do_elemento(filho)

    texto_final = normalizar_espacos("\n\n".join(partes))
    paginas_rag = [{"pagina": 1, "texto": texto_final}] if texto_final else []
    meta_info["diagramas"] = ctx.diagramas
    return texto_final, ctx.qtd, 1, meta_info, paginas_rag, ctx.confianca, ctx.diag


def formatar_valor_celula(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return f"{round(v, 6):.6f}".rstrip("0").rstrip(".")
    return str(v).strip()


def _matriz_aparada(linhas: list[list[str]]) -> list[list[str]]:
    linhas = [ln for ln in linhas if any(c for c in ln)]
    if not linhas:
        return []
    n = max(len(ln) for ln in linhas)
    for ln in linhas:
        ln.extend([""] * (n - len(ln)))
    usadas = [c for c in range(n) if any(ln[c] for ln in linhas)]
    return [[ln[c] for c in usadas] for ln in linhas]


def extrair_xlsx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
    salvar_diagramas: bool = True,
    min_dim_diagrama: int = 150,
    pasta_cache: Path | None = None,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Lê cada aba como tabela Markdown (valores calculados) e aplica OCR nas imagens da aba."""
    import openpyxl

    ctx = _ContextoOCR(min_dim, limiar_confianca, proporcao_max_aspecto, salvar_diagramas=salvar_diagramas, min_dim_diagrama=min_dim_diagrama, pasta_cache=pasta_cache)
    lbl_aba = tx(lang, "Planilha", "Sheet")
    lbl_oculta = tx(lang, "oculta", "hidden")
    lbl_formula = tx(lang, "fórmula", "formula")
    partes, paginas_rag = [], []

    wb = openpyxl.load_workbook(str(caminho), data_only=True)
    # Fórmulas só são exibidas quando a célula não tem valor calculado salvo (ex.: arquivo gerado por sistema).
    formulas: dict[str, dict] = {}
    try:
        wb_f = openpyxl.load_workbook(str(caminho), data_only=False, read_only=True)
        for ws_f in wb_f.worksheets:
            mapa = {}
            for row in ws_f.iter_rows():
                for cel in row:
                    v = getattr(cel, "value", None)
                    if isinstance(v, str) and v.startswith("="):
                        mapa[(cel.row, cel.column)] = v
            formulas[ws_f.title] = mapa
        wb_f.close()
    except Exception:
        formulas = {}

    meta_info = {}
    try:
        props = wb.properties
        if props.title:
            meta_info["title"] = props.title
        if props.creator:
            meta_info["author"] = props.creator
    except Exception:
        pass
    meta_info = limpar_meta_info(meta_info)

    try:
        for idx_aba, ws in enumerate(wb.worksheets, start=1):
            mapa_f = formulas.get(ws.title, {})
            linhas = []
            # min_row/min_col explícitos: mantém a numeração alinhada ao mapa de fórmulas
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, min_col=1, values_only=True), start=1):
                celulas = []
                for c_idx, v in enumerate(row, start=1):
                    s = formatar_valor_celula(v)
                    if not s and (r_idx, c_idx) in mapa_f:
                        s = f"[{lbl_formula}: {mapa_f[(r_idx, c_idx)]}]"
                    celulas.append(s)
                linhas.append(celulas)
            matriz = _matriz_aparada(linhas)

            blocos_aba = []
            if matriz:
                blocos_aba.append(tabela_para_markdown(matriz))
            if usar_ocr_imagens:
                for im in getattr(ws, "_images", []) or []:
                    try:
                        dados = im._data()
                    except Exception:
                        continue
                    txt = ctx.ocr(dados, num_local=idx_aba, tipo_local="aba")
                    if txt:
                        if "<!--DIAGRAMA_REF:" in txt:
                            blocos_aba.append(txt)
                        else:
                            blocos_aba.append(citar_ocr(tx(lang, "Imagem na planilha — texto via OCR", "Image in sheet — OCR text"), txt))
            if blocos_aba:
                oculta = f" ({lbl_oculta})" if getattr(ws, "sheet_state", "visible") != "visible" else ""
                bloco = f"### {lbl_aba}: {ws.title}{oculta}\n\n" + "\n\n".join(blocos_aba)
                partes.append(bloco)
                paginas_rag.append({"pagina": idx_aba, "texto": bloco})
        total_abas = len(wb.worksheets)
    finally:
        wb.close()

    meta_info["diagramas"] = ctx.diagramas
    return "\n\n".join(partes), ctx.qtd, total_abas, meta_info, paginas_rag, ctx.confianca, ctx.diag


def extrair_xls_legado(caminho: Path, lang: str) -> tuple[str, int, int, dict, list, float, dict]:
    import pandas as pd
    lbl_aba = tx(lang, "Planilha", "Sheet")
    partes, paginas_rag = [], []
    xl = pd.ExcelFile(str(caminho))
    for idx, nome in enumerate(xl.sheet_names, start=1):
        df = xl.parse(nome, header=None)
        linhas = [[formatar_valor_celula(v) for v in row] for row in df.itertuples(index=False, name=None)]
        matriz = _matriz_aparada(linhas)
        if matriz:
            bloco = f"### {lbl_aba}: {nome}\n\n{tabela_para_markdown(matriz)}"
            partes.append(bloco)
            paginas_rag.append({"pagina": idx, "texto": bloco})
    return "\n\n".join(partes), 0, len(xl.sheet_names), {}, paginas_rag, 0.0, {"ok": 0, "empty": 0, "error": 0}


def _iterar_shapes(shapes):
    """Percorre shapes do slide, incluindo os que estão dentro de grupos, em ordem de leitura."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def chave(s):
        try:
            return (int(s.top or 0), int(s.left or 0))
        except Exception:
            return (0, 0)

    for shape in sorted(list(shapes), key=chave):
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            yield from _iterar_shapes(shape.shapes)
        else:
            yield shape


def _grafico_para_markdown(chart, lang: str) -> str:
    titulo = ""
    try:
        if chart.has_title and chart.chart_title.has_text_frame:
            titulo = chart.chart_title.text_frame.text.strip()
    except Exception:
        pass
    cabecalho = f"**[{tx(lang, 'Gráfico', 'Chart')}{': ' + titulo if titulo else ''}]**"
    try:
        plot = chart.plots[0]
        categorias = [str(c) for c in plot.categories]
        series = list(plot.series)
        if categorias and series:
            linhas = [[tx(lang, "Categoria", "Category")] + [s.name or f"S{i + 1}" for i, s in enumerate(series)]]
            for i, cat in enumerate(categorias):
                linha_v = [cat]
                for s in series:
                    vals = list(s.values)
                    linha_v.append(formatar_valor_celula(vals[i]) if i < len(vals) else "")
                linhas.append(linha_v)
            return cabecalho + "\n\n" + tabela_para_markdown(linhas)
    except Exception:
        pass
    return cabecalho


def extrair_pptx_contextual(
    caminho: Path,
    usar_ocr_imagens: bool = True,
    min_dim: int = 150,
    limiar_confianca: float = 0.50,
    proporcao_max_aspecto: float = 20.0,
    lang: str = "pt",
    salvar_diagramas: bool = True,
    min_dim_diagrama: int = 150,
    pasta_cache: Path | None = None,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Extrai, por slide: título, textos (com níveis), tabelas, dados de gráficos, imagens (OCR) e notas."""
    from pptx import Presentation

    prs = Presentation(str(caminho))
    ctx = _ContextoOCR(min_dim, limiar_confianca, proporcao_max_aspecto, salvar_diagramas=salvar_diagramas, min_dim_diagrama=min_dim_diagrama, pasta_cache=pasta_cache)
    partes, paginas_rag = [], []
    meta_info = {}
    try:
        cp = prs.core_properties
        if cp.title:
            meta_info["title"] = cp.title
        if cp.author:
            meta_info["author"] = cp.author
    except Exception:
        pass
    meta_info = limpar_meta_info(meta_info)

    for idx_s, slide in enumerate(prs.slides, start=1):
        titulo_shape = None
        titulo = ""
        try:
            titulo_shape = slide.shapes.title
            if titulo_shape is not None and titulo_shape.has_text_frame:
                titulo = titulo_shape.text_frame.text.strip().replace("\n", " ")
        except Exception:
            titulo_shape = None
        cab = f"### Slide {idx_s}" + (f": {titulo}" if titulo else "")
        corpo: list[str] = []

        for shape in _iterar_shapes(slide.shapes):
            if titulo_shape is not None and shape.shape_id == titulo_shape.shape_id:
                continue
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    txt = "".join(r.text for r in p.runs).strip() or p.text.strip()
                    if txt:
                        corpo.append("  " * int(p.level or 0) + "- " + txt.replace("\n", " "))
            if getattr(shape, "has_table", False) and shape.has_table:
                linhas = [[cel.text.strip() for cel in row.cells] for row in shape.table.rows]
                md = tabela_para_markdown(linhas)
                if md:
                    corpo.append("\n" + md + "\n")
            if getattr(shape, "has_chart", False) and shape.has_chart:
                corpo.append("\n" + _grafico_para_markdown(shape.chart, lang) + "\n")
            if usar_ocr_imagens:
                blob = None
                try:
                    blob = shape.image.blob
                except Exception:
                    blob = None
                if blob:
                    txt = ctx.ocr(blob, num_local=idx_s, tipo_local="s")
                    if txt:
                        if "<!--DIAGRAMA_REF:" in txt:
                            corpo.append("\n" + txt + "\n")
                        else:
                            corpo.append("\n" + citar_ocr(tx(lang, "Imagem do slide — texto via OCR", "Slide image — OCR text"), txt) + "\n")

        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                notas = slide.notes_slide.notes_text_frame.text.strip()
                if notas:
                    corpo.append(f"\n**{tx(lang, 'Notas do apresentador', 'Speaker notes')}:** {notas}")
        except Exception:
            pass

        bloco = normalizar_espacos(cab + "\n\n" + "\n".join(corpo))
        partes.append(bloco)
        paginas_rag.append({"pagina": idx_s, "texto": bloco})

    meta_info["diagramas"] = ctx.diagramas
    return "\n\n".join(partes), ctx.qtd, len(prs.slides), meta_info, paginas_rag, ctx.confianca, ctx.diag


# ---------------------------------------------------------------------------
# Worker multiprocessado + cache
# ---------------------------------------------------------------------------
def calcular_hash_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with caminho.open("rb") as f:
        while chunk := f.read(1024 * 1024):
            h.update(chunk)
    return h.hexdigest()


def parametros_extracao(cfg: dict) -> dict:
    """Tudo o que altera o conteúdo extraído entra na chave do cache."""
    return {
        "versao": VERSAO_EXTRATOR,
        "idioma": obter_idioma(cfg),
        "ocr_pdf": bool(cfg.get("usar_ocr_em_pdf", True)),
        "ocr_img": bool(cfg.get("usar_ocr_em_imagens_embutidas", True)),
        "salvar_diag": bool(cfg.get("salvar_diagramas_em_pasta", True)),
        "min_dim_diag": int(cfg.get("min_dimensao_diagrama", 150)),
        "dpi": int(cfg.get("dpi_ocr_pdf", 200)),
        "min_dim": int(cfg.get("min_dimensao_imagem_ocr", 150)),
        "limiar": round(float(cfg.get("limiar_confianca_ocr", 0.50)), 3),
        "aspecto": round(float(cfg.get("proporcao_max_aspecto_ocr", 20.0)), 2),
        "pii": bool(cfg.get("mascarar_pii", False)),
        "cab_rod": bool(cfg.get("remover_cabecalhos_rodapes", True)),
    }


def chave_cache(cfg: dict) -> str:
    bruto = json.dumps(parametros_extracao(cfg), sort_keys=True)
    return hashlib.md5(bruto.encode("utf-8")).hexdigest()[:12]


def _despachar_extrator(
    arq: Path,
    cfg: dict,
    pasta_cache: Path | None,
    permitir_anexos_eml: bool = True,
) -> tuple[str, int, int, dict, list, float, dict]:
    """Despacha a extração para o leitor adequado conforme a extensão do arquivo."""
    lang = obter_idioma(cfg)
    ext = arq.suffix.lower()
    min_dim = int(cfg.get("min_dimensao_imagem_ocr", 150))
    limiar = float(cfg.get("limiar_confianca_ocr", 0.50))
    prop_max = float(cfg.get("proporcao_max_aspecto_ocr", 20.0))
    ocr_img = bool(cfg.get("usar_ocr_em_imagens_embutidas", True))
    salvar_diag = bool(cfg.get("salvar_diagramas_em_pasta", True))
    min_dim_diag = int(cfg.get("min_dimensao_diagrama", 150))
    vazio = {"ok": 0, "empty": 0, "error": 0}

    if ext == ".pdf":
        return extrair_pdf_hibrido(
            arq, usar_ocr=bool(cfg.get("usar_ocr_em_pdf", True)), dpi=int(cfg.get("dpi_ocr_pdf", 200)),
            min_dim=min_dim, limiar_confianca=limiar, proporcao_max_aspecto=prop_max, lang=lang,
            remover_repetidos=bool(cfg.get("remover_cabecalhos_rodapes", True)),
            salvar_diagramas=salvar_diag, min_dim_diagrama=min_dim_diag, pasta_cache=pasta_cache,
        )
    if ext in EXTENSOES_IMAGEM:
        if cfg.get("usar_ocr_em_pdf", True):
            return extrair_imagem_avulsa(arq, min_dim=min_dim, limiar_confianca=limiar,
                                         proporcao_max_aspecto=prop_max, lang=lang,
                                         salvar_diagramas=salvar_diag, min_dim_diagrama=min_dim_diag,
                                         pasta_cache=pasta_cache)
        return ("", 0, 1, {"aviso": tx(lang, "OCR desativado.", "OCR is disabled.")}, [], 0.0, vazio)
    if ext == ".docx":
        return extrair_docx_contextual(arq, ocr_img, min_dim, limiar, prop_max, lang,
                                       salvar_diagramas=salvar_diag, min_dim_diagrama=min_dim_diag,
                                       pasta_cache=pasta_cache)
    if ext in (".xlsx", ".xlsm"):
        return extrair_xlsx_contextual(arq, ocr_img, min_dim, limiar, prop_max, lang,
                                       salvar_diagramas=salvar_diag, min_dim_diagrama=min_dim_diag,
                                       pasta_cache=pasta_cache)
    if ext == ".xls":
        return extrair_xls_legado(arq, lang)
    if ext == ".pptx":
        return extrair_pptx_contextual(arq, ocr_img, min_dim, limiar, prop_max, lang,
                                       salvar_diagramas=salvar_diag, min_dim_diagrama=min_dim_diag,
                                       pasta_cache=pasta_cache)
    if ext == ".csv":
        txt = extrair_csv_estruturado(arq)
        return (txt, 0, 1, {}, [{"pagina": 1, "texto": txt}] if txt else [], 0.0, vazio)
    txt, meta = extrair_texto_generico(arq, ext, lang, cfg=cfg, pasta_cache=pasta_cache, permitir_anexos_eml=permitir_anexos_eml)
    meta = dict(meta or {})
    ocr_m = meta.pop("_ocr_count", 0)
    diag_m = meta.pop("_diag_ocr", vazio)
    conf_m = meta.pop("_confianca", 0.0)
    return (txt, ocr_m, 1, meta, [{"pagina": 1, "texto": txt}] if txt else [], conf_m, diag_m)


def worker_processar_arquivo(args: tuple) -> tuple[str, str, str, dict, str | None]:
    """Processa um arquivo (em processo separado) e grava o resultado no cache."""
    raiz_str, arq_str, pasta_cache_str, cfg = args
    raiz, arq, pasta_cache = Path(raiz_str), Path(arq_str), Path(pasta_cache_str)
    lang = obter_idioma(cfg)

    try:
        partes_rel = arq.relative_to(raiz).parts
        caminho_rel = "/".join((raiz.name,) + partes_rel)
        # Grupo (seção do índice e arquivo em por_pasta/): pasta raiz + primeira subpasta, de modo
        # que uma pasta mãe com departamentos (Financeiro/, Juridico/...) fique organizada por departamento.
        grupo = f"{raiz.name}/{partes_rel[0]}" if len(partes_rel) > 1 else raiz.name
    except Exception:
        caminho_rel = arq.name
        grupo = raiz.name

    try:
        file_hash = calcular_hash_arquivo(arq)
        arq_cache = pasta_cache / f"{file_hash}_{chave_cache(cfg)}.json"

        if arq_cache.exists():
            try:
                dados = json.loads(arq_cache.read_text(encoding="utf-8"))
                dados["cache"] = True
                dados["hash"] = file_hash
                return arq.name, caminho_rel, grupo, dados, None
            except Exception:
                pass  # cache corrompido: reprocessa

        ext = arq.suffix.lower()
        res = _despachar_extrator(arq, cfg, pasta_cache)

        conteudo, ocr_count, total_paginas, meta_info, paginas_rag, confianca, diag = res
        meta_info = dict(meta_info or {})
        diagramas_encontrados = meta_info.pop("diagramas", [])

        conteudo, n_mojibake = reparar_mojibake(conteudo or "")
        for p in paginas_rag:
            p["texto"], _ = reparar_mojibake(p.get("texto", ""))

        if cfg.get("mascarar_pii", False):
            conteudo = redigir_pii(conteudo)
            for p in paginas_rag:
                p["texto"] = redigir_pii(p["texto"])
            for k in ("author", "title", "subject"):
                if k in meta_info:
                    meta_info[k] = redigir_pii(meta_info[k])

        conteudo = conteudo.strip()
        if not conteudo and "aviso" not in meta_info:
            if ext in EXTENSOES_IMAGEM:
                meta_info["aviso"] = tx(lang, "Imagem sem texto legível.", "Image without readable text.")
            else:
                meta_info["aviso"] = tx(lang, "Documento sem texto extraível.", "Document without extractable text.")

        dados = {
            "arquivo": arq.name,
            "conteudo": conteudo,
            "ocr_pages": ocr_count,
            "total_paginas": total_paginas,
            "meta_info": meta_info,
            "confianca_ocr": confianca,
            "paginas_rag": paginas_rag,
            "diag_ocr": diag,
            "mojibake_repairs": n_mojibake,
            "diagramas": diagramas_encontrados,
            "data_cache": datetime.now().isoformat(),
        }

        # Gravação atômica com nome temporário exclusivo (dois arquivos idênticos podem ser
        # processados ao mesmo tempo por workers diferentes).
        try:
            pasta_cache.mkdir(parents=True, exist_ok=True)
            tmp = arq_cache.with_name(f"{arq_cache.stem}.{os.getpid()}.{time.time_ns()}.tmp")
            tmp.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp, arq_cache)
        except Exception:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass

        dados["cache"] = False
        dados["hash"] = file_hash
        return arq.name, caminho_rel, grupo, dados, None
    except Exception as e:
        return arq.name, caminho_rel, grupo, {}, f"{type(e).__name__}: {e}"


class ProcessamentoCancelado(Exception):
    pass


def processar_em_paralelo(tarefas: list[tuple], max_workers: int, ao_concluir) -> None:
    """
    Executa worker_processar_arquivo em paralelo. Se um arquivo derrubar um processo (falha
    nativa em biblioteca de PDF/OCR, falta de memória), os pendentes são reprocessados um a um
    para isolar o culpado, que é registrado como falha sem perder o restante do lote.
    """
    pendentes = list(tarefas)
    modo_isolado = False
    while pendentes:
        concluidos: set[int] = set()
        quebrou = False
        executor = ProcessPoolExecutor(max_workers=1 if modo_isolado else max(1, max_workers))
        cancelado = False
        try:
            futuros = {executor.submit(worker_processar_arquivo, tf): i for i, tf in enumerate(pendentes)}
            for fut in as_completed(futuros):
                i = futuros[fut]
                try:
                    resultado = fut.result()
                except BrokenProcessPool:
                    quebrou = True
                    break
                except Exception as e:
                    raiz, arq = Path(pendentes[i][0]), Path(pendentes[i][1])
                    resultado = (arq.name, arq.name, raiz.name, {}, f"{type(e).__name__}: {e}")
                concluidos.add(i)
                ao_concluir(resultado)
        except KeyboardInterrupt:
            cancelado = True
        finally:
            try:
                executor.shutdown(wait=not cancelado, cancel_futures=True)
            except Exception:
                pass
        if cancelado:
            raise ProcessamentoCancelado()

        restantes = [tf for i, tf in enumerate(pendentes) if i not in concluidos]
        if not quebrou:
            break
        if modo_isolado and restantes:
            # Com 1 worker as tarefas rodam em ordem: a primeira não concluída derrubou o processo.
            raiz, arq = Path(restantes[0][0]), Path(restantes[0][1])
            try:
                rel = f"{raiz.name}/{arq.relative_to(raiz).as_posix()}"
            except Exception:
                rel = arq.name
            ao_concluir((arq.name, rel, raiz.name, {}, "O processo de extração foi encerrado inesperadamente (arquivo corrompido ou memória insuficiente) / worker crashed"))
            restantes = restantes[1:]
        modo_isolado = True
        pendentes = restantes


# ---------------------------------------------------------------------------
# Varredura de pastas
# ---------------------------------------------------------------------------
_PASTAS_IGNORADAS = {"__pycache__", "node_modules", "$recycle.bin", "system volume information"}
_ARQUIVOS_IGNORADOS = {"thumbs.db", "desktop.ini", ".ds_store"}


def listar_arquivos_elegiveis(
    pastas: list[str],
    extensoes: list[str],
    excluir: list[Path] | None = None,
    ignorados: dict | None = None,
) -> tuple[list[tuple[Path, Path]], list[str]]:
    """
    Lista (raiz, arquivo) elegíveis. Ignora pastas ocultas, arquivos temporários do Office (~$),
    arquivos vazios e as pastas em `excluir` (ex.: a própria pasta de saída).
    `ignorados` (opcional) recebe a contagem de arquivos não suportados por extensão.
    """
    exts = {e.lower() for e in extensoes}
    excluir_chaves = {_chave_caminho(p) for p in (excluir or [])}
    arquivos: list[tuple[Path, Path]] = []
    vistos: set[str] = set()
    sem_permissao: list[str] = []

    def erro_walk(err):
        sem_permissao.append(str(getattr(err, "filename", err)))

    for p_str in pastas:
        raiz = resolver_caminho(p_str)
        if not raiz.is_dir():
            continue
        if _chave_caminho(raiz) in excluir_chaves:
            continue
        for dirpath, dirnames, filenames in os.walk(raiz, onerror=erro_walk):
            dirnames[:] = sorted(
                d for d in dirnames
                if not d.startswith((".", "~$"))
                and d.lower() not in _PASTAS_IGNORADAS
                and _chave_caminho(Path(dirpath) / d) not in excluir_chaves
            )
            for nome in filenames:
                if nome.startswith(("~$", ".")) or nome.lower() in _ARQUIVOS_IGNORADOS:
                    continue
                item = Path(dirpath) / nome
                ext = item.suffix.lower()
                if ext not in exts:
                    if ignorados is not None:
                        chave_ext = ext or "(sem extensão/no extension)"
                        ignorados[chave_ext] = ignorados.get(chave_ext, 0) + 1
                    continue
                try:
                    if item.stat().st_size == 0:
                        if ignorados is not None:
                            ignorados["(vazio/empty)"] = ignorados.get("(vazio/empty)", 0) + 1
                        continue
                    chave = _chave_caminho(item.resolve())
                except PermissionError:
                    sem_permissao.append(str(item))
                    continue
                except OSError:
                    continue
                if chave in vistos:
                    continue
                vistos.add(chave)
                arquivos.append((raiz, item))

    arquivos.sort(key=lambda x: chave_ordenacao_natural(f"{x[0].name}/{x[1].relative_to(x[0]).as_posix()}"))
    return arquivos, sem_permissao


def obter_estatisticas_arquivos(arquivos: list[tuple[Path, Path]]) -> dict:
    por_extensao: dict[str, dict] = {}
    total_bytes = 0
    for _, arq in arquivos:
        ext = arq.suffix.lower() or "(?)"
        info = por_extensao.setdefault(ext, {"qtd": 0, "bytes": 0})
        info["qtd"] += 1
        try:
            tam = arq.stat().st_size
        except OSError:
            tam = 0
        info["bytes"] += tam
        total_bytes += tam
    return {
        "total_arquivos": len(arquivos),
        "tamanho_total_bytes": total_bytes,
        "tamanho_formatado": formatar_tamanho_bytes(total_bytes),
        "por_extensao": por_extensao,
    }


_CACHE_VARREDURA: dict = {"chave": None, "dados": None}


def varrer_pastas(cfg: dict, forcar: bool = False) -> dict:
    """Varre as pastas de entrada (com cache para o menu não reler o disco a cada tela)."""
    saida = resolver_caminho(cfg.get("pasta_saida", "./saida_copilot"))
    chave = (tuple(cfg.get("pastas_entrada", [])), tuple(sorted(cfg.get("extensoes", []))), str(saida))
    if not forcar and _CACHE_VARREDURA["chave"] == chave and _CACHE_VARREDURA["dados"] is not None:
        return _CACHE_VARREDURA["dados"]
    ignorados: dict = {}
    arquivos, sem_perm = listar_arquivos_elegiveis(
        cfg.get("pastas_entrada", []), cfg.get("extensoes", EXTENSOES_PADRAO),
        excluir=[saida], ignorados=ignorados,
    )
    dados = {
        "arquivos": arquivos,
        "sem_permissao": sem_perm,
        "ignorados": ignorados,
        "stats": obter_estatisticas_arquivos(arquivos),
        "pastas_inexistentes": [p for p in cfg.get("pastas_entrada", []) if not Path(p).is_dir()],
    }
    _CACHE_VARREDURA.update(chave=chave, dados=dados)
    return dados


# ---------------------------------------------------------------------------
# Montagem da base para o Copilot
# ---------------------------------------------------------------------------
def sanitizar_nome_arquivo(nome: str) -> str:
    invalido = '<>:"/\\|?*'
    limpo = "".join("_" if (c in invalido or ord(c) < 32) else c for c in nome)
    return limpo.strip().strip(".") or "documento"


_ROTULOS_META = {
    "title": ("Título", "Title"),
    "author": ("Autor", "Author"),
    "subject": ("Assunto", "Subject"),
    "keywords": ("Palavras-chave", "Keywords"),
    "creationDate": ("Criado em", "Created"),
    "modDate": ("Modificado em", "Modified"),
}


def montar_documentos(resultados: list[tuple], lang: str) -> list[dict]:
    """Ordena por pasta/caminho (ordenação natural), atribui IDs e marca duplicados."""
    resultados = sorted(resultados, key=lambda r: chave_ordenacao_natural(r[1]))
    largura = max(3, len(str(len(resultados))))
    docs = []
    por_hash: dict[str, str] = {}
    por_texto: dict[str, str] = {}
    for n, (nome, rel, grupo, meta) in enumerate(resultados, start=1):
        doc_id = f"[{n:0{largura}d}]"
        nome_nfc = normalizar_unicode(nome)
        rel_nfc = normalizar_unicode(rel)
        grupo_nfc = normalizar_unicode(grupo)
        doc = {
            "id": doc_id, "num": n, "nome": nome_nfc, "rel": rel_nfc, "grupo": grupo_nfc,
            "ext": Path(nome_nfc).suffix.lower(), "meta": meta,
            "dup_de": None, "dup_tipo": None, "semelhantes": [],
        }
        f_hash = meta.get("hash", "")
        conteudo = (meta.get("conteudo") or "").strip()
        h_txt = hashlib.sha256(re.sub(r"\s+", " ", conteudo).encode("utf-8")).hexdigest() if conteudo else ""
        if f_hash and f_hash in por_hash:
            doc["dup_de"], doc["dup_tipo"] = por_hash[f_hash], "binario"
        elif h_txt and h_txt in por_texto:
            doc["dup_de"], doc["dup_tipo"] = por_texto[h_txt], "texto"
        else:
            if f_hash:
                por_hash[f_hash] = doc_id
            if h_txt:
                por_texto[h_txt] = doc_id
        docs.append(doc)
    return docs


def detectar_quase_duplicados(docs: list[dict], limite_docs: int = 20000) -> list[tuple[dict, dict, float]]:
    """Pares de documentos com >= 80% de vocabulário em comum e tamanhos parecidos (outras versões)."""
    itens = []
    for d in docs:
        if d["dup_de"]:
            continue
        txt = d["meta"].get("conteudo") or ""
        palavras = set(re.findall(r"\b[a-zA-ZÀ-ÿ]{4,}\b", txt.lower()))
        if len(palavras) >= 30:
            itens.append((len(txt), d, palavras))
    if len(itens) > limite_docs:
        return []
    itens.sort(key=lambda x: x[0])

    candidatos_por_i: dict[int, set[int]] | None = None
    if len(itens) > 150:
        # Índice invertido pelas palavras mais raras de cada documento (garante 0 falsos negativos para >= 80%)
        freq = Counter(p for _, _, w in itens for p in w)
        indice_inv: dict[str, list[int]] = {}
        for idx_i, (_, _, w) in enumerate(itens):
            raras = sorted(w, key=lambda x: (freq[x], x))[: max(2, int(len(w) * 0.20) + 2)]
            for pal in raras:
                indice_inv.setdefault(pal, []).append(idx_i)
        candidatos_por_i = {}
        for idx_j, (_, _, w_j) in enumerate(itens):
            for pal in w_j:
                for idx_i in indice_inv.get(pal, ()):
                    if idx_i < idx_j:
                        candidatos_por_i.setdefault(idx_i, set()).add(idx_j)
                    elif idx_j < idx_i:
                        candidatos_por_i.setdefault(idx_j, set()).add(idx_i)

    pares = []
    for i in range(len(itens)):
        l1, d1, w1 = itens[i]
        alvos = sorted(candidatos_por_i.get(i, ())) if candidatos_por_i is not None else range(i + 1, len(itens))
        for j in alvos:
            l2, d2, w2 = itens[j]
            if l1 < 0.6 * l2:
                break  # ordenado por tamanho: os próximos são ainda maiores
            sobreposicao = len(w1 & w2) / min(len(w1), len(w2))
            if sobreposicao >= 0.80:
                a, b = (d1, d2) if d1["num"] < d2["num"] else (d2, d1)
                pares.append((a, b, sobreposicao))
    pares.sort(key=lambda p: (p[0]["num"], p[1]["num"]))
    for a, b, _ in pares:
        a["semelhantes"].append(b["id"])
        b["semelhantes"].append(a["id"])
    return pares


def formatar_diagramas_no_texto(
    texto: str,
    diagramas: list[dict],
    doc: dict,
    lang: str,
    em_subpasta: bool = False,
    salvar_diag: bool = True,
) -> str:
    """Substitui marcadores <!--DIAGRAMA_REF:...--> por cards estruturados com links para a pasta ordenada."""
    if not diagramas or not texto:
        return texto
    for diag in diagramas:
        marcador = diag.get("marcador")
        if not marcador or marcador not in texto:
            continue
        if salvar_diag and diag.get("rel_path"):
            caminho_rel = diag["rel_path_sub"] if em_subpasta else diag["rel_path"]
            tipo_loc = diag.get("tipo_local", "p")
            num_loc = diag.get("num_local", 1)
            loc_nome = tx(lang, "Página", "Page") if tipo_loc == "p" else (tx(lang, "Slide", "Slide") if tipo_loc == "s" else tx(lang, "Aba", "Sheet"))
            w, h = diag.get("largura", 0), diag.get("altura", 0)
            dims_str = f"{w}x{h} px" if w and h else ""
            rotulos = diag.get("rotulos", [])
            rot_str = " · ".join(rotulos[:12]) if rotulos else ""

            linhas_card = [
                f"> 📐 **[{tx(lang, 'Figura / Diagrama Referenciado', 'Referenced Diagram / Figure')}: `{caminho_rel}`]**",
                f"> - **{tx(lang, 'Origem', 'Source')}:** `{doc.get('rel', doc.get('nome', ''))}` ({loc_nome} {num_loc})",
            ]
            if dims_str:
                linhas_card.append(f"> - **{tx(lang, 'Dimensões', 'Dimensions')}:** {dims_str}")
            txt_diag = (diag.get("texto_ocr") or "").strip()
            if txt_diag:
                linhas_card.append(f"> - **{tx(lang, 'Texto / Rótulos detectados', 'Detected text / labels')}:**")
                for ln_ocr in txt_diag.splitlines():
                    linhas_card.append(f">   {ln_ocr}" if ln_ocr.strip() else ">")
            elif rot_str:
                linhas_card.append(f"> - **{tx(lang, 'Rótulos / Textos detectados:', 'Detected labels / text:')}** *{rot_str}*")
            linhas_card.append(">")
            linhas_card.append(f"> ![{tx(lang, 'Diagrama', 'Diagram')}: {doc.get('nome', '')} - {loc_nome} {num_loc}]({caminho_rel})")
            card_md = "\n".join(linhas_card)
            texto = texto.replace(marcador, card_md)
        else:
            rotulos = diag.get("rotulos", [])
            txt_diag = (diag.get("texto_ocr") or "").strip() or "\n".join(rotulos)
            if txt_diag:
                lbl = tx(lang, "Imagem embutida — texto via OCR", "Embedded image — OCR text")
                texto = texto.replace(marcador, citar_ocr(lbl, txt_diag))
            else:
                texto = texto.replace(marcador, "")
    return texto


def render_catalogo_diagramas(catalogo: list[tuple[dict, dict]], lang: str, em_subpasta: bool = False) -> str:
    """Gera tabela consolidada de todos os diagramas e esquemas encontrados."""
    if not catalogo:
        return ""
    saida = [
        f"### {tx(lang, 'Catálogo de Diagramas e Esquemas Técnicos', 'Catalog of Technical Diagrams and Schematics')} ({len(catalogo)} {tx(lang, 'encontrados', 'found')})",
        "",
        f"| Ref | {tx(lang, 'Documento de Origem', 'Source Document')} | {tx(lang, 'Localização', 'Location')} | {tx(lang, 'Arquivo da Imagem', 'Image File')} | {tx(lang, 'Rótulos / Descrição', 'Labels / Description')} |",
        "| :---: | :--- | :---: | :--- | :--- |",
    ]
    for n, (d, diag) in enumerate(catalogo, start=1):
        tipo_loc = diag.get("tipo_local", "p")
        num_loc = diag.get("num_local", 1)
        loc_nome = tx(lang, "Pág.", "Page") if tipo_loc == "p" else (tx(lang, "Slide", "Slide") if tipo_loc == "s" else tx(lang, "Aba", "Sheet"))
        rel_val = diag.get("rel_path_sub" if em_subpasta else "rel_path", "")
        rel_link = f"[`{rel_val}`]({rel_val})"
        rotulos = " · ".join(diag.get("rotulos", [])[:6])
        desc = f"*{escapar_celula(rotulos)}*" if rotulos else f"*{tx(lang, 'Esquema visual sem texto', 'Visual schematic without text')}*"
        saida.append(f"| `FIG-{n:03d}` | {d.get('id', '')} {escapar_celula(d.get('nome', ''))} | {loc_nome} {num_loc} | {rel_link} | {desc} |")
    saida.append("")
    return "\n".join(saida) + "\n"


def _linha_tipo_doc(doc: dict, lang: str) -> str:
    m = doc["meta"]
    ext = doc["ext"]
    meta_info = m.get("meta_info", {}) or {}
    partes = []
    tipo = nome_tipo(ext, lang)
    if ext in EXTENSOES_IMAGEM:
        extra = ", ".join(v for v in (meta_info.get("formato"), meta_info.get("resolucao")) if v)
        tipo += f" ({extra})" if extra else ""
    partes.append(f"**{tx(lang, 'Tipo', 'Type')}:** {tipo}")
    total = int(m.get("total_paginas", 1) or 1)
    if ext == ".pdf":
        partes.append(f"**{tx(lang, 'Páginas', 'Pages')}:** {total}")
    elif ext == ".pptx":
        partes.append(f"**Slides:** {total}")
    elif ext in (".xlsx", ".xlsm", ".xls"):
        partes.append(f"**{tx(lang, 'Abas', 'Sheets')}:** {total}")
    elif ext in EXTENSOES_IMAGEM and total > 1:
        partes.append(f"**{tx(lang, 'Páginas', 'Pages')}:** {total}")
    if m.get("ocr_pages", 0):
        conf = m.get("confianca_ocr", 0.0) or 0.0
        txt_ocr = f"**OCR:** {m['ocr_pages']} {tx(lang, 'trecho(s)', 'segment(s)')}"
        if meta_info.get("paginas_ocr"):
            txt_ocr += f" ({tx(lang, 'páginas digitalizadas', 'scanned pages')}: {meta_info['paginas_ocr']})"
        if conf:
            txt_ocr += f", {tx(lang, 'confiança média', 'average confidence')} {conf * 100:.0f}%"
        partes.append(txt_ocr)
    if m.get("diagramas"):
        partes.append(f"**{tx(lang, 'Diagramas', 'Diagrams')}:** {len(m['diagramas'])}")
    return " · ".join(partes)


def render_documento(doc: dict, lang: str, omitir_duplicados: bool = True, em_subpasta: bool = False, salvar_diag: bool = True) -> str:
    m = doc["meta"]
    meta_info = m.get("meta_info", {}) or {}
    linhas = [f"## {doc['id']} {doc['nome']}", ""]
    linhas.append(f"- **{tx(lang, 'Origem', 'Source')}:** `{doc['rel']}`")
    linhas.append(f"- {_linha_tipo_doc(doc, lang)}")
    detalhes = [f"{tx(lang, *_ROTULOS_META[k])}: {escapar_celula(v)}" for k, v in meta_info.items() if k in _ROTULOS_META]
    if detalhes:
        linhas.append(f"- **{tx(lang, 'Metadados', 'Metadata')}:** " + " · ".join(detalhes))
    if doc["dup_de"]:
        motivo = tx(lang, "arquivo idêntico a", "identical file to") if doc["dup_tipo"] == "binario" else tx(lang, "mesmo texto que", "same text as")
        linhas.append(f"- **{tx(lang, 'Duplicado', 'Duplicate')}:** {motivo} {doc['dup_de']}")
    if doc["semelhantes"]:
        linhas.append(f"- **{tx(lang, 'Atenção', 'Note')}:** "
                      + tx(lang, "conteúdo muito semelhante a ", "content very similar to ")
                      + ", ".join(doc["semelhantes"])
                      + tx(lang, " (possível outra versão do mesmo documento).", " (possibly another version of the same document)."))
    if meta_info.get("aviso"):
        linhas.append(f"- **{tx(lang, 'Aviso', 'Warning')}:** {meta_info['aviso']}")
    linhas.append("")

    conteudo = (m.get("conteudo") or "").strip()
    if doc["dup_de"] and omitir_duplicados:
        linhas.append(tx(lang, f"*(Conteúdo omitido para evitar redundância: consulte o documento {doc['dup_de']}.)*",
                         f"*(Content omitted to avoid redundancy: see document {doc['dup_de']}.)*"))
    elif conteudo:
        conteudo_fmt = formatar_diagramas_no_texto(conteudo, m.get("diagramas", []), doc, lang, em_subpasta, salvar_diag)
        linhas.append(conteudo_fmt)
    else:
        linhas.append(tx(lang, "*(Nenhum texto extraível neste documento.)*", "*(No extractable text in this document.)*"))
    linhas += ["", f"*[{tx(lang, 'Fim do documento', 'End of document')} {doc['id']}]*", "", "---", ""]
    return "\n".join(linhas) + "\n"


def _observacoes_doc(doc: dict, lang: str) -> str:
    m = doc["meta"]
    obs = []
    if doc["dup_de"]:
        obs.append(tx(lang, f"duplicado de {doc['dup_de']}", f"duplicate of {doc['dup_de']}"))
    elif not (m.get("conteudo") or "").strip():
        obs.append(tx(lang, "sem texto", "no text"))
    if m.get("ocr_pages", 0):
        obs.append("OCR")
    if doc["semelhantes"]:
        obs.append(tx(lang, "semelhante a ", "similar to ") + ", ".join(doc["semelhantes"][:3]))
    return "; ".join(obs)


def render_indice(docs: list[dict], lang: str, titulo: str | None = None) -> str:
    """Índice agrupado por pasta de origem, em tabelas."""
    grupos: dict[str, list[dict]] = {}
    for d in docs:
        grupos.setdefault(d["grupo"], []).append(d)
    titulo = titulo or tx(lang, "ÍNDICE DE DOCUMENTOS", "DOCUMENT INDEX")
    saida = [f"## {titulo}", ""]
    for grupo, lista in grupos.items():
        qtd = f"{len(lista)} " + tx(lang, "documento(s)", "document(s)")
        saida += [f"### {tx(lang, 'Pasta', 'Folder')}: {grupo} ({qtd})", ""]
        saida.append(f"| ID | {tx(lang, 'Documento', 'Document')} | {tx(lang, 'Tipo', 'Type')} | {tx(lang, 'Observações', 'Notes')} |")
        saida.append("| --- | --- | --- | --- |")
        for d in lista:
            prefixo = f"{grupo}/"
            caminho = d["rel"][len(prefixo):] if d["rel"].startswith(prefixo) else d["rel"]
            saida.append(f"| {d['id']} | {escapar_celula(caminho)} | {nome_tipo(d['ext'], lang)} | {escapar_celula(_observacoes_doc(d, lang))} |")
        saida.append("")
    return "\n".join(saida) + "\n"


def render_preambulo(lang: str, resumo: dict, titulo_extra: str = "", instrucoes: bool = True) -> str:
    titulo = tx(lang, "Base documental consolidada", "Consolidated document base") + titulo_extra
    saida = [f"# {titulo}", ""]
    if instrucoes:
        saida.append(f"> **{tx(lang, 'Instruções para o assistente (Copilot)', 'Instructions for the assistant (Copilot)')}**")
        saida.append(">")
        for i, ln in enumerate(_PROMPT_LINHAS_EN if lang == "en" else _PROMPT_LINHAS_PT):
            if i:
                saida.append(">")
            saida.append(f"> {ln}")
        saida.append("")
    saida.append(f"- **{tx(lang, 'Gerado em', 'Generated on')}:** {resumo['data']}")
    saida.append(f"- **{tx(lang, 'Documentos', 'Documents')}:** {fmt_int(resumo['docs'], lang)} "
                 f"({tx(lang, 'pastas de origem', 'source folders')}: {resumo['grupos']})")
    if resumo.get("duplicados"):
        saida.append(f"- **{tx(lang, 'Duplicados com conteúdo omitido', 'Duplicates with omitted content')}:** {resumo['duplicados']}")
    saida.append(f"- **{tx(lang, 'Volume original', 'Original volume')}:** {resumo['volume']}")
    saida.append(f"- **{tx(lang, 'Tamanho estimado', 'Estimated size')}:** ~{fmt_int(resumo['tokens'], lang)} tokens")
    saida.append("")
    return "\n".join(saida) + "\n"


def dividir_texto(texto: str, limite_bytes: int) -> list[str]:
    """Divide um texto em pedaços <= limite_bytes, preferindo parágrafos, depois linhas."""
    pedacos, atual, tam = [], [], 0

    def fechar():
        nonlocal atual, tam
        if atual:
            pedacos.append("".join(atual))
        atual, tam = [], 0

    for par in re.split(r"(?<=\n\n)", texto):
        b = len(par.encode("utf-8"))
        if b > limite_bytes:
            fechar()
            for ln in par.splitlines(keepends=True):
                bl = len(ln.encode("utf-8"))
                if bl > limite_bytes:
                    fechar()
                    passo = max(1, limite_bytes // 4)
                    for i in range(0, len(ln), passo):
                        pedacos.append(ln[i:i + passo])
                    continue
                if tam + bl > limite_bytes:
                    fechar()
                atual.append(ln)
                tam += bl
            fechar()
            continue
        if tam + b > limite_bytes:
            fechar()
        atual.append(par)
        tam += b
    fechar()
    return pedacos or [texto]


def particionar(blocos: list[tuple[dict, str]], limite_bytes: int, lang: str) -> list[list[tuple[dict, str]]]:
    """Agrupa os documentos em partes <= limite, em ordem; documentos maiores viram trechos."""
    partes, atual, tam = [], [], 0
    for doc, bloco in blocos:
        b = len(bloco.encode("utf-8"))
        if b > limite_bytes:
            if atual:
                partes.append(atual)
                atual, tam = [], 0
            pedacos = dividir_texto(bloco, limite_bytes)
            n = len(pedacos)
            for k, ped in enumerate(pedacos, start=1):
                if k > 1:
                    ped = (f"## {doc['id']} {doc['nome']} — "
                           + tx(lang, f"continuação (trecho {k} de {n})", f"continued (part {k} of {n})")
                           + "\n\n" + ped)
                partes.append([(doc, ped)])
            continue
        if atual and tam + b > limite_bytes:
            partes.append(atual)
            atual, tam = [], 0
        atual.append((doc, bloco))
        tam += b
    if atual:
        partes.append(atual)
    return partes


def compactar_arquivos(arquivos: list[Path], base: Path, destino_zip: Path) -> bool:
    try:
        with zipfile.ZipFile(destino_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for arq in arquivos:
                if arq.is_file() and arq != destino_zip:
                    z.write(arq, arq.relative_to(base).as_posix())
                elif arq.is_dir():
                    for sub in sorted(arq.rglob("*")):
                        if sub.is_file() and sub != destino_zip:
                            z.write(sub, sub.relative_to(base).as_posix())
        with zipfile.ZipFile(destino_zip, "r") as z:
            if z.testzip() is not None:
                return False
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}[ZIP] {e}{Cor.RESET}")
        return False


def compactar_pasta(origem: Path, destino_zip: Path) -> bool:
    """Compatibilidade: compacta todos os arquivos da pasta (exceto .zip)."""
    arquivos = [p for p in origem.rglob("*") if p.is_file() and p.suffix.lower() != ".zip"]
    return compactar_arquivos(arquivos, origem, destino_zip)


def abrir_no_explorador(caminho: Path):
    try:
        if platform.system() == "Windows":
            os.startfile(str(caminho))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(caminho)])
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception as e:
        print(f"{Cor.AMARELO}{e}{Cor.RESET}")


def gerar_saidas(docs: list[dict], cfg: dict, pasta_destino: Path, resumo: dict) -> tuple[list[tuple[Path, str]], dict]:
    """Grava todos os arquivos de saída. Retorna ([(caminho, descricao)], info)."""
    lang = obter_idioma(cfg)
    omitir = bool(cfg.get("omitir_duplicados", True))
    salvar_diag = bool(cfg.get("salvar_diagramas_em_pasta", True))
    nome_pasta_diag = cfg.get("pasta_diagramas", "imagens")
    pasta_diag = pasta_destino / nome_pasta_diag
    pasta_saida_base = resolver_caminho(cfg.get("pasta_saida", "./saida_copilot"))
    pasta_cache = pasta_saida_base / ".cache_extracao"
    blobs_dir = pasta_cache / "blobs"

    gerados: list[tuple[Path, str]] = []
    info = {"partes": 0}

    total_diag_salvos = 0
    catalogo_diagramas = []

    if salvar_diag:
        pasta_diag.mkdir(parents=True, exist_ok=True)
        for antigo in pasta_diag.glob("*.png"):
            antigo.unlink(missing_ok=True)

        for d in docs:
            m = d["meta"]
            for diag in m.get("diagramas", []):
                doc_num = d["num"]
                tipo_loc = diag.get("tipo_local", "p")
                num_loc = diag.get("num_local", 1)
                idx_fig = diag.get("indice", 1)
                nome_arq = f"doc{doc_num:03d}_{tipo_loc}{num_loc:02d}_fig{idx_fig:02d}.png"
                dest_file = pasta_diag / nome_arq
                src_blob = blobs_dir / f"{diag['id_blob']}.png"
                if src_blob.exists():
                    try:
                        shutil.copy2(src_blob, dest_file)
                        total_diag_salvos += 1
                    except Exception:
                        pass
                diag["nome_arquivo"] = nome_arq
                diag["rel_path"] = f"{nome_pasta_diag}/{nome_arq}"
                diag["rel_path_sub"] = f"../{nome_pasta_diag}/{nome_arq}"
                if not (d["dup_de"] and omitir):
                    catalogo_diagramas.append((d, diag))

    blocos = [(d, render_documento(d, lang, omitir, em_subpasta=False, salvar_diag=salvar_diag)) for d in docs]
    total_chars = sum(len(b) for _, b in blocos)
    resumo["tokens"] = total_chars // 4

    indice = render_indice(docs, lang)
    if catalogo_diagramas:
        indice += "\n" + render_catalogo_diagramas(catalogo_diagramas, lang)
    preambulo = render_preambulo(lang, resumo)

    def gravar(caminho: Path, conteudo: str, desc: str):
        caminho.parent.mkdir(parents=True, exist_ok=True)
        caminho.write_text(conteudo, encoding="utf-8")
        gerados.append((caminho, desc))

    # Limpa saídas antigas (relevante quando não se usa pasta com data/hora)
    for padrao in ("corpus_parte_*.md", "por_pasta/*.md", "por_lotes/*.md", "individual/*.md"):
        for antigo in pasta_destino.glob(padrao):
            antigo.unlink(missing_ok=True)

    gravar(pasta_destino / "00_INDICE_MESTRE.md", preambulo + indice,
           tx(lang, "catálogo de todos os documentos", "catalog of all documents"))

    if cfg.get("gerar_arquivo_unico", True):
        caminho = pasta_destino / "corpus_completo.md"
        with caminho.open("w", encoding="utf-8") as f:
            f.write(preambulo)
            f.write(indice)
            f.write("\n---\n\n")
            for _, b in blocos:
                f.write(b)
        gerados.append((caminho, tx(lang, "BASE ÚNICA com todos os documentos", "SINGLE BASE with all documents")))

    if cfg.get("dividir_por_tamanho", True) and blocos:
        limite = int(float(cfg.get("limite_mb_por_parte", 8.0)) * 1024 * 1024 * 0.95)
        partes = particionar(blocos, max(limite, 64 * 1024), lang)
        if len(partes) > 1 or not cfg.get("gerar_arquivo_unico", True):
            n = len(partes)
            info["partes"] = n
            for k, parte in enumerate(partes, start=1):
                docs_parte = []
                for d, _ in parte:
                    if not docs_parte or docs_parte[-1] is not d:
                        docs_parte.append(d)
                ids_parte = {id(d) for d in docs_parte}
                diag_parte = [(d, diag) for d, diag in catalogo_diagramas if id(d) in ids_parte]
                titulo = tx(lang, f" — Parte {k} de {n}", f" — Part {k} of {n}")
                cab = render_preambulo(lang, resumo, titulo_extra=titulo)
                cab += (tx(lang, f"> Esta é a parte {k} de {n}. O catálogo completo está em `00_INDICE_MESTRE.md`.",
                           f"> This is part {k} of {n}. The full catalog is in `00_INDICE_MESTRE.md`.") + "\n\n")
                cab += render_indice(docs_parte, lang, titulo=tx(lang, "DOCUMENTOS NESTA PARTE", "DOCUMENTS IN THIS PART"))
                if diag_parte:
                    cab += "\n" + render_catalogo_diagramas(diag_parte, lang, em_subpasta=False)
                gravar(pasta_destino / f"corpus_parte_{k:02d}.md", cab + "\n---\n\n" + "".join(b for _, b in parte),
                       tx(lang, f"parte {k} de {n}", f"part {k} of {n}"))

    if cfg.get("gerar_por_pasta_raiz", True) and docs:
        grupos: dict[str, list[tuple[dict, str]]] = {}
        for d in docs:
            b_sub = render_documento(d, lang, omitir, em_subpasta=True, salvar_diag=salvar_diag)
            grupos.setdefault(d["grupo"], []).append((d, b_sub))
        for i, (grupo, lista) in enumerate(grupos.items(), start=1):
            docs_grupo = [d for d, _ in lista]
            ids_grupo = {id(d) for d in docs_grupo}
            diag_grupo = [(d, diag) for d, diag in catalogo_diagramas if id(d) in ids_grupo]
            titulo = f" — {tx(lang, 'Pasta', 'Folder')}: {grupo}"
            cab = render_preambulo(lang, dict(resumo, docs=len(lista), grupos=1,
                                              duplicados=sum(1 for d, _ in lista if d["dup_de"]) if omitir else 0,
                                              tokens=sum(len(b) for _, b in lista) // 4), titulo_extra=titulo)
            cab += render_indice(docs_grupo, lang)
            if diag_grupo:
                cab += "\n" + render_catalogo_diagramas(diag_grupo, lang, em_subpasta=True)
            gravar(pasta_destino / "por_pasta" / f"{i:02d}_{sanitizar_nome_arquivo(grupo)}.md",
                   cab + "\n---\n\n" + "".join(b for _, b in lista), tx(lang, f"pasta {grupo}", f"folder {grupo}"))

    if cfg.get("dividir_por_quantidade", False) and docs:
        qtd = max(1, int(cfg.get("qtd_docs_por_parte", 50)))
        for k, i in enumerate(range(0, len(blocos), qtd), start=1):
            fatia = blocos[i:i + qtd]
            docs_lote = [d for d, _ in fatia]
            ids_lote = {id(d) for d in docs_lote}
            diag_lote = [(d, diag) for d, diag in catalogo_diagramas if id(d) in ids_lote]
            titulo = tx(lang, f" — Lote {k}", f" — Batch {k}")
            cab = render_preambulo(lang, dict(resumo, docs=len(fatia), grupos=len({d["grupo"] for d, _ in fatia}),
                                              duplicados=sum(1 for d, _ in fatia if d["dup_de"]) if omitir else 0,
                                              tokens=sum(len(b) for _, b in fatia) // 4), titulo_extra=titulo)
            cab += render_indice(docs_lote, lang)
            if diag_lote:
                cab += "\n" + render_catalogo_diagramas(diag_lote, lang, em_subpasta=True)
            nome = f"lote_{k:02d}_docs_{fatia[0][0]['num']:03d}_a_{fatia[-1][0]['num']:03d}.md"
            gravar(pasta_destino / "por_lotes" / nome, cab + "\n---\n\n" + "".join(b for _, b in fatia),
                   tx(lang, f"lote {k}", f"batch {k}"))

    if cfg.get("gerar_individual_por_arquivo", False) and docs:
        for d in docs:
            b_ind = render_documento(d, lang, omitir, em_subpasta=True, salvar_diag=salvar_diag)
            nome = sanitizar_nome_arquivo(f"{d['id'].strip('[]')}_{d['nome']}") + ".md"
            gravar(pasta_destino / "individual" / nome, b_ind, d["rel"])

    if cfg.get("gerar_jsonl", False):
        caminho = pasta_destino / "corpus_rag.jsonl"
        with caminho.open("w", encoding="utf-8") as f:
            for d in docs:
                if d["dup_de"] and omitir:
                    continue
                m = d["meta"]
                diagramas_doc = m.get("diagramas", [])
                paginas = m.get("paginas_rag") or ([{"pagina": 1, "texto": m.get("conteudo", "")}] if m.get("conteudo") else [])
                for p in paginas:
                    pag_num = p.get("pagina", 1)
                    imgs_rag = [
                        {
                            "arquivo": diag.get("rel_path", ""),
                            "rotulos": diag.get("rotulos", []),
                            "largura": diag.get("largura", 0),
                            "altura": diag.get("altura", 0),
                        }
                        for diag in diagramas_doc if diag.get("num_local") == pag_num and diag.get("rel_path")
                    ]
                    txt_p = formatar_diagramas_no_texto(p.get("texto", ""), diagramas_doc, d, lang, em_subpasta=False, salvar_diag=salvar_diag)
                    registro = {
                        "id": d["id"],
                        "arquivo": d["nome"],
                        "caminho_relativo": d["rel"],
                        "departamento": d["grupo"],
                        "tipo": nome_tipo(d["ext"], lang),
                        "pagina": pag_num,
                        "total_paginas": m.get("total_paginas", 1),
                        "metadados": {k: v for k, v in (m.get("meta_info") or {}).items() if k in _ROTULOS_META},
                        "semelhante_a": d["semelhantes"],
                        "texto": txt_p,
                    }
                    if imgs_rag:
                        registro["imagens_referenciadas"] = imgs_rag
                    f.write(json.dumps(registro, ensure_ascii=False) + "\n")
        gerados.append((caminho, tx(lang, "base vetorial para RAG (uma linha por página)", "RAG dataset (one line per page)")))

    if salvar_diag and total_diag_salvos > 0:
        gerados.append((pasta_diag, tx(lang, f"pasta de diagramas ({total_diag_salvos} imagens)", f"diagrams folder ({total_diag_salvos} images)")))

    nome_prompt = "COPILOT_PROMPT.txt" if lang == "en" else "PROMPT_PARA_COPILOT.txt"
    gravar(pasta_destino / nome_prompt, PROMPT_SISTEMA_COPILOT_EN if lang == "en" else PROMPT_SISTEMA_COPILOT_PT,
           tx(lang, "instruções para colar no chat do Copilot", "instructions to paste into the Copilot chat"))
    return gerados, info


def construir_relatorio(lang: str, contexto: dict) -> str:
    c = contexto
    docs = c["docs"]
    L = []
    L.append(f"# {tx(lang, 'Relatório de auditoria da consolidação', 'Consolidation audit report')}")
    L.append("")
    L.append(f"- **{tx(lang, 'Data', 'Date')}:** {agora_fmt(lang)}")
    L.append(f"- **{tx(lang, 'Tempo de execução', 'Run time')}:** {c['duracao']:.1f} s ({c['workers']} {tx(lang, 'processos', 'workers')})")
    L.append(f"- **{tx(lang, 'Arquivos encontrados', 'Files found')}:** {c['total']} ({c['volume']})")
    L.append(f"- **{tx(lang, 'Indexados com sucesso', 'Indexed successfully')}:** {len(docs)}")
    L.append(f"- **{tx(lang, 'Reaproveitados do cache', 'Reused from cache')}:** {c['cache']}")
    L.append(f"- **{tx(lang, 'Falhas de leitura', 'Read failures')}:** {len(c['erros'])}")
    L.append(f"- **{tx(lang, 'Documentos sem texto', 'Documents without text')}:** {len(c['sem_texto'])}")
    L.append(f"- **{tx(lang, 'Duplicados', 'Duplicates')}:** {sum(1 for d in docs if d['dup_de'])}"
             + (tx(lang, " (conteúdo omitido na base)", " (content omitted from the base)") if c["omitir"] else ""))
    L.append(f"- **{tx(lang, 'Pares de versões semelhantes', 'Similar version pairs')}:** {len(c['quase'])}")
    L.append(f"- **OCR:** {c['ocr']['ok']} {tx(lang, 'com texto', 'with text')} · {c['ocr']['empty']} "
             f"{tx(lang, 'sem texto legível', 'without readable text')} · {c['ocr']['error']} {tx(lang, 'falhas técnicas', 'technical failures')}")
    if c["mojibake"]:
        L.append(f"- **{tx(lang, 'Acentuação corrompida corrigida (mojibake)', 'Corrupted encoding fixed (mojibake)')}:** {c['mojibake']}")
    L.append(f"- **{tx(lang, 'Anonimização de dados pessoais (PII)', 'Personal data redaction (PII)')}:** "
             + (tx(lang, "ATIVADA", "ENABLED") if c["pii"] else tx(lang, "desativada (dados originais preservados)", "disabled (raw data preserved)")))
    L.append(f"- **{tx(lang, 'Tamanho estimado da base', 'Estimated base size')}:** ~{fmt_int(c['tokens'], lang)} tokens "
             + tx(lang, "(estimativa: 4 caracteres por token)", "(estimate: 4 characters per token)"))
    L.append("")

    L.append(f"## {tx(lang, 'Qual arquivo enviar ao Copilot', 'Which file to upload to Copilot')}")
    L.append("")
    if c["partes"] > 1:
        L.append(tx(lang,
                    f"- A base foi dividida em **{c['partes']} partes** (`corpus_parte_XX.md`) de até {c['limite_mb']} MB. "
                    "Envie as partes ao Copilot; se o Copilot aceitar o arquivo inteiro, use `corpus_completo.md`.",
                    f"- The base was split into **{c['partes']} parts** (`corpus_parte_XX.md`) of up to {c['limite_mb']} MB. "
                    "Upload the parts to Copilot; if Copilot accepts the full file, use `corpus_completo.md`."))
    else:
        L.append(tx(lang, "- Envie `corpus_completo.md`: ele contém o índice e todos os documentos.",
                    "- Upload `corpus_completo.md`: it contains the index and every document."))
    L.append(tx(lang, "- Cole o texto de `PROMPT_PARA_COPILOT.txt` no chat para reforçar as regras de citação.",
                "- Paste the text of `COPILOT_PROMPT.txt` into the chat to reinforce the citation rules."))
    if c["tokens"] > 128_000:
        L.append(tx(lang,
                    f"- ⚠️ A base tem ~{fmt_int(c['tokens'], lang)} tokens. Para perguntas sobre um tema específico, "
                    "anexar só a pasta relevante (`por_pasta/`) costuma dar respostas mais precisas.",
                    f"- ⚠️ The base has ~{fmt_int(c['tokens'], lang)} tokens. For questions on a specific topic, "
                    "attaching only the relevant folder file (`por_pasta/`) usually gives more precise answers."))
    L.append("")

    L.append(f"## {tx(lang, 'Arquivos gerados', 'Generated files')}")
    L.append("")
    subpastas: dict[str, list[Path]] = {}
    for caminho, desc in c["gerados"]:
        if caminho.parent != c["pasta"]:
            subpastas.setdefault(caminho.parent.name, []).append(caminho)
            continue
        try:
            tam = formatar_tamanho_bytes(caminho.stat().st_size)
        except OSError:
            tam = "?"
        L.append(f"- `{caminho.name}` ({tam}): {desc}")
    for nome_sp, lista in subpastas.items():
        tam = formatar_tamanho_bytes(sum(p.stat().st_size for p in lista if p.exists()))
        L.append(f"- `{nome_sp}/` ({len(lista)} {tx(lang, 'arquivos', 'files')}, {tam})")
    L.append("")

    L.append(f"## {tx(lang, 'Por tipo de arquivo', 'By file type')}")
    L.append("")
    por_tipo: dict[str, list] = {}
    for d in docs:
        por_tipo.setdefault(d["ext"], []).append(d)
    linhas_t = [[tx(lang, "Extensão", "Extension"), tx(lang, "Documentos", "Documents"), tx(lang, "Com OCR", "With OCR"), tx(lang, "Sem texto", "No text")]]
    for ext, lista in sorted(por_tipo.items(), key=lambda x: -len(x[1])):
        linhas_t.append([ext, str(len(lista)), str(sum(1 for d in lista if d["meta"].get("ocr_pages"))),
                         str(sum(1 for d in lista if not (d["meta"].get("conteudo") or "").strip()))])
    L.append(tabela_para_markdown(linhas_t))
    L.append("")

    if c["erros"]:
        L.append(f"## {tx(lang, 'Falhas de leitura', 'Read failures')}")
        L.append("")
        for rel, err in sorted(c["erros"]):
            L.append(f"- `{rel}`: {err}")
        L.append("")
    if c["sem_texto"]:
        L.append(f"## {tx(lang, 'Documentos sem texto aproveitável', 'Documents without usable text')}")
        L.append("")
        for d in c["sem_texto"]:
            L.append(f"- {d['id']} `{d['rel']}`: {d['meta'].get('meta_info', {}).get('aviso', '')}")
        L.append("")
    baixa = [d for d in docs if d["meta"].get("ocr_pages") and 0 < (d["meta"].get("confianca_ocr") or 0) < 0.75]
    if baixa:
        L.append(f"## {tx(lang, 'OCR com baixa confiança (revise)', 'Low-confidence OCR (review)')}")
        L.append("")
        for d in baixa:
            L.append(f"- {d['id']} `{d['rel']}`: {d['meta']['confianca_ocr'] * 100:.0f}%")
        L.append("")
    dups = [d for d in docs if d["dup_de"]]
    if dups:
        L.append(f"## {tx(lang, 'Duplicados', 'Duplicates')}")
        L.append("")
        for d in dups:
            L.append(f"- {d['id']} `{d['rel']}` = {d['dup_de']}")
        L.append("")
    if c["quase"]:
        L.append(f"## {tx(lang, 'Versões semelhantes (≥ 80% do vocabulário em comum)', 'Similar versions (≥ 80% shared vocabulary)')}")
        L.append("")
        for a, b, s in c["quase"][:300]:
            L.append(f"- {a['id']} `{a['rel']}` ↔ {b['id']} `{b['rel']}`: {s * 100:.0f}%")
        L.append("")
    if c["ignorados"]:
        L.append(f"## {tx(lang, 'Arquivos ignorados (formato não suportado)', 'Skipped files (unsupported format)')}")
        L.append("")
        for ext, qtd in sorted(c["ignorados"].items(), key=lambda x: -x[1])[:25]:
            dica = DICAS_FORMATOS_NAO_SUPORTADOS.get(ext)
            dica_txt = f" — {dica[1] if lang == 'en' else dica[0]}" if dica else ""
            L.append(f"- `{ext}`: {qtd}{dica_txt}")
        L.append("")
    if c["sem_permissao"]:
        L.append(f"## {tx(lang, 'Sem permissão de leitura', 'No read permission')}")
        L.append("")
        for p in c["sem_permissao"][:100]:
            L.append(f"- `{p}`")
        L.append("")
    return "\n".join(L)


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------
def _cabecalho(cfg: dict, titulo: str):
    limpar_tela()
    print(linha())
    print(f"{Cor.TITULO}  {titulo}{Cor.RESET}")
    print(linha())


def _resumo_config_execucao(cfg: dict, lang: str, max_workers: int) -> list[tuple[str, str]]:
    s, n = t("sim", cfg), t("nao", cfg)
    saidas = []
    if cfg.get("gerar_arquivo_unico", True):
        saidas.append(tx(lang, "base única", "single base"))
    if cfg.get("dividir_por_tamanho", True):
        saidas.append(tx(lang, f"partes de {cfg.get('limite_mb_por_parte', 8.0)} MB", f"{cfg.get('limite_mb_por_parte', 8.0)} MB parts"))
    if cfg.get("gerar_por_pasta_raiz", True):
        saidas.append(tx(lang, "por pasta", "per folder"))
    if cfg.get("dividir_por_quantidade", False):
        saidas.append(tx(lang, f"lotes de {cfg.get('qtd_docs_por_parte', 50)}", f"batches of {cfg.get('qtd_docs_por_parte', 50)}"))
    if cfg.get("gerar_individual_por_arquivo", False):
        saidas.append("1:1")
    if cfg.get("gerar_jsonl", False):
        saidas.append("RAG JSONL")
    ocr = (tx(lang, f"ativado (DPI {cfg.get('dpi_ocr_pdf', 200)})", f"enabled (DPI {cfg.get('dpi_ocr_pdf', 200)})")
           if cfg.get("usar_ocr_em_pdf", True) else tx(lang, "desativado", "disabled"))
    return [
        (tx(lang, "Saídas", "Outputs"), ", ".join(saidas) or "-"),
        ("OCR", ocr),
        (tx(lang, "Dados pessoais (PII)", "Personal data (PII)"),
         tx(lang, "MASCARADOS", "MASKED") if cfg.get("mascarar_pii", False) else tx(lang, "preservados (originais)", "preserved (raw)")),
        (tx(lang, "Duplicados", "Duplicates"), tx(lang, "conteúdo omitido", "content omitted") if cfg.get("omitir_duplicados", True) else tx(lang, "mantidos", "kept")),
        (tx(lang, "Processos paralelos", "Parallel workers"), str(max_workers)),
        (tx(lang, "Compactar .zip", "Create .zip"), s if cfg.get("gerar_arquivo_zip", False) else n),
        (tx(lang, "Destino", "Destination"), str(resolver_caminho(cfg.get("pasta_saida", "./saida_copilot")))),
    ]


def executar_pipeline(cfg: dict, pausar_no_final: bool = True, interativo: bool | None = None) -> dict | None:
    """Varre, extrai, consolida e grava a base. Retorna um resumo (ou None se cancelado/vazio)."""
    if interativo is None:
        interativo = "--auto" not in sys.argv
    lang = obter_idioma(cfg)
    _cabecalho(cfg, t("titulo_app", cfg))

    pastas = cfg.get("pastas_entrada", [])
    if not pastas:
        print(f"{Cor.VERMELHO}{t('sem_pastas', cfg)}{Cor.RESET}")
        if pausar_no_final and interativo:
            pausar(lang)
        return None

    print(tx(lang, "Varrendo pastas...", "Scanning folders..."))
    varredura = varrer_pastas(cfg, forcar=True)
    arquivos = varredura["arquivos"]
    stats = varredura["stats"]
    total = len(arquivos)
    if total == 0:
        print(f"{Cor.VERMELHO}{t('sem_arquivos', cfg)}{Cor.RESET}")
        if pausar_no_final and interativo:
            pausar(lang)
        return None

    num_cpus = os.cpu_count() or 4

    def workers_efetivos():
        return max(1, min(int(cfg.get("max_workers", min(4, num_cpus))), num_cpus)) if cfg.get("multicore", True) else 1

    if interativo:
        while True:
            _cabecalho(cfg, tx(lang, "CONFIRMAR GERAÇÃO DA BASE", "CONFIRM BASE GENERATION"))
            print(f"\n {Cor.NEGRITO}{tx(lang, 'O que será processado', 'What will be processed')}{Cor.RESET}")
            print(f"  • {tx(lang, 'Pastas de entrada', 'Input folders'):<24}: {len(pastas)}")
            print(f"  • {tx(lang, 'Arquivos', 'Files'):<24}: {Cor.VERDE}{Cor.NEGRITO}{fmt_int(total, lang)}{Cor.RESET} ({stats['tamanho_formatado']})")
            for ext, inf in sorted(stats["por_extensao"].items(), key=lambda x: -x[1]["qtd"]):
                print(f"      {ext:<6} {inf['qtd']:>6}  ({formatar_tamanho_bytes(inf['bytes'])})")
            if varredura["ignorados"]:
                qtd_ign = sum(varredura["ignorados"].values())
                print(f"  • {Cor.AMARELO}{tx(lang, 'Ignorados (formato não suportado/vazios)', 'Skipped (unsupported/empty)')}: {qtd_ign}{Cor.RESET}")
            if varredura["pastas_inexistentes"]:
                print(f"  • {Cor.VERMELHO}{tx(lang, 'Pastas inexistentes', 'Missing folders')}: {len(varredura['pastas_inexistentes'])}{Cor.RESET}")
            print(f"\n {Cor.NEGRITO}{tx(lang, 'Configuração', 'Settings')}{Cor.RESET}")
            for rot, val in _resumo_config_execucao(cfg, lang, workers_efetivos()):
                print(f"  • {rot:<24}: {val}")
            print(linha(char="-"))
            print(f" {Cor.VERDE}{Cor.NEGRITO}[ENTER] {tx(lang, 'Iniciar', 'Start')}{Cor.RESET}")
            print(f" [M] {tx(lang, 'Tamanho máximo de cada parte (MB)', 'Maximum size of each part (MB)')}")
            print(f" [W] {tx(lang, 'Processos paralelos', 'Parallel workers')}")
            print(f" [P] {tx(lang, 'Ligar/desligar anonimização de PII', 'Toggle PII redaction')}")
            print(f" [O] {tx(lang, 'Ligar/desligar OCR', 'Toggle OCR')}")
            print(f" [Z] {tx(lang, 'Ligar/desligar .zip', 'Toggle .zip')}")
            print(f" [0] {tx(lang, 'Cancelar', 'Cancel')}")
            print(linha())
            op = perguntar(f"\n{t('prompt_opcao', cfg)}").strip().upper()
            if op in ("", "C", "S", "Y"):
                break
            if op in ("0", "V"):
                return None
            if op == "M":
                valor = perguntar(tx(lang, "MB por parte (ex.: 8, 20, 50): ", "MB per part (e.g. 8, 20, 50): ")).strip().replace(",", ".")
                try:
                    if float(valor) >= 0.5:
                        cfg["limite_mb_por_parte"] = float(valor)
                        cfg["dividir_por_tamanho"] = True
                except ValueError:
                    pass
            elif op == "W":
                valor = perguntar(tx(lang, f"Processos (1 a {num_cpus}): ", f"Workers (1 to {num_cpus}): ")).strip()
                if valor.isdigit() and 1 <= int(valor) <= num_cpus:
                    cfg["max_workers"] = int(valor)
                    cfg["multicore"] = True
            elif op == "P":
                cfg["mascarar_pii"] = not cfg.get("mascarar_pii", False)
            elif op == "O":
                novo = not cfg.get("usar_ocr_em_pdf", True)
                cfg["usar_ocr_em_pdf"] = novo
                cfg["usar_ocr_em_imagens_embutidas"] = novo
            elif op == "Z":
                cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", False)
            salvar_config(cfg)

    max_workers = workers_efetivos()
    pasta_saida_base = resolver_caminho(cfg.get("pasta_saida", "./saida_copilot"))
    pasta_cache = pasta_saida_base / ".cache_extracao"
    pasta_cache.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    pasta_destino = pasta_saida_base / f"processado_{carimbo}" if cfg.get("usar_pasta_timestamp", True) else pasta_saida_base
    pasta_destino.mkdir(parents=True, exist_ok=True)

    _cabecalho(cfg, tx(lang, "GERANDO A BASE PARA O COPILOT", "BUILDING THE COPILOT BASE"))
    print(f"{tx(lang, 'Arquivos', 'Files'):<20}: {fmt_int(total, lang)} ({stats['tamanho_formatado']})")
    print(f"{tx(lang, 'Processos paralelos', 'Parallel workers'):<20}: {max_workers}")
    print(f"{t('pasta_execucao', cfg):<20}: {pasta_destino}")
    print(f"{Cor.CINZA}{tx(lang, 'Ctrl+C cancela.', 'Ctrl+C cancels.')}{Cor.RESET}\n")

    inicio = time.time()
    resultados: list[tuple] = []
    erros: list[tuple[str, str]] = []
    cont = {"n": 0}

    def ao_concluir(res):
        nome, rel, grupo, meta, erro = res
        cont["n"] += 1
        n = cont["n"]
        decorrido = time.time() - inicio
        restante = int((total - n) * decorrido / max(n, 1))
        eta = f"{restante // 60:02d}:{restante % 60:02d}"
        prefixo = f"({n:>{len(str(total))}}/{total})"
        if erro:
            erros.append((rel, erro))
            print(f"{prefixo} {Cor.VERMELHO}[{tx(lang, 'FALHA', 'FAIL')}]{Cor.RESET} {rel} -> {erro}")
            return
        resultados.append((nome, rel, grupo, meta))
        status = f"{Cor.AZUL}[CACHE]{Cor.RESET}" if meta.get("cache") else f"{Cor.VERDE}[OK]{Cor.RESET}"
        if meta.get("ocr_pages"):
            status += f" {Cor.AMARELO}(+{meta['ocr_pages']} OCR){Cor.RESET}"
        if not (meta.get("conteudo") or "").strip():
            status += f" {Cor.AMARELO}({tx(lang, 'sem texto', 'no text')}){Cor.RESET}"
        print(f"{prefixo} {status} {rel} {Cor.CINZA}| {tx(lang, 'restante', 'ETA')} {eta}{Cor.RESET}")

    tarefas = [(str(raiz), str(arq), str(pasta_cache), dict(cfg)) for raiz, arq in arquivos]
    try:
        processar_em_paralelo(tarefas, max_workers, ao_concluir)
    except ProcessamentoCancelado:
        print(f"\n{Cor.AMARELO}{tx(lang, 'Processamento cancelado pelo usuário.', 'Processing cancelled by the user.')}{Cor.RESET}")
        if pasta_destino != pasta_saida_base:
            try:
                if not any(pasta_destino.iterdir()):
                    pasta_destino.rmdir()
            except Exception:
                pass
        if pausar_no_final and interativo:
            pausar(lang)
        return None

    print(f"\n{tx(lang, 'Montando a base consolidada...', 'Assembling the consolidated base...')}")
    docs = montar_documentos(resultados, lang)
    quase = detectar_quase_duplicados(docs)
    grupos = len({d["grupo"] for d in docs})
    resumo = {
        "data": agora_fmt(lang),
        "docs": len(docs),
        "grupos": grupos,
        "volume": stats["tamanho_formatado"],
        "duplicados": sum(1 for d in docs if d["dup_de"]) if cfg.get("omitir_duplicados", True) else 0,
        "tokens": 0,
    }
    gerados, info = gerar_saidas(docs, cfg, pasta_destino, resumo)

    ocr_total = {"ok": 0, "empty": 0, "error": 0}
    for d in docs:
        for k in ocr_total:
            ocr_total[k] += (d["meta"].get("diag_ocr") or {}).get(k, 0)
    duracao = time.time() - inicio
    contexto = {
        "docs": docs, "erros": erros, "quase": quase, "gerados": list(gerados), "pasta": pasta_destino,
        "duracao": duracao, "workers": max_workers, "total": total, "volume": stats["tamanho_formatado"],
        "cache": sum(1 for d in docs if d["meta"].get("cache")), "ocr": ocr_total,
        "mojibake": sum(d["meta"].get("mojibake_repairs", 0) for d in docs),
        "pii": bool(cfg.get("mascarar_pii", False)), "tokens": resumo["tokens"], "partes": info["partes"],
        "limite_mb": cfg.get("limite_mb_por_parte", 8.0), "omitir": bool(cfg.get("omitir_duplicados", True)),
        "sem_texto": [d for d in docs if not d["dup_de"] and not (d["meta"].get("conteudo") or "").strip()],
        "ignorados": varredura["ignorados"], "sem_permissao": varredura["sem_permissao"],
    }
    nome_rel = "audit_report.md" if lang == "en" else "relatorio_auditoria.md"
    caminho_rel = pasta_destino / nome_rel
    caminho_rel.write_text(construir_relatorio(lang, contexto), encoding="utf-8")
    gerados.append((caminho_rel, tx(lang, "relatório de auditoria", "audit report")))

    arquivo_zip = None
    if cfg.get("gerar_arquivo_zip", False):
        print(tx(lang, "Compactando .zip...", "Creating .zip..."))
        caminho_zip = pasta_destino / f"base_copilot_{carimbo}.zip"
        if compactar_arquivos([p for p, _ in gerados], pasta_destino, caminho_zip):
            arquivo_zip = caminho_zip

    if cfg.get("usar_pasta_timestamp", True):
        ultimo = pasta_saida_base / "ultimo_processamento"
        try:
            shutil.rmtree(ultimo, ignore_errors=True)
            shutil.copytree(pasta_destino, ultimo, ignore=shutil.ignore_patterns("*.zip"))
        except Exception:
            pass

    # Resumo final
    print()
    print(linha(Cor.VERDE))
    print(f"{Cor.VERDE}{Cor.NEGRITO}  {t('concluido_sucesso', cfg, tempo=duracao)}{Cor.RESET}")
    print(linha(Cor.VERDE))
    print(f"  {tx(lang, 'Documentos na base', 'Documents in base'):<22}: {Cor.NEGRITO}{len(docs)}{Cor.RESET} / {total}"
          f"  ({tx(lang, 'cache', 'cache')}: {contexto['cache']} · OCR: {ocr_total['ok']})")
    if erros:
        print(f"  {Cor.VERMELHO}{tx(lang, 'Falhas', 'Failures'):<22}: {len(erros)}{Cor.RESET}")
    if contexto["sem_texto"]:
        print(f"  {Cor.AMARELO}{tx(lang, 'Sem texto', 'Without text'):<22}: {len(contexto['sem_texto'])}{Cor.RESET}")
    dups = sum(1 for d in docs if d["dup_de"])
    if dups:
        print(f"  {tx(lang, 'Duplicados', 'Duplicates'):<22}: {dups}")
    print(f"  {tx(lang, 'Tamanho estimado', 'Estimated size'):<22}: ~{fmt_int(resumo['tokens'], lang)} tokens")
    print(f"  {t('pasta_execucao', cfg):<22}: {Cor.NEGRITO}{pasta_destino}{Cor.RESET}")
    print(f"\n {Cor.NEGRITO}{tx(lang, 'Arquivos gerados', 'Generated files')}:{Cor.RESET}")
    for caminho, desc in gerados:
        if caminho.parent != pasta_destino:
            continue
        try:
            tam = formatar_tamanho_bytes(caminho.stat().st_size)
        except OSError:
            tam = ""
        destaque = Cor.VERDE + Cor.NEGRITO if caminho.name == "corpus_completo.md" else Cor.TITULO
        print(f"  -> {destaque}{caminho.name}{Cor.RESET} ({tam}) {Cor.CINZA}{desc}{Cor.RESET}")
    subpastas = sorted({p.parent.name for p, _ in gerados if p.parent != pasta_destino})
    for sp in subpastas:
        qtd = sum(1 for p, _ in gerados if p.parent.name == sp and p.parent != pasta_destino)
        print(f"  -> {Cor.TITULO}{sp}/{Cor.RESET} ({qtd} {tx(lang, 'arquivos', 'files')})")
    if arquivo_zip:
        print(f"  -> {Cor.VERDE}{arquivo_zip.name}{Cor.RESET} ({formatar_tamanho_bytes(arquivo_zip.stat().st_size)})")
    if info["partes"] > 1:
        print(f"\n {Cor.AMARELO}{tx(lang, 'A base foi dividida em', 'The base was split into')} {info['partes']} "
              f"{tx(lang, 'partes (corpus_parte_XX.md).', 'parts (corpus_parte_XX.md).')}{Cor.RESET}")
    print(f"\n {Cor.NEGRITO}{tx(lang, 'Próximo passo', 'Next step')}:{Cor.RESET} "
          + tx(lang, "envie corpus_completo.md (ou as partes) ao Copilot e cole o PROMPT_PARA_COPILOT.txt no chat.",
               "upload corpus_completo.md (or the parts) to Copilot and paste COPILOT_PROMPT.txt into the chat."))

    _CACHE_VARREDURA["dados"] = None
    resultado = {"pasta": pasta_destino, "docs": docs, "erros": erros, "arquivos": [p for p, _ in gerados], "zip": arquivo_zip}

    if pausar_no_final and interativo:
        op = perguntar(f"\n[A] {tx(lang, 'Abrir a pasta do resultado', 'Open the result folder')}   [ENTER] {tx(lang, 'Voltar ao menu', 'Back to menu')}: ").strip().upper()
        if op == "A":
            abrir_no_explorador(pasta_destino)
    return resultado


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
def ler_pastas_txt() -> list[str]:
    pastas = []
    if PASTAS_TXT_FILE.exists():
        for ln in PASTAS_TXT_FILE.read_text(encoding="utf-8-sig").splitlines():
            ln = ln.strip()
            if ln and not ln.startswith("#"):
                pastas.append(str(resolver_caminho(ln)))
    return pastas


def salvar_config(cfg: dict):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
        linhas_txt = ["# Pastas de entrada do Consolidador (uma por linha) / Input folders (one per line)"]
        linhas_txt += list(cfg.get("pastas_entrada", []))
        PASTAS_TXT_FILE.write_text("\n".join(linhas_txt) + "\n", encoding="utf-8")
    except Exception as e:
        print(f"{Cor.AMARELO}[config] {e}{Cor.RESET}")


def _normalizar_pastas(lista: list[str]) -> list[str]:
    saida, vistos = [], set()
    for p in lista:
        if not str(p).strip():
            continue
        p_res = str(resolver_caminho(p))
        chave = os.path.normcase(p_res)
        if chave not in vistos:
            vistos.add(chave)
            saida.append(p_res)
    return saida


def carregar_config() -> dict:
    cfg = json.loads(json.dumps(CONFIG_PADRAO))
    if CONFIG_FILE.exists():
        try:
            dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
            if isinstance(dados, dict):
                cfg.update(dados)
        except Exception as e:
            print(f"{Cor.AMARELO}[config.json] {e} -> usando padrões / using defaults{Cor.RESET}")

    # Validação de tipos/limites (config editada à mão)
    def num(chave, tipo, minimo, maximo):
        try:
            cfg[chave] = min(max(tipo(cfg.get(chave, CONFIG_PADRAO[chave])), minimo), maximo)
        except (TypeError, ValueError):
            cfg[chave] = CONFIG_PADRAO[chave]

    num("limite_mb_por_parte", float, 0.5, 10_000)
    num("qtd_docs_por_parte", int, 1, 100_000)
    num("max_workers", int, 1, max(1, os.cpu_count() or 4))
    num("dpi_ocr_pdf", int, 72, 600)
    num("min_dimensao_imagem_ocr", int, 20, 5000)
    num("limiar_confianca_ocr", float, 0.05, 0.99)
    num("proporcao_max_aspecto_ocr", float, 2.0, 1000.0)
    if cfg.get("idioma") not in ("pt", "en"):
        cfg["idioma"] = "pt"

    exts = cfg.get("extensoes") or EXTENSOES_PADRAO
    exts = sorted({("." + e.lower().lstrip(".")) for e in exts if str(e).strip()})
    if set(exts) == _EXTENSOES_PADRAO_V6_0:
        exts = EXTENSOES_PADRAO  # atualiza a lista padrão antiga com os novos formatos
    cfg["extensoes"] = exts

    cfg["pastas_entrada"] = _normalizar_pastas(list(cfg.get("pastas_entrada", [])) + ler_pastas_txt())
    return cfg


def aplicar_argumentos(cfg: dict) -> dict:
    """Aplica as opções de linha de comando à configuração da sessão (sem salvar)."""
    argv = sys.argv
    lang = idioma_linha_comando()
    if lang:
        cfg["idioma"] = lang
    if "--redact-pii" in argv:
        cfg["mascarar_pii"] = True
    elif "--no-pii" in argv:
        cfg["mascarar_pii"] = False
    if "--no-ocr" in argv:
        cfg["usar_ocr_em_pdf"] = False
        cfg["usar_ocr_em_imagens_embutidas"] = False
    for i, a in enumerate(argv):
        if a == "--workers" and i + 1 < len(argv) and argv[i + 1].isdigit():
            cfg["max_workers"] = max(1, int(argv[i + 1]))
        elif a == "--output" and i + 1 < len(argv):
            # caminhos da linha de comando são relativos ao diretório atual do terminal
            cfg["pasta_saida"] = str(Path(argv[i + 1]).expanduser().resolve())
    entradas = [str(Path(argv[i + 1]).expanduser().resolve())
                for i, a in enumerate(argv) if a == "--input" and i + 1 < len(argv)]
    if entradas:
        cfg["pastas_entrada"] = _normalizar_pastas(entradas)
    return cfg


# ---------------------------------------------------------------------------
# Menus
# ---------------------------------------------------------------------------
def selecionar_pasta_gui(lang: str = "pt") -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta = filedialog.askdirectory(title=tx(lang, "Selecione a pasta com os documentos", "Select the folder with the documents"))
        root.destroy()
        if pasta:
            return str(Path(pasta).resolve())
    except Exception as e:
        print(f"{Cor.AMARELO}{tx(lang, 'Não foi possível abrir a janela de seleção', 'Could not open the folder dialog')}: {e}{Cor.RESET}")
    return None


def _contar_por_pasta(cfg: dict) -> dict[str, int]:
    contagem: dict[str, int] = {}
    for raiz, _ in varrer_pastas(cfg)["arquivos"]:
        chave = os.path.normcase(str(raiz))
        contagem[chave] = contagem.get(chave, 0) + 1
    return contagem


def _adicionar_pasta(cfg: dict, caminho: str, lang: str) -> bool:
    p = resolver_caminho(caminho)
    if not p.is_dir():
        print(f"{Cor.VERMELHO}{tx(lang, 'Pasta não encontrada', 'Folder not found')}: {p}{Cor.RESET}")
        return False
    saida = resolver_caminho(cfg.get("pasta_saida", "./saida_copilot"))
    if os.path.normcase(str(p)) == os.path.normcase(str(saida)):
        print(f"{Cor.VERMELHO}{tx(lang, 'Esta é a pasta de saída; não pode ser entrada.', 'This is the output folder; it cannot be an input.')}{Cor.RESET}")
        return False
    if os.path.normcase(str(p)) in {os.path.normcase(x) for x in cfg["pastas_entrada"]}:
        print(f"{Cor.AMARELO}{tx(lang, 'Pasta já está na lista.', 'Folder is already in the list.')}{Cor.RESET}")
        return False
    cfg["pastas_entrada"].append(str(p))
    print(f"{Cor.VERDE}{tx(lang, 'Adicionada', 'Added')}: {p}{Cor.RESET}")
    return True


def menu_pastas(cfg: dict):
    """Gerencia as pastas de entrada: ver, adicionar, remover e sincronizar com pastas.txt."""
    while True:
        lang = obter_idioma(cfg)
        _cabecalho(cfg, tx(lang, "PASTAS DE ENTRADA", "INPUT FOLDERS"))
        contagem = _contar_por_pasta(cfg)
        if cfg["pastas_entrada"]:
            for i, p in enumerate(cfg["pastas_entrada"], start=1):
                if Path(p).is_dir():
                    qtd = contagem.get(os.path.normcase(str(resolver_caminho(p))), 0)
                    status = f"{Cor.VERDE}{qtd:>6} {tx(lang, 'arq.', 'files')}{Cor.RESET}"
                else:
                    status = f"{Cor.VERMELHO}{tx(lang, 'INEXISTENTE', '   MISSING')}{Cor.RESET}"
                print(f" [{i:>2}] {status}  {p}")
        else:
            print(f" {Cor.AMARELO}{tx(lang, '(nenhuma pasta configurada)', '(no folders configured)')}{Cor.RESET}")
        print(linha(char="-"))
        janela = "Finder" if platform.system() == "Darwin" else "Windows Explorer"
        print(f" [A] {tx(lang, f'Adicionar pasta (janela do {janela})', f'Add folder ({janela} dialog)')}")
        print(f" [D] {tx(lang, 'Adicionar pasta digitando/colando o caminho', 'Add folder by typing/pasting the path')}")
        print(f" [M] {tx(lang, 'Adicionar todas as subpastas de uma pasta mãe', 'Add every subfolder of a parent folder')}")
        print(f" [R] {tx(lang, 'Remover pasta (pelo número)', 'Remove folder (by number)')}")
        print(f" [X] {tx(lang, 'Remover pastas inexistentes', 'Remove missing folders')}")
        print(f" [T] {tx(lang, 'Remover todas', 'Remove all')}")
        print(f" [S] {tx(lang, 'Recarregar a lista a partir do arquivo pastas.txt', 'Reload the list from pastas.txt')}")
        print(f" [0] {tx(lang, 'Voltar', 'Back')}")
        print(linha())
        op = perguntar(f"\n{t('prompt_opcao', cfg)}").strip().upper()
        mudou = False
        if op in ("0", ""):
            break
        if op == "A":
            print(tx(lang, "Abrindo a janela de seleção...", "Opening the folder dialog..."))
            sel = selecionar_pasta_gui(lang)
            mudou = bool(sel) and _adicionar_pasta(cfg, sel, lang)
        elif op == "D":
            caminho = perguntar(tx(lang, "Caminho da pasta (pode arrastar a pasta para cá): ", "Folder path (you can drag the folder here): ")).strip().strip('"').strip("'")
            mudou = bool(caminho) and _adicionar_pasta(cfg, caminho, lang)
        elif op == "M":
            caminho = perguntar(tx(lang, "Caminho da pasta mãe: ", "Parent folder path: ")).strip().strip('"').strip("'")
            mae = resolver_caminho(caminho) if caminho else None
            if mae and mae.is_dir():
                subpastas = sorted(sp for sp in mae.iterdir() if sp.is_dir() and not sp.name.startswith("."))
                antes = len(cfg["pastas_entrada"])
                for sp in subpastas:
                    if os.path.normcase(str(sp)) not in {os.path.normcase(x) for x in cfg["pastas_entrada"]}:
                        cfg["pastas_entrada"].append(str(sp))
                # arquivos soltos direto na pasta mãe também entram (sem duplicar os das subpastas)
                soltos = [f for f in mae.iterdir() if f.is_file() and f.suffix.lower() in set(cfg["extensoes"])]
                if soltos and os.path.normcase(str(mae)) not in {os.path.normcase(x) for x in cfg["pastas_entrada"]}:
                    cfg["pastas_entrada"].append(str(mae))
                adicionadas = len(cfg["pastas_entrada"]) - antes
                print(f"{Cor.VERDE}{adicionadas} {tx(lang, 'pasta(s) adicionada(s).', 'folder(s) added.')}{Cor.RESET}")
                mudou = adicionadas > 0
            else:
                print(f"{Cor.VERMELHO}{tx(lang, 'Pasta mãe inválida.', 'Invalid parent folder.')}{Cor.RESET}")
        elif op == "R" or op.isdigit():
            valor = op if op.isdigit() else perguntar(tx(lang, "Número da pasta a remover: ", "Number of the folder to remove: ")).strip()
            if valor.isdigit() and 1 <= int(valor) <= len(cfg["pastas_entrada"]):
                removida = cfg["pastas_entrada"].pop(int(valor) - 1)
                print(f"{Cor.AMARELO}{tx(lang, 'Removida', 'Removed')}: {removida}{Cor.RESET}")
                mudou = True
        elif op == "X":
            antes = len(cfg["pastas_entrada"])
            cfg["pastas_entrada"] = [p for p in cfg["pastas_entrada"] if Path(p).is_dir()]
            mudou = len(cfg["pastas_entrada"]) != antes
        elif op == "T":
            if perguntar(tx(lang, "Remover TODAS as pastas? (s/N): ", "Remove ALL folders? (y/N): ")).strip().lower() in ("s", "y", "sim", "yes"):
                cfg["pastas_entrada"] = []
                mudou = True
        elif op == "S":
            cfg["pastas_entrada"] = _normalizar_pastas(ler_pastas_txt())
            print(f"{Cor.VERDE}{len(cfg['pastas_entrada'])} {tx(lang, 'pasta(s) carregada(s) de pastas.txt.', 'folder(s) loaded from pastas.txt.')}{Cor.RESET}")
            mudou = True
        if mudou:
            salvar_config(cfg)
            varrer_pastas(cfg, forcar=True)
        if op not in ("0", ""):
            time.sleep(0.9)


def menu_listar_remover_pastas(cfg: dict):
    """Compatibilidade com versões anteriores."""
    menu_pastas(cfg)


def menu_adicionar_pastas(cfg: dict):
    """Compatibilidade com versões anteriores."""
    menu_pastas(cfg)


def _ler_numero(msg: str, tipo, minimo, maximo):
    valor = perguntar(msg).strip().replace(",", ".")
    try:
        v = tipo(float(valor)) if tipo is int else tipo(valor)
        if minimo <= v <= maximo:
            return v
    except ValueError:
        pass
    return None


def menu_opcoes_saida_e_ocr(cfg: dict):
    """Configurações agrupadas por finalidade."""
    while True:
        lang = obter_idioma(cfg)
        s, n = t("sim", cfg), t("nao", cfg)
        num_cpus = os.cpu_count() or 4

        def sn(chave, padrao=True):
            return f"{Cor.VERDE}{s}{Cor.RESET}" if cfg.get(chave, padrao) else f"{Cor.CINZA}{n}{Cor.RESET}"

        _cabecalho(cfg, tx(lang, "CONFIGURAÇÕES", "SETTINGS"))
        print(f"\n {Cor.NEGRITO}{tx(lang, 'Arquivos gerados', 'Generated files')}{Cor.RESET}")
        print(f"  [1] {tx(lang, 'Base única (corpus_completo.md)', 'Single base (corpus_completo.md)'):<46}: {sn('gerar_arquivo_unico')}")
        print(f"  [2] {tx(lang, 'Dividir em partes por tamanho', 'Split into parts by size'):<46}: {sn('dividir_por_tamanho')}")
        print(f"  [3] {tx(lang, '   Tamanho máximo de cada parte', '   Maximum size of each part'):<46}: {cfg.get('limite_mb_por_parte', 8.0)} MB")
        print(f"  [4] {tx(lang, 'Um arquivo por pasta de origem (por_pasta/)', 'One file per source folder (por_pasta/)'):<46}: {sn('gerar_por_pasta_raiz')}")
        print(f"  [5] {tx(lang, 'Lotes por quantidade de documentos', 'Batches by number of documents'):<46}: {sn('dividir_por_quantidade', False)}"
              f"  ({cfg.get('qtd_docs_por_parte', 50)} docs, [6] {tx(lang, 'alterar', 'change')})")
        print(f"  [7] {tx(lang, 'Um Markdown por documento (individual/)', 'One Markdown per document (individual/)'):<46}: {sn('gerar_individual_por_arquivo', False)}")
        print(f"  [8] {tx(lang, 'Base RAG em JSONL (corpus_rag.jsonl)', 'RAG dataset in JSONL (corpus_rag.jsonl)'):<46}: {sn('gerar_jsonl', False)}")
        print(f"  [9] {tx(lang, 'Compactar resultado em .zip', 'Compress result into .zip'):<46}: {sn('gerar_arquivo_zip', False)}")
        print(f" [10] {tx(lang, 'Nova pasta com data/hora a cada execução', 'New timestamped folder on each run'):<46}: {sn('usar_pasta_timestamp')}")
        print(f" [11] {tx(lang, 'Pasta de saída', 'Output folder'):<46}: {cfg.get('pasta_saida')}")

        print(f"\n {Cor.NEGRITO}{tx(lang, 'Qualidade do conteúdo', 'Content quality')}{Cor.RESET}")
        print(f" [12] {tx(lang, 'Omitir conteúdo de documentos duplicados', 'Omit content of duplicate documents'):<46}: {sn('omitir_duplicados')}")
        print(f" [13] {tx(lang, 'Remover cabeçalhos/rodapés repetidos (PDF)', 'Remove repeated headers/footers (PDF)'):<46}: {sn('remover_cabecalhos_rodapes')}")
        print(f" [14] {tx(lang, 'Tipos de arquivo lidos', 'File types read'):<46}: {len(cfg.get('extensoes', []))} {tx(lang, 'extensões', 'extensions')}")

        print(f"\n {Cor.NEGRITO}OCR{Cor.RESET}")
        print(f" [15] {tx(lang, 'OCR em PDFs digitalizados e imagens', 'OCR on scanned PDFs and images'):<46}: {sn('usar_ocr_em_pdf')}")
        print(f" [16] {tx(lang, 'OCR em imagens no Word/Excel/PowerPoint', 'OCR on images in Word/Excel/PowerPoint'):<46}: {sn('usar_ocr_em_imagens_embutidas')}")
        print(f" [17] {tx(lang, 'Resolução do OCR em PDF (DPI)', 'PDF OCR resolution (DPI)'):<46}: {cfg.get('dpi_ocr_pdf', 200)}")
        print(f" [18] {tx(lang, 'Tamanho mínimo de imagem para OCR (px)', 'Minimum image size for OCR (px)'):<46}: {cfg.get('min_dimensao_imagem_ocr', 150)}")
        print(f" [19] {tx(lang, 'Confiança mínima do OCR (0.05 a 0.99)', 'Minimum OCR confidence (0.05 to 0.99)'):<46}: {float(cfg.get('limiar_confianca_ocr', 0.5)):.2f}")
        print(f" [20] {tx(lang, 'Salvar diagramas/esquemas em pasta (imagens/)', 'Save diagrams/schematics to folder (imagens/)'):<46}: {sn('salvar_diagramas_em_pasta')}")

        print(f"\n {Cor.NEGRITO}{tx(lang, 'Privacidade e desempenho', 'Privacy and performance')}{Cor.RESET}")
        print(f" [21] {tx(lang, 'Mascarar dados pessoais (CPF, RG, nomes...)', 'Mask personal data (CPF, RG, names...)'):<46}: {sn('mascarar_pii', False)}")
        print(f" [22] {tx(lang, 'Processos paralelos', 'Parallel workers'):<46}: {cfg.get('max_workers')} ({tx(lang, 'máx.', 'max')} {num_cpus})")
        print(linha(char="-"))
        print(f"  [P] {tx(lang, 'Restaurar configuração padrão', 'Restore default settings')}   [0] {tx(lang, 'Voltar', 'Back')}")
        print(linha())

        op = perguntar(f"\n{t('prompt_opcao', cfg)}").strip().upper()
        alternar = {"1": "gerar_arquivo_unico", "2": "dividir_por_tamanho", "4": "gerar_por_pasta_raiz",
                    "5": "dividir_por_quantidade", "7": "gerar_individual_por_arquivo", "8": "gerar_jsonl",
                    "9": "gerar_arquivo_zip", "10": "usar_pasta_timestamp", "12": "omitir_duplicados",
                    "13": "remover_cabecalhos_rodapes", "15": "usar_ocr_em_pdf", "16": "usar_ocr_em_imagens_embutidas",
                    "20": "salvar_diagramas_em_pasta", "21": "mascarar_pii"}
        if op in ("0", ""):
            break
        if op in alternar:
            chave = alternar[op]
            cfg[chave] = not cfg.get(chave, CONFIG_PADRAO.get(chave, False))
        elif op == "3":
            v = _ler_numero(tx(lang, "MB por parte (ex.: 8, 20, 50): ", "MB per part (e.g. 8, 20, 50): "), float, 0.5, 10_000)
            if v:
                cfg["limite_mb_por_parte"] = v
        elif op == "6":
            v = _ler_numero(tx(lang, "Documentos por lote: ", "Documents per batch: "), int, 1, 100_000)
            if v:
                cfg["qtd_docs_por_parte"] = v
        elif op == "11":
            novo = perguntar(tx(lang, "Nova pasta de saída: ", "New output folder: ")).strip().strip('"').strip("'")
            if novo:
                destino = resolver_caminho(novo)
                if os.path.normcase(str(destino)) in {os.path.normcase(p) for p in cfg["pastas_entrada"]}:
                    print(f"{Cor.VERMELHO}{tx(lang, 'A saída não pode ser uma pasta de entrada.', 'The output cannot be an input folder.')}{Cor.RESET}")
                    time.sleep(1.5)
                else:
                    cfg["pasta_saida"] = str(destino)
        elif op == "14":
            print(f"\n{tx(lang, 'Atuais', 'Current')}: {' '.join(cfg.get('extensoes', []))}")
            print(f"{tx(lang, 'Suportadas', 'Supported')}: {' '.join(EXTENSOES_PADRAO)}")
            novo = perguntar(tx(lang, "Digite as extensões separadas por espaço (ENTER mantém, 'padrao' restaura): ",
                                "Type the extensions separated by spaces (ENTER keeps, 'default' restores): ")).strip().lower()
            if novo in ("padrao", "padrão", "default"):
                cfg["extensoes"] = EXTENSOES_PADRAO
            elif novo:
                escolhidas = sorted({"." + e.strip().lstrip(".") for e in re.split(r"[\s,;]+", novo) if e.strip()})
                validas = [e for e in escolhidas if e in EXTENSOES_PADRAO]
                if validas:
                    cfg["extensoes"] = validas
        elif op == "17":
            v = _ler_numero("DPI (100-600): ", int, 100, 600)
            if v:
                cfg["dpi_ocr_pdf"] = v
        elif op == "18":
            v = _ler_numero("px (20-2000): ", int, 20, 2000)
            if v:
                cfg["min_dimensao_imagem_ocr"] = v
        elif op == "19":
            v = _ler_numero("0.05 - 0.99: ", float, 0.05, 0.99)
            if v:
                cfg["limiar_confianca_ocr"] = v
        elif op == "22":
            v = _ler_numero(tx(lang, f"Processos (1 a {num_cpus}): ", f"Workers (1 to {num_cpus}): "), int, 1, num_cpus)
            if v:
                cfg["max_workers"] = v
                cfg["multicore"] = True
        elif op == "P":
            if perguntar(tx(lang, "Restaurar padrões (pastas e idioma são mantidos)? (s/N): ",
                            "Restore defaults (folders and language are kept)? (y/N): ")).strip().lower() in ("s", "y", "sim", "yes"):
                pastas, idioma = cfg.get("pastas_entrada", []), cfg.get("idioma", "pt")
                cfg.clear()
                cfg.update(json.loads(json.dumps(CONFIG_PADRAO)))
                cfg["pastas_entrada"], cfg["idioma"] = pastas, idioma
        salvar_config(cfg)


def _tamanho_pasta(p: Path) -> int:
    total = 0
    for raiz, _, arqs in os.walk(p):
        for a in arqs:
            try:
                total += (Path(raiz) / a).stat().st_size
            except OSError:
                pass
    return total


def menu_principal():
    cfg = aplicar_argumentos(carregar_config())
    while True:
        lang = obter_idioma(cfg)
        varredura = varrer_pastas(cfg)
        stats = varredura["stats"]
        num_cpus = os.cpu_count() or 4
        workers = int(cfg.get("max_workers", min(4, num_cpus))) if cfg.get("multicore", True) else 1
        saida = resolver_caminho(cfg.get("pasta_saida", "./saida_copilot"))
        ultimo = saida / "ultimo_processamento"

        _cabecalho(cfg, t("titulo_app", cfg))
        print(f"{Cor.CINZA}  {t('subtitulo_app', cfg, workers=workers)}{Cor.RESET}\n")
        qtd_pastas = len(cfg["pastas_entrada"])
        inexist = len(varredura["pastas_inexistentes"])
        aviso_inexist = f"  {Cor.VERMELHO}({inexist} {tx(lang, 'inexistente(s)', 'missing')}){Cor.RESET}" if inexist else ""
        print(f"  {t('pastas_monitoradas', cfg):<22}: {Cor.NEGRITO}{qtd_pastas}{Cor.RESET}{aviso_inexist}")
        print(f"  {t('arquivos_localizados', cfg):<22}: {Cor.VERDE}{Cor.NEGRITO}{fmt_int(stats['total_arquivos'], lang)}{Cor.RESET} ({stats['tamanho_formatado']})")
        if stats["por_extensao"]:
            resumo_ext = "  ".join(f"{e}:{i['qtd']}" for e, i in sorted(stats["por_extensao"].items(), key=lambda x: -x[1]["qtd"])[:8])
            print(f"  {'':<22}  {Cor.CINZA}{resumo_ext}{Cor.RESET}")
        pii = tx(lang, "mascarados", "masked") if cfg.get("mascarar_pii", False) else tx(lang, "preservados", "preserved")
        ocr = tx(lang, "ligado", "on") if cfg.get("usar_ocr_em_pdf", True) else tx(lang, "desligado", "off")
        print(f"  {tx(lang, 'OCR / Dados pessoais', 'OCR / Personal data'):<22}: {ocr} / {pii}")
        print(f"  {tx(lang, 'Pasta de saída', 'Output folder'):<22}: {saida}")
        print(linha(char="-"))
        print(f" {Cor.VERDE}{Cor.NEGRITO}[1] {tx(lang, 'GERAR BASE CONSOLIDADA PARA O COPILOT', 'BUILD THE CONSOLIDATED COPILOT BASE')}{Cor.RESET}")
        print(f" [2] {tx(lang, 'Pastas de entrada (adicionar, remover, ver)', 'Input folders (add, remove, view)')} ({qtd_pastas})")
        print(f" [3] {tx(lang, 'Configurações (saídas, OCR, privacidade)', 'Settings (outputs, OCR, privacy)')}")
        print(f" [4] {tx(lang, 'Abrir a pasta do último resultado', 'Open the latest result folder')}")
        print(f" [5] {tx(lang, 'Limpar cache de extração (força reler tudo)', 'Clear extraction cache (re-read everything)')}")
        print(f" [L] {tx(lang, 'Language: English', 'Idioma: Português')}")
        print(f" [0] {tx(lang, 'Sair', 'Exit')}")
        print(linha())

        op = perguntar(f"\n{t('prompt_opcao', cfg)}").strip().upper()
        if op == "1":
            executar_pipeline(cfg)
            varrer_pastas(cfg, forcar=True)
        elif op == "2":
            menu_pastas(cfg)
        elif op == "3":
            menu_opcoes_saida_e_ocr(cfg)
        elif op == "4":
            alvo = ultimo if ultimo.is_dir() else saida
            if alvo.is_dir():
                abrir_no_explorador(alvo)
            else:
                print(f"{Cor.AMARELO}{tx(lang, 'Nenhum resultado gerado ainda.', 'No result generated yet.')}{Cor.RESET}")
                time.sleep(1.2)
        elif op == "5":
            pasta_cache = saida / ".cache_extracao"
            tam = formatar_tamanho_bytes(_tamanho_pasta(pasta_cache)) if pasta_cache.is_dir() else "0 B"
            if perguntar(tx(lang, f"Apagar o cache ({tam})? A próxima execução relerá todos os arquivos. (s/N): ",
                            f"Delete the cache ({tam})? The next run will re-read every file. (y/N): ")).strip().lower() in ("s", "y", "sim", "yes"):
                shutil.rmtree(pasta_cache, ignore_errors=True)
                print(f"{Cor.VERDE}{tx(lang, 'Cache apagado.', 'Cache deleted.')}{Cor.RESET}")
                time.sleep(1.0)
        elif op == "L":
            cfg["idioma"] = "en" if lang == "pt" else "pt"
            salvar_config(cfg)
        elif op == "0":
            break


AJUDA = f"""
Local Document Consolidator for Microsoft Copilot (v{VERSAO}) - 100% offline

Usage / Uso:
  python app.py [options]

Options / Opções:
  --lang en|pt        Interface and output language / idioma da interface e da base
  --auto              Run without menus (uses config.json) / executa sem menus
  --input PATH        Input folder (repeatable; overrides config) / pasta de entrada (pode repetir)
  --output PATH       Output folder / pasta de saída
  --workers N         Parallel worker processes / processos paralelos
  --redact-pii        Mask personal data (CPF, RG, e-mail, phone...) / mascara dados pessoais
  --no-pii            Keep raw data (default) / mantém dados originais (padrão)
  --no-ocr            Disable OCR / desativa OCR
  --help, -h          Show this help / mostra esta ajuda
""".strip()


if __name__ == "__main__":
    if "--help" in sys.argv or "-h" in sys.argv:
        print(AJUDA)
        sys.exit(0)

    if "--auto" in sys.argv:
        cfg_auto = aplicar_argumentos(carregar_config())
        res = executar_pipeline(cfg_auto, pausar_no_final=False, interativo=False)
        # 0 = base gerada (falhas individuais ficam no relatório); 1 = nada foi gerado
        sys.exit(0 if res is not None else 1)
    else:
        try:
            menu_principal()
        except (KeyboardInterrupt, EOFError):
            print("\nEncerrado / Closed.")
