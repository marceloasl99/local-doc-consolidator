#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidador Local de Documentos para Microsoft Copilot (v4 - Enterprise Edition)
Compatível com Windows 11 | 100% Offline | Multicore Paralelo | Processamento Local Seguro
"""

import csv
import hashlib
import io
import json
import os
import platform
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
warnings.filterwarnings("ignore")

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

CONFIG_PADRAO = {
    "pastas_entrada": [],
    "pasta_saida": str(BASE_DIR / "saida_copilot"),
    # Estratégias de particionamento e saída Markdown:
    "gerar_arquivo_unico": True,
    "dividir_por_tamanho": True,
    "limite_mb_por_parte": 8.0,
    "dividir_por_quantidade": False,
    "qtd_docs_por_parte": 50,
    "gerar_individual_por_arquivo": False,
    "gerar_por_pasta_raiz": True,
    # Compactação e Timestamp:
    "usar_pasta_timestamp": True,
    "gerar_arquivo_zip": True,
    # Configurações de OCR:
    "usar_ocr_em_pdf": True,
    "usar_ocr_em_imagens_embutidas": True,
    "dpi_ocr_pdf": 200,
    "min_dimensao_imagem_ocr": 80,
    # Configurações de hardware / multicore:
    "multicore": True,
    "extensoes": EXTENSOES_PADRAO,
}

# Variável de processo worker para o motor de OCR (Singleton por processo)
_OCR_ENGINE = None


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
        if not arq.exists():
            continue
        ext = arq.suffix.lower()
        if ext not in por_extensao:
            por_extensao[ext] = {"qtd": 0, "bytes": 0}
        por_extensao[ext]["qtd"] += 1
        por_extensao[ext]["bytes"] += arq.stat().st_size
    return {
        "total_arquivos": total_arquivos,
        "tamanho_total": tamanho_total,
        "tamanho_formatado": formatar_tamanho_bytes(tamanho_total),
        "por_extensao": por_extensao,
    }


def ler_pastas_txt() -> list[str]:
    if not PASTAS_TXT_FILE.exists():
        return []
    caminhos = []
    for linha in PASTAS_TXT_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        limpa = linha.strip().strip('"').strip("'")
        if limpa and not limpa.startswith("#"):
            p = resolver_caminho(limpa)
            caminhos.append(str(p))
    return caminhos


def carregar_config() -> dict:
    cfg = CONFIG_PADRAO.copy()
    if CONFIG_FILE.exists():
        try:
            salvo = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            cfg.update(salvo)
        except Exception:
            pass
    else:
        salvar_config(cfg)

    atuais = [str(resolver_caminho(p)) for p in cfg.get("pastas_entrada", [])]
    for p in ler_pastas_txt():
        if p not in atuais:
            atuais.append(p)
    cfg["pastas_entrada"] = atuais
    return cfg


def salvar_config(cfg: dict):
    CONFIG_FILE.write_text(
        json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def compactar_pasta(pasta_origem: Path, caminho_zip: Path):
    """Compacta a pasta gerada em um arquivo .zip otimizado para o usuário."""
    with zipfile.ZipFile(caminho_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in sorted(pasta_origem.rglob("*")):
            if item.is_file() and not item.name.endswith(".zip") and ".cache_extracao" not in str(item):
                rel = item.relative_to(pasta_origem)
                zf.write(item, arcname=str(rel))


def ler_texto_com_fallback(caminho: Path) -> str:
    dados = caminho.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return dados.decode(enc)
        except UnicodeDecodeError:
            continue
    return dados.decode("utf-8", errors="replace")


def tabela_para_markdown(linhas: list[list[str]]) -> str:
    if not linhas:
        return ""
    linhas_limpas = [
        [str(celula if celula is not None else "").replace("\n", " ").replace("|", "\\|").strip() for celula in l]
        for l in linhas
        if any(str(c if c is not None else "").strip() for c in l)
    ]
    if not linhas_limpas:
        return ""
    num_cols = max(len(l) for l in linhas_limpas)
    normalizadas = [l + [""] * (num_cols - len(l)) for l in linhas_limpas]
    cabecalho = "| " + " | ".join(normalizadas[0]) + " |"
    separador = "| " + " | ".join(["---"] * num_cols) + " |"
    corpo = ["| " + " | ".join(l) + " |" for l in normalizadas[1:]]
    return "\n".join([cabecalho, separador] + corpo)


def executar_ocr_em_bytes(img_bytes: bytes, min_dim: int = 80) -> str:
    """Executa OCR em bytes de imagem (PNG, JPEG, etc.), filtrando ruídos."""
    if not img_bytes:
        return ""
    engine = obter_ocr_worker()
    if not engine:
        return ""
    try:
        from PIL import Image
        with Image.open(io.BytesIO(img_bytes)) as im:
            w, h = im.size
            if w < min_dim or h < min_dim:
                return ""
            max_edge = max(w, h)
            if max_edge > 2400:
                ratio = 2400 / max_edge
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                img_bytes = buf.getvalue()

        resultado, _ = engine(img_bytes)
        if not resultado:
            return ""
        linhas = [item[1].strip() for item in resultado if item[1].strip()]
        return "\n".join(linhas)
    except Exception:
        return ""


def extrair_imagens_zip_office(caminho: Path, prefixo: str, min_dim: int = 80) -> list[tuple[str, str]]:
    """Extrai imagens embutidas em pacotes OpenXML (DOCX, XLSX, PPTX) e aplica OCR."""
    resultados = []
    try:
        with zipfile.ZipFile(caminho, "r") as z:
            media_files = [n for n in z.namelist() if n.startswith(prefixo)]
            media_files.sort()
            for nome in media_files:
                ext = Path(nome).suffix.lower()
                if ext in {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tiff", ".tif"}:
                    dados = z.read(nome)
                    texto_ocr = executar_ocr_em_bytes(dados, min_dim=min_dim)
                    if texto_ocr:
                        nome_base = Path(nome).name
                        resultados.append((nome_base, texto_ocr))
    except Exception:
        pass
    return resultados


def extrair_imagem_avulsa(caminho: Path, min_dim: int = 40) -> tuple[str, int]:
    """Processa arquivos de imagem avulsos (.png, .jpg, .tiff, etc.) diretamente via OCR."""
    try:
        from PIL import Image
        with Image.open(caminho) as im:
            w, h = im.size
            formato = im.format or caminho.suffix.replace(".", "").upper()
    except Exception as e:
        return f"*(Erro ao abrir imagem: {e})*", 0

    texto = executar_ocr_em_bytes(caminho.read_bytes(), min_dim=min_dim)
    if texto:
        bloco = (
            f"**[Imagem: {caminho.name}]** (Resolução: {w}x{h} px | Formato: {formato})\n\n"
            f"> **Texto Extraído via OCR Local:**\n"
            + "\n".join(f"> {linha}" for linha in texto.splitlines())
        )
        return bloco, 1
    else:
        bloco = f"**[Imagem: {caminho.name}]** (Resolução: {w}x{h} px | Formato: {formato})\n*(Imagem sem texto legível detectado)*"
        return bloco, 0


def extrair_pdf_hibrido(caminho_pdf: Path, usar_ocr: bool = True, dpi: int = 200, min_dim: int = 80) -> tuple[str, int]:
    """Lê PDFs nativamente e aciona OCR tanto em páginas escaneadas quanto em imagens de páginas mistas."""
    cont_ocr = 0
    blocos_paginas = []

    try:
        import pymupdf
        pymupdf.no_recommend_layout()
        if hasattr(pymupdf, "TOOLS"):
            pymupdf.TOOLS.mupdf_display_warnings(False)
            pymupdf.TOOLS.mupdf_display_errors(False)
        doc = pymupdf.open(caminho_pdf)

        for num_pag, pagina in enumerate(doc, start=1):
            texto_nativo = pagina.get_text("text").strip()
            tabelas_md = ""
            try:
                tabs = pagina.find_tables()
                if tabs and tabs.tables:
                    tabelas_md = "\n\n" + "\n\n".join(t.to_markdown() for t in tabs.tables)
            except Exception:
                pass

            imagens_pagina = pagina.get_images()
            texto_ocr_extra = []

            # 1. Caso A: Página escaneada
            if usar_ocr and len(texto_nativo) < 40 and len(imagens_pagina) > 0:
                engine = obter_ocr_worker()
                if engine:
                    try:
                        pix = pagina.get_pixmap(dpi=dpi)
                        resultado, _ = engine(pix.tobytes("png"))
                        if resultado:
                            linhas = [item[1].strip() for item in resultado if item[1].strip()]
                            texto_pagina = f"*(Página {num_pag} escaneada - extraída via OCR local)*\n" + "\n".join(linhas)
                            cont_ocr += 1
                        else:
                            texto_pagina = f"*(Página {num_pag} escaneada sem texto legível)*"
                    except Exception:
                        texto_pagina = texto_nativo or f"*(Página {num_pag} escaneada)*"
                else:
                    texto_pagina = texto_nativo or f"*(Página {num_pag} escaneada)*"

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
                                if w >= min_dim and h >= min_dim:
                                    img_bytes = img_dict.get("image", b"")
                                    txt_img = executar_ocr_em_bytes(img_bytes, min_dim=min_dim)
                                    if txt_img:
                                        texto_ocr_extra.append(
                                            f"> **[OCR Imagem Embutida #{img_idx} ({w}x{h} px)]:**\n"
                                            + "\n".join(f"> {l}" for l in txt_img.splitlines())
                                        )
                                        cont_ocr += 1
                        except Exception:
                            continue

            bloco_final = texto_pagina
            if texto_ocr_extra:
                bloco_final += "\n\n" + "\n\n".join(texto_ocr_extra)

            if bloco_final:
                blocos_paginas.append(f"### Página {num_pag}\n{bloco_final}")

        doc.close()
        return "\n\n".join(blocos_paginas), cont_ocr
    except ImportError:
        pass

    # Fallback com pypdf
    from pypdf import PdfReader
    reader = PdfReader(str(caminho_pdf))
    for num_pag, pagina in enumerate(reader.pages, start=1):
        txt = (pagina.extract_text() or "").strip()
        if txt:
            blocos_paginas.append(f"### Página {num_pag}\n{txt}")
        else:
            blocos_paginas.append(f"### Página {num_pag}\n*(Página sem camada de texto selecionável)*")
    return "\n\n".join(blocos_paginas), cont_ocr


def extrair_docx_nativo(caminho: Path, usar_ocr_imagens: bool = True, min_dim: int = 80) -> tuple[str, int]:
    """Extrai títulos, parágrafos, tabelas e imagens embutidas com OCR em documentos Word."""
    import docx
    doc = docx.Document(str(caminho))
    partes = []

    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        estilo = (p.style.name or "").lower() if p.style else ""
        if "heading 1" in estilo or "título 1" in estilo:
            partes.append(f"### {t}")
        elif "heading 2" in estilo or "título 2" in estilo:
            partes.append(f"#### {t}")
        else:
            partes.append(t)

    for idx_t, tab in enumerate(doc.tables, start=1):
        linhas = [[cel.text.strip() for cel in row.cells] for row in tab.rows]
        md_tab = tabela_para_markdown(linhas)
        if md_tab:
            partes.append(f"\n**Tabela {idx_t}:**\n{md_tab}")

    cont_ocr = 0
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(caminho, "word/media/", min_dim=min_dim)
        if imagens_ocr:
            blocos_img = ["\n### Imagens Embutidas no Documento (OCR)"]
            for nome_img, txt in imagens_ocr:
                cont_ocr += 1
                blocos_img.append(
                    f"> **[OCR Imagem: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    return "\n\n".join(partes), cont_ocr


def extrair_xlsx_nativo(caminho: Path, usar_ocr_imagens: bool = True, min_dim: int = 80) -> tuple[str, int]:
    """Converte planilhas Excel em tabelas Markdown e extrai imagens embutidas com OCR."""
    import openpyxl
    wb = openpyxl.load_workbook(str(caminho), data_only=True, read_only=True)
    partes = []

    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        linhas = []
        for row in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() != "" for c in row):
                linhas.append(["" if c is None else str(c) for c in row])
        md_tab = tabela_para_markdown(linhas)
        if md_tab:
            partes.append(f"### Planilha: {nome_aba}\n\n{md_tab}")
    wb.close()

    cont_ocr = 0
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(caminho, "xl/media/", min_dim=min_dim)
        if imagens_ocr:
            blocos_img = ["\n### Imagens / Comprovantes Embutidos na Planilha (OCR)"]
            for nome_img, txt in imagens_ocr:
                cont_ocr += 1
                blocos_img.append(
                    f"> **[OCR Imagem Planilha: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    return "\n\n".join(partes), cont_ocr


def extrair_pptx_nativo(caminho: Path, usar_ocr_imagens: bool = True, min_dim: int = 80) -> tuple[str, int]:
    """Extrai slides, anotações, tabelas e imagens com OCR de apresentações PowerPoint."""
    from pptx import Presentation
    prs = Presentation(str(caminho))
    partes = []

    for idx_s, slide in enumerate(prs.slides, start=1):
        textos_slide = [f"### Slide {idx_s}"]
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t:
                        textos_slide.append(f"- {t}")
            if getattr(shape, "has_table", False):
                linhas = [[cel.text.strip() for cel in row.cells] for row in shape.table.rows]
                md_t = tabela_para_markdown(linhas)
                if md_t:
                    textos_slide.append(f"\n{md_t}")
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                textos_slide.append(f"\n> **Notas do Apresentador:** {notas}")
        if len(textos_slide) > 1:
            partes.append("\n".join(textos_slide))

    cont_ocr = 0
    if usar_ocr_imagens:
        imagens_ocr = extrair_imagens_zip_office(caminho, "ppt/media/", min_dim=min_dim)
        if imagens_ocr:
            blocos_img = ["\n### Imagens / Diagramas Embutidos nos Slides (OCR)"]
            for nome_img, txt in imagens_ocr:
                cont_ocr += 1
                blocos_img.append(
                    f"> **[OCR Slide Imagem: `{nome_img}`]:**\n"
                    + "\n".join(f"> {l}" for l in txt.splitlines())
                )
            partes.append("\n\n".join(blocos_img))

    return "\n\n".join(partes), cont_ocr


def extrair_csv_nativo(caminho: Path) -> str:
    bruto = ler_texto_com_fallback(caminho)
    try:
        dialeto = csv.Sniffer().sniff(bruto[:4096], delimiters=";,\t,")
        leitor = csv.reader(io.StringIO(bruto), dialeto)
    except Exception:
        leitor = csv.reader(io.StringIO(bruto))
    linhas = list(leitor)
    return tabela_para_markdown(linhas) or bruto


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

    try:
        caminho_rel = f"{raiz.name}/{arq.relative_to(raiz)}".replace("\\", "/")
    except Exception:
        caminho_rel = arq.name

    try:
        pasta_cache.mkdir(parents=True, exist_ok=True)
        file_hash = calcular_hash_arquivo(arq)
        flag_ocr = f"ocr_{int(cfg.get('usar_ocr_em_pdf', True))}_{int(cfg.get('usar_ocr_em_imagens_embutidas', True))}"
        arq_cache = pasta_cache / f"{file_hash}_{flag_ocr}.json"

        if arq_cache.exists():
            try:
                dados = json.loads(arq_cache.read_text(encoding="utf-8"))
                return arq.name, caminho_rel, raiz.name, {
                    "cache": True,
                    "ocr_pages": dados.get("ocr_pages", 0),
                    "hash": file_hash,
                    "conteudo": dados["conteudo"],
                }, None
            except Exception:
                pass

        ext = arq.suffix.lower()
        ocr_count = 0
        conteudo = ""
        min_dim = int(cfg.get("min_dimensao_imagem_ocr", 80))

        if ext == ".pdf":
            conteudo, ocr_count = extrair_pdf_hibrido(
                arq,
                usar_ocr=cfg.get("usar_ocr_em_pdf", True),
                dpi=int(cfg.get("dpi_ocr_pdf", 200)),
                min_dim=min_dim,
            )
        elif ext in EXTENSOES_IMAGEM:
            conteudo, ocr_count = extrair_imagem_avulsa(arq, min_dim=min_dim)
        elif ext == ".docx":
            conteudo, ocr_count = extrair_docx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
            )
        elif ext == ".xlsx":
            conteudo, ocr_count = extrair_xlsx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
            )
        elif ext == ".pptx":
            conteudo, ocr_count = extrair_pptx_nativo(
                arq,
                usar_ocr_imagens=cfg.get("usar_ocr_em_imagens_embutidas", True),
                min_dim=min_dim,
            )
        elif ext == ".csv":
            conteudo = extrair_csv_nativo(arq)
        elif ext in {".txt", ".md"}:
            conteudo = ler_texto_com_fallback(arq)
        else:
            try:
                import pandas as pd
                abas = pd.read_excel(arq, sheet_name=None)
                blocos = []
                for nome, df in abas.items():
                    linhas = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
                    blocos.append(f"### Planilha: {nome}\n\n{tabela_para_markdown(linhas)}")
                conteudo = "\n\n".join(blocos)
            except Exception:
                conteudo = ler_texto_com_fallback(arq)

        conteudo = conteudo.strip()
        arq_cache.write_text(
            json.dumps(
                {
                    "arquivo": arq.name,
                    "conteudo": conteudo,
                    "ocr_pages": ocr_count,
                    "data_cache": datetime.now().isoformat(),
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return arq.name, caminho_rel, raiz.name, {
            "cache": False,
            "ocr_pages": ocr_count,
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
        raiz = resolver_caminho(p_str)
        if not raiz.exists() or not raiz.is_dir():
            continue
        for arq in sorted(raiz.rglob("*")):
            if not arq.is_file():
                continue
            if arq.name.startswith("~$") or arq.name.startswith("."):
                continue
            if arq.suffix.lower() in exts:
                caminho_real = str(arq.resolve())
                if caminho_real not in vistos_abs:
                    vistos_abs.add(caminho_real)
                    arquivos.append((raiz, arq))
    return arquivos


def sanitizar_nome_arquivo(nome: str) -> str:
    invalido = '<>:"/\\|?*'
    limpo = "".join("_" if c in invalido else c for c in nome)
    return limpo.strip() or "documento"


def executar_pipeline(cfg: dict, pausar_no_final: bool = True):
    limpar_tela()
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
    print(f"{Cor.TITULO}   CONVERSÃO E CONSOLIDAÇÃO MULTICORE PARA MICROSOFT COPILOT (v4){Cor.RESET}")
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")

    pastas = cfg.get("pastas_entrada", [])
    if not pastas:
        print(f"{Cor.VERMELHO}[!] Nenhuma pasta configurada! Adicione pastas na opção [2] ou no arquivo pastas.txt.{Cor.RESET}")
        if pausar_no_final:
            input("\nPressione ENTER para voltar ao menu...")
        return

    arquivos = listar_arquivos_elegiveis(pastas, cfg.get("extensoes", EXTENSOES_PADRAO))
    total = len(arquivos)

    if total == 0:
        print(f"{Cor.VERMELHO}[!] Nenhum arquivo compatível foi encontrado nas pastas informadas.{Cor.RESET}")
        if pausar_no_final:
            input("\nPressione ENTER para voltar ao menu...")
        return

    stats = obter_estatisticas_arquivos(arquivos)
    pasta_saida_base = resolver_caminho(cfg["pasta_saida"])
    pasta_cache = pasta_saida_base / ".cache_extracao"
    pasta_saida_base.mkdir(parents=True, exist_ok=True)

    # Cria pasta com timestamp para esta execução
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    if cfg.get("usar_pasta_timestamp", True):
        pasta_destino = pasta_saida_base / f"processado_{timestamp_str}"
    else:
        pasta_destino = pasta_saida_base
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # TELA DE CONFIRMAÇÃO PRÉVIA (caso interativo)
    if "--auto" not in sys.argv:
        while True:
            limpar_tela()
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
            print(f"{Cor.TITULO}   CONFIRMAÇÃO DO PROCESSAMENTO LOCAL{Cor.RESET}")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")
            print(f" {Cor.NEGRITO}Resumo do Corpus a Processar:{Cor.RESET}")
            print(f"  • Pastas monitoradas      : {len(pastas)}")
            print(f"  • Arquivos localizados    : {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET} arquivos")
            print(f"  • Volume total de dados   : {Cor.AMARELO}{Cor.NEGRITO}{stats['tamanho_formatado']}{Cor.RESET}")
            print(f"  • Detalhamento por formato:")
            for ext, info in sorted(stats["por_extensao"].items(), key=lambda x: x[1]["qtd"], reverse=True):
                print(f"     - {ext.upper():<6} : {info['qtd']:>3} arquivos ({formatar_tamanho_bytes(info['bytes'])})")

            print(f"\n {Cor.NEGRITO}Configurações desta Execução:{Cor.RESET}")
            print(f"  • Pasta de destino        : {pasta_destino}")
            print(f"  • Dividir em MB           : {'SIM (' + str(cfg.get('limite_mb_por_parte', 8.0)) + ' MB)' if cfg.get('dividir_por_tamanho', True) else 'NÃO'}")
            print(f"  • Gerar arquivo único     : {'SIM (corpus_completo.md)' if cfg.get('gerar_arquivo_unico', True) else 'NÃO'}")
            print(f"  • Gerar por quantidade    : {'SIM (' + str(cfg.get('qtd_docs_por_parte', 50)) + ' por lote)' if cfg.get('dividir_por_quantidade', False) else 'NÃO'}")
            print(f"  • Gerar 1:1 Individual    : {'SIM' if cfg.get('gerar_individual_por_arquivo', False) else 'NÃO'}")
            print(f"  • Gerar por Pasta         : {'SIM' if cfg.get('gerar_por_pasta_raiz', True) else 'NÃO'}")
            print(f"  • Compactar em .ZIP final : {'SIM' if cfg.get('gerar_arquivo_zip', True) else 'NÃO'}")
            print(f"  • OCR em Imagens/PDF      : {'ATIVADO' if cfg.get('usar_ocr_em_pdf', True) else 'DESATIVADO'}")
            print(f"{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
            print(" [ENTER] ou [C] Confirmar e Iniciar o Processamento")
            print(" [M] Alterar limite de separação em MB (ex: 8, 50, 100, 200 MB)")
            print(" [Z] Alternar geração de arquivo compactado .ZIP")
            print(" [V] Cancelar e voltar ao menu")
            print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

            escolha_conf = input("\nSua opção [ENTER para confirmar]: ").strip().upper()
            if escolha_conf in ("", "C", "S", "SIM"):
                break
            elif escolha_conf in ("V", "0"):
                return
            elif escolha_conf == "M":
                novo_mb = input("\nDigite o novo limite em MB por arquivo (ex: 8, 50, 200): ").strip().replace(",", ".")
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
    print(f"{Cor.TITULO}  INICIANDO CONVERSÃO LOCAL PARA MICROSOFT COPILOT{Cor.RESET}")
    print(f"{Cor.TITULO}============================================================================{Cor.RESET}\n")
    print(f"Pastas monitoradas   : {len(pastas)}")
    print(f"Arquivos localizados : {Cor.NEGRITO}{total}{Cor.RESET} ({stats['tamanho_formatado']})")
    print(f"Modo Multicore CPU    : {Cor.VERDE}ATIVADO ({max_workers} núcleos paralelos){Cor.RESET}")
    print(f"Pasta de destino      : {pasta_destino}\n")

    tarefas = [
        (str(raiz), str(arq), str(pasta_cache), cfg)
        for raiz, arq in arquivos
    ]

    resultados_processados = []
    erros = []

    print("Processando documentos em paralelo...")
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futuros = {executor.submit(worker_processar_arquivo, tarefa): tarefa for tarefa in tarefas}
        for idx, futuro in enumerate(as_completed(futuros), start=1):
            nome_arq, caminho_rel, raiz_nome, meta_res, erro_str = futuro.result()
            if erro_str:
                print(f"({idx:03d}/{total:03d}) {Cor.VERMELHO}[ERRO]{Cor.RESET} {caminho_rel} -> {erro_str}")
                erros.append(f"- **[{idx:03d}]** `{caminho_rel}` -> `{erro_str}`")
            else:
                resultados_processados.append((idx, nome_arq, caminho_rel, raiz_nome, meta_res))
                status_str = f"{Cor.AZUL}[CACHE]{Cor.RESET}" if meta_res["cache"] else f"{Cor.VERDE}[OK]{Cor.RESET}"
                if meta_res.get("ocr_pages", 0) > 0:
                    status_str += f" {Cor.AMARELO}(+{meta_res['ocr_pages']} OCR){Cor.RESET}"
                print(f"({idx:03d}/{total:03d}) {status_str} {caminho_rel}")

    resultados_processados.sort(key=lambda x: x[0])

    indice_geral = [
        "# BASE DOCUMENTAL CORPORATIVA CONSOLIDADA PARA COPILOT",
        f"- **Data de consolidação:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"- **Total de documentos indexados:** {len(resultados_processados)}",
        f"- **Volume original de dados:** {stats['tamanho_formatado']}",
        f"- **Pastas de origem:** {len(pastas)}",
        "\n## ÍNDICE MESTRE DE DOCUMENTOS (ID -> ORIGEM)\n",
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
            duplicados.append(f"- `{caminho_rel}` ({doc_id}) possui conteúdo idêntico a {orig_id}")
            aviso_dup = f" *(Idêntico a {orig_id})*"
        elif f_hash:
            hashes_vistos[f_hash] = f"{doc_id} (`{caminho_rel}`)"

        tag_ocr = f" *(OCR: {meta['ocr_pages']})*" if meta.get("ocr_pages", 0) > 0 else ""
        indice_geral.append(f"- **{doc_id}** `{caminho_rel}`{tag_ocr}{aviso_dup}")

        bloco_md = (
            f"\n\n{'=' * 80}\n"
            f"## {doc_id} {nome_arq}\n"
            f"- **ID de Referência:** {doc_id}\n"
            f"- **Pasta Departamental:** `{raiz_nome}`\n"
            f"- **Caminho Relativo:** `{caminho_rel}`\n"
            f"{'=' * 80}\n\n"
            f"{meta['conteudo'] if meta.get('conteudo') else '*(Documento sem texto extraível)*'}\n"
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

    # 2. Estratégia B: Divisão por Tamanho em MB (ex: 8MB, 50MB, até 200MB)
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
                    f"*(Consulte o ÍNDICE MESTRE na Parte 01 para localizar todos os IDs de documentos)*\n\n"
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
                f"# LOTE {lote_idx:02d} DE DOCUMENTOS ({len(fatia)} arquivos)\n"
                f"- **Faixa de IDs:** [{i+1:03d}] até [{min(i+qtd_por_lote, len(blocos_gerais)):03d}]\n\n"
            )
            caminho_lote.write_text(topo_lote + "".join(fatia), encoding="utf-8")
        arquivos_gerados.append(f"por_lotes/*.md ({len(range(0, len(blocos_gerais), qtd_por_lote))} lotes de até {qtd_por_lote} docs)")

    # 4. Estratégia D: 1 Arquivo Markdown por Arquivo Original (Individual 1:1)
    if cfg.get("gerar_individual_por_arquivo", False) and docs_individuais:
        dir_indiv = pasta_destino / "individual"
        dir_indiv.mkdir(parents=True, exist_ok=True)
        for antigo in dir_indiv.glob("*.md"):
            antigo.unlink(missing_ok=True)
        for doc_id, nome_arq, caminho_rel, bloco_md in docs_individuais:
            nome_limpo = sanitizar_nome_arquivo(f"{doc_id}_{nome_arq}.md")
            (dir_indiv / nome_limpo).write_text(bloco_md.strip(), encoding="utf-8")
        arquivos_gerados.append(f"individual/*.md ({len(docs_individuais)} arquivos convertidos 1:1)")

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
                f"# CONSOLIDADO DO DEPARTAMENTO / PASTA: {nome_pasta}\n"
                f"- **Total de documentos indexados:** {len(lista_blocos)}\n\n"
            )
            arq_p.write_text(topo + "".join(lista_blocos), encoding="utf-8")
        arquivos_gerados.append(f"por_pasta/*.md ({len(blocos_por_pasta)} pastas departamentais)")

    duracao = time.time() - inicio
    relatorio = [
        "# RELATÓRIO DE AUDITORIA E INDEXAÇÃO LOCAL (MULTICORE)",
        f"- **Data/Hora:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"- **Tempo total de execução:** {duracao:.2f} segundos ({max_workers} núcleos ativos)",
        f"- **Total de documentos encontrados:** {total}",
        f"- **Volume total original:** {stats['tamanho_formatado']}",
        f"- **Indexados com sucesso:** {len(resultados_processados)}",
        f"- **Recuperados instantaneamente do cache SHA-256:** {cont_cache}",
        f"- **Documentos/Páginas/Imagens com OCR aplicado:** {cont_ocr_itens}",
        f"- **Documentos com conteúdo idêntico (duplicados):** {len(duplicados)}",
        f"- **Falhas de leitura:** {len(erros)}\n",
        "## Estrutura de Saída Gerada:",
    ]
    for ag in arquivos_gerados:
        relatorio.append(f"- `{ag}`")

    if duplicados:
        relatorio.append("\n## Documentos Idênticos Detectados:")
        relatorio.extend(duplicados)

    if erros:
        relatorio.append("\n## Falhas de Extração:")
        relatorio.extend(erros)

    (pasta_destino / "relatorio_auditoria.md").write_text("\n".join(relatorio), encoding="utf-8")

    prompt_copilot = (
        "Analise a base documental corporativa em anexo. Todos os documentos estão catalogados no "
        "ÍNDICE MESTRE DE DOCUMENTOS no início do arquivo, identificados por IDs únicos entre colchetes como [001], [002], etc.\n\n"
        "Regras obrigatórias para suas respostas:\n"
        "1. Para cada informação, valor, número ou regra citada, cite obrigatoriamente o ID e o caminho de origem (ex: 'Conforme [001] 01_Financeiro/Relatorio_2026.pdf...').\n"
        "2. Se existirem dados divergentes entre dois ou mais documentos, aponte ambos os IDs e destaque a diferença com precisão.\n"
        "3. Não invente nenhuma informação ausente nos documentos anexados; caso algo não conste na base, afirme expressamente."
    )
    (pasta_destino / "PROMPT_PARA_COPILOT.txt").write_text(prompt_copilot, encoding="utf-8")

    # Atualiza pasta de acesso rápido 'ultimo_processamento'
    if cfg.get("usar_pasta_timestamp", True):
        link_ultimo = pasta_saida_base / "ultimo_processamento"
        try:
            if link_ultimo.exists():
                shutil.rmtree(link_ultimo, ignore_errors=True)
            shutil.copytree(pasta_destino, link_ultimo, ignore=shutil.ignore_patterns("*.zip"))
        except Exception:
            pass

    # Compactação opcional em .ZIP
    arquivo_zip_gerado = None
    if cfg.get("gerar_arquivo_zip", True):
        caminho_zip = pasta_saida_base / f"consolidado_{timestamp_str}.zip"
        print(f"\n{Cor.AZUL}Gerando arquivo compactado .ZIP...{Cor.RESET}")
        compactar_pasta(pasta_destino, caminho_zip)
        tam_zip_mb = caminho_zip.stat().st_size / (1024 * 1024)
        arquivo_zip_gerado = f"{caminho_zip.name} ({tam_zip_mb:.2f} MB)"

    print(f"\n{Cor.VERDE}============================================================================{Cor.RESET}")
    print(f"{Cor.VERDE}  CONSOLIDAÇÃO CONCLUÍDA COM SUCESSO EM {duracao:.2f}s!{Cor.RESET}")
    print(f"{Cor.VERDE}============================================================================{Cor.RESET}")
    print(f"Pasta da execução    : {Cor.NEGRITO}{pasta_destino}{Cor.RESET}")
    print(f"Sucesso / Total      : {len(resultados_processados)} / {total} (Cache: {cont_cache} | OCR itens: {cont_ocr_itens})")
    print(f"Volume processado    : {stats['tamanho_formatado']}")

    print("\nArquivos disponíveis para o Microsoft Copilot:")
    for ag in arquivos_gerados:
        print(f"  -> {Cor.TITULO}{ag}{Cor.RESET}")
    print(f"  -> {Cor.AMARELO}relatorio_auditoria.md{Cor.RESET}")
    print(f"  -> {Cor.AMARELO}PROMPT_PARA_COPILOT.txt{Cor.RESET}")

    if arquivo_zip_gerado:
        print(f"\n  -> {Cor.VERDE}{Cor.NEGRITO}Pacote Compactado (.ZIP): {arquivo_zip_gerado}{Cor.RESET}")

    if pausar_no_final and "--auto" not in sys.argv:
        input("\nPressione ENTER para voltar ao menu principal...")


def selecionar_pasta_gui() -> str | None:
    """Abre janela visual nativa do Windows Explorer para o usuário navegar e selecionar a pasta."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta = filedialog.askdirectory(title="Selecione uma pasta para indexar no Copilot")
        root.destroy()
        if pasta:
            return str(Path(pasta))
    except Exception:
        pass
    return None


def menu_adicionar_pastas(cfg: dict):
    while True:
        limpar_tela()
        print(f"{Cor.TITULO}=== ADICIONAR PASTAS PARA INDEXAÇÃO ==={Cor.RESET}\n")
        print(f"Pastas atualmente configuradas: {Cor.NEGRITO}{len(cfg['pastas_entrada'])}{Cor.RESET}\n")
        print("[1] Digitar ou colar caminho(s) no terminal (separe vários com ';')")
        print("[2] Abrir janela do Windows Explorer para navegar e selecionar")
        print("[3] Informar Pasta Mãe e importar todas as suas subpastas (ex: 30 pastas de uma vez)")
        print("[0] Voltar ao menu principal")

        op = input("\nEscolha uma opção: ").strip()
        if op == "0":
            break
        elif op == "1":
            entrada = input("\nCole o(s) caminho(s) da(s) pasta(s): ").strip()
            if entrada:
                partes = [p.strip().strip('"').strip("'") for p in entrada.split(";") if p.strip()]
                adicionadas = 0
                for p in partes:
                    caminho = resolver_caminho(p)
                    if caminho.exists() and caminho.is_dir():
                        arqs = listar_arquivos_elegiveis([str(caminho)], cfg["extensoes"])
                        st = obter_estatisticas_arquivos(arqs)
                        print(f"\n  Pasta: {caminho}")
                        print(f"  Arquivos encontrados: {st['total_arquivos']} ({st['tamanho_formatado']})")
                        conf = input("  Confirmar adição desta pasta? (S/N) [S]: ").strip().upper()
                        if conf in ("", "S", "SIM"):
                            c_str = str(caminho)
                            if c_str not in cfg["pastas_entrada"]:
                                cfg["pastas_entrada"].append(c_str)
                                adicionadas += 1
                    else:
                        print(f"{Cor.VERMELHO}  [!] Pasta não encontrada: {p}{Cor.RESET}")
                salvar_config(cfg)
                if adicionadas:
                    print(f"\n{Cor.VERDE}  [+] {adicionadas} pasta(s) adicionada(s) com sucesso!{Cor.RESET}")
                time.sleep(1.2)
        elif op == "2":
            print(f"\n{Cor.AZUL}Abrindo janela do Windows Explorer para seleção...{Cor.RESET}")
            escolhida = selecionar_pasta_gui()
            if escolhida:
                caminho = resolver_caminho(escolhida)
                arqs = listar_arquivos_elegiveis([str(caminho)], cfg["extensoes"])
                st = obter_estatisticas_arquivos(arqs)
                print(f"\n  Pasta selecionada: {Cor.NEGRITO}{caminho}{Cor.RESET}")
                print(f"  Arquivos compatíveis encontrados: {Cor.VERDE}{st['total_arquivos']}{Cor.RESET} ({st['tamanho_formatado']})")
                conf = input("  Confirmar adição desta pasta? (S/N) [S]: ").strip().upper()
                if conf in ("", "S", "SIM"):
                    c_str = str(caminho)
                    if c_str not in cfg["pastas_entrada"]:
                        cfg["pastas_entrada"].append(c_str)
                        salvar_config(cfg)
                        print(f"{Cor.VERDE}  [+] Pasta adicionada com sucesso!{Cor.RESET}")
            time.sleep(1.2)
        elif op == "3":
            entrada = input("\nCole o caminho da Pasta Mãe (contém as subpastas): ").strip().strip('"').strip("'")
            if entrada:
                mae = resolver_caminho(entrada)
                if mae.exists() and mae.is_dir():
                    subpastas = [d for d in sorted(mae.iterdir()) if d.is_dir() and not d.name.startswith(".")]
                    if not subpastas:
                        subpastas = [mae]
                    print(f"\n  Localizadas {len(subpastas)} subpastas dentro de {mae.name}")
                    conf = input("  Confirmar adição de todas as subpastas? (S/N) [S]: ").strip().upper()
                    if conf in ("", "S", "SIM"):
                        adicionadas = 0
                        for sp in subpastas:
                            c_str = str(sp.resolve())
                            if c_str not in cfg["pastas_entrada"]:
                                cfg["pastas_entrada"].append(c_str)
                                adicionadas += 1
                        salvar_config(cfg)
                        print(f"{Cor.VERDE}  [+] {adicionadas} subpasta(s) adicionada(s)!{Cor.RESET}")
                else:
                    print(f"{Cor.VERMELHO}  [!] Pasta mãe não encontrada.{Cor.RESET}")
                time.sleep(1.2)


def menu_listar_remover_pastas(cfg: dict):
    while True:
        limpar_tela()
        print(f"{Cor.TITULO}=== PASTAS CONFIGURADAS ({len(cfg['pastas_entrada'])}) ==={Cor.RESET}\n")
        if not cfg["pastas_entrada"]:
            print("Nenhuma pasta cadastrada.")
        else:
            for i, p_str in enumerate(cfg["pastas_entrada"], start=1):
                p = resolver_caminho(p_str)
                if p.exists():
                    arqs = listar_arquivos_elegiveis([str(p)], cfg["extensoes"])
                    st = obter_estatisticas_arquivos(arqs)
                    status = f"{Cor.VERDE}OK ({st['total_arquivos']} arqs | {st['tamanho_formatado']}){Cor.RESET}"
                else:
                    status = f"{Cor.VERMELHO}NÃO ENCONTRADA{Cor.RESET}"
                print(f"  [{i:02d}] {p}  -> {status}")

        print("\nOpções:")
        print("  Digite o NÚMERO da pasta para removê-la")
        print("  Digite 'T' para limpar todas as pastas")
        print("  Digite '0' ou ENTER para voltar")

        escolha = input("\nSua escolha: ").strip().upper()
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
        print(f"{Cor.TITULO}=== CONFIGURAÇÃO DE SAÍDAS, PARTICIONAMENTO E OCR ==={Cor.RESET}\n")
        print(f"{Cor.NEGRITO}--- Modalidades de Saída Markdown ---{Cor.RESET}")
        print(f"[1] Gerar Arquivo Único (corpus_completo.md)       : {'SIM' if cfg.get('gerar_arquivo_unico', True) else 'NÃO'}")
        print(f"[2] Dividir por Tamanho Máximo em MB              : {'SIM' if cfg.get('dividir_por_tamanho', True) else 'NÃO'}")
        print(f"[3]   ↳ Limite de Tamanho por Parte (MB)          : {cfg.get('limite_mb_por_parte', 8.0)} MB")
        print(f"[4] Dividir por Quantidade de Documentos por Lote : {'SIM' if cfg.get('dividir_por_quantidade', False) else 'NÃO'}")
        print(f"[5]   ↳ Quantidade de Docs por Lote               : {cfg.get('qtd_docs_por_parte', 50)} docs")
        print(f"[6] Gerar 1 Arquivo Markdown por Documento (1:1)  : {'SIM' if cfg.get('gerar_individual_por_arquivo', False) else 'NÃO'}")
        print(f"[7] Gerar Consolidado por Pasta Departamental      : {'SIM' if cfg.get('gerar_por_pasta_raiz', True) else 'NÃO'}")
        print(f"\n{Cor.NEGRITO}--- Compactação & Organização com Timestamp ---{Cor.RESET}")
        print(f"[8] Salvar em Pasta com Timestamp por Execução    : {'SIM' if cfg.get('usar_pasta_timestamp', True) else 'NÃO'}")
        print(f"[9] Gerar Pacote Compactado (.ZIP) ao final       : {'SIM' if cfg.get('gerar_arquivo_zip', True) else 'NÃO'}")
        print(f"\n{Cor.NEGRITO}--- Reconhecimento Óptico de Caracteres (OCR Local) ---{Cor.RESET}")
        print(f"[10] OCR em PDFs (Escaneados e Imagens Mistas)    : {'ATIVADO' if cfg.get('usar_ocr_em_pdf', True) else 'DESATIVADO'}")
        print(f"[11] OCR em Imagens Embutidas (Word, Excel, PPTX) : {'ATIVADO' if cfg.get('usar_ocr_em_imagens_embutidas', True) else 'DESATIVADO'}")
        print(f"[12] Resolução DPI para OCR em PDF               : {cfg.get('dpi_ocr_pdf', 200)} DPI")
        print(f"[13] Dimensão mínima de imagem para OCR (px)     : {cfg.get('min_dimensao_imagem_ocr', 80)} px")
        print(f"\n{Cor.NEGRITO}--- Desempenho & Destino ---{Cor.RESET}")
        print(f"[14] Processamento Multicore Concorrente (CPU)   : {'ATIVADO' if cfg.get('multicore', True) else 'DESATIVADO'}")
        print(f"[15] Alterar Pasta de Saída (Atual: {cfg.get('pasta_saida')})")
        print(f"\n{Cor.AMARELO}[R] Restaurar Configuração Padrão Original{Cor.RESET}")
        print("[0] Voltar ao menu principal")

        op = input("\nEscolha a opção para alternar ou editar: ").strip().upper()
        if op == "0" or not op:
            break
        elif op == "1":
            cfg["gerar_arquivo_unico"] = not cfg.get("gerar_arquivo_unico", True)
        elif op == "2":
            cfg["dividir_por_tamanho"] = not cfg.get("dividir_por_tamanho", True)
        elif op == "3":
            val = input("Digite o limite em MB por arquivo (ex: 8, 50, 100, 200): ").strip().replace(",", ".")
            try:
                mb = float(val)
                if mb >= 0.5:
                    cfg["limite_mb_por_parte"] = mb
            except ValueError:
                pass
        elif op == "4":
            cfg["dividir_por_quantidade"] = not cfg.get("dividir_por_quantidade", False)
        elif op == "5":
            val = input("Digite a quantidade de documentos por lote (ex: 25, 50, 100): ").strip()
            if val.isdigit() and int(val) > 0:
                cfg["qtd_docs_por_parte"] = int(val)
        elif op == "6":
            cfg["gerar_individual_por_arquivo"] = not cfg.get("gerar_individual_por_arquivo", False)
        elif op == "7":
            cfg["gerar_por_pasta_raiz"] = not cfg.get("gerar_por_pasta_raiz", True)
        elif op == "8":
            cfg["usar_pasta_timestamp"] = not cfg.get("usar_pasta_timestamp", True)
        elif op == "9":
            cfg["gerar_arquivo_zip"] = not cfg.get("gerar_arquivo_zip", True)
        elif op == "10":
            cfg["usar_ocr_em_pdf"] = not cfg.get("usar_ocr_em_pdf", True)
        elif op == "11":
            cfg["usar_ocr_em_imagens_embutidas"] = not cfg.get("usar_ocr_em_imagens_embutidas", True)
        elif op == "12":
            val = input("Digite o DPI de renderização para OCR em PDF (ex: 150, 200, 300): ").strip()
            if val.isdigit() and int(val) >= 100:
                cfg["dpi_ocr_pdf"] = int(val)
        elif op == "13":
            val = input("Digite a dimensão mínima em pixels para OCR de imagem (ex: 60, 80, 100): ").strip()
            if val.isdigit() and int(val) >= 20:
                cfg["min_dimensao_imagem_ocr"] = int(val)
        elif op == "14":
            cfg["multicore"] = not cfg.get("multicore", True)
        elif op == "15":
            novo = input("Digite o novo caminho para a pasta de saída: ").strip().strip('"')
            if novo:
                cfg["pasta_saida"] = str(resolver_caminho(novo))
        elif op == "R":
            pastas_atuais = cfg.get("pastas_entrada", [])
            cfg.clear()
            cfg.update(CONFIG_PADRAO.copy())
            cfg["pastas_entrada"] = pastas_atuais
            print(f"\n{Cor.VERDE}[OK] Configurações restauradas para o padrão!{Cor.RESET}")
            time.sleep(1.0)

        salvar_config(cfg)


def menu_principal():
    cfg = carregar_config()
    while True:
        limpar_tela()
        qtd_pastas = len(cfg["pastas_entrada"])
        arquivos = listar_arquivos_elegiveis(cfg["pastas_entrada"], cfg["extensoes"]) if qtd_pastas else []
        stats = obter_estatisticas_arquivos(arquivos) if arquivos else {"total_arquivos": 0, "tamanho_formatado": "0 B"}
        num_cpus = os.cpu_count() or 4

        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f"{Cor.TITULO}   CONSOLIDADOR LOCAL DE DOCUMENTOS PARA MICROSOFT COPILOT (v4){Cor.RESET}")
        print(f"{Cor.TITULO}   Ambiente: Windows 11 | 100% Offline | {num_cpus} Núcleos CPU Concorrentes{Cor.RESET}")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")
        print(f" Status Atual:")
        print(f"  • Pastas monitoradas      : {Cor.NEGRITO}{qtd_pastas}{Cor.RESET}")
        print(f"  • Arquivos localizados    : {Cor.VERDE}{Cor.NEGRITO}{stats['total_arquivos']}{Cor.RESET} ({Cor.AMARELO}{stats['tamanho_formatado']}{Cor.RESET})")
        print(f"  • Multicore Paralelo      : {Cor.VERDE + 'ATIVADO' if cfg.get('multicore', True) else Cor.AMARELO + 'DESATIVADO'} ({num_cpus} cores){Cor.RESET}")
        print(f"  • OCR em PDFs e Imagens   : {Cor.VERDE + 'ATIVADO' if cfg.get('usar_ocr_em_pdf', True) else Cor.AMARELO + 'DESATIVADO'}{Cor.RESET}")
        print(f"  • OCR Imagens Embutidas   : {Cor.VERDE + 'ATIVADO' if cfg.get('usar_ocr_em_imagens_embutidas', True) else Cor.AMARELO + 'DESATIVADO'} (Word/Excel/PPTX){Cor.RESET}")

        partes_status = []
        if cfg.get("gerar_arquivo_unico", True):
            partes_status.append("Único")
        if cfg.get("dividir_por_tamanho", True):
            partes_status.append(f"Limite {cfg.get('limite_mb_por_parte', 8.0)}MB")
        if cfg.get("dividir_por_quantidade", False):
            partes_status.append(f"Lotes de {cfg.get('qtd_docs_por_parte', 50)} docs")
        if cfg.get("gerar_individual_por_arquivo", False):
            partes_status.append("1:1 Individual")
        if cfg.get("gerar_por_pasta_raiz", True):
            partes_status.append("Por Pasta")
        div_str = " | ".join(partes_status) if partes_status else "Nenhuma"

        print(f"  • Divisão de Arquivo      : {div_str}")
        print(f"  • Organização / Timestamp : {'Pasta com Data/Hora ' if cfg.get('usar_pasta_timestamp', True) else 'Direta '}"
              f"{'+ Compactar .ZIP' if cfg.get('gerar_arquivo_zip', True) else ''}")
        print(f"  • Destino base            : {cfg.get('pasta_saida')}")
        print(f"{Cor.TITULO}----------------------------------------------------------------------------{Cor.RESET}")
        print(f" {Cor.VERDE}{Cor.NEGRITO}[1] EXECUTAR CONVERSÃO E GERAR ARQUIVOS PARA O COPILOT{Cor.RESET}")
        print(f" [2] Adicionar pastas (Windows Explorer, digitar caminhos ou Pasta Mãe)")
        print(f" [3] Listar / Verificar / Remover pastas configuradas ({qtd_pastas})")
        print(f" [4] Recarregar lista de pastas do arquivo 'pastas.txt'")
        print(f" [5] Configurar opções de saída (Tamanhos/MB, Lotes, ZIP, Timestamp, OCR)")
        print(f" [6] Limpar cache de extração SHA-256 (forçar releitura completa)")
        print(f" [0] Sair")
        print(f"{Cor.TITULO}============================================================================{Cor.RESET}")

        opcao = input("\nDigite o número da opção desejada: ").strip()
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
            print(f"\n{Cor.VERDE}[OK] Cache SHA-256 limpo com sucesso!{Cor.RESET}")
            time.sleep(1.0)
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
            print("\nEncerrado pelo usuário.")
