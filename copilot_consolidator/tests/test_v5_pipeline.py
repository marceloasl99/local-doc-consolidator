"""Unit tests for local-doc-consolidator v5.0 offline pipeline."""

import hashlib
import os
import sys
import tempfile
import unittest
from pathlib import Path

# Adiciona o diretório do projeto ao sys.path
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import app


class TestOfflinePipelineV5(unittest.TestCase):
    """Testes de conformidade, qualidade textual e integridade da versão 5.0."""

    def test_zero_network_imports_in_app(self):
        """Garante que app.py não importa nenhuma biblioteca que faça chamadas de rede."""
        with open(BASE_DIR / "app.py", "r", encoding="utf-8") as f:
            code = f.read()

        proibidos = [
            "requests", "urllib.request", "httpx", "aiohttp",
            "openai", "deepseek", "litellm", "google.generativeai"
        ]
        for lib in proibidos:
            self.assertNotIn(f"import {lib}", code, f"Biblioteca de rede '{lib}' não é permitida.")
            self.assertNotIn(f"from {lib}", code, f"Biblioteca de rede '{lib}' não é permitida.")

    def test_mojibake_repair(self):
        """Valida que textos com dupla codificação CP1252/UTF-8 são restaurados com acentuação correta."""
        corrompido = (
            "ObservaÃ§Ã£o: Cloudflare Proxy pode retornar IPs da borda, nÃ£o do origin. "
            "MÃ©todo: GET somente leitura. ComunicaÃ§Ã£o efetuada."
        )
        reparado, qtd = app.reparar_mojibake(corrompido)
        self.assertGreater(qtd, 0)
        self.assertIn("Observação", reparado)
        self.assertIn("não", reparado)
        self.assertIn("Método", reparado)
        self.assertIn("Comunicação", reparado)
        self.assertNotIn("Ã§", reparado)
        self.assertNotIn("Ã£", reparado)

    def test_ocr_spurious_space_debreaking(self):
        """Valida a recombinação fonotática de palavras quebradas por segmentação OCR."""
        amostra_ocr = (
            "Aute nticacao efetuada com comprovra nte de pagamento. "
            "Verifique o docu mento e o estabeleci mento."
        )
        corrigido, qtd = app.corrigir_espacos_ocr(amostra_ocr)
        self.assertGreater(qtd, 0)
        self.assertIn("Autenticacao", corrigido)
        self.assertIn("documento", corrigido)
        self.assertIn("estabelecimento", corrigido)

    def test_pii_redaction(self):
        """Testa a redação e mascaramento de dados sensíveis (LGPD / Compliance) quando a opção está ativada."""
        texto_original = (
            "Cliente: João da Silva Sauro, CPF: 123.456.789-00, RG: 12.345.678-9, "
            "CNPJ: 12.345.678/0001-90, Endereço: Rua das Flores, 123 - Apto 4, CEP: 01310-100, "
            "contato: joao.silva@empresa.com.br, telefone: (11) 98765-4321."
        )
        redigido = app.redigir_pii(texto_original)

        self.assertNotIn("123.456.789-00", redigido)
        self.assertNotIn("12.345.678/0001-90", redigido)
        self.assertNotIn("12.345.678-9", redigido)
        self.assertNotIn("Rua das Flores, 123", redigido)
        self.assertNotIn("01310-100", redigido)
        self.assertNotIn("joao.silva@empresa.com.br", redigido)
        self.assertNotIn("98765-4321", redigido)
        self.assertIn("[CPF REDIGIDO]", redigido)
        self.assertIn("[CNPJ REDIGIDO]", redigido)
        self.assertIn("[RG REDIGIDO]", redigido)
        self.assertIn("[ENDEREÇO REDIGIDO]", redigido)
        self.assertIn("[CEP REDIGIDO]", redigido)
        self.assertIn("[EMAIL REDIGIDO]", redigido)
        self.assertIn("[TELEFONE REDIGIDO]", redigido)
        self.assertIn("[NOME REDIGIDO]", redigido)

    def test_pii_default_is_disabled(self):
        """Valida que por padrão os dados sensíveis NÃO são excluídos (preservando dados originais)."""
        cfg_padrao = app.CONFIG_PADRAO
        self.assertFalse(cfg_padrao.get("mascarar_pii", True), "Por padrão mascarar_pii deve ser False.")

    def test_csv_structured_extraction(self):
        """Valida que CSV com diferentes delimitadores é convertido em tabela Markdown."""
        csv_conteudo = "ID;Servidor;IP;Status\n1;SRV-APP-01;192.168.1.10;Ativo\n2;SRV-DB-01;192.168.1.20;Ativo"
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="utf-8") as tmp:
            tmp.write(csv_conteudo)
            tmp_path = Path(tmp.name)

        try:
            texto_md = app.extrair_csv_estruturado(tmp_path)
            self.assertIn("| ID | Servidor | IP | Status |", texto_md)
            self.assertIn("| --- | --- | --- | --- |", texto_md)
            self.assertIn("| 1 | SRV-APP-01 | 192.168.1.10 | Ativo |", texto_md)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def test_cache_config_hash_invalidation(self):
        """A chave real do cache muda com DPI, PII e idioma (rótulos do conteúdo dependem do idioma)."""
        base = dict(app.CONFIG_PADRAO, idioma="pt")
        chave = app.chave_cache(base)
        self.assertNotEqual(chave, app.chave_cache(dict(base, dpi_ocr_pdf=300)), "DPI deve invalidar o cache.")
        self.assertNotEqual(chave, app.chave_cache(dict(base, mascarar_pii=True)), "PII deve invalidar o cache.")
        self.assertNotEqual(chave, app.chave_cache(dict(base, idioma="en")), "Idioma deve invalidar o cache.")
        self.assertEqual(chave, app.chave_cache(dict(base)), "Mesma configuração deve gerar a mesma chave.")
        self.assertEqual(app.parametros_extracao(base)["versao"], app.VERSAO_EXTRATOR)

    def test_i18n_dictionary_completeness(self):
        """Garante que todas as chaves do dicionário em português existem em inglês."""
        pt_keys = set(app.I18N["pt"].keys())
        en_keys = set(app.I18N["en"].keys())

        missing_in_en = pt_keys - en_keys
        self.assertEqual(len(missing_in_en), 0, f"Chaves faltantes em EN: {missing_in_en}")

    def test_language_argument_parsing(self):
        """Valida que obter_idioma reconhece flags de idioma via sys.argv."""
        orig_argv = sys.argv.copy()
        try:
            sys.argv = ["app.py", "--lang", "en"]
            self.assertEqual(app.obter_idioma({}), "en")

            sys.argv = ["app.py", "--lang", "pt"]
            self.assertEqual(app.obter_idioma({}), "pt")

            sys.argv = ["app.py", "--english"]
            self.assertEqual(app.obter_idioma({}), "en")
        finally:
            sys.argv = orig_argv


class TestCorrecoesV61(unittest.TestCase):
    """Regressões dos bugs corrigidos na v6.1."""

    def test_idioma_do_menu_prevalece_sobre_argv(self):
        """O launcher passa --lang en, mas a troca de idioma no menu precisa funcionar."""
        orig_argv = sys.argv.copy()
        try:
            sys.argv = ["app.py", "--lang", "en"]
            self.assertEqual(app.obter_idioma({"idioma": "pt"}), "pt")
            cfg = app.aplicar_argumentos({"idioma": "pt"})
            self.assertEqual(cfg["idioma"], "en")
        finally:
            sys.argv = orig_argv

    def test_ocr_nao_cola_siglas_nem_palavras_reais(self):
        texto = "Fornecedor Cloud Services LTDA e Comércio NT sobre proteção que lhe disse nossa mente contra todos"
        corrigido, _ = app.corrigir_espacos_ocr(texto)
        self.assertEqual(corrigido, texto)

    def test_mojibake_nao_altera_texto_legitimo(self):
        texto = "SÃO PAULO - Ação de cobrança nº 123 — “aspas” corretas"
        reparado, qtd = app.reparar_mojibake(texto)
        self.assertEqual(reparado, texto)
        self.assertEqual(qtd, 0)

    def test_layout_ocr_preserva_linhas(self):
        itens = [
            ([[10, 10], [100, 10], [100, 20], [10, 20]], "Valor:", 0.9),
            ([[110, 11], [200, 11], [200, 21], [110, 21]], "R$ 45.890,00", 0.9),
            ([[10, 30], [120, 30], [120, 40], [10, 40]], "Data: 12/03/2026", 0.9),
        ]
        self.assertEqual(app.reconstruir_layout_ocr(itens), "Valor: R$ 45.890,00\nData: 12/03/2026")

    def test_tabela_escapa_pipe_e_remove_colunas_vazias(self):
        md = app.tabela_para_markdown([["A", "", "B"], ["x|y", "", "2"]])
        self.assertEqual(md, "| A | B |\n| --- | --- |\n| x\\|y | 2 |")

    def test_markdown_rebaixa_titulos(self):
        md = app.rebaixar_titulos_markdown("# Título\n```\n# comentário\n```\n## Seção")
        self.assertEqual(md, "### Título\n```\n# comentário\n```\n#### Seção")

    def test_cabecalhos_rodapes_repetidos(self):
        paginas = [f"EMPRESA X - CONFIDENCIAL\nConteúdo da página {i} com texto único {i * 7}\nPágina {i} de 5" for i in range(1, 6)]
        limpas, removidas = app.remover_cabecalhos_rodapes(paginas)
        self.assertEqual(limpas[0], paginas[0], "A primeira página fica intacta.")
        self.assertGreater(removidas, 0)
        for p in limpas[1:]:
            self.assertNotIn("CONFIDENCIAL", p)
            self.assertNotIn("de 5", p)
            self.assertIn("Conteúdo da página", p)

    def test_varredura_ignora_pasta_de_saida_e_temporarios(self):
        with tempfile.TemporaryDirectory() as tmp:
            raiz = Path(tmp)
            (raiz / "saida").mkdir()
            (raiz / "saida" / "corpus_completo.md").write_text("velho", encoding="utf-8")
            (raiz / "doc.txt").write_text("conteúdo", encoding="utf-8")
            (raiz / "~$temp.docx").write_text("x", encoding="utf-8")
            (raiz / "vazio.txt").write_text("", encoding="utf-8")
            (raiz / "antigo.doc").write_text("x", encoding="utf-8")
            ignorados = {}
            arquivos, _ = app.listar_arquivos_elegiveis([str(raiz)], [".txt", ".md", ".docx"],
                                                         excluir=[raiz / "saida"], ignorados=ignorados)
            self.assertEqual([a.name for _, a in arquivos], ["doc.txt"])
            self.assertEqual(ignorados.get(".doc"), 1)

    def test_pipeline_ponta_a_ponta(self):
        """Gera a base a partir de arquivos variados e valida estrutura, IDs e duplicados."""
        import docx

        with tempfile.TemporaryDirectory() as tmp:
            entrada = Path(tmp) / "Empresa"
            (entrada / "Financeiro").mkdir(parents=True)
            (entrada / "Juridico").mkdir(parents=True)
            (entrada / "Financeiro" / "contas.csv").write_text("Conta;Valor\nLuz;100\nAgua;50", encoding="utf-8")
            (entrada / "Financeiro" / "copia.csv").write_text("Conta;Valor\nLuz;100\nAgua;50", encoding="utf-8")
            (entrada / "Juridico" / "notas.md").write_text("# Contrato\nCláusula | única", encoding="utf-8")
            d = docx.Document()
            d.add_heading("Política Interna", level=1)
            d.add_paragraph("Texto da política.")
            tab = d.add_table(rows=2, cols=2)
            tab.cell(0, 0).text, tab.cell(0, 1).text = "Nível", "Prazo"
            tab.cell(1, 0).text, tab.cell(1, 1).text = "Interno", "5 anos"
            d.save(entrada / "Juridico" / "politica.docx")

            cfg = dict(app.CONFIG_PADRAO, idioma="pt", pastas_entrada=[str(entrada)],
                       pasta_saida=str(Path(tmp) / "saida"), max_workers=2,
                       gerar_arquivo_zip=True, usar_ocr_em_pdf=False, usar_ocr_em_imagens_embutidas=False)
            res = app.executar_pipeline(cfg, pausar_no_final=False, interativo=False)
            self.assertIsNotNone(res)
            self.assertEqual(res["erros"], [])
            corpus = (res["pasta"] / "corpus_completo.md").read_text(encoding="utf-8")

            self.assertIn("## ÍNDICE DE DOCUMENTOS", corpus)
            self.assertIn("### Pasta: Empresa/Financeiro", corpus)
            self.assertIn("### Pasta: Empresa/Juridico", corpus)
            for doc_id in ("[001]", "[002]", "[003]", "[004]"):
                self.assertIn(f"*[Fim do documento {doc_id}]*", corpus)
            # copia.csv é idêntica a contas.csv: conteúdo omitido e marcado
            self.assertIn("duplicado de [001]", corpus)
            self.assertIn("Conteúdo omitido para evitar redundância", corpus)
            self.assertEqual(corpus.count("| Luz | 100 |"), 1)
            # títulos do .md rebaixados e pipe escapado; tabela do Word convertida
            self.assertIn("### Contrato", corpus)
            self.assertNotIn("\n# Contrato", corpus)
            self.assertIn("### Política Interna", corpus)
            self.assertIn("| Interno | 5 anos |", corpus)
            # saídas auxiliares
            self.assertTrue((res["pasta"] / "relatorio_auditoria.md").exists())
            self.assertTrue((res["pasta"] / "PROMPT_PARA_COPILOT.txt").exists())
            self.assertEqual(len(list((res["pasta"] / "por_pasta").glob("*.md"))), 2)
            self.assertIsNotNone(res["zip"])
            linhas_jsonl = (res["pasta"] / "corpus_rag.jsonl").read_text(encoding="utf-8").strip().splitlines()
            self.assertEqual(len(linhas_jsonl), 3, "O duplicado não entra na base RAG.")

    def test_particionamento_preserva_ordem_e_divide_documento_grande(self):
        docs = [{"id": f"[{i:03d}]", "nome": f"d{i}.txt"} for i in range(1, 4)]
        blocos = [(docs[0], "a" * 100), (docs[1], ("linha\n\n" * 400)), (docs[2], "c" * 100)]
        partes = app.particionar(blocos, 1000, "pt")
        ids = [d["id"] for parte in partes for d, _ in parte]
        self.assertEqual(ids[0], "[001]")
        self.assertEqual(ids[-1], "[003]")
        self.assertGreater(ids.count("[002]"), 1, "Documento maior que o limite vira vários trechos.")
        self.assertTrue(all(len(b.encode("utf-8")) <= 1200 for parte in partes for _, b in parte))


class TestRecursosV62(unittest.TestCase):
    """Testes da nova funcionalidade de salvar e referenciar diagramas/esquemas na v6.2."""

    def test_classificar_imagem_diagrama(self):
        from PIL import Image, ImageDraw
        import io

        # 1. Imagem que é diagrama geométrico (linhas e caixas)
        im_diag = Image.new("RGB", (600, 400), (255, 255, 255))
        d = ImageDraw.Draw(im_diag)
        d.rectangle([50, 50, 200, 150], outline=(0, 0, 0), width=4)
        d.rectangle([350, 50, 500, 150], outline=(0, 0, 0), width=4)
        d.line([200, 100, 350, 100], fill=(0, 0, 0), width=4)
        buf = io.BytesIO()
        im_diag.save(buf, format="PNG")
        eh_diag, motivo, rotulos, png, dims = app.classificar_e_extrair_diagrama(buf.getvalue(), min_dim_diag=150)
        self.assertTrue(eh_diag, "Desenho geométrico com linhas deve ser identificado como diagrama/esquema.")
        self.assertIsNotNone(png)
        self.assertEqual(dims, (600, 400))

        # 2. Ícone pequeno (< 50px) deve ser rejeitado
        im_pequena = Image.new("RGB", (40, 40), (255, 255, 255))
        buf_p = io.BytesIO()
        im_pequena.save(buf_p, format="PNG")
        eh_p, motivo_p, _, _, _ = app.classificar_e_extrair_diagrama(buf_p.getvalue(), min_dim_diag=150)
        self.assertFalse(eh_p)
        self.assertEqual(motivo_p, "small_icon")

        # 3. Banner/divisor com proporção extrema (> 12:1) deve ser rejeitado
        im_banner = Image.new("RGB", (800, 30), (255, 255, 255))
        buf_b = io.BytesIO()
        im_banner.save(buf_b, format="PNG")
        eh_b, motivo_b, _, _, _ = app.classificar_e_extrair_diagrama(buf_b.getvalue(), min_dim_diag=150, prop_max=12.0)
        self.assertFalse(eh_b)
        self.assertEqual(motivo_b, "divider_or_banner")

    def test_pipeline_com_salvamento_de_diagramas(self):
        """Valida que diagramas são salvos em imagens/, catalogados e referenciados no Markdown."""
        from PIL import Image, ImageDraw
        import io
        import json

        with tempfile.TemporaryDirectory() as tmp:
            entrada = Path(tmp) / "Projeto"
            entrada.mkdir(parents=True)

            # Cria um diagrama sintético
            im = Image.new("RGB", (500, 350), (255, 255, 255))
            d = ImageDraw.Draw(im)
            d.rectangle([40, 40, 200, 150], outline=(0, 0, 0), width=4)
            d.rectangle([280, 40, 440, 150], outline=(0, 0, 0), width=4)
            d.line([200, 95, 280, 95], fill=(0, 0, 0), width=4)
            im.save(entrada / "fluxo_arquitetura.png")

            # 1. Executa com salvar_diagramas_em_pasta=True
            cfg_ativo = dict(
                app.CONFIG_PADRAO,
                idioma="pt",
                pastas_entrada=[str(entrada)],
                pasta_saida=str(Path(tmp) / "saida_com_diag"),
                max_workers=2,
                salvar_diagramas_em_pasta=True,
                gerar_arquivo_zip=True,
            )
            res_ativo = app.executar_pipeline(cfg_ativo, pausar_no_final=False, interativo=False)
            self.assertIsNotNone(res_ativo)
            pasta_ativo = res_ativo["pasta"]

            # Valida que a pasta imagens/ existe e contém o arquivo ordenado
            pasta_img = pasta_ativo / "imagens"
            self.assertTrue(pasta_img.exists(), "A pasta imagens/ deve ser criada.")
            arquivos_img = list(pasta_img.glob("*.png"))
            self.assertEqual(len(arquivos_img), 1, "Deve existir exatamente 1 diagrama salvo.")
            self.assertEqual(arquivos_img[0].name, "doc001_p01_fig01.png")

            # Valida card e link no corpus_completo.md
            corpus_md = (pasta_ativo / "corpus_completo.md").read_text(encoding="utf-8")
            self.assertIn("![Diagrama:", corpus_md)
            self.assertIn("imagens/doc001_p01_fig01.png", corpus_md)
            self.assertIn("Figura / Diagrama Referenciado", corpus_md)

            # Valida tabela no 00_INDICE_MESTRE.md
            indice_md = (pasta_ativo / "00_INDICE_MESTRE.md").read_text(encoding="utf-8")
            self.assertIn("Catálogo de Diagramas e Esquemas Técnicos", indice_md)
            self.assertIn("FIG-001", indice_md)
            self.assertIn("doc001_p01_fig01.png", indice_md)

            # Valida registro no corpus_rag.jsonl
            rag_linhas = (pasta_ativo / "corpus_rag.jsonl").read_text(encoding="utf-8").strip().splitlines()
            self.assertEqual(len(rag_linhas), 1)
            rag_item = json.loads(rag_linhas[0])
            self.assertIn("imagens_referenciadas", rag_item)
            self.assertEqual(rag_item["imagens_referenciadas"][0]["arquivo"], "imagens/doc001_p01_fig01.png")

            # 2. Executa com salvar_diagramas_em_pasta=False
            cfg_inativo = dict(
                app.CONFIG_PADRAO,
                idioma="pt",
                pastas_entrada=[str(entrada)],
                pasta_saida=str(Path(tmp) / "saida_sem_diag"),
                max_workers=2,
                salvar_diagramas_em_pasta=False,
                gerar_arquivo_zip=False,
            )
            res_inativo = app.executar_pipeline(cfg_inativo, pausar_no_final=False, interativo=False)
            pasta_inativo = res_inativo["pasta"]
            self.assertFalse((pasta_inativo / "imagens").exists(), "A pasta imagens/ NÃO deve ser criada quando desativada.")
            corpus_inativo = (pasta_inativo / "corpus_completo.md").read_text(encoding="utf-8")
            self.assertNotIn("imagens/doc001_p01_fig01.png", corpus_inativo)


if __name__ == "__main__":
    unittest.main()
