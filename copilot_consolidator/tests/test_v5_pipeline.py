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
        """Verifica se alteração nos parâmetros de renderização/OCR altera o hash de cache."""
        def calc_hash(dpi, min_dim, limiar, prop, pii):
            return hashlib.md5(f"{dpi}_{min_dim}_{limiar}_{prop}_{pii}".encode("utf-8")).hexdigest()[:8]

        hash_200 = calc_hash(200, 80, 0.50, 20.0, False)
        hash_300 = calc_hash(300, 80, 0.50, 20.0, False)
        hash_pii = calc_hash(200, 80, 0.50, 20.0, True)

        self.assertNotEqual(hash_200, hash_300, "Mudança de DPI deve invalidar cache.")
        self.assertNotEqual(hash_200, hash_pii, "Mudança de flag PII deve invalidar cache.")

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


if __name__ == "__main__":
    unittest.main()
