"""Executor da Bateria de Testes com Todos os Cenários.

Executa o pipeline em cada um dos 8 arquivos de teste gerados,
validando as regras de negócio em SQLite e na planilha Excel gerada.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import openpyxl

from src.core.cep_service import listar_ceps_pendentes
from src.core.database import get_cep_db, get_triagem_db, initialize_schema_if_needed
from src.runner.main_seco import processar_fluxo_seco

TEST_DATA_DIR = BASE_DIR / "tests" / "test_data"



class TestBateriaCenarios(unittest.TestCase):
    """Executa a validação de cada arquivo de teste."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="vre_bateria_")
        self.triagem_db = Path(self.temp_dir) / "triagem_vre.db"
        self.cep_db = Path(self.temp_dir) / "cep_areas.db"
        self.saida_xlsx = Path(self.temp_dir) / "saida_teste.xlsx"

        # Copia os bancos já semeados com dados oficiais para o diretório de teste
        db_real_triagem = BASE_DIR / "database" / "triagem_vre.db"
        db_real_cep = BASE_DIR / "database" / "cep_areas.db"

        if db_real_triagem.exists():
            shutil.copy(db_real_triagem, self.triagem_db)
            conn = get_triagem_db(self.triagem_db)
            conn.execute("DELETE FROM vre_backlog;")
            conn.commit()
            conn.close()
        else:
            conn = get_triagem_db(self.triagem_db)
            initialize_schema_if_needed(conn, is_cep=False)
            conn.close()

        if db_real_cep.exists():
            shutil.copy(db_real_cep, self.cep_db)
            conn = get_cep_db(self.cep_db)
            conn.execute("DELETE FROM cep_pendentes;")
            conn.commit()
            conn.close()
        else:
            conn = get_cep_db(self.cep_db)
            initialize_schema_if_needed(conn, is_cep=True)
            conn.close()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_01_cenario_novas_empresas(self):
        """Valida que empresas novas entram sem fiscal e com status PENDENTE."""
        csv_file = TEST_DATA_DIR / "01_teste_empresas_novas.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        self.assertEqual(ws.max_column, 32)
        self.assertIn(ws.cell(3, 2).value, ("", None), "Status pendente deve permanecer vazio na planilha ativa")
        self.assertIn(ws.cell(3, 3).value, (None, ""))  # Fiscal vazio
        self.assertEqual(ws.cell(3, 4).value, "Sem Restrição Livre")
        self.assertEqual(ws.cell(3, 5).value, "Sem Restrição")

    def test_02_cenario_historico_fiscais(self):
        """Valida que CNPJ com histórico prévio recebe o fiscal correto."""
        csv_file = TEST_DATA_DIR / "02_teste_historico_fiscais.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        # Protocolo com fiscal histórico e não noturno -> deve receber fiscal e 'Diurno Livre'
        fiscal_atribuido = ws.cell(3, 3).value
        self.assertTrue(bool(fiscal_atribuido), "O fiscal histórico deve ser preenchido.")
        self.assertEqual(ws.cell(3, 4).value, "Diurno Livre")

    def test_03_cenario_noturno_e_heranca(self):
        """Valida que restaurantes e bares herdam a marcação noturna em Tipo de Serviço e Restrição."""
        csv_file = TEST_DATA_DIR / "03_teste_cnaes_noturnos_heranca.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        # Linha 3: Churrascaria/Restaurante (5611-2/01)
        self.assertEqual(ws.cell(3, 4).value, "Noturno Livre")
        self.assertEqual(ws.cell(3, 5).value, "Noturno")

        # Linha 4: Bar (5611-2/02)
        self.assertEqual(ws.cell(4, 4).value, "Noturno Livre")
        self.assertEqual(ws.cell(4, 5).value, "Noturno")

        # Linha 5: Danceteria (9329-8/01)
        self.assertEqual(ws.cell(5, 4).value, "Noturno Livre")
        self.assertEqual(ws.cell(5, 5).value, "Noturno")

    def test_04_cenario_restricao(self):
        """Valida que sucatas e frigoríficos recebem tipo de serviço e restrição corretos."""
        csv_file = TEST_DATA_DIR / "04_teste_cnaes_restricao.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        # Sucatas (4687-7/01)
        self.assertEqual(ws.cell(3, 4).value, "Com Restrição Livre")
        self.assertEqual(ws.cell(3, 5).value, "Com Restrição")

    def test_05_cenario_combinado(self):
        """Valida empresa com Noturno e Restrição simultâneos (Noturno prevalece)."""
        csv_file = TEST_DATA_DIR / "05_teste_noturno_e_restricao_combinados.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        self.assertEqual(ws.cell(3, 4).value, "Noturno Livre")
        self.assertEqual(ws.cell(3, 5).value, "Noturno")

    def test_06_cenario_fallback_cep_e_pendencias(self):
        """Valida fallback de CEP para área e fila de CEPs pendentes."""
        csv_file = TEST_DATA_DIR / "06_teste_fallback_cep_e_pendencias.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active

        # Linha 3: Bairro vazio, mas CEP 15042015 -> CIDADE DA CRIANÇA na coluna 8 (H: ÁREA)
        area_resolvida = ws.cell(3, 8).value
        self.assertIn(area_resolvida, ("CIDADE DA CRIANÇA", "CIDADE CRIANÇA"))

        # Linha 4: CEP desconhecido 15099888 -> Área em branco
        area_desconhecida = ws.cell(4, 8).value
        self.assertIn(area_desconhecida, (None, ""))

        # Verifica se o CEP 15099888 foi registrado na tabela de pendentes
        pendentes = listar_ceps_pendentes(db_path=self.cep_db)
        ceps_pends = [p["cep_limpo"] for p in pendentes]
        self.assertIn("15099888", ceps_pends)

    def test_07_cenario_encodings(self):
        """Valida integridade de strings acentuadas em UTF-8."""
        csv_file = TEST_DATA_DIR / "07_teste_encodings_e_acentuacao.csv"
        processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        nome_empresa = ws.cell(3, 19).value  # Coluna 19 (S: NomeEmpresarial)
        self.assertIn("SÃO JOÃO", nome_empresa)
        self.assertIn("AÇOUGUE", nome_empresa)

    def test_08_cenario_volume_500(self):
        """Valida que lote de 500 registros processa rapidamente e sem corrupção."""
        csv_file = TEST_DATA_DIR / "08_teste_carga_volume_500_linhas.csv"
        total = processar_fluxo_seco(
            caminho_csv=csv_file,
            saida_xlsx=self.saida_xlsx,
            triagem_db_path=self.triagem_db,
            cep_db_path=self.cep_db,
            silencioso=True,
        )
        self.assertEqual(total, 500)

        wb = openpyxl.load_workbook(self.saida_xlsx)
        ws = wb.active
        self.assertEqual(ws.max_row, 502)  # Linha 1 + Cabeçalho (Linha 2) + 500 dados
        self.assertEqual(ws.max_column, 32)
        self.assertEqual(ws.freeze_panes, "A3")
        self.assertTrue(ws.auto_filter.ref.startswith("A2:AF502"))


if __name__ == "__main__":
    unittest.main()
