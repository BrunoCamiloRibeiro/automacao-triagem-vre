"""Testes unitários para o módulo cep_service.py."""

from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from src.core.cep_service import (
    consultar_area_por_cep,
    ensure_cep_tables,
    listar_ceps_mapeados,
    listar_ceps_pendentes,
    registrar_area_cep,
    registrar_cep_pendente,
    resolver_area,
)
from src.core.database import connect_db, initialize_schema_if_needed


class TestCepService(unittest.TestCase):
    """Testa cadastro, consulta, fallback e controle de pendências no cep_areas.db."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.cep_db_path = Path(self.tmp_dir.name) / "test_cep.db"
        self.triagem_db_path = Path(self.tmp_dir.name) / "test_triagem.db"

        self.conn_cep = connect_db(self.cep_db_path)
        ensure_cep_tables(self.conn_cep)

        self.conn_triagem = connect_db(self.triagem_db_path)
        initialize_schema_if_needed(self.conn_triagem, is_cep=False)

    def tearDown(self) -> None:
        self.conn_cep.close()
        self.conn_triagem.close()
        self.tmp_dir.cleanup()

    def test_registrar_e_consultar_cep(self) -> None:
        # Consulta inicial não existente
        self.assertIsNone(consultar_area_por_cep("15050090", conn=self.conn_cep))

        # Cadastro com sucesso
        ok = registrar_area_cep(
            cep="15050090",
            area="REPRESA",
            bairro="Vila Anchieta",
            origem="MANUAL",
            conn=self.conn_cep,
        )
        self.assertTrue(ok)

        # Consulta com sucesso
        area = consultar_area_por_cep("15050-090", conn=self.conn_cep)
        self.assertEqual(area, "REPRESA")

        # Atualização de Área para o mesmo CEP
        registrar_area_cep(
            cep="15050090",
            area="CENTRAL",
            bairro="Vila Anchieta",
            origem="MANUAL",
            conn=self.conn_cep,
        )
        area_atualizada = consultar_area_por_cep("15050090", conn=self.conn_cep)
        self.assertEqual(area_atualizada, "CENTRAL")

    def test_registro_e_resolucao_pendente(self) -> None:
        # Registra pendência
        registrar_cep_pendente("15010000", bairro="Centro", logradouro="Rua Bernardino", conn=self.conn_cep)
        registrar_cep_pendente("15010000", bairro="Centro", conn=self.conn_cep)

        pendentes = listar_ceps_pendentes(conn=self.conn_cep)
        self.assertEqual(len(pendentes), 1)
        self.assertEqual(pendentes[0]["cep_limpo"], "15010000")
        self.assertEqual(pendentes[0]["ocorrencias"], 2)

        # Ao mapear o CEP pendente, ele deve sumir de cep_pendentes
        registrar_area_cep("15010000", "CENTRAL", conn=self.conn_cep)
        pendentes_apos = listar_ceps_pendentes(conn=self.conn_cep)
        self.assertEqual(len(pendentes_apos), 0)

    def test_pipeline_fallback_resolver_area(self) -> None:
        # Prepara tabela bairros_areas no triagem_db
        cur = self.conn_triagem.cursor()
        cur.execute("INSERT INTO bairros_areas (bairro_normalizado, area) VALUES ('VILA ANCHIETA', 'REPRESA');")

        # Caso 1: Bairro conhecido -> Resolve via BAIRRO
        area1, origem1 = resolver_area(
            bairro_original="Vila Anchieta",
            cep_original="15050090",
            triagem_conn=self.conn_triagem,
            cep_conn=self.conn_cep,
        )
        self.assertEqual(area1, "REPRESA")
        self.assertEqual(origem1, "BAIRRO")

        # Caso 2: Bairro desconhecido, mas CEP mapeado -> Resolve via CEP (Fallback)
        registrar_area_cep("15080000", "HB", conn=self.conn_cep)
        area2, origem2 = resolver_area(
            bairro_original="Bairro Desconhecido XYZ",
            cep_original="15080-000",
            triagem_conn=self.conn_triagem,
            cep_conn=self.conn_cep,
        )
        self.assertEqual(area2, "HB")
        self.assertEqual(origem2, "CEP")

        # Caso 3: Bairro desconhecido e CEP desconhecido -> Registra Pendência
        area3, origem3 = resolver_area(
            bairro_original="Outro Bairro Inexistente",
            cep_original="15099999",
            triagem_conn=self.conn_triagem,
            cep_conn=self.conn_cep,
        )
        self.assertEqual(area3, "")
        self.assertEqual(origem3, "PENDENTE")

        # Verifica se caiu em cep_pendentes
        pendentes = listar_ceps_pendentes(conn=self.conn_cep)
        self.assertTrue(any(p["cep_limpo"] == "15099999" for p in pendentes))


if __name__ == "__main__":
    unittest.main()
