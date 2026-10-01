"""Testes para o algoritmo de herança hierárquica e regras de CNAE."""

from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from src.core.database import connect_db, initialize_schema_if_needed


class TestCnaeSeedAndMatching(unittest.TestCase):
    """Verifica herança de classes e grupos e curto-circuito de matching."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmp_dir.name) / "test_cnae.db"
        self.conn = connect_db(self.db_path)
        initialize_schema_if_needed(self.conn, is_cep=False)

    def tearDown(self) -> None:
        self.conn.close()
        self.tmp_dir.cleanup()

    def test_regras_cnae_banco_real(self) -> None:
        """Verifica se o banco real de triagem contém as regras herdadas esperadas."""
        from src.core.database import get_triagem_db

        conn_real = get_triagem_db()
        cur = conn_real.cursor()

        cur.execute("SELECT COUNT(*) FROM regras_cnae;")
        if cur.fetchone()[0] == 0:
            from database.seed_cnaes import carregar_e_herdar_cnaes
            carregar_e_herdar_cnaes()

        # 1. Subclasse 4723-7/00 (Bebidas) deve herdar Noturno = 1 da Classe 47.23-7
        cur.execute("SELECT noturno, restricao FROM regras_cnae WHERE cnae_7digitos = '4723700';")
        row_bebidas = cur.fetchone()
        self.assertIsNotNone(row_bebidas, "Subclasse de bebidas (4723700) deve existir no banco.")
        self.assertEqual(row_bebidas["noturno"], 1, "4723700 deve ter Noturno=1 via herança hierárquica.")

        # 2. Subclasse 5611-2/01 (Restaurantes) deve herdar Noturno = 1 do Grupo 56.1 / Classe 56.11-2
        cur.execute("SELECT noturno FROM regras_cnae WHERE cnae_7digitos = '5611201';")
        row_restaurante = cur.fetchone()
        self.assertIsNotNone(row_restaurante)
        self.assertEqual(row_restaurante["noturno"], 1, "5611201 deve ter Noturno=1 via herança do grupo 56.1.")

        # 3. Subclasse 4731-8/00 (Combustíveis) deve ter Restrição = 1
        cur.execute("SELECT restricao FROM regras_cnae WHERE cnae_7digitos = '4731800';")
        row_combustivel = cur.fetchone()
        self.assertIsNotNone(row_combustivel)
        self.assertEqual(row_combustivel["restricao"], 1)

        conn_real.close()

    def test_curto_circuito_matching(self) -> None:
        """Testa o comportamento de parada rápida quando ambas as flags são 1."""
        cur = self.conn.cursor()
        # CNAE 1: Apenas restrição
        cur.execute("INSERT INTO regras_cnae VALUES ('1111111', '1111-1/11', 'Teste 1', 1, 0);")
        # CNAE 2: Apenas noturno
        cur.execute("INSERT INTO regras_cnae VALUES ('2222222', '2222-2/22', 'Teste 2', 0, 1);")
        # CNAE 3: Não deve ser consultado
        cur.execute("INSERT INTO regras_cnae VALUES ('3333333', '3333-3/33', 'Teste 3', 0, 0);")

        cnaes = ["1111111", "2222222", "3333333"]

        consultas_realizadas = 0
        flag_r = 0
        flag_n = 0

        for c in cnaes:
            consultas_realizadas += 1
            cur.execute("SELECT restricao, noturno FROM regras_cnae WHERE cnae_7digitos = ?;", (c,))
            r = cur.fetchone()
            if r:
                if r["restricao"]:
                    flag_r = 1
                if r["noturno"]:
                    flag_n = 1
            if flag_r == 1 and flag_n == 1:
                break

        self.assertEqual(flag_r, 1)
        self.assertEqual(flag_n, 1)
        self.assertEqual(consultas_realizadas, 2, "Curto-circuito deve parar no segundo CNAE sem consultar o terceiro.")


if __name__ == "__main__":
    unittest.main()
