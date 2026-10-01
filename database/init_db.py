"""Script de Inicialização dos Bancos de Dados SQLite.

Cria e inicializa as estruturas DDL completas dos dois bancos:
1. `database/triagem_vre.db` (Central: Backlog, Histórico de Fiscais, CNAEs, Bairros)
2. `database/cep_areas.db` (Dedicado: Mapeamento Dinâmico de CEP x Área)

Opcionalmente executa o seeding inicial completo a partir dos arquivos fonte.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.config import (
    CEP_DB_PATH,
    TRIAGEM_DB_PATH,
    ensure_directories,
)
from src.core.database import get_cep_db, get_triagem_db, initialize_schema_if_needed


def inicializar_bancos(seed: bool = False) -> None:
    """Instancia ambos os bancos de dados aplicando DDL e pragmas de alta performance."""
    ensure_directories()
    print("[*] Inicializando ecossistema de bancos de dados SQLite da SMAURB...")

    # 1. triagem_vre.db
    print(f"[*] Aplicando schema em: {TRIAGEM_DB_PATH}")
    conn_triagem = get_triagem_db()
    initialize_schema_if_needed(conn_triagem, is_cep=False)
    conn_triagem.close()
    print("    [✓] triagem_vre.db pronto (WAL mode ativado).")

    # 2. cep_areas.db
    print(f"[*] Aplicando schema em: {CEP_DB_PATH}")
    conn_cep = get_cep_db()
    initialize_schema_if_needed(conn_cep, is_cep=True)
    conn_cep.close()
    print("    [✓] cep_areas.db pronto (WAL mode ativado).")

    # 3. templates/template_triagem.xlsx
    from src.core.excel_writer import criar_template_padrao
    criar_template_padrao()
    print("    [✓] template_triagem.xlsx validado.")

    if seed:
        print("\n[*] Iniciando procedimento de carga inicial (Seeding)...")
        try:
            from database.seed_cnaes import carregar_e_herdar_cnaes
            carregar_e_herdar_cnaes()
        except Exception as e:
            print(f"    [!] Erro ao semear CNAEs: {e}")

        try:
            from database.seed_bairros import carregar_bairros_areas
            carregar_bairros_areas()
        except Exception as e:
            print(f"    [!] Erro ao semear Bairros: {e}")

        try:
            from database.seed_historico import carregar_historico_e_ceps
            carregar_historico_e_ceps()
        except Exception as e:
            print(f"    [!] Erro ao semear Histórico e CEPs: {e}")

        print("[✓] Processo de Seeding concluído!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Inicializador de Bancos de Dados VRE")
    parser.add_argument(
        "--seed",
        action="store_true",
        help="Executa a carga automática de CNAEs, Bairros e Histórico",
    )
    args = parser.parse_args()
    inicializar_bancos(seed=args.seed)
