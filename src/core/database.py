"""Módulo de Acesso e Gerenciamento do Banco de Dados SQLite.

Fornece conexões thread-safe configuradas com WAL mode, cache estendido
e pragmas de alta performance, além de context managers para transações atômicas.
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

import src.core.config as config


def connect_db(db_path: Path | str, pragmas: list[str] | None = None) -> sqlite3.Connection:
    """Abre uma conexão com o SQLite aplicando pragmas de performance."""
    config.ensure_directories()
    path_obj = Path(db_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(
        str(path_obj),
        timeout=30.0,
        isolation_level=None,  # Autocommit por padrão; transações manuais via BEGIN
    )
    conn.row_factory = sqlite3.Row

    # Executa os pragmas configurados
    cur = conn.cursor()
    active_pragmas = pragmas if pragmas is not None else config.SQLITE_PRAGMAS
    for pragma in active_pragmas:
        try:
            cur.execute(pragma)
        except sqlite3.DatabaseError:
            pass
    cur.close()

    return conn


def get_triagem_db(db_path: Path | str | None = None) -> sqlite3.Connection:
    """Retorna uma conexão ativa e otimizada para o banco central triagem_vre.db."""
    target_path = db_path if db_path is not None else config.TRIAGEM_DB_PATH
    return connect_db(target_path)


def get_cep_db(db_path: Path | str | None = None) -> sqlite3.Connection:
    """Retorna uma conexão ativa e otimizada para o banco satélite cep_areas.db."""
    target_path = db_path if db_path is not None else config.CEP_DB_PATH
    return connect_db(target_path)


@contextmanager
def db_transaction(conn: sqlite3.Connection) -> Generator[sqlite3.Cursor, None, None]:
    """Context manager para execução de transações atômicas com rollback automático."""
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE;")
    try:
        yield cursor
        cursor.execute("COMMIT;")
    except Exception:
        cursor.execute("ROLLBACK;")
        raise
    finally:
        cursor.close()


def initialize_schema_if_needed(conn: sqlite3.Connection, is_cep: bool = False) -> None:
    """Aplica o arquivo DDL de criação caso as tabelas ainda não existam."""
    schema_file = config.resolve_schema_path(is_cep=is_cep)
    if not schema_file.exists():
        raise FileNotFoundError(f"Arquivo de schema DDL não encontrado: {schema_file}")

    schema_sql = schema_file.read_text(encoding="utf-8")
    conn.executescript(schema_sql)
