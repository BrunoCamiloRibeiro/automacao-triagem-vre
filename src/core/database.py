"""Módulo de Acesso e Gerenciamento do Banco de Dados SQLite.

Fornece infraestrutura transacional robusta para os dois bancos de dados do ecossistema:
1. `triagem_vre.db`: Banco central de backlog de pedidos, regras de CNAE, bairros e histórico.
2. `cep_areas.db`: Banco satélite de geolocalização por CEP e log de CEPs pendentes.

Recursos de Alta Performance e Resiliência:
- Modo WAL (Write-Ahead Logging) habilitado por padrão em todas as conexões.
- Transações atômicas com `BEGIN IMMEDIATE` via context manager (`db_transaction`),
  prevenindo Deadlocks (SQLITE_BUSY) ao obter o write-lock logo na abertura da transação.
- Timeout estendido para 30.0 segundos para suportar concorrência sem falhas de I/O.
- Configuração de `row_factory = sqlite3.Row` para acesso por nome de coluna.
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

import src.core.config as config


def connect_db(db_path: Path | str, pragmas: list[str] | None = None) -> sqlite3.Connection:
    """Abre uma conexão otimizada com o banco de dados SQLite.

    Garante que os diretórios pais existam fisicamente e aplica os PRAGMAs de performance
    configurados (WAL mode, cache estendido de 64MB, synchronous NORMAL).

    Args:
        db_path: Caminho no sistema de arquivos para o banco de dados (.db).
        pragmas: Lista opcional de comandos SQL PRAGMA a serem executados na inicialização.
                 Caso seja None, utiliza `config.SQLITE_PRAGMAS`.

    Returns:
        sqlite3.Connection: Instância de conexão ativa com `row_factory = sqlite3.Row`
                            e modo autocommit habilitado (`isolation_level=None`).

    Raises:
        sqlite3.OperationalError: Caso o caminho seja inválido ou haja bloqueio persistente.

    Complexity:
        Tempo: O(1) - inicialização e execução de 4 pragmas.
        Espaço: O(1) - estrutura de controle da conexão SQLite.
    """
    config.ensure_directories()
    path_obj = Path(db_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    # isolation_level=None habilita o autocommit nativo do SQLite;
    # transações atômicas seguras são controladas manualmente via BEGIN IMMEDIATE.
    conn = sqlite3.connect(
        str(path_obj),
        timeout=30.0,
        isolation_level=None,
    )
    conn.row_factory = sqlite3.Row

    # Execução segura dos PRAGMAs com supressão defensiva de falhas não-críticas
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
    """Retorna uma conexão ativa e otimizada para o banco central triagem_vre.db.

    Args:
        db_path: Caminho alternativo opcional para o banco de dados (útil em testes unitários).
                 Se omitido, utiliza `config.TRIAGEM_DB_PATH`.

    Returns:
        sqlite3.Connection: Conexão configurada para o banco central.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    target_path = db_path if db_path is not None else config.TRIAGEM_DB_PATH
    return connect_db(target_path)


def get_cep_db(db_path: Path | str | None = None) -> sqlite3.Connection:
    """Retorna uma conexão ativa e otimizada para o banco satélite cep_areas.db.

    Args:
        db_path: Caminho alternativo opcional para o banco de dados (útil em testes unitários).
                 Se omitido, utiliza `config.CEP_DB_PATH`.

    Returns:
        sqlite3.Connection: Conexão configurada para o banco satélite de CEPs.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    target_path = db_path if db_path is not None else config.CEP_DB_PATH
    return connect_db(target_path)


@contextmanager
def db_transaction(conn: sqlite3.Connection) -> Generator[sqlite3.Cursor, None, None]:
    """Context manager para execução de transações atômicas com rollback automático.

    Implementa o padrão Unit of Work com proteção estrita:
    1. Abre a transação com `BEGIN IMMEDIATE;`, reservando imediatamente a trava de escrita
       e impedindo que outras threads/processos iniciem escritas conflitantes.
    2. Cede um cursor ativo para o bloco chamador (`yield cursor`).
    3. Em caso de sucesso, efetiva as alterações com `COMMIT;`.
    4. Se ocorrer qualquer exceção não tratada, desfaz todas as alterações com `ROLLBACK;`
       e propaga o erro original.
    5. Garante o fechamento do cursor no bloco `finally`.

    Args:
        conn: Conexão SQLite ativa sobre a qual a transação será executada.

    Yields:
        sqlite3.Cursor: Cursor aberto para execução dos comandos DML/DDL.

    Raises:
        Exception: Qualquer exceção levantada dentro do bloco transacional.

    Complexity:
        Tempo: O(1) de overhead transacional (BEGIN e COMMIT/ROLLBACK).
        Espaço: O(1).
    """
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
    """Aplica o arquivo DDL de criação caso as tabelas ainda não existam no banco.

    Executa de forma idempotente o script SQL de schema (`schema.sql` ou `schema_cep.sql`),
    garantindo que tabelas fundamentais (`vre_backlog`, `historico_fiscais`, `regras_cnae`,
    `bairros_areas`, `cep_mapeamento`, `cep_pendentes`) e índices estejam criados.

    Args:
        conn: Conexão SQLite ativa para o banco alvo.
        is_cep: Se True, aplica `schema_cep.sql`. Se False, aplica `schema.sql`.

    Raises:
        FileNotFoundError: Se o arquivo DDL correspondente não for localizado no disco.
        sqlite3.OperationalError: Se houver sintaxe SQL inválida no script DDL.

    Complexity:
        Tempo: O(N) onde N é o número de instruções DDL do arquivo de schema.
        Espaço: O(M) onde M é o tamanho do arquivo SQL lido na memória (~10KB).
    """
    schema_file = config.resolve_schema_path(is_cep=is_cep)
    if not schema_file.exists():
        raise FileNotFoundError(f"Arquivo de schema DDL não encontrado: {schema_file}")

    schema_sql = schema_file.read_text(encoding="utf-8")
    conn.executescript(schema_sql)
