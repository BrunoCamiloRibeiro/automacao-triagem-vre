"""Módulo Dedicado para Mapeamento Dinâmico de CEP x Área.

Gerencia o banco satélite `database/cep_areas.db`, implementando:
1. Consulta de Área por CEP (`consultar_area_por_cep`).
2. Cadastro e atualização de correlações (`registrar_area_cep`).
3. Registro de CEPs não resolvidos (`registrar_cep_pendente`).
4. Pipeline de fallback: Bairro normalizado -> CEP mapeado -> Registro de pendência.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

from src.core.config import CEP_DB_PATH
from src.core.database import db_transaction, get_cep_db, initialize_schema_if_needed
from src.core.normalizer import formatar_cep, normalizar_bairro, normalizar_cep


def ensure_cep_tables(conn: sqlite3.Connection) -> None:
    """Garante que as tabelas de cep_mapeamento e cep_pendentes existam."""
    initialize_schema_if_needed(conn, is_cep=True)


def consultar_area_por_cep(
    cep: str | None,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> str | None:
    """Consulta a Área cadastrada para o CEP informado.

    Retorna o nome da Área (ex: 'CENTRAL', 'HB') ou None se não encontrado.
    """
    cep_limpo = normalizar_cep(cep)
    if not cep_limpo or len(cep_limpo) != 8:
        return None

    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT area FROM cep_mapeamento WHERE cep_limpo = ? LIMIT 1;",
            (cep_limpo,),
        )
        row = cursor.fetchone()
        if row and row["area"]:
            return str(row["area"]).strip()
        return None
    finally:
        if close_conn:
            conn.close()


HARMONIZAR_AREAS: dict[str, str] = {
    "CIDADE CRIANÇA": "CIDADE DA CRIANÇA",
    "SCHIMITT": "SCHMITT",
}


def harmonizar_nome_area(area: str | None) -> str:
    """Padroniza grafias de áreas (ex: 'CIDADE CRIANÇA' -> 'CIDADE DA CRIANÇA')."""
    if not area:
        return ""
    a = str(area).strip().upper()
    return HARMONIZAR_AREAS.get(a, a)


def registrar_area_cep(
    cep: str | None,
    area: str,
    bairro: str | None = None,
    origem: str = "MANUAL",
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
    triagem_conn: sqlite3.Connection | None = None,
    triagem_db_path: Path | str | None = None,
) -> bool:
    """Insere ou atualiza uma correlação de CEP x Área.

    Remove o CEP da lista de pendências e atualiza registros pendentes no backlog.
    """
    cep_limpo = normalizar_cep(cep)
    if not cep_limpo or len(cep_limpo) != 8:
        raise ValueError(f"CEP inválido para cadastro: {cep}")

    area_limpa = str(area).strip()
    if not area_limpa:
        raise ValueError("A Área não pode ser vazia.")

    area_harmonizada = harmonizar_nome_area(area_limpa)
    cep_fmt = formatar_cep(cep_limpo)
    bairro_ref = str(bairro).strip() if bairro else None

    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        with db_transaction(conn) as cur:
            cur.execute(
                """
                INSERT INTO cep_mapeamento (
                    cep_limpo, cep_formatado, area, bairro_referencia, origem, atualizado_em
                )
                VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(cep_limpo) DO UPDATE SET
                    area = excluded.area,
                    bairro_referencia = COALESCE(excluded.bairro_referencia, cep_mapeamento.bairro_referencia),
                    origem = excluded.origem,
                    atualizado_em = CURRENT_TIMESTAMP;
                """,
                (cep_limpo, cep_fmt, area_harmonizada, bairro_ref, origem),
            )
            # Remove da fila de pendentes se existir
            cur.execute("DELETE FROM cep_pendentes WHERE cep_limpo = ?;", (cep_limpo,))

        # Tenta atualizar protocolos pendentes no backlog central que tenham esse CEP
        try:
            backfill_area_backlog_por_cep(
                cep_limpo,
                area_harmonizada,
                triagem_conn=triagem_conn,
                triagem_db_path=triagem_db_path,
            )
        except Exception:
            pass

        return True
    finally:
        if close_conn:
            conn.close()


def backfill_area_backlog_por_cep(
    cep_limpo: str,
    area: str,
    triagem_conn: sqlite3.Connection | None = None,
    triagem_db_path: Path | str | None = None,
) -> int:
    """Atualiza a Área no backlog para registros pendentes que possuam esse CEP e Área vazia."""
    import json
    from src.core.database import db_transaction, get_triagem_db

    close_conn = False
    if triagem_conn is None:
        try:
            triagem_conn = get_triagem_db(triagem_db_path)
            close_conn = True
        except Exception:
            return 0

    try:
        cur = triagem_conn.cursor()
        cur.execute(
            "SELECT protocolo_licenca, dados_extras_json FROM vre_backlog WHERE status = 'PENDENTE' AND (area IS NULL OR area = '');"
        )
        pendentes = cur.fetchall()
        prots_para_atualizar = []
        for r in pendentes:
            extras_raw = r["dados_extras_json"]
            if extras_raw:
                try:
                    ex = json.loads(extras_raw)
                    cep_ex = normalizar_cep(ex.get("Cep"))
                    if cep_ex == cep_limpo:
                        prots_para_atualizar.append(r["protocolo_licenca"])
                except Exception:
                    pass

        if prots_para_atualizar:
            with db_transaction(triagem_conn) as t_cur:
                t_cur.executemany(
                    "UPDATE vre_backlog SET area = ? WHERE protocolo_licenca = ?;",
                    [(harmonizar_nome_area(area), p) for p in prots_para_atualizar],
                )
            return len(prots_para_atualizar)
        return 0
    finally:
        if close_conn:
            triagem_conn.close()


def registrar_cep_pendente(
    cep: str | None,
    bairro: str | None = None,
    logradouro: str | None = None,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> None:
    """Registra um CEP inconclusivo para posterior resolução administrativa."""
    cep_limpo = normalizar_cep(cep)
    if not cep_limpo or len(cep_limpo) != 8:
        return

    cep_fmt = formatar_cep(cep_limpo)
    bairro_ref = str(bairro).strip() if bairro else None
    logr_ref = str(logradouro).strip() if logradouro else None

    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        # Se o CEP já estiver mapeado em cep_mapeamento, não registra como pendente
        chk_cur = conn.cursor()
        chk_cur.execute("SELECT 1 FROM cep_mapeamento WHERE cep_limpo = ? LIMIT 1;", (cep_limpo,))
        if chk_cur.fetchone():
            return

        with db_transaction(conn) as cur:
            cur.execute(
                """
                INSERT INTO cep_pendentes (
                    cep_limpo, cep_formatado, bairro_referencia, logradouro, ocorrencias, criado_em, atualizado_em
                )
                VALUES (?, ?, ?, ?, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                ON CONFLICT(cep_limpo) DO UPDATE SET
                    ocorrencias = cep_pendentes.ocorrencias + 1,
                    bairro_referencia = COALESCE(excluded.bairro_referencia, cep_pendentes.bairro_referencia),
                    logradouro = COALESCE(excluded.logradouro, cep_pendentes.logradouro),
                    atualizado_em = CURRENT_TIMESTAMP;
                """,
                (cep_limpo, cep_fmt, bairro_ref, logr_ref),
            )
    finally:
        if close_conn:
            conn.close()


def remover_cep_mapeado(
    cep: str | None,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> bool:
    """Remove uma correlação de CEP mapeado."""
    cep_limpo = normalizar_cep(cep)
    if not cep_limpo:
        return False

    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        with db_transaction(conn) as cur:
            cur.execute("DELETE FROM cep_mapeamento WHERE cep_limpo = ?;", (cep_limpo,))
        return True
    finally:
        if close_conn:
            conn.close()


def listar_ceps_mapeados(
    filtro: str | None = None,
    area_filtro: str | None = None,
    bairro_filtro: str | None = None,
    limit: int = 1000,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> list[dict[str, Any]]:
    """Lista CEPs mapeados com suporte a múltiplos filtros (termo, área e bairro)."""
    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        cursor = conn.cursor()

        query = """
            SELECT cep_limpo, cep_formatado, area, bairro_referencia, origem, atualizado_em
            FROM cep_mapeamento
            WHERE 1=1
        """
        params: list[Any] = []

        if filtro and filtro.strip():
            termo = f"%{filtro.strip().upper()}%"
            query += " AND (UPPER(cep_limpo) LIKE ? OR UPPER(cep_formatado) LIKE ? OR UPPER(COALESCE(bairro_referencia, '')) LIKE ?)"
            params.extend([termo, termo, termo])

        if area_filtro and area_filtro.strip() and not area_filtro.startswith("Todas"):
            query += " AND UPPER(area) = ?"
            params.append(area_filtro.strip().upper())

        if bairro_filtro and bairro_filtro.strip() and not bairro_filtro.startswith("Todos"):
            query += " AND UPPER(bairro_referencia) = ?"
            params.append(bairro_filtro.strip().upper())

        query += " ORDER BY cep_limpo ASC LIMIT ?;"
        params.append(limit)

        cursor.execute(query, tuple(params))
        return [dict(row) for row in cursor.fetchall()]
    finally:
        if close_conn:
            conn.close()


def listar_bairros_mapeados(
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> list[str]:
    """Retorna lista alfabética de todos os bairros distintos cadastrados em cep_areas.db."""
    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT DISTINCT UPPER(TRIM(bairro_referencia)) as b
            FROM cep_mapeamento
            WHERE bairro_referencia IS NOT NULL AND TRIM(bairro_referencia) != ''
            ORDER BY b ASC;
            """
        )
        return [r["b"] for r in cursor.fetchall() if r["b"]]
    finally:
        if close_conn:
            conn.close()


def listar_areas_mapeadas(
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> list[str]:
    """Retorna lista alfabética de todas as áreas distintas cadastradas em cep_areas.db."""
    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT DISTINCT UPPER(TRIM(area)) as a
            FROM cep_mapeamento
            WHERE area IS NOT NULL AND TRIM(area) != ''
            ORDER BY a ASC;
            """
        )
        return [r["a"] for r in cursor.fetchall() if r["a"]]
    finally:
        if close_conn:
            conn.close()


def listar_ceps_pendentes(
    limit: int = 500,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> list[dict[str, Any]]:
    """Lista CEPs pendentes ordenados pelo número de ocorrências."""
    close_conn = False
    if conn is None:
        conn = get_cep_db(db_path)
        close_conn = True

    try:
        ensure_cep_tables(conn)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT cep_limpo, cep_formatado, bairro_referencia, logradouro, ocorrencias, criado_em, atualizado_em
            FROM cep_pendentes
            ORDER BY ocorrencias DESC, atualizado_em DESC
            LIMIT ?;
            """,
            (limit,),
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        if close_conn:
            conn.close()


def resolver_area(
    bairro_original: str | None,
    cep_original: str | None,
    triagem_conn: sqlite3.Connection,
    cep_conn: sqlite3.Connection | None = None,
    logradouro: str | None = None,
) -> tuple[str, str]:
    """Resolve a Área municipal aplicando o pipeline com fallback rigoroso.

    Ordem de resolução:
    1. Tenta casar o Bairro normalizado na tabela `bairros_areas` de `triagem_vre.db`.
       Se encontrar, retorna (area, 'BAIRRO').
    2. Se inconclusivo ou não encontrado, recorre ao banco `cep_areas.db`.
       Se encontrar, retorna (area, 'CEP').
    3. Se ambos falharem:
       Registra o CEP pendente em `cep_pendentes` (caso CEP seja válido).
       Retorna ("", "PENDENTE").
    """
    # 1. Tentativa via Bairro Normalizado
    bairro_norm = normalizar_bairro(bairro_original)
    if bairro_norm:
        cur = triagem_conn.cursor()
        cur.execute(
            "SELECT area FROM bairros_areas WHERE bairro_normalizado = ? LIMIT 1;",
            (bairro_norm,),
        )
        row = cur.fetchone()
        if row and row["area"]:
            return harmonizar_nome_area(str(row["area"])), "BAIRRO"

    # 2. Fallback via CEP Mapeado
    cep_limpo = normalizar_cep(cep_original)
    if cep_limpo and len(cep_limpo) == 8:
        area_cep = consultar_area_por_cep(cep_limpo, conn=cep_conn)
        if area_cep:
            return harmonizar_nome_area(area_cep), "CEP"

        # 3. Não resolvido por nenhum dos métodos: registrar na fila de pendentes
        registrar_cep_pendente(
            cep=cep_limpo,
            bairro=bairro_original,
            logradouro=logradouro,
            conn=cep_conn,
        )

    return "", "PENDENTE"


def reconciliar_areas_backlog(
    triagem_conn: sqlite3.Connection,
    cep_conn: sqlite3.Connection | None = None,
) -> int:
    """Varre protocolos pendentes com Área vazia e resolve via Bairro ou CEP."""
    import json
    from src.core.database import db_transaction

    cur = triagem_conn.cursor()
    cur.execute(
        """
        SELECT protocolo_licenca, bairro_original, dados_extras_json
        FROM vre_backlog
        WHERE status = 'PENDENTE' AND (area IS NULL OR area = '');
        """
    )
    rows = cur.fetchall()
    atualizados: list[tuple[str, str]] = []

    for r in rows:
        bairro = r["bairro_original"]
        extras_raw = r["dados_extras_json"]
        cep = ""
        logr = ""
        if extras_raw:
            try:
                ex = json.loads(extras_raw)
                if ex.get("AreaLimpaManualmente"):
                    continue
                cep = ex.get("Cep", "")
                logr = ex.get("Logradouro", "")
            except Exception:
                pass

        area_res, origem = resolver_area(
            bairro_original=bairro,
            cep_original=cep,
            triagem_conn=triagem_conn,
            cep_conn=cep_conn,
            logradouro=logr,
        )
        if area_res:
            atualizados.append((area_res, r["protocolo_licenca"]))

    if atualizados:
        with db_transaction(triagem_conn) as t_cur:
            t_cur.executemany(
                "UPDATE vre_backlog SET area = ? WHERE protocolo_licenca = ?;",
                atualizados,
            )

    return len(atualizados)

