"""Módulo Dedicado para Mapeamento Dinâmico de CEP x Área.

Gerencia o banco de dados satélite `database/cep_areas.db` e implementa o pipeline
geográfico de 3 níveis da Inspetoria Fiscal da SMAURB:
1. Nível 1 - Bairro Normalizado: Busca casamento exato do bairro na tabela `bairros_areas` do banco central.
2. Nível 2 - CEP Mapeado: Fallback para consulta direta do CEP no banco `cep_areas.db`.
3. Nível 3 - Fila de Pendências: Registro do CEP inconclusivo em `cep_pendentes` com contagem
   de ocorrências e atualização retroativa automática (`backfill`) assim que o operador cadastra a área.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any

from src.core.database import db_transaction, get_cep_db, get_triagem_db, initialize_schema_if_needed
from src.core.normalizer import formatar_cep, normalizar_bairro, normalizar_cep

# Dicionário canônico de harmonização de grafias de áreas fiscais
HARMONIZAR_AREAS: dict[str, str] = {
    "CIDADE CRIANÇA": "CIDADE DA CRIANÇA",
    "SCHIMITT": "SCHMITT",
}


def ensure_cep_tables(conn: sqlite3.Connection) -> None:
    """Garante que as tabelas cep_mapeamento e cep_pendentes existam no banco satélite.

    Aplica o script DDL `schema_cep.sql` caso as tabelas ainda não tenham sido criadas.

    Args:
        conn: Conexão ativa com o banco SQLite `cep_areas.db`.

    Complexity:
        Tempo: O(1) - verificação de tabelas existentes no sqlite_master.
        Espaço: O(1).
    """
    initialize_schema_if_needed(conn, is_cep=True)


def harmonizar_nome_area(area: str | None) -> str:
    """Padroniza e harmoniza grafias de áreas fiscais para nomes oficiais institucionais.

    Corrige desvios fonéticos ou variações históricas (ex: 'CIDADE CRIANÇA' -> 'CIDADE DA CRIANÇA',
    'SCHIMITT' -> 'SCHMITT') mantendo integridade com as áreas de atuação dos fiscais.

    Args:
        area: Nome bruto ou formatado da Área Fiscal.

    Returns:
        str: Nome padronizado em caixa alta ou string vazia se nulo.

    Complexity:
        Tempo: O(1) - consulta em dicionário hash de 2 elementos.
        Espaço: O(1).
    """
    if not area:
        return ""
    a = str(area).strip().upper()
    return HARMONIZAR_AREAS.get(a, a)


def consultar_area_por_cep(
    cep: str | None,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> str | None:
    """Consulta a Área Fiscal cadastrada para o CEP informado.

    Realiza normalização prévia para 8 dígitos numéricos e executa busca indexada
    em `cep_mapeamento`. Se não for fornecida conexão, gerencia abertura e fechamento
    automático de forma segura.

    Args:
        cep: Código postal nos formatos '15050-090', '15050090' ou numérico.
        conn: Conexão SQLite opcional pré-existente.
        db_path: Caminho alternativo opcional para o banco de dados.

    Returns:
        str | None: Nome da Área cadastrada (ex: 'CENTRAL', 'HB') ou None se não encontrado/inválido.

    Complexity:
        Tempo: O(1) devido ao índice único B-Tree na coluna `cep_limpo`.
        Espaço: O(1).
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


def backfill_area_backlog_por_cep(
    cep_limpo: str,
    area: str,
    triagem_conn: sqlite3.Connection | None = None,
    triagem_db_path: Path | str | None = None,
) -> int:
    """Atualiza a Área no backlog para protocolos pendentes que possuam esse CEP e Área vazia.

    Quando um operador cadastra a área de um CEP novo (seja via interface gráfica ou sincronização),
    esta função varre o backlog central `vre_backlog` em busca de processos pendentes cujo endereço
    possua esse mesmo CEP e atualiza atomicamente a coluna `area`.

    Args:
        cep_limpo: CEP normalizado com exatamente 8 dígitos.
        area: Nome da Área Fiscal a ser atribuída.
        triagem_conn: Conexão ativa opcional com `triagem_vre.db`.
        triagem_db_path: Caminho alternativo para o banco central.

    Returns:
        int: Quantidade de protocolos que receberam a Área retroativamente.

    Complexity:
        Tempo: O(P) onde P é o número de protocolos pendentes sem área no backlog.
        Espaço: O(P) para a lista de protocolos atualizados em memória.
    """
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
        prots_para_atualizar: list[str] = []
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
            area_final = harmonizar_nome_area(area)
            with db_transaction(triagem_conn) as t_cur:
                t_cur.executemany(
                    "UPDATE vre_backlog SET area = ? WHERE protocolo_licenca = ?;",
                    [(area_final, p) for p in prots_para_atualizar],
                )
            return len(prots_para_atualizar)
        return 0
    finally:
        if close_conn:
            triagem_conn.close()


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
    """Insere ou atualiza uma correlação de CEP x Área Fiscal.

    Ações executadas atomicamente:
    1. Normaliza e formata o CEP (8 dígitos e máscara).
    2. Harmoniza o nome da área.
    3. Executa UPSERT na tabela `cep_mapeamento` (atualizando `area`, `origem` e `atualizado_em`).
    4. Remove o CEP da fila de `cep_pendentes` caso ele estivesse aguardando resolução.
    5. Executa backfill retroativo no backlog central para preencher protocolos pendentes com esse CEP.

    Args:
        cep: CEP a ser registrado.
        area: Nome da Área Fiscal correspondente.
        bairro: Bairro de referência opcional.
        origem: Identificador da fonte de cadastro ('MANUAL', 'IMPORTACAO', 'PLANILHA_ATIVA').
        conn: Conexão opcional com `cep_areas.db`.
        db_path: Caminho alternativo do banco de CEPs.
        triagem_conn: Conexão opcional com `triagem_vre.db` para backfill.
        triagem_db_path: Caminho alternativo do banco de triagem.

    Returns:
        bool: True em caso de sucesso.

    Raises:
        ValueError: Caso o CEP seja inválido (diferente de 8 dígitos) ou a Área seja vazia.

    Complexity:
        Tempo: O(1) de inserção no banco de CEPs + O(P) de backfill no backlog central.
        Espaço: O(1).
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


def registrar_cep_pendente(
    cep: str | None,
    bairro: str | None = None,
    logradouro: str | None = None,
    conn: sqlite3.Connection | None = None,
    db_path: Path | str | None = None,
) -> None:
    """Registra um CEP inconclusivo para posterior resolução administrativa.

    Caso o CEP não seja resolvido nem por Bairro nem por CEP, esta rotina é acionada:
    1. Verifica se o CEP já se encontra mapeado em `cep_mapeamento` (evitando duplicações).
    2. Se não estiver mapeado, insere ou incrementa o contador `ocorrencias` em `cep_pendentes`.
    3. Atualiza o carimbo de data/hora (`atualizado_em`) e dados de logradouro/bairro mais recentes.

    Args:
        cep: CEP não localizado.
        bairro: Bairro original constante no pedido.
        logradouro: Logradouro constante no pedido.
        conn: Conexão com `cep_areas.db`.
        db_path: Caminho alternativo do banco.

    Complexity:
        Tempo: O(1) de busca e upsert indexado.
        Espaço: O(1).
    """
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
    """Remove uma correlação de CEP mapeado do banco satélite.

    Args:
        cep: CEP a ser removido da tabela `cep_mapeamento`.
        conn: Conexão opcional com `cep_areas.db`.
        db_path: Caminho alternativo do banco.

    Returns:
        bool: True se o comando de exclusão foi executado com sucesso, False se CEP inválido.

    Complexity:
        Tempo: O(1) de deleção indexada.
        Espaço: O(1).
    """
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
    """Lista correlações de CEP mapeados com suporte a múltiplos filtros combinados.

    Permite consulta textual por dígito de CEP ou nome de bairro, combinada com
    filtros seletivos de Área Fiscal e Bairro.

    Args:
        filtro: Substring a buscar no CEP puro, CEP formatado ou bairro.
        area_filtro: Área Fiscal específica ('CENTRAL', 'HB', etc.) ou 'Todas'.
        bairro_filtro: Bairro específico ou 'Todos'.
        limit: Limite máximo de registros a retornar (padrão 1000).
        conn: Conexão SQLite opcional.
        db_path: Caminho alternativo do banco.

    Returns:
        list[dict[str, Any]]: Lista de dicionários representando as linhas de `cep_mapeamento`.

    Complexity:
        Tempo: O(N) onde N é o número de registros filtrados (limitado a 1000).
        Espaço: O(N).
    """
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
    """Retorna lista alfabética de todos os bairros distintos cadastrados em `cep_areas.db`.

    Utilizado para popular dropdowns e comboboxes da interface gráfica administrativa.

    Args:
        conn: Conexão opcional com `cep_areas.db`.
        db_path: Caminho alternativo do banco.

    Returns:
        list[str]: Lista ordenada com nomes de bairros em caixa alta.

    Complexity:
        Tempo: O(B log B) onde B é o número de bairros únicos em `cep_mapeamento`.
        Espaço: O(B).
    """
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
    """Retorna lista alfabética de todas as áreas fiscais distintas cadastradas em `cep_areas.db`.

    Args:
        conn: Conexão opcional com `cep_areas.db`.
        db_path: Caminho alternativo do banco.

    Returns:
        list[str]: Lista ordenada com nomes de áreas em caixa alta.

    Complexity:
        Tempo: O(A log A) onde A é o número de áreas cadastradas.
        Espaço: O(A).
    """
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
    """Lista CEPs pendentes ordenados decrescentemente pelo número de ocorrências.

    Prioriza a exibição de CEPs que mais aparecem nos arquivos do VRE para que o
    administrador resolva primeiro os pontos de maior impacto geográfico.

    Args:
        limit: Quantidade máxima de registros a retornar (padrão 500).
        conn: Conexão opcional com `cep_areas.db`.
        db_path: Caminho alternativo do banco.

    Returns:
        list[dict[str, Any]]: Lista de dicionários representando as linhas de `cep_pendentes`.

    Complexity:
        Tempo: O(P log P) de ordenação por ocorrências.
        Espaço: O(P).
    """
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
       Se encontrar, retorna (area_harmonizada, 'BAIRRO').
    2. Se inconclusivo ou não encontrado, recorre ao banco satélite `cep_areas.db`.
       Se encontrar, retorna (area_harmonizada, 'CEP').
    3. Se ambos falharem:
       Registra o CEP pendente em `cep_pendentes` (caso CEP seja válido).
       Retorna ("", "PENDENTE").

    Args:
        bairro_original: Nome do bairro extraído do cadastro do VRE.
        cep_original: CEP extraído do cadastro do VRE.
        triagem_conn: Conexão aberta com o banco central de triagem.
        cep_conn: Conexão aberta opcional com o banco satélite de CEPs.
        logradouro: Logradouro do pedido para registro auxiliar em caso de pendência.

    Returns:
        tuple[str, str]: Tupla contendo (nome_da_area, origem_da_resolucao).
                         Exemplos: ('CENTRAL', 'BAIRRO'), ('HB', 'CEP'), ('', 'PENDENTE').

    Complexity:
        Tempo: O(1) - duas buscas indexadas em B-Tree no SQLite.
        Espaço: O(1).
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
    """Varre protocolos pendentes com Área vazia e resolve via Bairro ou CEP.

    Executada periodicamente ou após importações para auto-completar áreas de
    protocolos pendentes cuja área foi descoberta após o cadastro original.
    Respeita a flag `AreaLimpaManualmente` no JSON para não reescrever decisões
    onde o usuário explicitamente apagou a área no Excel.

    Args:
        triagem_conn: Conexão ativa com o banco central de triagem.
        cep_conn: Conexão ativa com o banco de CEPs.

    Returns:
        int: Total de protocolos cuja Área foi reconciliada e preenchida.

    Complexity:
        Tempo: O(P) onde P é o número de protocolos pendentes sem área no backlog.
        Espaço: O(P) para a lista de tuplas de atualização.
    """
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
