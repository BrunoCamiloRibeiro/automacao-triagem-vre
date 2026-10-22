"""Serviço de Negócio para Gestão de Arquivamento e Visibilidade de Competências.

Implementa a rotina institucional de Arquivamento por Mês de Competência:
- "Arquivado = Sim": Oculto da planilha de trabalho atual (Planilha_Controle_VRE_Ativa.xlsx).
- "Arquivado = Não" (Ativo): Visível na planilha de trabalho atual.
- Ao aplicar alterações:
  1. A planilha de trabalho que estava sendo usada é salva como backup em /Planilhas anteriores.
  2. Gera uma nova planilha ativa com todos os meses selecionados como ativos,
     juntamente com quaisquer novos registros recebidos via CSV.
  3. Os registros permanecem sempre íntegros e preservados no banco SQLite (vre_backlog).
- Pesquisa multi-critério de fiscais históricos por CNPJ ou Razão Social com proteção contra injeção de wildcards.
"""

from __future__ import annotations

import datetime
import shutil
import sqlite3
from pathlib import Path
from typing import Any

from src.core.config import (
    PLANILHAS_ANTERIORES_DIR,
    resolve_output_xlsx_path,
    resolve_template_path,
)
from src.core.database import db_transaction, get_triagem_db, initialize_schema_if_needed
from src.core.excel_writer import (
    gerar_planilha_trabalho,
    limpar_indeferidos_e_normalizar_backlog,
    registrar_hash_planilha_gerada,
    sincronizar_planilha_ativa_com_banco,
)
from src.core.normalizer import formatar_cnpj, normalizar_cnpj


def garantir_tabela_controle(conn: sqlite3.Connection) -> None:
    """Garante que a tabela controle_competencias e seus índices existam no banco central.

    Sincroniza automaticamente quaisquer competências identificadas em `vre_backlog`
    que ainda não tenham sido cadastradas no controle, definindo-as inicialmente
    como ativas (`arquivado = 0`).

    Args:
        conn: Conexão SQLite ativa com `triagem_vre.db`.

    Complexity:
        Tempo: O(C) onde C é o número de competências distintas no backlog.
        Espaço: O(1).
    """
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS controle_competencias (
            mes_competencia TEXT PRIMARY KEY,
            arquivado INTEGER DEFAULT 0,
            atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_comp_arquivado ON controle_competencias(arquivado);")
    # Sincroniza qualquer competência existente no backlog que ainda não conste no controle
    cur.execute(
        """
        INSERT OR IGNORE INTO controle_competencias (mes_competencia, arquivado)
        SELECT DISTINCT mes_competencia, 0
        FROM vre_backlog
        WHERE mes_competencia IS NOT NULL AND TRIM(mes_competencia) != '';
        """
    )


def sincronizar_competencias(conn: sqlite3.Connection | None = None) -> None:
    """Sincroniza os meses do backlog com a tabela controle_competencias.

    Função utilitária exposta para invocação manual ou pré-carga de interfaces.

    Args:
        conn: Conexão SQLite opcional. Se None, gerencia ciclo de vida próprio.

    Complexity:
        Tempo: O(C).
        Espaço: O(1).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True
    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
    finally:
        if close_conn:
            conn.close()


def obter_estatisticas_gerais(conn: sqlite3.Connection | None = None) -> dict[str, int]:
    """Retorna contadores consolidados do backlog, competências e histórico de fiscais.

    Calcula métricas agregadas para o dashboard da interface gráfica:
    - Total de protocolos no backlog.
    - Protocolos pendentes, distribuídos e finalizados.
    - Total de registros em `historico_fiscais`.
    - Meses cadastrados, ativos e arquivados.
    - Protocolos que constam na planilha de trabalho ativa atual.

    Args:
        conn: Conexão SQLite opcional.

    Returns:
        dict[str, int]: Dicionário com 9 contadores inteiros fundamentais.

    Complexity:
        Tempo: O(1) a O(N) com consultas otimizadas por índice.
        Espaço: O(1).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
        cur = conn.cursor()

        cur.execute("SELECT COUNT(*) FROM vre_backlog;")
        total_backlog = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM vre_backlog WHERE UPPER(COALESCE(status, 'PENDENTE')) = 'PENDENTE';")
        total_pendentes = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM vre_backlog WHERE UPPER(status) IN ('DISTRIBUIDO', 'DISTRIBUÍDO', 'EM ANDAMENTO');")
        total_distribuidos = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM vre_backlog WHERE UPPER(status) IN ('FINALIZADO', 'DEFERIDO', 'INDEFERIDO', 'CANCELADO', 'CASSADO');")
        total_finalizados = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM historico_fiscais;")
        total_historico = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM controle_competencias;")
        meses_cadastrados = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM controle_competencias WHERE arquivado = 0;")
        meses_ativos = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM controle_competencias WHERE arquivado = 1;")
        meses_arquivados = cur.fetchone()[0]

        cur.execute(
            """
            SELECT COUNT(*) FROM vre_backlog
            WHERE mes_competencia NOT IN (
                SELECT mes_competencia FROM controle_competencias WHERE arquivado = 1
            );
            """
        )
        protocolos_ativos_visiveis = cur.fetchone()[0]

        return {
            "total_backlog": total_backlog,
            "total_pendentes": total_pendentes,
            "total_distribuidos": total_distribuidos,
            "total_finalizados": total_finalizados,
            "total_historico": total_historico,
            "meses_cadastrados": meses_cadastrados,
            "meses_ativos": meses_ativos,
            "meses_arquivados": meses_arquivados,
            "protocolos_ativos_visiveis": protocolos_ativos_visiveis,
        }
    finally:
        if close_conn:
            conn.close()


def listar_competencias_status(
    conn: sqlite3.Connection | None = None,
) -> list[dict[str, Any]]:
    """Lista todos os meses de competência e sua condição de arquivamento (visibilidade).

    Retorna tabela agregada com totais, pendentes, distribuídos, finalizados
    e percentual de conclusão de cada lote mensal.

    Args:
        conn: Conexão SQLite opcional.

    Returns:
        list[dict[str, Any]]: Lista de competências ordenadas decrescentemente (mais recentes primeiro).

    Complexity:
        Tempo: O(C * log C) onde C é o número de meses cadastrados.
        Espaço: O(C).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
        cur = conn.cursor()

        cur.execute(
            """
            SELECT
                c.mes_competencia,
                c.arquivado,
                COUNT(b.protocolo_licenca) as total,
                SUM(CASE WHEN UPPER(COALESCE(b.status, 'PENDENTE')) = 'PENDENTE' THEN 1 ELSE 0 END) as pendentes,
                SUM(CASE WHEN UPPER(b.status) IN ('DISTRIBUIDO', 'DISTRIBUÍDO', 'EM ANDAMENTO') THEN 1 ELSE 0 END) as distribuidos,
                SUM(CASE WHEN UPPER(b.status) IN ('FINALIZADO', 'DEFERIDO', 'INDEFERIDO', 'CANCELADO', 'CASSADO') THEN 1 ELSE 0 END) as finalizados
            FROM controle_competencias c
            LEFT JOIN vre_backlog b ON b.mes_competencia = c.mes_competencia
            GROUP BY c.mes_competencia
            ORDER BY c.mes_competencia DESC;
            """
        )
        linhas: list[dict[str, Any]] = []
        for r in cur.fetchall():
            tot = r["total"] or 0
            fin = r["finalizados"] or 0
            pend = r["pendentes"] or 0
            dist = r["distribuidos"] or 0
            arq = bool(r["arquivado"])
            taxa_conclusao = (fin / tot * 100) if tot > 0 else 0.0

            status_desc = "📦 Arquivado (Oculto da Planilha)" if arq else "✅ Ativo (Visível na Planilha)"

            linhas.append(
                {
                    "mes_competencia": r["mes_competencia"],
                    "arquivado": arq,
                    "status_texto": status_desc,
                    "total": tot,
                    "finalizados": fin,
                    "pendentes": pend,
                    "distribuidos": dist,
                    "taxa_conclusao": round(taxa_conclusao, 1),
                    "pode_arquivar": True,
                }
            )
        return linhas
    finally:
        if close_conn:
            conn.close()


def definir_status_competencia(
    mes_competencia: str,
    arquivado: bool,
    conn: sqlite3.Connection | None = None,
) -> bool:
    """Define se uma competência deve estar arquivada (oculta) ou ativa (visível).

    Args:
        mes_competencia: Mês no formato 'AAAA-MM' (ex: '2026-05').
        arquivado: True para ocultar da planilha ativa, False para torná-lo visível.
        conn: Conexão SQLite opcional.

    Returns:
        bool: True se a alteração foi persistida com sucesso, False se parâmetro inválido.

    Complexity:
        Tempo: O(1) de upsert indexado.
        Espaço: O(1).
    """
    comp_limpa = str(mes_competencia).strip()
    if not comp_limpa:
        return False

    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
        val_arq = 1 if arquivado else 0
        with db_transaction(conn) as cur:
            cur.execute(
                """
                INSERT INTO controle_competencias (mes_competencia, arquivado, atualizado_em)
                VALUES (?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(mes_competencia) DO UPDATE SET
                    arquivado = excluded.arquivado,
                    atualizado_em = CURRENT_TIMESTAMP;
                """,
                (comp_limpa, val_arq),
            )
        return True
    finally:
        if close_conn:
            conn.close()


def alternar_status_competencia(
    mes_competencia: str,
    conn: sqlite3.Connection | None = None,
) -> bool:
    """Alterna o status de arquivamento (Ativo <-> Arquivado) de um mês de competência.

    Args:
        mes_competencia: Mês a alternar (ex: '2026-06').
        conn: Conexão SQLite opcional.

    Returns:
        bool: Novo estado de arquivamento (True = arquivado/oculto, False = ativo/visível).

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    comp_limpa = str(mes_competencia).strip()
    if not comp_limpa:
        return False

    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
        cur = conn.cursor()
        cur.execute("SELECT arquivado FROM controle_competencias WHERE mes_competencia = ?;", (comp_limpa,))
        row = cur.fetchone()
        novo_status = 1 if (not row or row["arquivado"] == 0) else 0

        with db_transaction(conn) as t_cur:
            t_cur.execute(
                """
                INSERT INTO controle_competencias (mes_competencia, arquivado, atualizado_em)
                VALUES (?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(mes_competencia) DO UPDATE SET
                    arquivado = excluded.arquivado,
                    atualizado_em = CURRENT_TIMESTAMP;
                """,
                (comp_limpa, novo_status),
            )
        return novo_status == 1
    finally:
        if close_conn:
            conn.close()


def arquivar_competencia(
    mes_competencia: str,
    forcar: bool = False,
    conn: sqlite3.Connection | None = None,
) -> tuple[int, str]:
    """Marca a competência como arquivada (oculta da planilha de trabalho atual).

    Ações adicionais de integridade:
    - Sincroniza todos os atendimentos com fiscal atribuído daquela competência
      diretamente para `historico_fiscais`, garantindo que esses fiscais continuem
      sendo reconhecidos automaticamente em futuros pedidos de mesmo CNPJ.

    Args:
        mes_competencia: Mês a arquivar.
        forcar: Flag de compatibilidade retroativa.
        conn: Conexão SQLite opcional.

    Returns:
        tuple[int, str]: Quantidade de protocolos arquivados e mensagem descritiva.

    Complexity:
        Tempo: O(P) onde P é o número de protocolos daquele mês.
        Espaço: O(1).
    """
    comp_limpa = str(mes_competencia).strip()
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM vre_backlog WHERE mes_competencia = ?;", (comp_limpa,))
        total_prot = cur.fetchone()[0]

        # Sincroniza atendimentos com fiscal para historico_fiscais
        with db_transaction(conn) as t_cur:
            t_cur.execute(
                """
                INSERT INTO historico_fiscais (cnpj_limpo, fiscal, data_atendimento, origem_registro)
                SELECT
                    cnpj_limpo,
                    TRIM(fiscal_atribuido),
                    COALESCE(data_finalizacao, data_importacao, DATE('now')),
                    ('Mes_' || mes_competencia)
                FROM vre_backlog
                WHERE mes_competencia = ?
                  AND cnpj_limpo != ''
                  AND fiscal_atribuido IS NOT NULL
                  AND TRIM(fiscal_atribuido) != ''
                  AND fiscal_atribuido != 'Não Informado'
                ON CONFLICT(cnpj_limpo, fiscal) DO UPDATE SET
                    data_atendimento = excluded.data_atendimento;
                """,
                (comp_limpa,),
            )

        definir_status_competencia(comp_limpa, arquivado=True, conn=conn)
        return total_prot, f"Mês {comp_limpa} arquivado com sucesso ({total_prot} protocolos arquivados da planilha ativa)."
    finally:
        if close_conn:
            conn.close()


def reativar_competencia(
    mes_competencia: str,
    conn: sqlite3.Connection | None = None,
) -> tuple[int, str]:
    """Reativa a competência (passa a constar na planilha de trabalho atual).

    Args:
        mes_competencia: Mês a reativar.
        conn: Conexão SQLite opcional.

    Returns:
        tuple[int, str]: Código de sucesso (1 ou 0) e mensagem descritiva.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    comp_limpa = str(mes_competencia).strip()
    ok = definir_status_competencia(comp_limpa, arquivado=False, conn=conn)
    if ok:
        return 1, f"Mês {comp_limpa} reativado com sucesso (seus registros constarão na planilha de trabalho ativa)."
    return 0, f"Falha ao reativar competência {comp_limpa}."


def fazer_backup_planilha_ativa(
    caminho_planilha: Path | str | None = None,
) -> Path | None:
    """Gera cópia de segurança datada da planilha de trabalho ativa em /Planilhas anteriores.

    Utiliza `copy2` para manter os carimbos de tempo originais. Em caso de bloqueio temporário
    (ex: planilha aberta pelo Excel do usuário), tenta cópia via streaming de bytes de baixo nível.

    Args:
        caminho_planilha: Caminho da planilha a copiar ou None para `OUTPUT_XLSX_PATH`.

    Returns:
        Path | None: Caminho do arquivo de backup criado ou None se origem inexistente/bloqueada.

    Complexity:
        Tempo: O(B) onde B é o tamanho do arquivo em bytes.
        Espaço: O(1).
    """
    p_origem = Path(caminho_planilha) if caminho_planilha else resolve_output_xlsx_path()
    if not p_origem.exists() or p_origem.stat().st_size == 0:
        return None

    PLANILHAS_ANTERIORES_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_backup = f"Planilha_Controle_VRE_Ativa_anterior_{timestamp}.xlsx"
    p_destino = PLANILHAS_ANTERIORES_DIR / nome_backup

    try:
        shutil.copy2(str(p_origem), str(p_destino))
        return p_destino
    except Exception:
        try:
            with open(p_origem, "rb") as f_in, open(p_destino, "wb") as f_out:
                shutil.copyfileobj(f_in, f_out)
            return p_destino
        except Exception:
            return None


def aplicar_reorganizacao_planilha(
    conn: sqlite3.Connection | None = None,
    saida_xlsx: Path | str | None = None,
    caminho_template: Path | str | None = None,
) -> dict[str, Any]:
    """Aplica as alterações de visibilidade e gera a nova planilha de trabalho ativa.

    Passos executados:
    1. Salva backup preventivo da planilha em uso na pasta `Planilhas anteriores/`.
    2. Sincroniza qualquer edição manual feita diretamente no Excel para o SQLite.
    3. Se houver novo CSV pendente, processa-o via runner seco.
    4. Consulta todos os protocolos dos meses ativos (não arquivados).
    5. Gera atomicamente a nova `Planilha_Controle_VRE_Ativa.xlsx` institucional.

    Args:
        conn: Conexão SQLite opcional.
        saida_xlsx: Caminho do arquivo de saída.
        caminho_template: Caminho do template Excel.

    Returns:
        dict[str, Any]: Sumário com contadores de linhas, meses e mensagem de status.

    Complexity:
        Tempo: O(N_ativos * Colunas) para escrita no Excel.
        Espaço: O(N_ativos * Colunas).
    """
    # 1. Backup da planilha em uso
    backup_path = fazer_backup_planilha_ativa(caminho_planilha=saida_xlsx)

    # 2. Se houver novo CSV pendente, processa-o via runner seco
    from src.runner.main_seco import localizar_arquivo_csv, processar_fluxo_seco

    csv_pendente = localizar_arquivo_csv()
    csv_nome = None
    close_conn = False

    if csv_pendente is not None and csv_pendente.exists():
        csv_nome = csv_pendente.name
        linhas_geradas = processar_fluxo_seco(
            caminho_csv=csv_pendente,
            saida_xlsx=saida_xlsx,
            template_path=caminho_template,
            silencioso=True,
        )
    else:
        if conn is None:
            conn = get_triagem_db()
            close_conn = True

        try:
            initialize_schema_if_needed(conn, is_cep=False)
            garantir_tabela_controle(conn)

            p_saida = Path(saida_xlsx) if saida_xlsx else resolve_output_xlsx_path()
            p_tpl = Path(caminho_template) if caminho_template else resolve_template_path()

            sincronizar_planilha_ativa_com_banco(
                caminho_xlsx=p_saida,
                triagem_conn=conn,
            )
            limpar_indeferidos_e_normalizar_backlog(conn)

            cur = conn.cursor()
            cur.execute(
                """
                SELECT
                    protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
                    nome_empresarial, mes_competencia, data_importacao,
                    bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
                    flag_restricao, flag_noturno, tipo_servico, status,
                    data_finalizacao, dados_extras_json
                FROM vre_backlog
                WHERE mes_competencia NOT IN (
                    SELECT mes_competencia FROM controle_competencias WHERE arquivado = 1
                )
                ORDER BY mes_competencia ASC, data_importacao ASC, protocolo_sil ASC;
                """
            )
            linhas_backlog = [dict(r) for r in cur.fetchall()]

            gerar_planilha_trabalho(
                registros_backlog=linhas_backlog,
                caminho_saida=p_saida,
                caminho_template=p_tpl,
            )
            registrar_hash_planilha_gerada(p_saida, conn)
            linhas_geradas = len(linhas_backlog)
        finally:
            if close_conn:
                conn.close()

    # Coleta meses ativos e arquivados para resumo
    lista_lotes = listar_competencias_status(conn=conn if not close_conn else None)
    meses_ativos = [l["mes_competencia"] for l in lista_lotes if not l["arquivado"]]
    meses_arquivados = [l["mes_competencia"] for l in lista_lotes if l["arquivado"]]

    msg = f"Planilha de trabalho atualizada com sucesso!\n"
    msg += f"• Protocolos na planilha ativa: {linhas_geradas}\n"
    msg += f"• Meses visíveis: {len(meses_ativos)} ({', '.join(meses_ativos) if meses_ativos else 'Nenhum'})\n"
    msg += f"• Meses arquivados: {len(meses_arquivados)}"
    if backup_path:
        msg += f"\n• Cópia da planilha anterior guardada em: Planilhas anteriores/{backup_path.name}"
    if csv_nome:
        msg += f"\n• Novos pedidos do CSV '{csv_nome}' incorporados com sucesso."

    return {
        "status": "SUCESSO",
        "backup_criado": str(backup_path) if backup_path else None,
        "linhas_geradas": linhas_geradas,
        "meses_ativos": meses_ativos,
        "meses_arquivados": meses_arquivados,
        "csv_processado": csv_nome,
        "mensagem": msg,
    }


def listar_fiscais_cadastrados(conn: sqlite3.Connection | None = None) -> list[str]:
    """Retorna lista alfabética e sem duplicatas de todos os fiscais registrados.

    Combina registros de `historico_fiscais` e `vre_backlog` via `UNION` SQL.

    Args:
        conn: Conexão SQLite opcional.

    Returns:
        list[str]: Lista alfabética de nomes de fiscais em caixa alta.

    Complexity:
        Tempo: O(F log F) onde F é a quantidade de fiscais distintos.
        Espaço: O(F).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT DISTINCT UPPER(TRIM(fiscal)) as f
            FROM historico_fiscais
            WHERE fiscal IS NOT NULL AND TRIM(fiscal) != '' AND fiscal != 'Não Informado'
            UNION
            SELECT DISTINCT UPPER(TRIM(fiscal_atribuido)) as f
            FROM vre_backlog
            WHERE fiscal_atribuido IS NOT NULL AND TRIM(fiscal_atribuido) != '' AND fiscal_atribuido != 'Não Informado'
            ORDER BY f ASC;
            """
        )
        return [r["f"] for r in cur.fetchall() if r["f"]]
    finally:
        if close_conn:
            conn.close()


def obter_ultimos_atendimentos(
    limit: int = 50,
    fiscal_filtro: str | None = None,
    conn: sqlite3.Connection | None = None,
) -> list[dict[str, Any]]:
    """Retorna os atendimentos mais recentes gravados no histórico e no backlog.

    Args:
        limit: Quantidade máxima de linhas a retornar (padrão 50).
        fiscal_filtro: Nome específico de fiscal para filtrar ou 'Todos'.
        conn: Conexão SQLite opcional.

    Returns:
        list[dict[str, Any]]: Lista de dicionários representando os atendimentos.

    Complexity:
        Tempo: O(limit).
        Espaço: O(limit).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        cur = conn.cursor()
        resultados: list[dict[str, Any]] = []

        query = """
            SELECT cnpj_limpo, fiscal, data_atendimento, origem_registro
            FROM historico_fiscais
            WHERE fiscal IS NOT NULL AND TRIM(fiscal) != '' AND fiscal != 'Não Informado'
        """
        params: list[Any] = []
        if fiscal_filtro and fiscal_filtro.strip() and not fiscal_filtro.startswith("Todos"):
            query += " AND UPPER(fiscal) = ?"
            params.append(fiscal_filtro.strip().upper())

        query += " ORDER BY data_atendimento DESC, id DESC LIMIT ?;"
        params.append(limit)

        cur.execute(query, tuple(params))
        for r in cur.fetchall():
            resultados.append(
                {
                    "cnpj_formatado": formatar_cnpj(r["cnpj_limpo"]),
                    "cnpj_limpo": r["cnpj_limpo"],
                    "fiscal": r["fiscal"],
                    "data": r["data_atendimento"] or "-",
                    "origem": r["origem_registro"] or "Histórico",
                    "status": "Histórico Definitivo",
                }
            )

        # Se ainda houver espaço no limit, busca também no backlog ativo
        if len(resultados) < limit:
            restante = limit - len(resultados)
            q_back = """
                SELECT cnpj_limpo, fiscal_atribuido, nome_empresarial, data_importacao, status
                FROM vre_backlog
                WHERE fiscal_atribuido IS NOT NULL AND TRIM(fiscal_atribuido) != '' AND fiscal_atribuido != 'Não Informado'
            """
            p_back: list[Any] = []
            if fiscal_filtro and fiscal_filtro.strip() and not fiscal_filtro.startswith("Todos"):
                q_back += " AND UPPER(fiscal_atribuido) = ?"
                p_back.append(fiscal_filtro.strip().upper())
            q_back += " ORDER BY data_importacao DESC LIMIT ?;"
            p_back.append(restante)

            cur.execute(q_back, tuple(p_back))
            for r in cur.fetchall():
                resultados.append(
                    {
                        "cnpj_formatado": formatar_cnpj(r["cnpj_limpo"]),
                        "cnpj_limpo": r["cnpj_limpo"],
                        "fiscal": r["fiscal_atribuido"],
                        "data": r["data_importacao"] or "-",
                        "origem": f"Backlog ({r['nome_empresarial'] or ''})".strip(),
                        "status": r["status"],
                    }
                )

        return resultados
    finally:
        if close_conn:
            conn.close()


def buscar_historico_fiscais(
    termo_busca: str = "",
    fiscal_filtro: str | None = None,
    limit: int = 100,
    conn: sqlite3.Connection | None = None,
) -> list[dict[str, Any]]:
    """Busca fiscais históricos associados a um CNPJ, Razão Social ou filtrados por Fiscal.

    Tratamento de Pesquisa:
    1. Se `termo_busca` for vazio, retorna os últimos atendimentos.
    2. Se `termo_busca` normalizar para um CNPJ de 14 dígitos, executa busca exata por CNPJ
       em `historico_fiscais` e no `vre_backlog`.
    3. Se for busca textual por Razão Social, aplica sanitização rigorosa de caracteres curinga
       do SQL LIKE (escapando `%` e `_` com `ESCAPE '\\'`) para evitar queries custosas ou erros de sintaxe.

    Args:
        termo_busca: CNPJ, Razão Social ou substring a buscar.
        fiscal_filtro: Nome específico do fiscal ou 'Todos'.
        limit: Limite máximo de resultados (padrão 100).
        conn: Conexão SQLite opcional.

    Returns:
        list[dict[str, Any]]: Lista de correspondências com CNPJ, Razão Social, Fiscal e Origem.

    Complexity:
        Tempo: O(1) indexado para CNPJ de 14 dígitos; O(N) para busca textual LIKE por Razão Social.
        Espaço: O(limit).
    """
    termo = str(termo_busca or "").strip()
    fisc = str(fiscal_filtro or "").strip()
    if fisc.startswith("Todos"):
        fisc = ""

    if not termo and not fisc:
        return obter_ultimos_atendimentos(limit=limit, conn=conn)

    if not termo and fisc:
        return obter_ultimos_atendimentos(limit=limit, fiscal_filtro=fisc, conn=conn)

    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        cur = conn.cursor()

        cnpj_limpo = normalizar_cnpj(termo)
        resultados: list[dict[str, Any]] = []

        if cnpj_limpo and len(cnpj_limpo) == 14:
            # Busca exata por CNPJ em historico_fiscais
            q_hist = """
                SELECT cnpj_limpo, fiscal, data_atendimento, origem_registro
                FROM historico_fiscais
                WHERE cnpj_limpo = ?
                  AND fiscal IS NOT NULL
                  AND TRIM(fiscal) != ''
                  AND fiscal != 'Não Informado'
            """
            p_hist: list[Any] = [cnpj_limpo]
            if fisc:
                q_hist += " AND UPPER(fiscal) = ?"
                p_hist.append(fisc.upper())
            q_hist += " ORDER BY data_atendimento DESC, id DESC LIMIT ?;"
            p_hist.append(limit)

            cur.execute(q_hist, tuple(p_hist))
            for r in cur.fetchall():
                resultados.append(
                    {
                        "cnpj_formatado": formatar_cnpj(r["cnpj_limpo"]),
                        "cnpj_limpo": r["cnpj_limpo"],
                        "fiscal": r["fiscal"],
                        "data": r["data_atendimento"] or "-",
                        "origem": r["origem_registro"] or "Histórico",
                        "status": "Histórico Definitivo",
                    }
                )

            # Também consulta se existe atendimento ou registro no backlog ativo
            q_back = """
                SELECT cnpj_limpo, fiscal_atribuido, nome_empresarial, data_importacao, status
                FROM vre_backlog
                WHERE cnpj_limpo = ?
            """
            p_back = [cnpj_limpo]
            if fisc:
                q_back += " AND UPPER(fiscal_atribuido) = ?"
                p_back.append(fisc.upper())
            q_back += " ORDER BY data_importacao DESC LIMIT ?;"
            p_back.append(limit)

            cur.execute(q_back, tuple(p_back))
            for r in cur.fetchall():
                fisc_text = (
                    r["fiscal_atribuido"].strip()
                    if (r["fiscal_atribuido"] and r["fiscal_atribuido"].strip())
                    else "Não atribuído (Pendente)"
                )
                origem_text = f"Backlog ({r['nome_empresarial'] or ''})".strip()
                resultados.append(
                    {
                        "cnpj_formatado": formatar_cnpj(r["cnpj_limpo"]),
                        "cnpj_limpo": r["cnpj_limpo"],
                        "fiscal": fisc_text,
                        "data": r["data_importacao"] or "-",
                        "origem": origem_text,
                        "status": r["status"],
                    }
                )
        else:
            # Busca textual por nome ou termo (com escape seguro de wildcards LIKE)
            termo_sanitizado = termo.replace("%", "\\%").replace("_", "\\_").upper()
            busca_like = f"%{termo_sanitizado}%"

            q_back = """
                SELECT b.cnpj_limpo, b.fiscal_atribuido, b.nome_empresarial, b.data_importacao, b.status
                FROM vre_backlog b
                WHERE (UPPER(b.nome_empresarial) LIKE ? ESCAPE '\\'
                   OR UPPER(COALESCE(b.fiscal_atribuido, '')) LIKE ? ESCAPE '\\')
            """
            p_back = [busca_like, busca_like]
            if fisc:
                q_back += " AND UPPER(b.fiscal_atribuido) = ?"
                p_back.append(fisc.upper())
            q_back += " LIMIT ?;"
            p_back.append(limit)

            cur.execute(q_back, tuple(p_back))
            for r in cur.fetchall():
                resultados.append(
                    {
                        "cnpj_formatado": formatar_cnpj(r["cnpj_limpo"]),
                        "cnpj_limpo": r["cnpj_limpo"],
                        "fiscal": r["fiscal_atribuido"] or "Não atribuído",
                        "data": r["data_importacao"] or "-",
                        "origem": f"Backlog: {r['nome_empresarial'] or ''}",
                        "status": r["status"],
                    }
                )

        return resultados
    finally:
        if close_conn:
            conn.close()


def normalizar_data_filtro(data_str: str) -> str:
    """Normaliza datas nos formatos DD/MM/AAAA ou AAAA-MM-DD para o padrão ISO AAAA-MM-DD.

    Args:
        data_str: String representando uma data (ex: '21/10/2026' ou '2026-10-21').

    Returns:
        str: Data harmonizada no formato 'AAAA-MM-DD' ou string vazia se nula/inválida.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if not data_str:
        return ""
    d = str(data_str).strip()
    if "/" in d:
        partes = d.split("/")
        if len(partes) == 3:
            if len(partes[0]) == 4:
                return f"{partes[0]}-{partes[1].zfill(2)}-{partes[2].zfill(2)}"
            elif len(partes[2]) == 4:
                return f"{partes[2]}-{partes[1].zfill(2)}-{partes[0].zfill(2)}"
    elif "-" in d:
        partes = d.split("-")
        if len(partes) == 3:
            if len(partes[0]) == 4:
                return f"{partes[0]}-{partes[1].zfill(2)}-{partes[2].zfill(2)}"
            elif len(partes[2]) == 4:
                return f"{partes[2]}-{partes[1].zfill(2)}-{partes[0].zfill(2)}"
    return d


def listar_dias_importados(conn: sqlite3.Connection | None = None) -> list[dict[str, Any]]:
    """Lista todas as datas de importação registradas no backlog com seus totais consolidados.

    Args:
        conn: Conexão SQLite opcional. Se None, gerencia ciclo de vida próprio.

    Returns:
        list[dict[str, Any]]: Lista ordenada por data decrescente contendo:
            - data_importacao: Data em formato ISO 'AAAA-MM-DD'
            - data_formatada: Data em formato brasileiro 'DD/MM/AAAA'
            - total: Total de protocolos da data
            - pendentes: Total de protocolos com status 'PENDENTE'
            - distribuidos: Total de protocolos com status 'DISTRIBUIDO'
            - finalizados: Total de protocolos com status 'FINALIZADO'
            - competencias: Mês(es) de competência associados

    Complexity:
        Tempo: O(D) onde D é o número de dias distintos no backlog.
        Espaço: O(D).
    """
    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                data_importacao,
                COUNT(*) as total,
                SUM(CASE WHEN status = 'PENDENTE' THEN 1 ELSE 0 END) as pendentes,
                SUM(CASE WHEN status = 'DISTRIBUIDO' THEN 1 ELSE 0 END) as distribuidos,
                SUM(CASE WHEN status = 'FINALIZADO' THEN 1 ELSE 0 END) as finalizados,
                GROUP_CONCAT(DISTINCT mes_competencia) as competencias
            FROM vre_backlog
            WHERE data_importacao IS NOT NULL AND TRIM(data_importacao) != ''
            GROUP BY data_importacao
            ORDER BY data_importacao DESC;
            """
        )
        resultados: list[dict[str, Any]] = []
        for r in cur.fetchall():
            d_raw = str(r["data_importacao"]).strip()
            # Converte para formato brasileiro se for ISO
            if len(d_raw) == 10 and d_raw[4] == "-" and d_raw[7] == "-":
                d_fmt = f"{d_raw[8:10]}/{d_raw[5:7]}/{d_raw[:4]}"
            else:
                d_fmt = d_raw

            resultados.append(
                {
                    "data_importacao": d_raw,
                    "data_formatada": d_fmt,
                    "total": int(r["total"] or 0),
                    "pendentes": int(r["pendentes"] or 0),
                    "distribuidos": int(r["distribuidos"] or 0),
                    "finalizados": int(r["finalizados"] or 0),
                    "competencias": str(r["competencias"] or ""),
                }
            )
        return resultados
    finally:
        if close_conn:
            conn.close()


def apagar_dia_banco(
    data_alvo: str,
    conn: sqlite3.Connection | None = None,
    atualizar_planilha: bool = True,
    saida_xlsx: Path | str | None = None,
    caminho_template: Path | str | None = None,
) -> dict[str, Any]:
    """Exclui todos os protocolos de uma data específica do banco central (`triagem_vre.db`).

    Projetado especificamente para expurgar remessas diárias (arquivos CSV) que foram
    carregadas incorretamente, de forma duplicada ou com dados corrompidos.

    Fluxo Operacional Transacional:
    1. Harmonização da data informada para ISO 'AAAA-MM-DD' e formato nacional 'DD/MM/AAAA'.
    2. Sincronização prévia de edições manuais feitas no Excel ativo para garantir a preservação
       de eventuais alterações em outros dias antes do expurgo.
    3. Exclusão atômica de todos os protocolos daquela data em `vre_backlog`.
    4. Remoção de competências órfãs de `controle_competencias` que ficaram sem protocolos.
    5. Se `atualizar_planilha=True`:
       - Gera backup datado da planilha em uso em `Planilhas anteriores/`.
       - Reconstrói a `Planilha_Controle_VRE_Ativa.xlsx` refletindo a exclusão imediatamente.
       - Atualiza o registro de hash institucional.

    Args:
        data_alvo: Data a excluir (aceita 'AAAA-MM-DD' ou 'DD/MM/AAAA').
        conn: Conexão SQLite opcional. Se None, gerencia ciclo de vida próprio.
        atualizar_planilha: Se True, recria a planilha ativa no disco.
        saida_xlsx: Caminho da planilha de trabalho ativa.
        caminho_template: Caminho do template Excel.

    Returns:
        dict[str, Any]: Sumário contendo status, total de registros removidos, datas e mensagem.

    Complexity:
        Tempo: O(N_dia + N_ativos * Colunas) para exclusão e regeração da planilha ativa.
        Espaço: O(N_ativos * Colunas).
    """
    data_limpa = str(data_alvo or "").strip()
    data_iso = normalizar_data_filtro(data_limpa)
    if not data_iso:
        return {
            "sucesso": False,
            "removidos": 0,
            "data_alvo": data_limpa,
            "mensagem": "Data informada é inválida ou vazia.",
        }

    # Gera representações possíveis no banco para garantir casamento exato
    if len(data_iso) == 10 and data_iso[4] == "-" and data_iso[7] == "-":
        data_br = f"{data_iso[8:10]}/{data_iso[5:7]}/{data_iso[:4]}"
    else:
        data_br = data_iso

    candidatos = list({data_iso, data_br, data_limpa})
    placeholders = ", ".join(["?"] * len(candidatos))
    params = tuple(candidatos)

    close_conn = False
    if conn is None:
        conn = get_triagem_db()
        close_conn = True

    try:
        initialize_schema_if_needed(conn, is_cep=False)
        garantir_tabela_controle(conn)

        p_saida = Path(saida_xlsx) if saida_xlsx else resolve_output_xlsx_path()
        p_tpl = Path(caminho_template) if caminho_template else resolve_template_path()

        # 1. Sincroniza edições manuais existentes antes da exclusão
        if atualizar_planilha:
            sincronizar_planilha_ativa_com_banco(caminho_xlsx=p_saida, triagem_conn=conn)

        # 2. Verifica quantos registros casam com a data informada
        cur = conn.cursor()
        cur.execute(f"SELECT COUNT(*) FROM vre_backlog WHERE data_importacao IN ({placeholders});", params)
        total_remover = cur.fetchone()[0]

        if total_remover == 0:
            return {
                "sucesso": False,
                "removidos": 0,
                "data_iso": data_iso,
                "data_formatada": data_br,
                "mensagem": f"Nenhum protocolo encontrado no banco para a data '{data_br}'.",
            }

        # 3. Exclusão atômica em transação protegida
        with db_transaction(conn) as cur_del:
            cur_del.execute(f"DELETE FROM vre_backlog WHERE data_importacao IN ({placeholders});", params)
            # Remove competências órfãs que possam ter ficado com 0 protocolos
            cur_del.execute(
                """
                DELETE FROM controle_competencias
                WHERE mes_competencia NOT IN (
                    SELECT DISTINCT mes_competencia FROM vre_backlog
                    WHERE mes_competencia IS NOT NULL AND TRIM(mes_competencia) != ''
                );
                """
            )

        # 4. Atualização da planilha ativa em disco
        backup_path = None
        linhas_geradas = 0
        planilha_atualizada = False

        if atualizar_planilha:
            backup_path = fazer_backup_planilha_ativa(caminho_planilha=p_saida)
            limpar_indeferidos_e_normalizar_backlog(conn)

            cur_sel = conn.cursor()
            cur_sel.execute(
                """
                SELECT
                    protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
                    nome_empresarial, mes_competencia, data_importacao,
                    bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
                    flag_restricao, flag_noturno, tipo_servico, status,
                    data_finalizacao, dados_extras_json
                FROM vre_backlog
                WHERE mes_competencia NOT IN (
                    SELECT mes_competencia FROM controle_competencias WHERE arquivado = 1
                )
                ORDER BY mes_competencia ASC, data_importacao ASC, protocolo_sil ASC;
                """
            )
            linhas_backlog = [dict(r) for r in cur_sel.fetchall()]
            linhas_geradas = len(linhas_backlog)

            try:
                gerar_planilha_trabalho(
                    registros_backlog=linhas_backlog,
                    caminho_saida=p_saida,
                    caminho_template=p_tpl,
                )
                registrar_hash_planilha_gerada(p_saida, conn)
                planilha_atualizada = True
            except PermissionError:
                planilha_atualizada = False

        msg = f"Sucesso: {total_remover} protocolo(s) da data {data_br} foram excluídos do banco de dados."
        if atualizar_planilha:
            if planilha_atualizada:
                msg += f"\n• Planilha de trabalho ativa atualizada ({linhas_geradas} protocolos restantes)."
            else:
                msg += "\n• [AVISO] A planilha ativa está aberta no Excel e não pôde ser sobrescrita. Feche o Excel e clique em 'Aplicar Alterações'."
        if backup_path:
            msg += f"\n• Cópia de segurança salva em: Planilhas anteriores/{backup_path.name}"

        return {
            "sucesso": True,
            "removidos": total_remover,
            "data_iso": data_iso,
            "data_formatada": data_br,
            "linhas_planilha_ativa": linhas_geradas,
            "backup_criado": str(backup_path) if backup_path else None,
            "planilha_atualizada": planilha_atualizada,
            "mensagem": msg,
        }
    finally:
        if close_conn:
            conn.close()
