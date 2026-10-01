"""Script de Carga e Herança Hierárquica da Tabela Normativa de CNAEs.

Lê a planilha oficial 'Planilhas dos CNAIs (essa é definitiva).xls' (BIFF8) via xlrd,
implementa o algoritmo de herança hierárquica (Divisão -> Grupo -> Classe -> Subclasse)
e persiste todas as subclasses fatiadas em 7 dígitos limpos na tabela `regras_cnae`.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Sequence

import xlrd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.config import BASE_DIR, CONFIGURACOES_DIR, TRIAGEM_DB_PATH
from src.core.database import db_transaction, get_triagem_db, initialize_schema_if_needed
from src.core.normalizer import formatar_cnae


def _parse_flag(valor: object) -> int:
    """Interpreta flags numéricas 1.0, 1 ou texto '1' como 1, e vazio/0 como 0."""
    if valor is None or valor == "":
        return 0
    try:
        f = float(str(valor).strip().replace(",", "."))
        return 1 if f > 0 else 0
    except (ValueError, TypeError):
        return 0


def localizar_arquivo_cnae(caminho_especifico: Path | str | None = None) -> Path:
    """Busca o arquivo normativo de CNAEs em locais padrão do projeto."""
    if caminho_especifico:
        p = Path(caminho_especifico)
        if p.exists():
            return p

    candidatos = [
        BASE_DIR / "database" / "dados_referencia" / "Planilhas dos CNAIs (essa é definitiva).xls",
        CONFIGURACOES_DIR / "database" / "dados_referencia" / "Planilhas dos CNAIs (essa é definitiva).xls",
        BASE_DIR / "database" / "Planilhas dos CNAIs (essa é definitiva).xls",
        PROJECT_ROOT / "database" / "dados_referencia" / "Planilhas dos CNAIs (essa é definitiva).xls",
        PROJECT_ROOT / "Planilhas dos CNAIs (essa é definitiva).xls",
    ]

    for cand in candidatos:
        if cand.exists():
            return cand

    # Busca recursiva por .xls com CNAE no nome
    for match in BASE_DIR.rglob("*.xls"):
        if "CNA" in match.name.upper():
            return match

    raise FileNotFoundError(
        "Arquivo de CNAEs (Planilhas dos CNAIs (essa é definitiva).xls) não encontrado."
    )


def carregar_e_herdar_cnaes(
    caminho_xls: Path | str | None = None,
    db_path: Path | str | None = None,
) -> int:
    """Executa o parser com herança hierárquica e alimenta regras_cnae."""
    arquivo_xls = localizar_arquivo_cnae(caminho_xls)
    conn = get_triagem_db(db_path)
    initialize_schema_if_needed(conn, is_cep=False)

    print(f"[*] Abrindo planilha de CNAEs: {arquivo_xls.name}")
    book = xlrd.open_workbook(str(arquivo_xls))
    sheet = book.sheet_by_index(0)

    # Estado da hierarquia ativa
    cur_div_restr = 0
    cur_div_noturno = 0

    cur_grp_restr = 0
    cur_grp_noturno = 0

    cur_cls_restr = 0
    cur_cls_noturno = 0

    registros_para_inserir: list[tuple[str, str, str, int, int]] = []

    for idx in range(sheet.nrows):
        row = sheet.row_values(idx)
        if len(row) < 9:
            continue

        col_sec = str(row[1]).strip()
        col_div = str(row[2]).strip()
        col_grp = str(row[3]).strip()
        col_cls = str(row[4]).strip()
        col_sub = str(row[5]).strip()
        col_denom = str(row[6]).strip()

        flag_restr = _parse_flag(row[7])
        flag_noturno = _parse_flag(row[8])

        # Ignora cabeçalhos repetidos de quebra de página
        if any(h in col_sec for h in ["Estrutura detalhada", "(continua", "código CNAE", "Seção"]):
            continue
        if not col_denom and not col_sub and not col_cls and not col_grp and not col_div and not col_sec:
            continue

        # 1. Nova Seção (Letra única A-U) -> Reseta ramos inferiores
        if len(col_sec) == 1 and col_sec.isalpha():
            cur_div_restr = 0
            cur_div_noturno = 0
            cur_grp_restr = 0
            cur_grp_noturno = 0
            cur_cls_restr = 0
            cur_cls_noturno = 0

        # 2. Nova Divisão (2 dígitos)
        if col_div:
            cur_div_restr = flag_restr
            cur_div_noturno = flag_noturno
            cur_grp_restr = 0
            cur_grp_noturno = 0
            cur_cls_restr = 0
            cur_cls_noturno = 0

        # 3. Novo Grupo (ex: 56.1 ou 01.1)
        if col_grp:
            cur_grp_restr = flag_restr | cur_div_restr
            cur_grp_noturno = flag_noturno | cur_div_noturno
            cur_cls_restr = 0
            cur_cls_noturno = 0

        # 4. Nova Classe (ex: 47.23-7 ou 56.11-2)
        if col_cls:
            cur_cls_restr = flag_restr | cur_grp_restr
            cur_cls_noturno = flag_noturno | cur_grp_noturno

        # 5. Subclasse (ex: 0111-3/01 ou 4723-7/00)
        if col_sub:
            cnae_limpo = re.sub(r"\D", "", col_sub)
            if len(cnae_limpo) == 7:
                # Aplica herança obrigatória de Restrição e Noturno
                final_restr = flag_restr | cur_cls_restr
                final_noturno = flag_noturno | cur_cls_noturno
                cnae_fmt = col_sub if "/" in col_sub else formatar_cnae(cnae_limpo)

                registros_para_inserir.append(
                    (cnae_limpo, cnae_fmt, col_denom, final_restr, final_noturno)
                )

    print(f"[*] Subclasses identificadas e herdadas: {len(registros_para_inserir)}")

    with db_transaction(conn) as cur:
        cur.executemany(
            """
            INSERT INTO regras_cnae (cnae_7digitos, cnae_formatado, denominacao, restricao, noturno)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(cnae_7digitos) DO UPDATE SET
                cnae_formatado = excluded.cnae_formatado,
                denominacao = excluded.denominacao,
                restricao = excluded.restricao,
                noturno = excluded.noturno;
            """,
            registros_para_inserir,
        )

    # Estatísticas
    total_restr = sum(1 for r in registros_para_inserir if r[3] == 1)
    total_noturno = sum(1 for r in registros_para_inserir if r[4] == 1)
    print("[✓] Carga de CNAEs concluída com sucesso no banco!")
    print(f"    - Total inserido: {len(registros_para_inserir)} subclasses")
    print(f"    - Subclasses com Restrição: {total_restr}")
    print(f"    - Subclasses com Noturno: {total_noturno}")

    conn.close()
    return len(registros_para_inserir)


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else None
    carregar_e_herdar_cnaes(caminho)
