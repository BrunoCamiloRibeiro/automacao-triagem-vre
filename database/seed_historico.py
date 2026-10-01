"""Script de Carga do Histórico Inicial de Fiscais e Mapeamentos Iniciais de CEP.

Suporta dois modos operacionais:
1. Modo Produção Interna: Lê a planilha de trabalho legada ('Planilha de Trabalho atual.xlsx')
   se ela estiver disponível no ambiente local.
2. Modo Demonstrativo / Open Source: Caso nenhuma planilha seja localizada, carrega
   automaticamente o dataset sintético de demonstração ('historico_demonstrativo.json'),
   garantindo inicialização e testes 100% funcionais sem expor dados reais de contribuintes.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.cep_service import registrar_area_cep
from src.core.config import BASE_DIR, CEP_DB_PATH, CONFIGURACOES_DIR, TRIAGEM_DB_PATH
from src.core.database import (
    db_transaction,
    get_cep_db,
    get_triagem_db,
    initialize_schema_if_needed,
)
from src.core.normalizer import normalizar_cep, normalizar_cnpj


def localizar_planilha_trabalho(caminho_especifico: Path | str | None = None) -> Path | None:
    """Localiza a planilha com histórico de trabalho se existir."""
    if caminho_especifico:
        p = Path(caminho_especifico)
        if p.exists() and p.is_file():
            return p

    candidatos = [
        BASE_DIR / "Planilha de Trabalho atual.xlsx",
        CONFIGURACOES_DIR / "Planilha de Trabalho atual.xlsx",
        PROJECT_ROOT / "Planilha de Trabalho atual.xlsx",
    ]

    for cand in candidatos:
        if cand.exists() and cand.is_file():
            return cand

    return None


def carregar_dados_demonstrativos(
    triagem_conn,
    cep_conn,
) -> tuple[int, int]:
    """Carrega dados demonstrativos a partir do arquivo JSON sintético."""
    json_path = BASE_DIR / "database" / "dados_referencia" / "historico_demonstrativo.json"
    if not json_path.exists():
        json_path = PROJECT_ROOT / "database" / "dados_referencia" / "historico_demonstrativo.json"

    if not json_path.exists():
        print("    [!] Arquivo demonstrativo não encontrado. Pulando carga de histórico.")
        return 0, 0

    print(f"[*] Carregando dataset sintético/demonstrativo de: {json_path.name}")
    dados = json.loads(json_path.read_text(encoding="utf-8"))

    fiscais = dados.get("fiscais", [])
    registros_fiscais = [
        (
            item["cnpj_limpo"],
            item["fiscal"],
            item["data_atendimento"],
            item.get("origem_registro", "DEMONSTRATIVO"),
        )
        for item in fiscais
    ]

    with db_transaction(triagem_conn) as cur:
        cur.executemany(
            """
            INSERT INTO historico_fiscais (cnpj_limpo, fiscal, data_atendimento, origem_registro)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(cnpj_limpo, fiscal) DO NOTHING;
            """,
            registros_fiscais,
        )

    ceps = dados.get("ceps", [])
    for c in ceps:
        registrar_area_cep(
            cep=c["cep_limpo"],
            area=c["area"],
            bairro=c.get("bairro"),
            origem=c.get("origem", "DEMONSTRATIVO"),
            conn=cep_conn,
        )

    print(f"[✓] Carga demonstrativa concluída:")
    print(f"    - Registros de fiscal por CNPJ adicionados: {len(registros_fiscais)}")
    print(f"    - Mapeamentos de CEP x Área cadastrados: {len(ceps)}")
    return len(registros_fiscais), len(ceps)


def carregar_historico_e_ceps(
    caminho_xlsx: Path | str | None = None,
    triagem_db_path: Path | str | None = None,
    cep_db_path: Path | str | None = None,
) -> tuple[int, int]:
    """Extrai histórico de fiscais para triagem_vre.db e CEPs para cep_areas.db."""
    conn_triagem = get_triagem_db(triagem_db_path)
    initialize_schema_if_needed(conn_triagem, is_cep=False)

    conn_cep = get_cep_db(cep_db_path)
    initialize_schema_if_needed(conn_cep, is_cep=True)

    arquivo_xlsx = localizar_planilha_trabalho(caminho_xlsx)

    if not arquivo_xlsx:
        print("[*] Nenhuma planilha de histórico local detectada. Utilizando modo demonstrativo...")
        fiscais, ceps = carregar_dados_demonstrativos(conn_triagem, conn_cep)
        conn_triagem.close()
        conn_cep.close()
        return fiscais, ceps

    print(f"[*] Abrindo planilha de histórico: {arquivo_xlsx.name}")
    wb = openpyxl.load_workbook(str(arquivo_xlsx), data_only=True)

    registros_fiscais: list[tuple[str, str, str, str]] = []
    mapa_ceps: dict[str, tuple[str, str]] = {}

    for sheet_name in ["Controle VRE", "Controle VRE 202X"]:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        print(f"[*] Processando aba: {sheet_name} ({ws.max_row} linhas)")

        for r in range(3, ws.max_row + 1):
            fiscal = ws.cell(row=r, column=3).value
            data_val = ws.cell(row=r, column=13).value or ws.cell(row=r, column=5).value
            area = ws.cell(row=r, column=7).value
            cnpj = ws.cell(row=r, column=17).value
            bairro = ws.cell(row=r, column=25).value
            cep = ws.cell(row=r, column=26).value

            data_str = str(data_val)[:10] if data_val else "2025-01-01"

            # 1. Histórico de Fiscal
            if fiscal and cnpj:
                fiscal_limpo = str(fiscal).strip()
                cnpj_limpo = normalizar_cnpj(str(cnpj))
                if len(cnpj_limpo) == 14 and fiscal_limpo:
                    registros_fiscais.append(
                        (cnpj_limpo, fiscal_limpo, data_str, f"Planilha_{sheet_name}")
                    )

            # 2. Correlação de CEP x Área
            if cep and area:
                cep_limpo = normalizar_cep(str(cep))
                area_limpa = str(area).strip().upper()
                if len(cep_limpo) == 8 and area_limpa:
                    bairro_str = str(bairro).strip() if bairro else ""
                    if cep_limpo not in mapa_ceps:
                        mapa_ceps[cep_limpo] = (area_limpa, bairro_str)

    # Inserção no banco de triagem (historico_fiscais)
    with db_transaction(conn_triagem) as cur:
        cur.executemany(
            """
            INSERT INTO historico_fiscais (cnpj_limpo, fiscal, data_atendimento, origem_registro)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(cnpj_limpo, fiscal) DO NOTHING;
            """,
            registros_fiscais,
        )
    fiscais_inseridos = len(registros_fiscais)

    # Inserção no banco de CEPs (cep_mapeamento)
    for cep_limpo, (area_limpa, bairro_str) in mapa_ceps.items():
        registrar_area_cep(
            cep=cep_limpo,
            area=area_limpa,
            bairro=bairro_str,
            origem="INFERIDO_PLANILHA",
            conn=conn_cep,
        )
    ceps_inseridos = len(mapa_ceps)

    print("[✓] Carga de histórico concluída:")
    print(f"    - Registros de fiscal por CNPJ adicionados: {fiscais_inseridos}")
    print(f"    - Mapeamentos iniciais de CEP x Área inferidos: {ceps_inseridos}")

    conn_triagem.close()
    conn_cep.close()
    return fiscais_inseridos, ceps_inseridos


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else None
    carregar_historico_e_ceps(caminho)
