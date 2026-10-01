"""Executável Seco Operacional da Triagem VRE (Single-Action Headless Runner).

Fluxo de Execução Silenciosa:
1. Resolução do CSV de entrada via sys.argv[1] (arraste/drag-and-drop) ou busca automática
   do arquivo .csv mais recente na pasta `entrada_csv/`.
2. Leitura com suporte a ISO-8859-1 / Windows-1252 e delimitador ';'.
3. Cruzamento no SQLite:
   - Identificação do fiscal anterior via histórico do CNPJ limpo.
   - Detecção de Restrição e Noturno com curto-circuito de CNAEs.
   - Atribuição de Área por Bairro com fallback para o banco `cep_areas.db`.
4. Inserção protegida no backlog (`INSERT OR IGNORE` no `protocolo_licenca`).
5. Recuperação de todos os protocolos ativos (`status IN ('PENDENTE', 'DISTRIBUIDO')`).
6. Exportação rápida e atômica da `Planilha_Controle_VRE_Ativa.xlsx`.
7. Encerramento limpo com sys.exit(0).
"""

from __future__ import annotations

import csv
import datetime
import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any

# Permite importações a partir da raiz do repositório
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.cep_service import reconciliar_areas_backlog, resolver_area
from src.core.config import (
    BASE_DIR,
    CSV_DELIMITER,
    CSV_ENCODINGS,
    ENTRADA_CSV_DIR,
    OUTPUT_XLSX_PATH,
    PROCESSADOS_CSV_DIR,
    ensure_directories,
)
from src.core.database import (
    db_transaction,
    get_cep_db,
    get_triagem_db,
    initialize_schema_if_needed,
)
from src.core.excel_writer import gerar_planilha_trabalho
from src.core.normalizer import (
    extrair_cnaes,
    formatar_cnpj,
    normalizar_cnpj,
)


from src.core.historico_importer import (
    escanear_e_importar_planilhas_antigas,
    normalizar_data_atendimento,
)


def extrair_mes_competencia(data_str: str | None) -> str:
    """Extrai 'AAAA-MM' a partir de datas nos formatos 'DD/MM/AAAA' ou 'AAAA-MM-DD'."""
    if not data_str:
        return datetime.date.today().strftime("%Y-%m")

    d_limpa = str(data_str).strip()
    # Formato DD/MM/AAAA
    if "/" in d_limpa:
        partes = d_limpa.split("/")
        if len(partes) == 3:
            dia, mes, ano = partes[0].zfill(2), partes[1].zfill(2), partes[2]
            if len(ano) == 4 and len(mes) == 2:
                return f"{ano}-{mes}"

    # Formato AAAA-MM-DD
    if "-" in d_limpa:
        partes = d_limpa.split("-")
        if len(partes) >= 2:
            ano, mes = partes[0], partes[1].zfill(2)
            if len(ano) == 4 and len(mes) == 2:
                return f"{ano}-{mes}"

    return datetime.date.today().strftime("%Y-%m")


def localizar_arquivo_csv(
    caminho_arg: str | None = None,
    pasta_entrada: Path | str | None = None,
) -> Path | None:
    """Localiza o arquivo CSV a ser processado com máxima conveniência para o usuário."""
    ensure_directories()

    # 1. Argumento direto de linha de comando ou Drag & Drop sobre o .bat
    if caminho_arg:
        p = Path(caminho_arg).resolve()
        if p.exists() and p.is_file():
            return p
        raise FileNotFoundError(f"Arquivo CSV especificado não encontrado: {caminho_arg}")

    # 2. Prioridade Máxima: busca arquivo .csv diretamente na pasta raiz (ao lado dos .bats)
    if BASE_DIR.exists():
        arquivos_raiz = [f for f in BASE_DIR.glob("*.csv") if f.is_file()]
        if arquivos_raiz:
            return max(arquivos_raiz, key=lambda f: f.stat().st_mtime)

    # 3. Busca o arquivo .csv mais recente na pasta de entrada configurada
    dir_entrada = Path(pasta_entrada) if pasta_entrada else ENTRADA_CSV_DIR
    if dir_entrada.exists():
        arquivos_entrada = [f for f in dir_entrada.glob("*.csv") if f.is_file()]
        if arquivos_entrada:
            return max(arquivos_entrada, key=lambda f: f.stat().st_mtime)

    # 4. Fallback: subpasta entrada_csv na raiz se existir
    raiz_entrada = BASE_DIR / "entrada_csv"
    if raiz_entrada.exists():
        arquivos_raiz_ent = [f for f in raiz_entrada.glob("*.csv") if f.is_file()]
        if arquivos_raiz_ent:
            return max(arquivos_raiz_ent, key=lambda f: f.stat().st_mtime)

    return None


def ler_linhas_csv(caminho_csv: Path) -> list[dict[str, str]]:
    """Lê o arquivo CSV testando encodings comuns (Latin-1 / CP1252 / UTF-8)."""
    conteudo = None
    encoding_usado = None

    for enc in CSV_ENCODINGS:
        try:
            with open(caminho_csv, "r", encoding=enc) as f:
                conteudo = f.read()
                encoding_usado = enc
                break
        except (UnicodeDecodeError, OSError):
            continue

    if conteudo is None:
        raise ValueError(f"Não foi possível decodificar o arquivo CSV: {caminho_csv}")

    linhas = conteudo.splitlines()
    if not linhas:
        return []

    reader = csv.reader(linhas, delimiter=CSV_DELIMITER)
    try:
        cabecalhos_brutos = next(reader)
    except StopIteration:
        return []

    # Higieniza cabeçalhos removendo espaços e caracteres vazios
    cabecalhos = [c.strip() for c in cabecalhos_brutos]

    registros: list[dict[str, str]] = []
    for row in reader:
        if not row or not any(c.strip() for c in row):
            continue
        d = {}
        for idx, col_nome in enumerate(cabecalhos):
            if col_nome:
                d[col_nome] = row[idx].strip() if idx < len(row) else ""
        registros.append(d)

    return registros


def processar_fluxo_seco(
    caminho_csv: Path | str | None = None,
    saida_xlsx: Path | str | None = None,
    silencioso: bool = True,
    triagem_db_path: Path | str | None = None,
    cep_db_path: Path | str | None = None,
    template_path: Path | str | None = None,
    entrada_dir: Path | str | None = None,
) -> int:
    """Executa a triagem completa e gera a planilha de trabalho ativa."""
    ensure_directories()
    arquivo_csv = None
    if caminho_csv:
        p = Path(caminho_csv).resolve()
        if p.exists() and p.is_file():
            arquivo_csv = p
    elif caminho_csv is None:
        arquivo_csv = localizar_arquivo_csv(pasta_entrada=entrada_dir)

    if not silencioso:
        if arquivo_csv:
            print(f"[*] Localizado arquivo CSV de entrada: {arquivo_csv.name}")
        else:
            print("    [AVISO] Nenhum arquivo .csv de novos pedidos encontrado para processar.")
            print("    -> Para realizar a triagem:")
            print("       1. Coloque o arquivo .csv exportado do VRE diretamente nesta pasta;")
            print("       OU 2. Arraste e solte o arquivo .csv sobre o ícone 'Executar_Triagem.bat'.")
            print("    -> Atualizando a planilha de trabalho com os protocolos ativos existentes...")

    conn_triagem = get_triagem_db(triagem_db_path)
    initialize_schema_if_needed(conn_triagem, is_cep=False)

    conn_cep = get_cep_db(cep_db_path)
    initialize_schema_if_needed(conn_cep, is_cep=True)

    total_importados = 0

    if arquivo_csv is not None and arquivo_csv.exists():
        linhas_csv = ler_linhas_csv(arquivo_csv)
        if not silencioso:
            print(f"[*] Registros encontrados no CSV: {len(linhas_csv)}")

        cur_triagem = conn_triagem.cursor()

        # Cache em memória de CNAEs e Fiscais para velocidade instantânea
        regras_cnae_cache: dict[str, tuple[int, int]] = {}
        cur_triagem.execute("SELECT cnae_7digitos, restricao, noturno FROM regras_cnae;")
        for r in cur_triagem.fetchall():
            regras_cnae_cache[r["cnae_7digitos"]] = (r["restricao"], r["noturno"])

        novos_backlog: list[tuple[Any, ...]] = []
        data_hoje = datetime.date.today().isoformat()

        for item in linhas_csv:
            prot_licenca = item.get("ProtocoloLicenca", "").strip()
            if not prot_licenca:
                continue

            prot_sil = item.get("ProtocoloSil", "").strip()
            cnpj_original = item.get("Cnpj", "").strip()
            cnpj_limpo = normalizar_cnpj(cnpj_original)
            cnpj_fmt = formatar_cnpj(cnpj_limpo) if cnpj_limpo else cnpj_original
            razao_social = item.get("NomeEmpresarial", "").strip()
            data_csv = item.get("Data", "").strip()
            # DEVE-SE USAR A DATA DO VRE NÃO A COLOCADA:
            data_vre_norm = normalizar_data_atendimento(data_csv, data_hoje)
            mes_comp = extrair_mes_competencia(data_vre_norm)

            bairro_orig = item.get("Bairro", "").strip()
            cep_orig = item.get("Cep", "").strip()
            logradouro = item.get("Logradouro", "").strip()

            # 1. Busca fiscal anterior por CNPJ limpo (histórico definitivo + backlog ativo)
            fiscal_historico = ""
            if cnpj_limpo:
                cur_triagem.execute(
                    """
                    SELECT fiscal FROM historico_fiscais
                    WHERE cnpj_limpo = ?
                      AND fiscal IS NOT NULL
                      AND TRIM(fiscal) != ''
                      AND fiscal != 'Não Informado'
                    ORDER BY data_atendimento DESC, id DESC
                    LIMIT 1;
                    """,
                    (cnpj_limpo,),
                )
                row_fisc = cur_triagem.fetchone()
                if row_fisc and row_fisc["fiscal"]:
                    fiscal_historico = str(row_fisc["fiscal"]).strip()
                else:
                    # Se não encontrado no histórico permanente, consulta atribuições no backlog ativo
                    cur_triagem.execute(
                        """
                        SELECT fiscal_atribuido FROM vre_backlog
                        WHERE cnpj_limpo = ?
                          AND fiscal_atribuido IS NOT NULL
                          AND TRIM(fiscal_atribuido) != ''
                          AND fiscal_atribuido != 'Não Informado'
                        ORDER BY data_importacao DESC
                        LIMIT 1;
                        """,
                        (cnpj_limpo,),
                    )
                    row_back = cur_triagem.fetchone()
                    if row_back and row_back["fiscal_atribuido"]:
                        fiscal_historico = str(row_back["fiscal_atribuido"]).strip()

            # 2. Análise de CNAEs com curto-circuito
            cnaes_str = item.get("CnaesLicenciadas", "")
            lista_cnaes_7 = extrair_cnaes(cnaes_str)

            flag_restr = 0
            flag_noturno = 0

            for cnae_7 in lista_cnaes_7:
                if cnae_7 in regras_cnae_cache:
                    r_val, n_val = regras_cnae_cache[cnae_7]
                    if r_val:
                        flag_restr = 1
                    if n_val:
                        flag_noturno = 1

                    # Curto-circuito: se ambas as flags atingiram 1, encerra a busca
                    if flag_restr == 1 and flag_noturno == 1:
                        break

            # 3. Sugestão de Tipo de Serviço
            if flag_noturno == 1:
                tipo_servico = "Noturno Livre"
            elif flag_restr == 1:
                tipo_servico = "Com Restrição Livre"
            else:
                tipo_servico = "Sem Restrição Livre"

            # 4. Resolução da Área com fallback para cep_areas.db
            area_resolvida, _ = resolver_area(
                bairro_original=bairro_orig,
                cep_original=cep_orig,
                triagem_conn=conn_triagem,
                cep_conn=conn_cep,
                logradouro=logradouro,
            )

            # 5. Preservação de todos os dados extras em JSON com a data oficial do VRE
            item["Data"] = data_vre_norm
            extras_json = json.dumps(item, ensure_ascii=False)

            novos_backlog.append(
                (
                    prot_licenca,
                    prot_sil,
                    cnpj_limpo,
                    cnpj_fmt,
                    razao_social,
                    mes_comp,
                    data_vre_norm,
                    bairro_orig,
                    area_resolvida,
                    cnaes_str,
                    fiscal_historico,
                    flag_restr,
                    flag_noturno,
                    tipo_servico,
                    extras_json,
                )
            )

        # 4. Upsert protegido no backlog ativo
        with db_transaction(conn_triagem) as cur:
            cur.executemany(
                """
                INSERT OR IGNORE INTO vre_backlog (
                    protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
                    nome_empresarial, mes_competencia, data_importacao,
                    bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
                    flag_restricao, flag_noturno, tipo_servico, status, dados_extras_json
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDENTE', ?);
                """,
                novos_backlog,
            )

        # Garante que os meses presentes no CSV fiquem como ativos (arquivado = 0) na planilha
        meses_novos = {item[5] for item in novos_backlog if item[5]}
        if meses_novos:
            with db_transaction(conn_triagem) as cur:
                for m in meses_novos:
                    cur.execute(
                        """
                        INSERT INTO controle_competencias (mes_competencia, arquivado, atualizado_em)
                        VALUES (?, 0, CURRENT_TIMESTAMP)
                        ON CONFLICT(mes_competencia) DO UPDATE SET
                            arquivado = 0,
                            atualizado_em = CURRENT_TIMESTAMP;
                        """,
                        (m,),
                    )

        total_importados = len(novos_backlog)

    # 5. Reconciliação automática: resolve Área para protocolos pendentes cujos CEPs ou Bairros foram mapeados
    reconciliar_areas_backlog(triagem_conn=conn_triagem, cep_conn=conn_cep)

    # 6. Recuperação de todos os protocolos em aberto nos meses ativos (não arquivados)
    cur_triagem = conn_triagem.cursor()
    cur_triagem.execute(
        """
        CREATE TABLE IF NOT EXISTS controle_competencias (
            mes_competencia TEXT PRIMARY KEY,
            arquivado INTEGER DEFAULT 0,
            atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        """
    )
    cur_triagem.execute(
        """
        SELECT
            protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
            nome_empresarial, mes_competencia, data_importacao,
            bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
            flag_restricao, flag_noturno, tipo_servico, status,
            data_finalizacao, dados_extras_json
        FROM vre_backlog
        WHERE status IN ('PENDENTE', 'DISTRIBUIDO')
          AND mes_competencia NOT IN (
              SELECT mes_competencia FROM controle_competencias WHERE arquivado = 1
          )
        ORDER BY mes_competencia ASC, data_importacao ASC, protocolo_sil ASC;
        """
    )
    linhas_backlog = [dict(r) for r in cur_triagem.fetchall()]

    # 7. Backup preventivo da planilha ativa anterior para /Planilhas anteriores
    saida_final = Path(saida_xlsx) if saida_xlsx else OUTPUT_XLSX_PATH
    if saida_final.exists() and saida_final.stat().st_size > 0:
        if "tests" not in saida_final.parts and "test_data" not in saida_final.parts:
            try:
                from src.core.config import PLANILHAS_ANTERIORES_DIR
                PLANILHAS_ANTERIORES_DIR.mkdir(parents=True, exist_ok=True)
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bkp_nome = f"Planilha_Controle_VRE_Ativa_anterior_{ts}.xlsx"
                shutil.copy2(str(saida_final), str(PLANILHAS_ANTERIORES_DIR / bkp_nome))
            except Exception:
                pass

    # 8. Geração da planilha de trabalho ativa
    gerar_planilha_trabalho(
        registros_backlog=linhas_backlog,
        caminho_saida=saida_final,
        caminho_template=template_path,
    )

    # Move para processados apenas em produção (nunca move arquivos de teste)
    if arquivo_csv is not None and arquivo_csv.exists():
        caminho_res = arquivo_csv.resolve()
        if "tests" not in caminho_res.parts and "test_data" not in caminho_res.parts:
            if (
                caminho_res.parent == BASE_DIR.resolve()
                or ENTRADA_CSV_DIR.resolve() in caminho_res.parents
                or caminho_res.parent == ENTRADA_CSV_DIR.resolve()
            ):
                PROCESSADOS_CSV_DIR.mkdir(parents=True, exist_ok=True)
                timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_dest = f"{arquivo_csv.stem}_{timestamp_str}{arquivo_csv.suffix}"
                try:
                    shutil.move(str(arquivo_csv), str(PROCESSADOS_CSV_DIR / nome_dest))
                    if not silencioso:
                        print(f"    [✓] Arquivo '{arquivo_csv.name}' arquivado em Configurações/entrada_csv/processados/")
                except Exception:
                    pass

    conn_triagem.close()
    conn_cep.close()

    if not silencioso:
        print(f"[✓] Triagem concluída:")
        print(f"    - Novos protocolos importados: {total_importados}")
        print(f"    - Linhas ativas no Excel gerado: {len(linhas_backlog)}")
        print(f"    - Planilha salva em: {saida_final}")

    return len(linhas_backlog)


def main() -> None:
    """Ponto de entrada do executável seco (triagem_vre.exe)."""
    arg_csv = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else None
    try:
        processar_fluxo_seco(caminho_csv=arg_csv, silencioso=False)
    except Exception as e:
        # Grava erro em arquivo de log na pasta do executável
        log_file = BASE_DIR / "triagem_erro.log"
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now().isoformat()}] Erro no processamento: {e}\n")
        print(f"\n[ERRO FATAL] {e}")
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    # Se chamado com --verbose, exibe logs no terminal para depuração
    if "--verbose" in sys.argv:
        caminho = None
        for a in sys.argv[1:]:
            if not a.startswith("--"):
                caminho = a
                break
        processar_fluxo_seco(caminho_csv=caminho, silencioso=False)
        sys.exit(0)
    else:
        main()
