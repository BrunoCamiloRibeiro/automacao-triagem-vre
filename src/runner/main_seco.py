"""Executável Seco Operacional da Triagem VRE (Single-Action Headless Runner).

Módulo principal de processamento em lote da Inspetoria Fiscal da SMAURB.
Projetado para execução instantânea, silenciosa e com zero dependências de interface gráfica:
1. Resolução do arquivo CSV de entrada:
   - Suporta argumento direto via `sys.argv[1]` (Drag & Drop de arquivo sobre o `.bat`).
   - Busca automática do `.csv` mais recente na raiz do projeto (`BASE_DIR`) ou em `entrada_csv/`.
2. Leitura resiliente com detecção automática de encodings (UTF-8 com/sem BOM, CP1252, Latin-1).
3. Cruzamento relacional com SQLite de alta velocidade:
   - Identificação do fiscal anterior histórico via CNPJ limpo de 14 dígitos.
   - Detecção de Restrição e Noturno com algoritmo de curto-circuito em cache de memória de CNAEs.
   - Atribuição geográfica de Área por Bairro com fallback para `cep_areas.db`.
4. Inserção protegida no banco (`INSERT OR IGNORE` no `protocolo_licenca`).
5. Recuperação dos protocolos de todas as competências ativas (não arquivadas).
6. Exportação atômica da `Planilha_Controle_VRE_Ativa.xlsx` e remoção do CSV processado.
"""

from __future__ import annotations

import csv
import datetime
import json
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
    PLANILHAS_ANTERIORES_DIR,
    PROCESSADOS_CSV_DIR,
    ensure_directories,
)
from src.core.database import (
    db_transaction,
    get_cep_db,
    get_triagem_db,
    initialize_schema_if_needed,
)
from src.core.excel_writer import (
    gerar_planilha_trabalho,
    limpar_indeferidos_e_normalizar_backlog,
    registrar_hash_planilha_gerada,
    sincronizar_planilha_ativa_com_banco,
)
from src.core.historico_importer import normalizar_data_atendimento
from src.core.normalizer import (
    extrair_cnaes,
    formatar_cnpj,
    normalizar_cnpj,
)


def extrair_mes_competencia(data_str: str | None) -> str:
    """Extrai a competência 'AAAA-MM' a partir de datas nos formatos 'DD/MM/AAAA' ou 'AAAA-MM-DD'.

    Se a data for nula ou inválida, adota o mês e ano correntes como fallback seguro.

    Args:
        data_str: String contendo a data informada no pedido ou arquivo.

    Returns:
        str: Competência no formato padrão 'AAAA-MM' (ex: '2026-06').

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if not data_str:
        return datetime.date.today().strftime("%Y-%m")

    d_limpa = str(data_str).strip()
    # Formato brasileiro DD/MM/AAAA
    if "/" in d_limpa:
        partes = d_limpa.split("/")
        if len(partes) == 3:
            dia, mes, ano = partes[0].zfill(2), partes[1].zfill(2), partes[2]
            if len(ano) == 4 and len(mes) == 2:
                return f"{ano}-{mes}"

    # Formato ISO AAAA-MM-DD
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
    """Localiza o arquivo CSV a ser processado com máxima conveniência operacional.

    Hierarquia de Busca:
    1. Argumento direto de linha de comando ou Drag & Drop sobre o script `.bat`.
    2. Arquivo `.csv` mais recente depositado diretamente na pasta raiz do projeto (`BASE_DIR`).
    3. Arquivo `.csv` mais recente na pasta configurada `ENTRADA_CSV_DIR`.
    4. Fallback na subpasta `entrada_csv` da raiz caso exista.

    Args:
        caminho_arg: Caminho explícito opcional fornecido por argumento CLI.
        pasta_entrada: Pasta opcional para busca de arquivos de entrada.

    Returns:
        Path | None: Caminho do arquivo CSV mais recente ou None se nenhum for encontrado.

    Raises:
        FileNotFoundError: Se um caminho foi explicitamente passado em `caminho_arg` mas não existe.

    Complexity:
        Tempo: O(F) onde F é a quantidade de arquivos .csv na pasta examinada.
        Espaço: O(F) para listagem transitória de arquivos.
    """
    ensure_directories()

    # 1. Argumento direto de linha de comando ou Drag & Drop sobre o .bat
    if caminho_arg:
        p = Path(caminho_arg).resolve()
        if p.exists() and p.is_file():
            return p
        raise FileNotFoundError(f"Arquivo CSV especificado não encontrado: {caminho_arg}")

    # 2. Prioridade Máxima: busca arquivo .csv diretamente na pasta raiz (ao lado dos executáveis)
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
    """Lê o arquivo CSV aplicando detecção em cascata de encodings e sanitização de cabeçalhos.

    Testa sequencialmente as codificações declaradas em `CSV_ENCODINGS`:
    `utf-8-sig` (com BOM), `utf-8`, `cp1252`, `iso-8859-1` e `latin-1`.
    Remove espaços em branco nos nomes de colunas e valores de cada registro.

    Args:
        caminho_csv: Caminho para o arquivo CSV no sistema de arquivos.

    Returns:
        list[dict[str, str]]: Lista de dicionários representando as linhas do CSV.

    Raises:
        ValueError: Se nenhuma das codificações suportadas conseguir decodificar o arquivo.

    Complexity:
        Tempo: O(K) onde K é o número total de caracteres do arquivo CSV.
        Espaço: O(R * C) onde R é o número de linhas e C o número de colunas.
    """
    conteudo = None

    for enc in CSV_ENCODINGS:
        try:
            with open(caminho_csv, "r", encoding=enc) as f:
                conteudo = f.read()
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
        d: dict[str, str] = {}
        for idx, col_nome in enumerate(cabecalhos):
            if col_nome:
                d[col_nome] = row[idx].strip() if idx < len(row) else ""
        registros.append(d)

    return registros


def limpar_csvs_processados() -> int:
    """Remove arquivos CSV antigos acumulados na pasta de processados ou de entrada.

    Garante que nenhum CSV legado permaneça no disco após a execução.

    Returns:
        int: Quantidade de arquivos CSV removidos.

    Complexity:
        Tempo: O(F) onde F é a quantidade de arquivos encontrados.
        Espaço: O(1).
    """
    total_removidos = 0
    if PROCESSADOS_CSV_DIR.exists():
        for f in PROCESSADOS_CSV_DIR.glob("*.csv"):
            if f.is_file() and "tests" not in f.parts and "test_data" not in f.parts:
                try:
                    f.unlink()
                    total_removidos += 1
                except Exception:
                    pass
    return total_removidos


def processar_fluxo_seco(
    caminho_csv: Path | str | None = None,
    saida_xlsx: Path | str | None = None,
    silencioso: bool = True,
    triagem_db_path: Path | str | None = None,
    cep_db_path: Path | str | None = None,
    template_path: Path | str | None = None,
    entrada_dir: Path | str | None = None,
) -> int:
    """Executa a triagem completa do VRE e gera a planilha de trabalho ativa.

    Pipeline Operacional Completo:
    1. Localização e resolução do arquivo CSV.
    2. Abertura das conexões com `triagem_vre.db` e `cep_areas.db`.
    3. Sincronização prévia de edições manuais feitas no Excel ativo existente para o banco.
    4. Processamento dos registros do novo CSV:
       - Normalização do CNPJ (14 dígitos).
       - Busca do fiscal anterior titular da empresa (histórico ou backlog).
       - Análise de CNAEs com curto-circuito de flags (`flag_restr`, `flag_noturno`).
       - Classificação de Tipo de Serviço (omissão estrita em indeferidos, 'Diurno Livre' se titular, etc.).
       - Resolução de Área via Bairro -> CEP -> Fila de pendências.
    5. Inserção protegida (`INSERT OR IGNORE`) no backlog central.
    6. Reconciliação automática de áreas e saneamento de indeferidos.
    7. Consulta de todos os protocolos dos meses ativos (não arquivados).
    8. Backup preventivo da planilha anterior em `Planilhas anteriores/`.
    9. Geração da nova `Planilha_Controle_VRE_Ativa.xlsx` e remoção do CSV processado.

    Args:
        caminho_csv: Caminho opcional do arquivo CSV. Se None, busca automaticamente.
        saida_xlsx: Caminho opcional da planilha de saída (padrão: `OUTPUT_XLSX_PATH`).
        silencioso: Se False, emite logs de progresso no console.
        triagem_db_path: Caminho alternativo do banco central.
        cep_db_path: Caminho alternativo do banco de CEPs.
        template_path: Caminho alternativo do template Excel.
        entrada_dir: Diretório de busca de CSVs de entrada.

    Returns:
        int: Quantidade total de linhas ativas inseridas na planilha Excel gerada.

    Complexity:
        Tempo: O(N_csv + N_backlog) de processamento linear.
        Espaço: O(N_backlog * Colunas) de memória para o Excel gerado.
    """
    ensure_directories()
    arquivo_csv: Path | None = None
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

    saida_final = Path(saida_xlsx) if saida_xlsx else OUTPUT_XLSX_PATH

    # 0. Sincroniza edições manuais feitas pelo usuário na planilha ativa existente ANTES de importar novo CSV
    sincronizar_planilha_ativa_com_banco(
        caminho_xlsx=saida_final,
        triagem_conn=conn_triagem,
        cep_conn=conn_cep,
    )

    total_importados = 0

    if arquivo_csv is not None and arquivo_csv.exists():
        linhas_csv = ler_linhas_csv(arquivo_csv)
        if not silencioso:
            print(f"[*] Registros encontrados no CSV: {len(linhas_csv)}")

        cur_triagem = conn_triagem.cursor()

        # Cache em memória de CNAEs para velocidade instantânea sem consultas repetitivas ao banco
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

                    # Curto-circuito: se ambas as flags atingiram 1, encerra a busca de imediato
                    if flag_restr == 1 and flag_noturno == 1:
                        break

            # 3. Sugestão de Tipo de Serviço (quando está indeferido, não preenche Tipo de Serviço)
            situacao_postura = item.get("SituacaoPostura", "").strip()
            if situacao_postura.upper() == "INDEFERIDO":
                tipo_servico = ""
            elif flag_noturno == 1:
                tipo_servico = "Noturno" if fiscal_historico else "Noturno Livre"
            elif fiscal_historico:
                tipo_servico = "Diurno Livre"
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

    # Limpa Tipo de Serviço e Restrição de todos os registros indeferidos (inclusive passados)
    # e garante 'Diurno Livre' quando tem dono e não é noturno
    limpar_indeferidos_e_normalizar_backlog(conn_triagem)

    # 6. Recuperação de todos os protocolos dos meses ativos (não arquivados)
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
        WHERE mes_competencia NOT IN (
            SELECT mes_competencia FROM controle_competencias WHERE arquivado = 1
        )
        ORDER BY mes_competencia ASC, data_importacao ASC, protocolo_sil ASC;
        """
    )
    linhas_backlog = [dict(r) for r in cur_triagem.fetchall()]

    # 7. Backup preventivo da planilha ativa anterior para /Planilhas anteriores
    if saida_final.exists() and saida_final.stat().st_size > 0:
        if "tests" not in saida_final.parts and "test_data" not in saida_final.parts:
            try:
                PLANILHAS_ANTERIORES_DIR.mkdir(parents=True, exist_ok=True)
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bkp_nome = f"Planilha_Controle_VRE_Ativa_anterior_{ts}.xlsx"
                shutil.copy2(str(saida_final), str(PLANILHAS_ANTERIORES_DIR / bkp_nome))
            except Exception:
                pass

    # 8. Geração da planilha de trabalho ativa e registro do hash oficial no banco
    gerar_planilha_trabalho(
        registros_backlog=linhas_backlog,
        caminho_saida=saida_final,
        caminho_template=template_path,
    )
    registrar_hash_planilha_gerada(saida_final, conn_triagem)

    # Remove o arquivo CSV após execução com sucesso (nunca remove arquivos de teste)
    if arquivo_csv is not None and arquivo_csv.exists():
        caminho_res = arquivo_csv.resolve()
        if "tests" not in caminho_res.parts and "test_data" not in caminho_res.parts:
            if (
                caminho_res.parent == BASE_DIR.resolve()
                or ENTRADA_CSV_DIR.resolve() in caminho_res.parents
                or caminho_res.parent == ENTRADA_CSV_DIR.resolve()
            ):
                try:
                    arquivo_csv.unlink()
                    if not silencioso:
                        print(f"    [✓] Arquivo '{arquivo_csv.name}' removido após a execução com sucesso.")
                except Exception:
                    pass

    # Garante a limpeza de quaisquer CSVs legados acumulados em processados/
    limpar_csvs_processados()

    conn_triagem.close()
    conn_cep.close()

    if not silencioso:
        print(f"[✓] Triagem concluída:")
        print(f"    - Novos protocolos importados: {total_importados}")
        print(f"    - Linhas ativas no Excel gerado: {len(linhas_backlog)}")
        print(f"    - Planilha salva em: {saida_final}")

    return len(linhas_backlog)


def main() -> None:
    """Ponto de entrada do executável seco (triagem_vre.exe).

    Captura argumentos de linha de comando ou arrastamento de arquivo (.csv)
    e executa o processamento com captura defensiva de exceções gravadas em log.

    Complexity:
        Tempo: O(C + R) onde C é a leitura/sanitização do CSV e R a sincronização/geração do Excel.
        Espaço: O(R) para armazenamento temporário das linhas em memória.
    """
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
