"""Módulo de Gravação e Otimização de Planilhas Excel com OpenPyXL.

Responsável por gerar o arquivo oficial de produção `Planilha_Controle_VRE_Ativa.xlsx`
reproduzindo com fidelidade milimétrica o padrão institucional da Inspetoria Fiscal:
- Arquitetura de 32 colunas oficiais (A a AF), incluindo a coluna única de Restrição (E).
- Injeção de Plain Values (valores literais pré-computados em Python, eliminando fórmulas
  lentas de VLOOKUP/PROCV/SE no Excel que causavam lentidão e travamento).
- Sincronização bidirecional inteligente baseada em hash SHA-256 para capturar edições manuais
  do operador diretamente das células do Excel de volta para o SQLite.
- Aplicação das 15 regras de formatação condicional na exata ordem de prioridade visual.
- Congelamento de painel em A3 e ativação de auto-filtro na linha 2.
- Backups preventivos datados na pasta `/Planilhas anteriores`.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import shutil
import sqlite3
from copy import copy
from pathlib import Path
from typing import Any, Sequence

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.core.cep_service import harmonizar_nome_area, registrar_area_cep
from src.core.config import (
    BASE_DIR,
    COR_CABECALHO,
    COR_EM_ANDAMENTO,
    COR_FINALIZADO_STATUS,
    COR_FISCAL_ATRIBUIDO,
    COR_INDEFERIDO,
    COR_NOTURNO_ENTREGUE,
    COR_NOTURNO_LIVRE,
    COR_RESTRICAO_LIVRE,
    COR_SEM_RESTRICAO,
    OUTPUT_XLSX_PATH,
    PLANILHAS_ANTERIORES_DIR,
    TEMPLATE_XLSX_PATH,
    ensure_directories,
    resolve_template_path,
)
from src.core.database import db_transaction, get_cep_db, get_triagem_db, initialize_schema_if_needed
from src.core.normalizer import (
    formatar_cnpj,
    montar_endereco_completo,
    normalizar_cep,
    normalizar_cnpj,
)

# Cache em memória do último hash SHA-256 gerado pelo sistema para cada caminho de planilha
_ULTIMO_HASH_GERADO: dict[str, str] = {}

# Chave de migração única para limpar as datas anteriores preenchidas automaticamente na Coluna F (Data)
_MIGRACAO_LIMPEZA_COL_F_KEY: str = "limpeza_unica_data_col_f_v15"


def _migracao_limpeza_col_f_concluida(triagem_conn: sqlite3.Connection | None) -> bool:
    """Verifica se a limpeza única da coluna Data (F) já foi executada no banco central.

    A versão 1.5 introduziu a separação definitiva entre a Data Oficial do VRE (Coluna N)
    e a Coluna de Data Manual de Distribuição (Coluna F). Esta rotina garante que uma migração
    única limpe datas automáticas antigas da Coluna F, preservando qualquer edição manual futura.

    Args:
        triagem_conn: Conexão SQLite ativa com `triagem_vre.db`.

    Returns:
        bool: True se a migração já foi registrada na tabela `controle_migracoes`.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if triagem_conn is None:
        return False
    try:
        cur = triagem_conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='controle_migracoes';"
        )
        if not cur.fetchone():
            return False
        cur.execute(
            "SELECT 1 FROM controle_migracoes WHERE chave = ? LIMIT 1;",
            (_MIGRACAO_LIMPEZA_COL_F_KEY,),
        )
        return cur.fetchone() is not None
    except Exception:
        return False


# ==============================================================================
# DEFINIÇÃO OFICIAL DAS 32 COLUNAS INSTITUCIONAIS (ORDEM E LARGURA EM CARACTERES)
# ==============================================================================
# Colunas padronizadas em estrita harmonia com o modelo da Inspetoria Fiscal da SMAURB:
COLUNAS_CABECALHO: list[tuple[str, float]] = [
    ("Número ordem ", 17.43),      # Col 1 (A): Sequencial sequencial 1, 2, 3...
    ("Status", 15.57),             # Col 2 (B): Em Andamento, Deferido, Indeferido, etc. (vazio se pendente)
    ("Fiscal", 15.57),             # Col 3 (C): Nome do Fiscal responsável
    ("Tipo de Serviço", 18.29),    # Col 4 (D): Diurno Livre, Noturno, Noturno entregue, etc.
    ("Restrição", 16.0),           # Col 5 (E): Coluna Única: Sem Restrição, Com Restrição, Noturno
    ("Data", 15.57),               # Col 6 (F): Data manual do fiscal (vazia por padrão)
    ("IPTU", 12.43),               # Col 7 (G): Número de cadastro imobiliário
    ("ÁREA", 20.0),                # Col 8 (H): Área fiscal municipal (CENTRAL, HB, etc.)
    ("Endereço", 85.14),           # Col 9 (I): Endereço completo estático harmonizado
    ("Protocolo Existente", 55.0), # Col 10 (J): Número de protocolos existentes
    ("ProtocoloSigm / OS", 21.0),  # Col 11 (K): Ordem de serviço no sistema Sigm
    ("ProtocoloSil", 11.29),       # Col 12 (L): Protocolo SIL estadual
    ("ProtocoloLicenca", 16.57),   # Col 13 (M): Identificador único do pedido no VRE
    ("Data", 10.71),               # Col 14 (N): Data oficial de envio pelo contribuinte no VRE
    ("GrauRiscoGeral", 7.0),       # Col 15 (O): Grau de risco (BAIXO, MEDIO, ALTO)
    ("SituacaoPostura", 13.14),    # Col 16 (P): Situação na Postura (DEFERIDO, INDEFERIDO, etc.)
    ("InscricaoMunicipal", 21.0),  # Col 17 (Q): Inscrição municipal cadastrada
    ("Cnpj", 18.43),               # Col 18 (R): CNPJ formatado da empresa
    ("NomeEmpresarial", 48.0),     # Col 19 (S): Razão Social da empresa
    ("Telefone", 11.57),           # Col 20 (T): Telefone de contato
    ("E-mail", 17.43),             # Col 21 (U): E-mail do solicitante/contador
    ("TipoEndereco", 5.71),        # Col 22 (V): RUA, AV, ALAMEDA, etc.
    ("Logradouro", 27.86),         # Col 23 (W): Nome da via pública
    ("Numero", 8.0),               # Col 24 (X): Número predial
    ("Complemento", 12.43),        # Col 25 (Y): Sala, bloco, galpão
    ("Bairro", 56.14),             # Col 26 (Z): Bairro declarado
    ("Cep", 11.57),                # Col 27 (AA): CEP do estabelecimento
    ("FinalidadeEndereco", 17.86), # Col 28 (AB): Finalidade cadastral do endereço
    ("CnaesLicenciadas", 50.71),   # Col 29 (AC): Relação de CNAEs licenciadas
    ("CnaesNaoLicenciadas", 11.57),# Col 30 (AD): CNAEs não licenciadas
    ("Iptu", 13.0),                # Col 31 (AE): IPTU auxiliar de conferência
    ("CEP4", 13.0),                # Col 32 (AF): CEP raiz (5 primeiros dígitos)
]

# ==============================================================================
# REGRAS DE FORMATAÇÃO CONDICIONAL NA EXATA ORDEM DE PRIORIDADE INSTITUCIONAL
# ==============================================================================
# No Excel, a ordem de avaliação das regras define a cor final da célula.
# Regras de Status explícito (Em Andamento, Indeferido, Finalizado) prevalecem
# sobre regras de fiscal atribuído e classificação de serviço.
REGRAS_CONDICIONAIS_ORDEM: list[tuple[str, str]] = [
    ('$B3="Em Andamento"', COR_EM_ANDAMENTO),
    ('$B3="Indeferido"', COR_FINALIZADO_STATUS),
    ('$B3="Cancelado"', COR_FINALIZADO_STATUS),
    ('$B3="Cassado"', COR_FINALIZADO_STATUS),
    ('$B3="Deferido"', COR_FINALIZADO_STATUS),
    ('$P3="INDEFERIDO"', COR_INDEFERIDO),
    ('$P3="CANCELADO"', COR_INDEFERIDO),
    ('$D3="Noturno entregue"', COR_NOTURNO_ENTREGUE),
    ('$C3<>""', COR_FISCAL_ATRIBUIDO),
    ('$D3="Noturno Livre"', COR_NOTURNO_LIVRE),
    ('$E3="Noturno"', COR_NOTURNO_LIVRE),
    ('$D3="Sem Restrição Livre"', COR_SEM_RESTRICAO),
    ('$E3="Sem Restrição"', COR_SEM_RESTRICAO),
    ('$D3="Com Restrição Livre"', COR_RESTRICAO_LIVRE),
    ('$E3="Com Restrição"', COR_RESTRICAO_LIVRE),
]


def _calcular_hash_xlsx(caminho: Path) -> str:
    """Calcula o hash SHA-256 de um arquivo Excel em disco.

    Args:
        caminho: Caminho absoluto para o arquivo .xlsx.

    Returns:
        str: Hash SHA-256 em hexadecimal ou string vazia se o arquivo não existir.

    Complexity:
        Tempo: O(B) onde B é o tamanho do arquivo em disco.
        Espaço: O(1).
    """
    if not caminho.exists() or not caminho.is_file():
        return ""
    hasher = hashlib.sha256()
    try:
        with open(caminho, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return ""


def registrar_hash_planilha_gerada(
    caminho_xlsx: Path | str,
    triagem_conn: sqlite3.Connection | None = None,
) -> str:
    """Registra o hash SHA-256 da planilha recém-gerada pelo sistema.

    Salva o hash em memória e na tabela `controle_planilha_ativa`.
    Dessa forma, o sistema sabe com exatidão matemática se o usuário editou ou não
    o arquivo no Excel antes de executar a próxima triagem.

    Args:
        caminho_xlsx: Caminho do arquivo Excel salvo.
        triagem_conn: Conexão SQLite opcional para persistência do hash em banco.

    Returns:
        str: Hash SHA-256 registrado.

    Complexity:
        Tempo: O(B) onde B é o tamanho do arquivo Excel.
        Espaço: O(1).
    """
    p = Path(caminho_xlsx).resolve()
    h = _calcular_hash_xlsx(p)
    if not h:
        return ""
    chave = str(p)
    _ULTIMO_HASH_GERADO[chave] = h

    if triagem_conn is not None:
        try:
            with db_transaction(triagem_conn) as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS controle_planilha_ativa (
                        caminho TEXT PRIMARY KEY,
                        hash_gerado TEXT,
                        atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
                    );
                    """
                )
                cur.execute(
                    """
                    INSERT INTO controle_planilha_ativa (caminho, hash_gerado, atualizado_em)
                    VALUES (?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(caminho) DO UPDATE SET
                        hash_gerado = excluded.hash_gerado,
                        atualizado_em = CURRENT_TIMESTAMP;
                    """,
                    (chave, h),
                )
        except Exception:
            pass
    return h


def _obter_hash_planilha_registrado(
    caminho_xlsx: Path,
    triagem_conn: sqlite3.Connection | None = None,
) -> str:
    """Recupera o último hash SHA-256 gerado pelo sistema para a planilha ativa.

    Args:
        caminho_xlsx: Caminho do arquivo Excel a verificar.
        triagem_conn: Conexão SQLite opcional com o banco de triagem.

    Returns:
        str: Hash registrado ou string vazia.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    chave = str(caminho_xlsx.resolve())
    if triagem_conn is not None:
        try:
            cur = triagem_conn.cursor()
            cur.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='controle_planilha_ativa';"
            )
            if cur.fetchone():
                cur.execute(
                    "SELECT hash_gerado FROM controle_planilha_ativa WHERE caminho = ? LIMIT 1;",
                    (chave,),
                )
                row = cur.fetchone()
                if row and row["hash_gerado"]:
                    return str(row["hash_gerado"])
        except Exception:
            pass
    return _ULTIMO_HASH_GERADO.get(chave, "")


def _limpar_valor_celula(val: Any) -> str:
    """Converte valor de célula do Excel em string limpa, preservando inteiros e datas.

    Trata:
    - Booleans, ints, floats (removendo sufixo .0 de inteiros).
    - Objetos `datetime.date` e `datetime.datetime` formatados em 'AAAA-MM-DD'.
    - Nulos e valores de erro do Excel ('#N/A', '#REF!') convertidos para `""`.

    Args:
        val: Valor extraído de `cell.value`.

    Returns:
        str: String sanitizada representativa do conteúdo.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val).strip()
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if s.upper() in ("NONE", "NULL", "#N/A", "#REF!"):
        return ""
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s


def _extrair_coluna_planilha(row: tuple[Any, ...] | list[Any], pos_1based: int) -> str:
    """Extrai e higieniza o valor textual de uma coluna posicional (1-based) da linha da planilha.

    Args:
        row: Tupla ou lista contendo os valores das células da linha.
        pos_1based: Índice posicional da coluna iniciando em 1 (ex: Coluna B = 2).

    Returns:
        str: String higienizada sem espaços supérfluos ou string vazia se ausente.

    Complexity:
        Tempo: O(1) - acesso indexado em tupla/lista.
        Espaço: O(1).
    """
    if pos_1based <= len(row):
        return _limpar_valor_celula(row[pos_1based - 1])
    return ""


def sincronizar_planilha_ativa_com_banco(
    caminho_xlsx: Path | str | None = None,
    triagem_conn: sqlite3.Connection | None = None,
    cep_conn: sqlite3.Connection | None = None,
    forcar: bool = False,
) -> int:
    """Lê a planilha ativa existente e sincroniza edições manuais de volta para o SQLite.

    Garante a regra de ouro de Zero Perda de Dados:
    Quando o fiscal ou operador abre o Excel e altera status, atribui fiscal, digita número
    de OS/Sigm, corrige IPTU ou ajusta Área, essas informações são salvas de volta no banco
    antes que novos arquivos CSV sejam processados ou que o arquivo seja recriado.

    Fluxo de Trabalho:
    1. Compara o hash atual do arquivo no disco com o hash gravado na última geração.
       Se idênticos e `forcar=False`, encerra imediatamente sem custo de processamento.
    2. Lê os dados em modo `read_only=True` e `data_only=True`.
    3. Mapeia colunas da linha 2 em diante.
    4. Identifica o protocolo de licença correspondente no banco.
    5. Detecta se houve modificação manual e atualiza `vre_backlog`, `historico_fiscais`
       e `cep_areas.db`.

    Args:
        caminho_xlsx: Caminho para a planilha ativa em disco.
        triagem_conn: Conexão SQLite opcional com `triagem_vre.db`.
        cep_conn: Conexão SQLite opcional com `cep_areas.db`.
        forcar: Se True, força a leitura mesmo que os hashes coincidam.

    Returns:
        int: Quantidade de registros modificados e sincronizados com sucesso.

    Complexity:
        Tempo: O(L) onde L é o número de linhas da planilha.
        Espaço: O(L) para armazenamento das tuplas de atualização.
    """
    alvo = Path(caminho_xlsx) if caminho_xlsx else OUTPUT_XLSX_PATH
    if not alvo.exists() or not alvo.is_file() or alvo.stat().st_size == 0:
        return 0

    close_triagem = False
    if triagem_conn is None:
        triagem_conn = get_triagem_db()
        initialize_schema_if_needed(triagem_conn, is_cep=False)
        close_triagem = True

    close_cep = False
    if cep_conn is None:
        try:
            cep_conn = get_cep_db()
            initialize_schema_if_needed(cep_conn, is_cep=True)
            close_cep = True
        except Exception:
            cep_conn = None

    try:
        hash_atual = _calcular_hash_xlsx(alvo)
        hash_registrado = _obter_hash_planilha_registrado(alvo, triagem_conn)

        # Se a planilha em disco tem exatamente o mesmo hash de quando foi gerada pelo sistema,
        # nenhuma edição humana ocorreu no arquivo.
        if not forcar and hash_atual and hash_registrado and hash_atual == hash_registrado:
            return 0

        try:
            wb = openpyxl.load_workbook(str(alvo), data_only=True, read_only=True)
        except Exception:
            return 0

        ws = wb["Controle VRE"] if "Controle VRE" in wb.sheetnames else wb.active
        linhas_planilha = list(ws.iter_rows(values_only=True))
        wb.close()

        # Precisa ter pelo menos cabeçalho na linha 2 e 1 linha de dados (linha 3+)
        if len(linhas_planilha) < 3:
            return 0

        cabecalho_l2 = [str(c or "").strip().upper() for c in linhas_planilha[1]]
        # Validação estrita: certifica-se de que é a planilha de trabalho oficial
        if not any(k in cabecalho_l2 for k in ("STATUS", "FISCAL", "PROTOCOLOLICENCA", "CNPJ")):
            return 0

        cur = triagem_conn.cursor()
        cur.execute(
            """
            SELECT
                protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
                nome_empresarial, mes_competencia, data_importacao,
                bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
                flag_restricao, flag_noturno, tipo_servico, status,
                dados_extras_json
            FROM vre_backlog;
            """
        )
        existentes_por_licenca: dict[str, dict[str, Any]] = {}
        existentes_por_sil: dict[str, dict[str, Any]] = {}
        for r in cur.fetchall():
            d = dict(r)
            lic = str(d.get("protocolo_licenca") or "").strip()
            sil = str(d.get("protocolo_sil") or "").strip()
            if lic:
                existentes_por_licenca[lic] = d
            if sil and sil not in existentes_por_sil:
                existentes_por_sil[sil] = d

        atualizacoes_backlog: list[tuple[Any, ...]] = []
        insercoes_backlog: list[tuple[Any, ...]] = []
        fiscais_para_historico: list[tuple[str, str, str, str]] = []
        ceps_para_mapear: list[tuple[str, str, str]] = []
        meses_ativos_detectados: set[str] = set()

        limpeza_col_f_concluida = _migracao_limpeza_col_f_concluida(triagem_conn)
        data_hoje = datetime.date.today().isoformat()

        for idx_linha, row in enumerate(linhas_planilha[2:], start=3):
            if not row or not any(row):
                continue

            _col = lambda pos_1based: _extrair_coluna_planilha(row, pos_1based)

            status_excel = _col(2)
            fiscal_excel = _col(3)
            tipo_servico_excel = _col(4)
            restricao_excel = _col(5)
            val_col6_raw = row[5] if len(row) >= 6 else None
            if isinstance(val_col6_raw, (datetime.date, datetime.datetime)):
                data_col6 = val_col6_raw.strftime("%d/%m/%Y")
            else:
                data_col6 = _col(6)
            iptu_col7 = _col(7)
            area_excel = _col(8)
            endereco_excel = _col(9)
            prot_existente_excel = _col(10)
            prot_sigm_excel = _col(11)
            prot_sil_excel = _col(12)
            prot_licenca_excel = _col(13)
            data_col14 = _col(14)
            grau_risco_excel = _col(15)
            situacao_postura_excel = _col(16)
            insc_mun_excel = _col(17)
            cnpj_excel = _col(18)
            nome_emp_excel = _col(19)
            telefone_excel = _col(20)
            email_excel = _col(21)
            tipo_end_excel = _col(22)
            logradouro_excel = _col(23)
            numero_excel = _col(24)
            complemento_excel = _col(25)
            bairro_excel = _col(26)
            cep_excel = _col(27)
            finalidade_excel = _col(28)
            cnaes_lic_excel = _col(29)
            cnaes_nao_lic_excel = _col(30)
            iptu_col31 = _col(31)

            # Ignora linhas sem nenhuma chave de identificação
            if not prot_licenca_excel and not prot_sil_excel and not cnpj_excel and not nome_emp_excel:
                continue

            reg_ant = None
            if prot_licenca_excel and prot_licenca_excel in existentes_por_licenca:
                reg_ant = existentes_por_licenca[prot_licenca_excel]
            elif prot_sil_excel and prot_sil_excel in existentes_por_sil:
                reg_ant = existentes_por_sil[prot_sil_excel]

            extras: dict[str, Any] = {}
            if reg_ant and reg_ant.get("dados_extras_json"):
                try:
                    extras = json.loads(reg_ant["dados_extras_json"])
                except Exception:
                    extras = {}

            prot_licenca_final = (
                reg_ant["protocolo_licenca"]
                if reg_ant
                else (prot_licenca_excel or f"MANUAL_{normalizar_cnpj(cnpj_excel) or idx_linha}_{data_hoje}")
            )
            prot_sil_final = prot_sil_excel or (reg_ant.get("protocolo_sil") if reg_ant else "") or ""

            cnpj_limpo = normalizar_cnpj(cnpj_excel) or (reg_ant.get("cnpj_limpo") if reg_ant else "") or ""
            cnpj_fmt = formatar_cnpj(cnpj_limpo) if len(cnpj_limpo) == 14 else (cnpj_excel or (reg_ant.get("cnpj_formatado") if reg_ant else "") or "")
            nome_emp_final = nome_emp_excel or (reg_ant.get("nome_empresarial") if reg_ant else "") or ""

            # Preserva competência e data_importacao do registro existente
            data_vre_final = data_col14 or extras.get("Data") or (reg_ant.get("data_importacao") if reg_ant else "") or data_col6 or data_hoje
            if reg_ant and reg_ant.get("mes_competencia"):
                mes_comp_final = reg_ant["mes_competencia"]
                data_imp_final = reg_ant.get("data_importacao") or data_vre_final
            else:
                # Extrai AAAA-MM
                if "/" in data_vre_final and len(data_vre_final.split("/")) == 3:
                    p_d = data_vre_final.split("/")
                    mes_comp_final = f"{p_d[2][:4]}-{p_d[1].zfill(2)}"
                    data_imp_final = f"{p_d[2][:4]}-{p_d[1].zfill(2)}-{p_d[0].zfill(2)}"
                elif "-" in data_vre_final and len(data_vre_final) >= 7:
                    mes_comp_final = data_vre_final[:7]
                    data_imp_final = data_vre_final[:10]
                else:
                    mes_comp_final = data_hoje[:7]
                    data_imp_final = data_hoje
            meses_ativos_detectados.add(mes_comp_final)

            # Higienização de Status: se o usuário preencheu Status (ex: 'Em Andamento', 'Deferido', etc.), preserva
            status_upper = status_excel.upper()
            if not status_excel or status_upper in (
                "PENDENTE",
                "DISTRIBUIDO",
                "DISTRIBUÍDO",
                "NÃO DISTRIBUÍDO",
                "NAO DISTRIBUIDO",
                "NÃO DISTRIBUIDO",
                "NAO DISTRIBUÍDO",
            ):
                status_db = "DISTRIBUIDO" if fiscal_excel else "PENDENTE"
                extras["Status"] = ""
            else:
                status_db = status_excel
                extras["Status"] = status_excel

            eh_indeferido_excel = (
                situacao_postura_excel.strip().upper() == "INDEFERIDO"
                or status_upper == "INDEFERIDO"
            )

            # Flags e Restrição (quando indeferido, não deve preencher Tipo de Serviço nem Restrição)
            flag_r = int(reg_ant.get("flag_restricao") or 0) if reg_ant else 0
            flag_n = int(reg_ant.get("flag_noturno") or 0) if reg_ant else 0
            if eh_indeferido_excel:
                extras.pop("RestricaoCustom", None)
                extras.pop("TipoServicoCustom", None)
                tipo_servico_final = ""
            else:
                if restricao_excel:
                    r_up = restricao_excel.upper()
                    if "NOTURNO" in r_up:
                        flag_n = 1
                    elif "RESTRI" in r_up and "SEM" not in r_up:
                        flag_r = 1
                        flag_n = 0
                    elif "SEM" in r_up:
                        flag_r = 0
                        flag_n = 0
                    extras["RestricaoCustom"] = restricao_excel
                else:
                    extras.pop("RestricaoCustom", None)

                # Tipo de Serviço (quando tem dono e não é noturno -> Diurno Livre, salvo se editado)
                tem_dono_excel = bool(
                    fiscal_excel and fiscal_excel.upper() not in ("NÃO INFORMADO", "NAO INFORMADO", "PENDENTE")
                )
                if tem_dono_excel and (
                    not tipo_servico_excel
                    or tipo_servico_excel in (
                        "Noturno Livre",
                        "Com Restrição Livre",
                        "Sem Restrição Livre",
                        "Com Restricao Livre",
                        "Sem Restricao Livre",
                    )
                ):
                    tipo_servico_final = "Noturno" if (flag_n or tipo_servico_excel == "Noturno Livre") else "Diurno Livre"
                else:
                    tipo_servico_final = tipo_servico_excel

                if tipo_servico_final:
                    extras["TipoServicoCustom"] = tipo_servico_final
                else:
                    extras.pop("TipoServicoCustom", None)

            # IPTU (prioriza edição na Coluna 7 ou Coluna 31)
            iptu_ant = str(extras.get("Iptu") or extras.get("IPTU") or "")
            if iptu_col7 != iptu_ant:
                iptu_final = iptu_col7
            elif iptu_col31 != iptu_ant:
                iptu_final = iptu_col31
            else:
                iptu_final = iptu_col7 or iptu_col31
            extras["Iptu"] = iptu_final
            extras["IPTU"] = iptu_final

            # Área (preserva edição ou marcação de limpeza manual)
            area_ant = str(reg_ant.get("area") or "").strip() if reg_ant else ""
            area_final = harmonizar_nome_area(area_excel) if area_excel else ""
            if area_ant and not area_final:
                extras["AreaLimpaManualmente"] = True
            elif area_final:
                extras["AreaLimpaManualmente"] = False

            # Campos complementares
            bairro_final = bairro_excel or (reg_ant.get("bairro_original") if reg_ant else "") or ""
            cnaes_lic_final = cnaes_lic_excel or (reg_ant.get("cnaes_licenciadas") if reg_ant else "") or ""

            extras["Data"] = data_vre_final
            if limpeza_col_f_concluida and data_col6:
                extras["DataCol6"] = data_col6
            else:
                extras.pop("DataCol6", None)

            extras["ProtocoloExistente"] = prot_existente_excel
            extras["ProtocoloSigm"] = prot_sigm_excel
            extras["ProtocoloSigm / OS"] = prot_sigm_excel
            extras["GrauRiscoGeral"] = grau_risco_excel
            extras["SituacaoPostura"] = situacao_postura_excel
            extras["InscricaoMunicipal"] = insc_mun_excel
            extras["Telefone"] = telefone_excel
            extras["E-mail"] = email_excel
            extras["TipoEndereco"] = tipo_end_excel
            extras["Logradouro"] = logradouro_excel
            extras["Numero"] = numero_excel
            extras["Complemento"] = complemento_excel
            extras["Bairro"] = bairro_final
            extras["Cep"] = cep_excel
            extras["FinalidadeEndereco"] = finalidade_excel
            extras["CnaesLicenciadas"] = cnaes_lic_final
            extras["CnaesNaoLicenciadas"] = cnaes_nao_lic_excel

            end_recalculado = montar_endereco_completo(
                tipo_end_excel,
                logradouro_excel,
                numero_excel,
                complemento_excel,
                bairro_final,
            )
            if endereco_excel and endereco_excel != end_recalculado:
                extras["EnderecoCustom"] = endereco_excel
            else:
                extras.pop("EnderecoCustom", None)

            extras_json = json.dumps(extras, ensure_ascii=False)

            if reg_ant:
                atualizacoes_backlog.append(
                    (
                        prot_sil_final,
                        cnpj_limpo,
                        cnpj_fmt,
                        nome_emp_final,
                        bairro_final,
                        area_final,
                        cnaes_lic_final,
                        fiscal_excel,
                        flag_r,
                        flag_n,
                        tipo_servico_final,
                        status_db,
                        extras_json,
                        prot_licenca_final,
                    )
                )
            else:
                insercoes_backlog.append(
                    (
                        prot_licenca_final,
                        prot_sil_final,
                        cnpj_limpo,
                        cnpj_fmt,
                        nome_emp_final,
                        mes_comp_final,
                        data_imp_final,
                        bairro_final,
                        area_final,
                        cnaes_lic_final,
                        fiscal_excel,
                        flag_r,
                        flag_n,
                        tipo_servico_final,
                        status_db,
                        extras_json,
                    )
                )

            # Aprendizado contínuo: fiscais e CEPs editados alimentam os bancos de histórico
            if fiscal_excel and fiscal_excel.upper() not in ("NÃO INFORMADO", "NAO INFORMADO", "PENDENTE") and len(cnpj_limpo) == 14:
                fiscais_para_historico.append((cnpj_limpo, fiscal_excel, data_imp_final, "Planilha_Ativa"))

            cep_limpo_calc = normalizar_cep(cep_excel)
            if area_final and len(cep_limpo_calc) == 8:
                ceps_para_mapear.append((cep_limpo_calc, area_final, bairro_final))

        if atualizacoes_backlog or insercoes_backlog:
            with db_transaction(triagem_conn) as t_cur:
                if atualizacoes_backlog:
                    t_cur.executemany(
                        """
                        UPDATE vre_backlog
                        SET protocolo_sil = ?,
                            cnpj_limpo = ?,
                            cnpj_formatado = ?,
                            nome_empresarial = ?,
                            bairro_original = ?,
                            area = ?,
                            cnaes_licenciadas = ?,
                            fiscal_atribuido = ?,
                            flag_restricao = ?,
                            flag_noturno = ?,
                            tipo_servico = ?,
                            status = ?,
                            dados_extras_json = ?
                        WHERE protocolo_licenca = ?;
                        """,
                        atualizacoes_backlog,
                    )
                if insercoes_backlog:
                    t_cur.executemany(
                        """
                        INSERT OR IGNORE INTO vre_backlog (
                            protocolo_licenca, protocolo_sil, cnpj_limpo, cnpj_formatado,
                            nome_empresarial, mes_competencia, data_importacao,
                            bairro_original, area, cnaes_licenciadas, fiscal_atribuido,
                            flag_restricao, flag_noturno, tipo_servico, status, dados_extras_json
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                        """,
                        insercoes_backlog,
                    )
                if fiscais_para_historico:
                    t_cur.executemany(
                        """
                        INSERT INTO historico_fiscais (cnpj_limpo, fiscal, data_atendimento, origem_registro)
                        VALUES (?, ?, ?, ?)
                        ON CONFLICT(cnpj_limpo, fiscal) DO UPDATE SET
                            data_atendimento = CASE
                                WHEN excluded.data_atendimento > historico_fiscais.data_atendimento
                                THEN excluded.data_atendimento
                                ELSE historico_fiscais.data_atendimento
                            END;
                        """,
                        fiscais_para_historico,
                    )
                for m in meses_ativos_detectados:
                    if m and len(m) == 7:
                        t_cur.execute(
                            """
                            INSERT OR IGNORE INTO controle_competencias (mes_competencia, arquivado)
                            VALUES (?, 0);
                            """,
                            (m,),
                        )

        if cep_conn is not None and ceps_para_mapear:
            for cep_l, area_f, bairro_f in ceps_para_mapear:
                try:
                    registrar_area_cep(
                        cep=cep_l,
                        area=area_f,
                        bairro=bairro_f,
                        origem="PLANILHA_ATIVA",
                        conn=cep_conn,
                        triagem_conn=triagem_conn,
                    )
                except Exception:
                    pass

        return len(atualizacoes_backlog) + len(insercoes_backlog)
    finally:
        if close_triagem:
            triagem_conn.close()
        if close_cep and cep_conn is not None:
            cep_conn.close()


def limpar_indeferidos_e_normalizar_backlog(
    triagem_conn: sqlite3.Connection,
    executar_migracao_col_f: bool = True,
) -> int:
    """Garante a sanitização e normalização de regras de negócio em todo o backlog.

    Regras aplicadas de forma abrangente:
    1. Protocolos indeferidos (na postura ou no status) têm Tipo de Serviço e Restrição omitidos/limpos.
    2. Protocolos com fiscal titular e não-noturnos recebem 'Diurno Livre' (ou 'Noturno' se noturnos).
    3. Executa a migração única de limpeza das datas residuais da Coluna F caso ainda pendente.

    Args:
        triagem_conn: Conexão SQLite ativa com `triagem_vre.db`.
        executar_migracao_col_f: Se True, efetua a limpeza única da Coluna F se necessário.

    Returns:
        int: Quantidade de registros ajustados no banco.

    Complexity:
        Tempo: O(N) onde N é o total de linhas do backlog.
        Espaço: O(N) para a lista de atualizações.
    """
    precisa_limpar_col_f = executar_migracao_col_f and not _migracao_limpeza_col_f_concluida(triagem_conn)

    cur = triagem_conn.cursor()
    cur.execute(
        """
        SELECT protocolo_licenca, fiscal_atribuido, flag_restricao, flag_noturno,
               tipo_servico, status, dados_extras_json
        FROM vre_backlog;
        """
    )
    linhas = cur.fetchall()
    atualizacoes: list[tuple[str, str, str]] = []

    for r in linhas:
        prot_lic = r["protocolo_licenca"]
        fisc = str(r["fiscal_atribuido"] or "").strip()
        flag_n = int(r["flag_noturno"] or 0)
        ts_atual = str(r["tipo_servico"] or "").strip()
        status_db = str(r["status"] or "").strip()
        raw_json = r["dados_extras_json"] or "{}"
        try:
            extras = json.loads(raw_json)
        except Exception:
            extras = {}

        mudou_json = False
        if precisa_limpar_col_f and "DataCol6" in extras:
            extras.pop("DataCol6", None)
            mudou_json = True

        postura_up = str(extras.get("SituacaoPostura") or "").strip().upper()
        status_ex_up = str(extras.get("Status") or "").strip().upper()
        eh_indeferido = (
            postura_up == "INDEFERIDO"
            or status_db.upper() == "INDEFERIDO"
            or status_ex_up == "INDEFERIDO"
        )

        if eh_indeferido:
            if "RestricaoCustom" in extras:
                extras.pop("RestricaoCustom", None)
                mudou_json = True
            if "TipoServicoCustom" in extras:
                extras.pop("TipoServicoCustom", None)
                mudou_json = True
            if ts_atual != "" or mudou_json:
                novo_json = json.dumps(extras, ensure_ascii=False) if mudou_json else raw_json
                atualizacoes.append(("", novo_json, prot_lic))
        else:
            tem_dono = bool(fisc and fisc.upper() not in ("NÃO INFORMADO", "NAO INFORMADO", "PENDENTE"))
            novo_ts = ts_atual
            if tem_dono and (
                not ts_atual
                or ts_atual in (
                    "Noturno Livre",
                    "Com Restrição Livre",
                    "Sem Restrição Livre",
                    "Com Restricao Livre",
                    "Sem Restricao Livre",
                )
            ):
                novo_ts = "Noturno" if (flag_n == 1 or ts_atual == "Noturno Livre") else "Diurno Livre"
            if novo_ts != ts_atual or mudou_json:
                novo_json = json.dumps(extras, ensure_ascii=False) if mudou_json else raw_json
                atualizacoes.append((novo_ts, novo_json, prot_lic))

    if atualizacoes or precisa_limpar_col_f:
        with db_transaction(triagem_conn) as t_cur:
            if atualizacoes:
                t_cur.executemany(
                    """
                    UPDATE vre_backlog
                    SET tipo_servico = ?,
                        dados_extras_json = ?
                    WHERE protocolo_licenca = ?;
                    """,
                    atualizacoes,
                )
            if precisa_limpar_col_f:
                t_cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS controle_migracoes (
                        chave TEXT PRIMARY KEY,
                        executado_em DATETIME DEFAULT CURRENT_TIMESTAMP
                    );
                    """
                )
                t_cur.execute(
                    """
                    INSERT OR IGNORE INTO controle_migracoes (chave, executado_em)
                    VALUES (?, CURRENT_TIMESTAMP);
                    """,
                    (_MIGRACAO_LIMPEZA_COL_F_KEY,),
                )
    return len(atualizacoes)


def aplicar_regras_condicionais(ws, sqref: str) -> None:
    """Aplica as regras de formatação condicional na exata ordem institucional do exemplo.

    Args:
        ws: Worksheet do OpenPyXL.
        sqref: Faixa de células no formato 'A3:AF5000'.

    Complexity:
        Tempo: O(R) onde R é o número de regras (15).
        Espaço: O(1).
    """
    for formula, cor in REGRAS_CONDICIONAIS_ORDEM:
        fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
        rule = FormulaRule(formula=[formula], fill=fill)
        ws.conditional_formatting.add(sqref, rule)


def criar_template_padrao(caminho_template: Path | str | None = None, forcar: bool = False) -> Path:
    """Cria e formata o arquivo template_triagem.xlsx caso ele não exista.

    Constrói a planilha protótipo de 32 colunas, linhas de cabeçalho, larguras
    oficiais, alinhamentos, congelamento de painel em A3 e regras de formatação condicional.

    Args:
        caminho_template: Caminho alvo onde salvar o template.
        forcar: Se True, recria o template mesmo se o arquivo já existir.

    Returns:
        Path: Caminho do arquivo de template gerado.

    Complexity:
        Tempo: O(1) - construção de planilha pequena de 2 linhas.
        Espaço: O(1).
    """
    ensure_directories()
    alvo = Path(caminho_template) if caminho_template else TEMPLATE_XLSX_PATH
    if alvo.exists() and not forcar:
        return alvo
    alvo.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Controle VRE"

    fonte_padrao = Font(name="Arial", size=10, bold=False)
    fundo_cabecalho = PatternFill(
        start_color=COR_CABECALHO, end_color=COR_CABECALHO, fill_type="solid"
    )

    # Linha 1: Título municipal na coluna W (Logradouro, agora coluna 23)
    ws.row_dimensions[1].height = 12.75
    ws.cell(row=1, column=23, value="SÃO JOSÉ DO RIO PRETO").font = fonte_padrao

    # Linha 2: Cabeçalhos oficiais (altura 12.75)
    ws.row_dimensions[2].height = 12.75

    for col_idx, (nome_coluna, largura) in enumerate(COLUNAS_CABECALHO, start=1):
        celula = ws.cell(row=2, column=col_idx, value=nome_coluna)
        celula.font = fonte_padrao
        celula.fill = fundo_cabecalho

        # Alinhamentos específicos do cabeçalho
        if col_idx in (2, 5, 10, 11, 12, 17):
            celula.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in (4, 8):
            celula.alignment = Alignment(horizontal="left", vertical="center")
        else:
            celula.alignment = Alignment(vertical="center")

        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = largura

    # Painel congelado em A3 e filtro ativo
    ws.freeze_panes = "A3"
    ultima_letra = get_column_letter(len(COLUNAS_CABECALHO))
    ws.auto_filter.ref = f"A2:{ultima_letra}2"

    # Regras de formatação condicional na ordem exata
    aplicar_regras_condicionais(ws, f"A3:{ultima_letra}5000")

    wb.save(str(alvo))
    return alvo


def gerar_planilha_trabalho(
    registros_backlog: Sequence[dict[str, Any]],
    caminho_saida: Path | str | None = None,
    caminho_template: Path | str | None = None,
) -> Path:
    """Gera o arquivo Planilha_Controle_VRE_Ativa.xlsx a partir dos dados do SQLite.

    Otimização Extrema (Zero Travamento):
    1. Executa backup preventivo da planilha anterior em `/Planilhas anteriores`.
    2. Clona a estrutura protótipo do template `template_triagem.xlsx`.
    3. Escreve exclusivamente Plain Values nas células de dados.
    4. Aplica formatação condicional em bloco abrangente (`A3:AF5000`).
    5. Configura congelamento em A3 e auto-filtro A2:AF.
    6. Salva o arquivo em disco e registra o hash SHA-256 no banco para controle de edições.

    Args:
        registros_backlog: Sequência de dicionários com os protocolos a incluir.
        caminho_saida: Caminho de saída desejado (padrão: `OUTPUT_XLSX_PATH`).
        caminho_template: Caminho do template base (padrão: `resolve_template_path()`).

    Returns:
        Path: Caminho do arquivo Excel gerado.

    Complexity:
        Tempo: O(N * C) onde N é o número de linhas do backlog e C = 32 colunas.
        Espaço: O(N * C) para o grafo em memória do Workbook OpenPyXL.
    """
    ensure_directories()
    dest_path = Path(caminho_saida) if caminho_saida else OUTPUT_XLSX_PATH
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    # Backup preventivo em /Planilhas anteriores antes de sobrescrever
    if dest_path.exists() and dest_path.stat().st_size > 0:
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = PLANILHAS_ANTERIORES_DIR / f"Planilha_Controle_VRE_Ativa_backup_{timestamp}.xlsx"
            PLANILHAS_ANTERIORES_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(dest_path), str(backup_file))
        except Exception:
            pass

    template_path = Path(caminho_template) if caminho_template else resolve_template_path()
    if not template_path.exists():
        criar_template_padrao(template_path)

    wb = openpyxl.load_workbook(str(template_path))
    if "Controle VRE" in wb.sheetnames:
        ws = wb["Controle VRE"]
    else:
        ws = wb.active
        ws.title = "Controle VRE"

    # Garante cabeçalho institucional na linha 1 (Coluna W - Logradouro)
    ws.cell(row=1, column=23, value="SÃO JOSÉ DO RIO PRETO").font = Font(name="Arial", size=10)

    # Identifica altura padrão e estilos protótipo da linha 3 (se definidos no template)
    altura_linha = 12.75
    estilos_colunas: dict[int, dict[str, Any]] = {}
    if ws.max_row >= 3:
        if ws.row_dimensions[3].height:
            altura_linha = ws.row_dimensions[3].height
        for c_idx in range(1, len(COLUNAS_CABECALHO) + 1):
            cell_proto = ws.cell(row=3, column=c_idx)
            if cell_proto.font or cell_proto.alignment or cell_proto.border or cell_proto.fill or cell_proto.number_format:
                estilos_colunas[c_idx] = {
                    "font": copy(cell_proto.font) if cell_proto.font else None,
                    "alignment": copy(cell_proto.alignment) if cell_proto.alignment else None,
                    "border": copy(cell_proto.border) if cell_proto.border else None,
                    "fill": copy(cell_proto.fill) if cell_proto.fill else None,
                    "number_format": cell_proto.number_format if cell_proto.number_format != "General" else None,
                }
        ws.delete_rows(3, ws.max_row - 2)

    fonte_dados = Font(name="Arial", size=10)
    alinhamento_padrao = Alignment(vertical="center")
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_esquerda = Alignment(horizontal="left", vertical="center")

    linha_atual = 3
    for seq_num, reg in enumerate(registros_backlog, start=1):
        extras: dict[str, Any] = {}
        dados_extras_raw = reg.get("dados_extras_json")
        if dados_extras_raw:
            try:
                extras = json.loads(dados_extras_raw)
            except Exception:
                extras = {}

        # Higienização de Status: não deve ser escrito não distribuído e/ou distribuído, apenas deixe vazio
        status_raw = str(extras.get("Status") or reg.get("status") or "").strip()
        status_upper = status_raw.upper()
        if status_upper in [
            "PENDENTE",
            "DISTRIBUIDO",
            "DISTRIBUÍDO",
            "NÃO DISTRIBUÍDO",
            "NAO DISTRIBUIDO",
            "NÃO DISTRIBUIDO",
            "NAO DISTRIBUÍDO",
            "FINALIZADO",
            "NONE",
            "NULL",
            "",
        ]:
            status_exibicao = ""
        else:
            status_exibicao = status_raw

        situacao_postura_val = str(extras.get("SituacaoPostura", "") or "").strip()
        eh_indeferido = (
            situacao_postura_val.upper() == "INDEFERIDO"
            or status_upper == "INDEFERIDO"
        )

        # Quando está indeferido (somente quando indeferido), não deve preencher Tipo de Serviço e Restrição
        flag_r = int(reg.get("flag_restricao") or 0)
        flag_n = int(reg.get("flag_noturno") or 0)
        if eh_indeferido:
            classificacao_restricao = ""
            tipo_servico = ""
        else:
            restricao_custom = str(extras.get("RestricaoCustom") or "").strip()
            if restricao_custom:
                classificacao_restricao = restricao_custom
            elif flag_n:
                classificacao_restricao = "Noturno"
            elif flag_r:
                classificacao_restricao = "Com Restrição"
            else:
                classificacao_restricao = "Sem Restrição"

            # Tipo de Serviço sugerido (quando tem dono: se for noturno -> 'Noturno', se não for noturno -> 'Diurno Livre')
            fiscal_atribuido_val = (reg.get("fiscal_atribuido") or "").strip()
            tem_dono = bool(fiscal_atribuido_val and fiscal_atribuido_val != "Não Informado")

            tipo_servico = (reg.get("tipo_servico") or "").strip()
            if tem_dono and (
                not tipo_servico
                or tipo_servico in (
                    "Noturno Livre",
                    "Com Restrição Livre",
                    "Sem Restrição Livre",
                    "Com Restricao Livre",
                    "Sem Restricao Livre",
                )
            ):
                if flag_n or tipo_servico == "Noturno Livre":
                    tipo_servico = "Noturno"
                else:
                    tipo_servico = "Diurno Livre"
            elif not tipo_servico:
                if flag_n:
                    tipo_servico = "Noturno Livre"
                elif flag_r:
                    tipo_servico = "Com Restrição Livre"
                else:
                    tipo_servico = "Sem Restrição Livre"

        # Endereço estático limpo (Plain Values), preservando edição manual se houver
        tipo_end = extras.get("TipoEndereco", "")
        logr = extras.get("Logradouro", "")
        num = extras.get("Numero", "")
        comp = extras.get("Complemento", "")
        bairro = reg.get("bairro_original", "") or extras.get("Bairro", "")
        endereco_calculado = extras.get("EnderecoCustom") or montar_endereco_completo(tipo_end, logr, num, comp, bairro)

        # CEP e CEP4 (5 dígitos)
        cep_bruto = str(extras.get("Cep", "") or "")
        cep4 = cep_bruto[:5] if len(cep_bruto) >= 5 else cep_bruto

        # Coluna N (14): Data oficial do VRE. Coluna F (6): vazia por padrão, preenchida apenas a mão (DataCol6).
        data_vre = extras.get("Data", "") or reg.get("data_importacao", "")
        data_col6 = extras.get("DataCol6") or ""

        iptu_val = extras.get("Iptu", "") or extras.get("IPTU", "") or ""

        # Mapeamento oficial das 32 colunas
        valores_linha = [
            seq_num,                                                  # Col 1 (A): Número de ordem
            status_exibicao,                                          # Col 2 (B): Status (vazio para pendente/distribuído)
            reg.get("fiscal_atribuido") or "",                        # Col 3 (C): Fiscal
            tipo_servico,                                             # Col 4 (D): Tipo de Serviço
            classificacao_restricao,                                  # Col 5 (E): Restrição (NOVA COLUNA ÚNICA)
            data_col6,                                                # Col 6 (F): Data manual
            iptu_val,                                                 # Col 7 (G): IPTU
            reg.get("area") or "",                                    # Col 8 (H): ÁREA
            endereco_calculado,                                       # Col 9 (I): Endereço
            extras.get("ProtocoloExistente", "") or "",               # Col 10 (J): Protocolo Existente
            extras.get("ProtocoloSigm", "") or extras.get("ProtocoloSigm / OS", "") or "", # Col 11 (K): ProtocoloSigm / OS
            reg.get("protocolo_sil") or "",                           # Col 12 (L): ProtocoloSil
            reg.get("protocolo_licenca") or "",                       # Col 13 (M): ProtocoloLicenca
            data_vre,                                                 # Col 14 (N): Data VRE
            extras.get("GrauRiscoGeral", ""),                         # Col 15 (O): GrauRiscoGeral
            extras.get("SituacaoPostura", ""),                        # Col 16 (P): SituacaoPostura
            extras.get("InscricaoMunicipal", ""),                     # Col 17 (Q): InscricaoMunicipal
            reg.get("cnpj_formatado") or "",                          # Col 18 (R): Cnpj
            reg.get("nome_empresarial") or "",                        # Col 19 (S): NomeEmpresarial
            extras.get("Telefone", ""),                               # Col 20 (T): Telefone
            extras.get("E-mail", ""),                                 # Col 21 (U): E-mail
            tipo_end,                                                 # Col 22 (V): TipoEndereco
            logr,                                                     # Col 23 (W): Logradouro
            num,                                                      # Col 24 (X): Numero
            comp,                                                     # Col 25 (Y): Complemento
            bairro,                                                   # Col 26 (Z): Bairro
            cep_bruto,                                                # Col 27 (AA): Cep
            extras.get("FinalidadeEndereco", ""),                     # Col 28 (AB): FinalidadeEndereco
            reg.get("cnaes_licenciadas") or "",                       # Col 29 (AC): CnaesLicenciadas
            extras.get("CnaesNaoLicenciadas", ""),                    # Col 30 (AD): CnaesNaoLicenciadas
            iptu_val,                                                 # Col 31 (AE): Iptu
            cep4,                                                     # Col 32 (AF): CEP4
        ]

        ws.row_dimensions[linha_atual].height = altura_linha
        for col_idx, valor in enumerate(valores_linha, start=1):
            celula = ws.cell(row=linha_atual, column=col_idx, value=valor)
            if col_idx in estilos_colunas:
                proto = estilos_colunas[col_idx]
                celula.font = copy(proto["font"]) if proto["font"] else fonte_dados
                if proto["alignment"]:
                    celula.alignment = copy(proto["alignment"])
                elif col_idx in (2, 5, 10, 11, 12, 17):
                    celula.alignment = alinhamento_centro
                elif col_idx in (4, 8):
                    celula.alignment = alinhamento_esquerda
                else:
                    celula.alignment = alinhamento_padrao

                if proto["border"]:
                    celula.border = copy(proto["border"])
                if proto["fill"]:
                    celula.fill = copy(proto["fill"])
                if proto["number_format"]:
                    celula.number_format = proto["number_format"]
            else:
                celula.font = fonte_dados
                # Alinhamentos em estrita conformidade com o padrão institucional
                if col_idx in (2, 5, 10, 11, 12, 17):
                    celula.alignment = alinhamento_centro
                elif col_idx in (4, 8):
                    celula.alignment = alinhamento_esquerda
                else:
                    celula.alignment = alinhamento_padrao

        linha_atual += 1

    ultima_linha = max(3, linha_atual - 1)
    total_colunas = len(COLUNAS_CABECALHO)
    ultima_letra = get_column_letter(total_colunas)

    limite_cf = max(ultima_linha, 5000)
    faixa_cf = f"A3:{ultima_letra}{limite_cf}"

    # Preserva integralmente as regras de formatação condicional definidas no template
    regras_template = list(ws.conditional_formatting._cf_rules.items())
    if regras_template:
        ws.conditional_formatting = ConditionalFormattingList()
        for cf_obj, rules in regras_template:
            for r in rules:
                ws.conditional_formatting.add(faixa_cf, r)
    else:
        # Fallback apenas se o template não tiver nenhuma regra definida
        aplicar_regras_condicionais(ws, faixa_cf)

    if len(registros_backlog) > 0:
        ws.auto_filter.ref = f"A2:{ultima_letra}{ultima_linha}"
    else:
        ws.auto_filter.ref = f"A2:{ultima_letra}2"

    ws.freeze_panes = "A3"
    wb.save(str(dest_path))

    # Registra hash em memória da planilha gerada pelo sistema
    registrar_hash_planilha_gerada(dest_path)

    # Sincroniza saída e template com a pasta fbuild se existir
    try:
        fbuild_dir = BASE_DIR / "fbuild"
        if fbuild_dir.exists():
            fbuild_saida = fbuild_dir / "Planilha_Controle_VRE_Ativa.xlsx"
            if dest_path.resolve() != fbuild_saida.resolve():
                shutil.copy2(str(dest_path), str(fbuild_saida))
                registrar_hash_planilha_gerada(fbuild_saida)
            fbuild_tmpl = fbuild_dir / "Configurações" / "templates" / "template_triagem.xlsx"
            if template_path.resolve() != fbuild_tmpl.resolve() and template_path.exists():
                fbuild_tmpl.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(template_path), str(fbuild_tmpl))
    except Exception:
        pass

    return dest_path
