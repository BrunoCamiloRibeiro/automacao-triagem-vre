"""Módulo de Configuração Centralizada da Automação VRE.

Centraliza todas as constantes de ambiente, caminhos dinâmicos do sistema de arquivos,
definições de banco de dados SQLite, parâmetros de formatação condicional do Excel,
configurações de decodificação de arquivos CSV e rotinas de bootstrap de infraestrutura.

Compatibilidade:
- Execução direta via interpretador Python padrão em ambiente de desenvolvimento.
- Execução via pacote portátil do Windows (pasta `Configurações/` com Python embutido).
- Execução como binário congelado autossuficiente via PyInstaller (`sys.frozen`).
- Suporte a sobrescrita dinâmica por variáveis de ambiente (`VRE_*`) para testes automatizados.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# ==============================================================================
# DETERMINAÇÃO DINÂMICA DA PASTA BASE (BASE_DIR E BUNDLE_DIR)
# ==============================================================================
# Avalia a origem da execução:
# 1. Se congelado via PyInstaller (`sys.frozen`), o executável reside no diretório do binário,
#    e os assets temporários extraídos residem em `sys._MEIPASS`.
# 2. Se em modo interpretado, retrocede 3 níveis a partir deste arquivo (`src/core/config.py`),
#    garantindo que se a pasta raiz for 'Configurações' (versão portátil), a raiz seja ajustada
#    para o diretório pai onde residem os executáveis e planilhas de trabalho.
if getattr(sys, "frozen", False):
    _default_base = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", _default_base))
else:
    _default_base = Path(__file__).resolve().parent.parent.parent
    if _default_base.name.lower() in ["configurações", "configuracoes", "config"]:
        _default_base = _default_base.parent
    BUNDLE_DIR = _default_base

# Suporte a override por variável de ambiente (essencial para isolar testes unitários em diretórios temporários)
BASE_DIR: Path = Path(os.getenv("VRE_BASE_DIR", str(_default_base))).resolve()


# ==============================================================================
# RESOLUÇÃO DE DIRETÓRIOS INSTITUCIONAIS
# ==============================================================================

def resolve_configuracoes_dir() -> Path:
    """Localiza o diretório de configurações (/Configurações ou configurado por ambiente).

    Verifica se existe override via variável de ambiente `VRE_CONFIG_DIR` ou busca
    variações nominais no disco para tolerar desvios de maiúsculas/minúsculas ou acentuação.

    Returns:
        Path: Caminho absoluto para a pasta de configurações do sistema.

    Complexity:
        Tempo: O(1) - varredura de lista estática de 4 nomes em disco.
        Espaço: O(1) - alocação de poucos objetos Path.
    """
    env_cfg = os.getenv("VRE_CONFIG_DIR")
    if env_cfg:
        return Path(env_cfg).resolve()

    for nome in ["Configurações", "Configuracoes", "configuracoes", "config"]:
        candidato = BASE_DIR / nome
        if candidato.exists() and candidato.is_dir():
            return candidato
    return BASE_DIR / "Configurações"


def resolve_planilhas_anteriores_dir() -> Path:
    """Localiza o diretório de backup e histórico de planilhas anteriores.

    Garante que os backups preventivos e importações legadas sejam depositados na
    pasta institucional `/Planilhas anteriores`, tolerando variações de nome.

    Returns:
        Path: Caminho absoluto para o diretório de planilhas anteriores.

    Complexity:
        Tempo: O(1) - checagem de existência de 3 nomes de pastas.
        Espaço: O(1).
    """
    env_dir = os.getenv("VRE_PLANILHAS_ANTERIORES_DIR")
    if env_dir:
        return Path(env_dir).resolve()

    for nome in ["Planilhas anteriores", "Planilhas Anteriores", "planilhas_anteriores"]:
        candidato = BASE_DIR / nome
        if candidato.exists() and candidato.is_dir():
            return candidato
    return BASE_DIR / "Planilhas anteriores"


CONFIGURACOES_DIR: Path = resolve_configuracoes_dir()
PLANILHAS_ANTERIORES_DIR: Path = resolve_planilhas_anteriores_dir()


def _resolve_subpasta(subpasta_nome: str, env_var: str | None = None) -> Path:
    """Busca subpasta operacional com suporte a env vars, pasta Configurações ou raiz.

    A hierarquia de resolução prioriza:
    1. Variável de ambiente informada em `env_var`.
    2. Subpasta dentro de `Configurações/` (padrão de produção da versão portátil).
    3. Subpasta na raiz do projeto `BASE_DIR/` (padrão de desenvolvimento).
    4. Fallback padrão: dentro de `Configurações/` se esta existir, senão na raiz.

    Args:
        subpasta_nome: Nome canônico da subpasta desejada (ex: 'database', 'templates').
        env_var: Nome opcional da variável de ambiente para override.

    Returns:
        Path: Caminho absoluto resolvido para a subpasta operacional.

    Complexity:
        Tempo: O(1) - checagens de existência de diretório no sistema de arquivos.
        Espaço: O(1).
    """
    if env_var:
        val = os.getenv(env_var)
        if val:
            return Path(val).resolve()

    dentro_cfg = CONFIGURACOES_DIR / subpasta_nome
    if dentro_cfg.exists():
        return dentro_cfg
    dentro_raiz = BASE_DIR / subpasta_nome
    if dentro_raiz.exists():
        return dentro_raiz
    if CONFIGURACOES_DIR.exists():
        return dentro_cfg
    return dentro_raiz


# ==============================================================================
# SUBPASTAS OPERACIONAIS DO SISTEMA
# ==============================================================================
DATABASE_DIR: Path = _resolve_subpasta("database", "VRE_DATABASE_DIR")
TEMPLATES_DIR: Path = _resolve_subpasta("templates", "VRE_TEMPLATES_DIR")
ENTRADA_CSV_DIR: Path = _resolve_subpasta("entrada_csv", "VRE_ENTRADA_CSV_DIR")
PROCESSADOS_CSV_DIR: Path = ENTRADA_CSV_DIR / "processados"
DOCUMENTACAO_DIR: Path = _resolve_subpasta("docs", "VRE_DOCS_DIR")
PLANILHAS_TRABALHO_DIR: Path = BASE_DIR / "planilhas_trabalho"


# ==============================================================================
# ARQUIVOS DE BANCO DE DADOS SQLITE
# ==============================================================================
DEFAULT_TRIAGEM_DB_PATH: Path = DATABASE_DIR / "triagem_vre.db"
DEFAULT_CEP_DB_PATH: Path = DATABASE_DIR / "cep_areas.db"


def resolve_triagem_db_path() -> Path:
    """Localiza o banco central triagem_vre.db respeitando a ordem de busca.

    Prioridade de resolução:
    1. Variável de ambiente `VRE_TRIAGEM_DB_PATH`.
    2. `DATABASE_DIR/triagem_vre.db`.
    3. `Configurações/database/triagem_vre.db`.
    4. `Configurações/triagem_vre.db`.
    5. `BASE_DIR/database/triagem_vre.db`.
    6. `BASE_DIR/triagem_vre.db`.
    7. Padrão: `DEFAULT_TRIAGEM_DB_PATH`.

    Returns:
        Path: Caminho absoluto para o banco central de triagem.

    Complexity:
        Tempo: O(1) - checagens locais de arquivos existentes.
        Espaço: O(1).
    """
    env_db = os.getenv("VRE_TRIAGEM_DB_PATH")
    if env_db:
        return Path(env_db).resolve()

    candidatos = [
        DATABASE_DIR / "triagem_vre.db",
        CONFIGURACOES_DIR / "database" / "triagem_vre.db",
        CONFIGURACOES_DIR / "triagem_vre.db",
        BASE_DIR / "database" / "triagem_vre.db",
        BASE_DIR / "triagem_vre.db",
    ]
    for c in candidatos:
        if c.exists():
            return c
    return DEFAULT_TRIAGEM_DB_PATH


def resolve_cep_db_path() -> Path:
    """Localiza o banco satélite cep_areas.db respeitando a ordem de busca.

    Prioridade de resolução:
    1. Variável de ambiente `VRE_CEP_DB_PATH`.
    2. `DATABASE_DIR/cep_areas.db`.
    3. `Configurações/database/cep_areas.db`.
    4. `Configurações/cep_areas.db`.
    5. `BASE_DIR/database/cep_areas.db`.
    6. `BASE_DIR/cep_areas.db`.
    7. Padrão: `DEFAULT_CEP_DB_PATH`.

    Returns:
        Path: Caminho absoluto para o banco de correlações CEP x Área.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    env_db = os.getenv("VRE_CEP_DB_PATH")
    if env_db:
        return Path(env_db).resolve()

    candidatos = [
        DATABASE_DIR / "cep_areas.db",
        CONFIGURACOES_DIR / "database" / "cep_areas.db",
        CONFIGURACOES_DIR / "cep_areas.db",
        BASE_DIR / "database" / "cep_areas.db",
        BASE_DIR / "cep_areas.db",
    ]
    for c in candidatos:
        if c.exists():
            return c
    return DEFAULT_CEP_DB_PATH


TRIAGEM_DB_PATH: Path = resolve_triagem_db_path()
CEP_DB_PATH: Path = resolve_cep_db_path()

# Arquivos de schema DDL oficiais
SCHEMA_SQL_PATH: Path = DATABASE_DIR / "schema.sql"
SCHEMA_CEP_SQL_PATH: Path = DATABASE_DIR / "schema_cep.sql"

# Schemas embutidos de fallback para empacotamento PyInstaller
BUNDLE_SCHEMA_SQL_PATH: Path = BUNDLE_DIR / "database" / "schema.sql"
BUNDLE_SCHEMA_CEP_SQL_PATH: Path = BUNDLE_DIR / "database" / "schema_cep.sql"

# Arquivos de template oficial do Excel
TEMPLATE_XLSX_PATH: Path = TEMPLATES_DIR / "template_triagem.xlsx"
BUNDLE_TEMPLATE_XLSX_PATH: Path = BUNDLE_DIR / "templates" / "template_triagem.xlsx"


def resolve_output_xlsx_path() -> Path:
    """Retorna o caminho canônico do arquivo ativo 'Planilha_Controle_VRE_Ativa.xlsx'.

    A planilha oficial deve ficar na raiz do projeto (`BASE_DIR`), local de fácil
    acesso para os fiscais e operadores. Permite override por `VRE_OUTPUT_XLSX_PATH`.

    Returns:
        Path: Caminho do arquivo Excel de saída em produção.

    Complexity:
        Tempo: O(1) - leitura de variável de ambiente ou resolução de Path.
        Espaço: O(1).
    """
    env_out = os.getenv("VRE_OUTPUT_XLSX_PATH")
    if env_out:
        return Path(env_out).resolve()
    return BASE_DIR / "Planilha_Controle_VRE_Ativa.xlsx"


OUTPUT_XLSX_PATH: Path = resolve_output_xlsx_path()


# ==============================================================================
# ENCODINGS E DELIMITADORES DE ARQUIVOS CSV
# ==============================================================================
# O portal VRE do Estado de SP costuma exportar arquivos em Latin-1 (ISO-8859-1 / Windows-1252),
# mas exportações customizadas ou edições no Excel podem gerar UTF-8 com ou sem BOM.
# A ordem prioriza UTF-8 com BOM e UTF-8 para evitar falsos positivos na decodificação Latin-1.
CSV_ENCODINGS: list[str] = ["utf-8-sig", "utf-8", "cp1252", "iso-8859-1", "latin-1"]
CSV_DELIMITER: str = ";"


# ==============================================================================
# PRAGMAS DE ALTA PERFORMANCE PARA SQLITE LOCAL
# ==============================================================================
# WAL: Write-Ahead Logging permite leitura e escrita simultâneas sem contenção de lock.
# synchronous = NORMAL: Garante consistência transacional enquanto minimiza fsync() em disco mecânico ou rede.
# cache_size = -64000: Aloca 64MB de cache em memória RAM para aceleração de buscas no backlog.
# temp_store = MEMORY: Aloca tabelas temporárias e ordenações exclusivamente em memória volátil.
SQLITE_PRAGMAS: list[str] = [
    "PRAGMA journal_mode = WAL;",
    "PRAGMA synchronous = NORMAL;",
    "PRAGMA cache_size = -64000;",
    "PRAGMA temp_store = MEMORY;",
]


# ==============================================================================
# PALETA DE CORES INSTITUCIONAIS SMAURB (FORMATAÇÃO CONDICIONAL E CABEÇALHOS)
# ==============================================================================
# Cores exatas padronizadas pela Inspetoria Fiscal da SMAURB de São José do Rio Preto:
COR_CABECALHO: str = "8EB4E3"          # Azul institucional SMAURB (Prefeitura)
COR_SEM_RESTRICAO: str = "FFFFC000"    # Dourado / Amarelo suave
COR_INDEFERIDO: str = "FFFF0000"       # Vermelho alerta (Postura indeferida)
COR_EM_ANDAMENTO: str = "FF9BBB59"     # Verde suave (Processo sob análise fiscal)
COR_FINALIZADO_STATUS: str = "FFD99694"# Rosa / Vermelho suave (Indeferido/Cancelado/Cassado/Deferido)
COR_NOTURNO_ENTREGUE: str = "FFA6A6A6" # Cinza neutro (Notificação noturna entregue)
COR_FISCAL_ATRIBUIDO: str = "FFB9CDE5" # Azul claro suave (Protocolo com fiscal titular)
COR_NOTURNO_LIVRE: str = "FF8064A2"    # Roxo suave (Atividade noturna livre para fiscalização)
COR_RESTRICAO_LIVRE: str = "FFF79646"  # Laranja suave (Atividade diurna restrita)

# Compatibilidade retroativa com módulos legados
COR_FUNDO_NOTURNO: str = "FF8064A2"
COR_FONTE_NOTURNO: str = "FFFFFF"
COR_FUNDO_RESTRICAO: str = "FFF79646"
COR_FONTE_RESTRICAO: str = "000000"


# ==============================================================================
# FUNÇÕES DE INFRAESTRUTURA E RESOLUÇÃO DE TEMPLATES
# ==============================================================================

def ensure_directories() -> None:
    """Garante a existência física de todos os diretórios do ecossistema VRE.

    Cria automaticamente, de forma idempotente (`exist_ok=True`), as pastas:
    - Configurações, Planilhas anteriores, database, templates, entrada_csv, processados.

    Complexity:
        Tempo: O(1) - chamadas a mkdir para conjunto fixo de 6 diretórios.
        Espaço: O(1).
    """
    for directory in [
        CONFIGURACOES_DIR,
        PLANILHAS_ANTERIORES_DIR,
        DATABASE_DIR,
        TEMPLATES_DIR,
        ENTRADA_CSV_DIR,
        PROCESSADOS_CSV_DIR,
    ]:
        directory.mkdir(parents=True, exist_ok=True)


def resolve_schema_path(is_cep: bool = False) -> Path:
    """Localiza o arquivo DDL de schema SQL correspondente no disco.

    Args:
        is_cep: Se True, busca `schema_cep.sql` (banco satélite de CEPs).
                Se False, busca `schema.sql` (banco central de triagem).

    Returns:
        Path: Caminho para o arquivo DDL do schema.

    Complexity:
        Tempo: O(1) - checagem de lista fixa de candidatos.
        Espaço: O(1).
    """
    nome = "schema_cep.sql" if is_cep else "schema.sql"
    candidatos = [
        CONFIGURACOES_DIR / "database" / nome,
        CONFIGURACOES_DIR / nome,
        DATABASE_DIR / nome,
        BASE_DIR / "database" / nome,
        BUNDLE_DIR / "database" / nome,
    ]
    for c in candidatos:
        if c.exists():
            return c
    return BUNDLE_SCHEMA_CEP_SQL_PATH if is_cep else BUNDLE_SCHEMA_SQL_PATH


def resolve_template_path() -> Path:
    """Localiza o template oficial do Excel, priorizando o arquivo modificado mais recentemente.

    Garante que se o usuário ou administrador atualizou a formatação de colunas
    em `Configurações/templates/template_triagem.xlsx` ou na raiz, a versão
    mais atual seja carregada.

    Returns:
        Path: Caminho absoluto para o template .xlsx a ser clonado.

    Complexity:
        Tempo: O(1) - stat() em até 5 arquivos candidatos.
        Espaço: O(1).
    """
    candidatos = [
        TEMPLATES_DIR / "template_triagem.xlsx",
        BASE_DIR / "templates" / "template_triagem.xlsx",
        CONFIGURACOES_DIR / "templates" / "template_triagem.xlsx",
        CONFIGURACOES_DIR / "template_triagem.xlsx",
        BUNDLE_TEMPLATE_XLSX_PATH,
    ]
    existentes = [c for c in candidatos if c.exists()]
    if not existentes:
        return TEMPLATES_DIR / "template_triagem.xlsx"

    return max(existentes, key=lambda p: p.stat().st_mtime)
