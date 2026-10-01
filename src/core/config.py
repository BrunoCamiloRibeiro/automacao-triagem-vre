"""Módulo de Configuração Centralizada da Automação VRE.

Define caminhos de diretórios, arquivos de banco de dados, encodings aceitos,
pragmas de performance do SQLite e rotinas de inicialização de estrutura.
Compatível com execução direta em Python, variáveis de ambiente (.env)
e binários empacotados via PyInstaller.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# Determinação dinâmica da pasta base
if getattr(sys, "frozen", False):
    # Execução via binário compilado PyInstaller (.exe)
    _default_base = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", _default_base))
else:
    # Execução normal em ambiente Python (dev ou pacote portátil)
    _default_base = Path(__file__).resolve().parent.parent.parent
    if _default_base.name.lower() in ["configurações", "configuracoes", "config"]:
        _default_base = _default_base.parent
    BUNDLE_DIR = _default_base

# Suporte a override por variável de ambiente
BASE_DIR = Path(os.getenv("VRE_BASE_DIR", str(_default_base))).resolve()


def resolve_configuracoes_dir() -> Path:
    """Localiza o diretório de configurações (/Configurações ou configurado)."""
    env_cfg = os.getenv("VRE_CONFIG_DIR")
    if env_cfg:
        return Path(env_cfg).resolve()

    for nome in ["Configurações", "Configuracoes", "configuracoes", "config"]:
        candidato = BASE_DIR / nome
        if candidato.exists() and candidato.is_dir():
            return candidato
    return BASE_DIR / "Configurações"


def resolve_planilhas_anteriores_dir() -> Path:
    """Localiza o diretório de backup de planilhas anteriores."""
    env_dir = os.getenv("VRE_PLANILHAS_ANTERIORES_DIR")
    if env_dir:
        return Path(env_dir).resolve()

    for nome in ["Planilhas anteriores", "Planilhas Anteriores", "planilhas_anteriores"]:
        candidato = BASE_DIR / nome
        if candidato.exists() and candidato.is_dir():
            return candidato
    return BASE_DIR / "Planilhas anteriores"


CONFIGURACOES_DIR = resolve_configuracoes_dir()
PLANILHAS_ANTERIORES_DIR = resolve_planilhas_anteriores_dir()


def _resolve_subpasta(subpasta_nome: str, env_var: str | None = None) -> Path:
    """Busca subpasta com suporte a env vars, pasta Configurações ou raiz."""
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


# Subpastas operacionais
DATABASE_DIR = _resolve_subpasta("database", "VRE_DATABASE_DIR")
TEMPLATES_DIR = _resolve_subpasta("templates", "VRE_TEMPLATES_DIR")
ENTRADA_CSV_DIR = _resolve_subpasta("entrada_csv", "VRE_ENTRADA_CSV_DIR")
PROCESSADOS_CSV_DIR = ENTRADA_CSV_DIR / "processados"
DOCUMENTACAO_DIR = _resolve_subpasta("docs", "VRE_DOCS_DIR")
PLANILHAS_TRABALHO_DIR = BASE_DIR / "planilhas_trabalho"

# Arquivos de banco de dados
DEFAULT_TRIAGEM_DB_PATH = DATABASE_DIR / "triagem_vre.db"
DEFAULT_CEP_DB_PATH = DATABASE_DIR / "cep_areas.db"


def resolve_triagem_db_path() -> Path:
    """Localiza triagem_vre.db respeitando prioridade de variáveis e pastas."""
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
    """Localiza cep_areas.db respeitando prioridade de variáveis e pastas."""
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


TRIAGEM_DB_PATH = resolve_triagem_db_path()
CEP_DB_PATH = resolve_cep_db_path()

# Arquivos de schema DDL
SCHEMA_SQL_PATH = DATABASE_DIR / "schema.sql"
SCHEMA_CEP_SQL_PATH = DATABASE_DIR / "schema_cep.sql"

# Caminhos de fallback para schemas embutidos no binário
BUNDLE_SCHEMA_SQL_PATH = BUNDLE_DIR / "database" / "schema.sql"
BUNDLE_SCHEMA_CEP_SQL_PATH = BUNDLE_DIR / "database" / "schema_cep.sql"

# Arquivos de templates e saída do Excel
TEMPLATE_XLSX_PATH = TEMPLATES_DIR / "template_triagem.xlsx"
BUNDLE_TEMPLATE_XLSX_PATH = BUNDLE_DIR / "templates" / "template_triagem.xlsx"


def resolve_output_xlsx_path() -> Path:
    """A planilha ativa oficial fica diretamente no Diretório Base (raiz)."""
    env_out = os.getenv("VRE_OUTPUT_XLSX_PATH")
    if env_out:
        return Path(env_out).resolve()
    return BASE_DIR / "Planilha_Controle_VRE_Ativa.xlsx"


OUTPUT_XLSX_PATH = resolve_output_xlsx_path()

# Encodings e delimitadores de CSV aceitos (UTF-8 primeiro para evitar falso positivo do Latin-1)
CSV_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "iso-8859-1", "latin-1"]
CSV_DELIMITER = ";"

# Pragmas de otimização para SQLite local de alta performance
SQLITE_PRAGMAS = [
    "PRAGMA journal_mode = WAL;",
    "PRAGMA synchronous = NORMAL;",
    "PRAGMA cache_size = -64000;",
    "PRAGMA temp_store = MEMORY;",
]

# Estilos Visuais Padrão SMAURB (Formatação Condicional e Cabeçalho)
COR_CABECALHO = "8EB4E3"          # Azul institucional SMAURB
COR_SEM_RESTRICAO = "FFFFC000"    # Dourado / Amarelo
COR_INDEFERIDO = "FFFF0000"       # Vermelho
COR_EM_ANDAMENTO = "FF9BBB59"     # Verde suave
COR_FINALIZADO_STATUS = "FFD99694"# Rosa / Vermelho suave (Indeferido/Cancelado/Cassado/Deferido)
COR_NOTURNO_ENTREGUE = "FFA6A6A6" # Cinza médio
COR_FISCAL_ATRIBUIDO = "FFB9CDE5" # Azul claro suave
COR_NOTURNO_LIVRE = "FF8064A2"    # Roxo suave
COR_RESTRICAO_LIVRE = "FFF79646"  # Laranja suave

# Mantidos para compatibilidade com outros módulos
COR_FUNDO_NOTURNO = "FF8064A2"
COR_FONTE_NOTURNO = "FFFFFF"
COR_FUNDO_RESTRICAO = "FFF79646"
COR_FONTE_RESTRICAO = "000000"


def ensure_directories() -> None:
    """Garante que todos os diretórios do ecossistema existam fisicamente."""
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
    """Localiza o arquivo de schema DDL correspondente."""
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
    """Localiza o template oficial do Excel, priorizando o arquivo modificado mais recentemente."""
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
