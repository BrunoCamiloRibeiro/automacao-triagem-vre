"""Subpacote Core com serviços fundamentais, persistência e regras de negócio."""

from src.core.config import (
    BASE_DIR,
    CEP_DB_PATH,
    ENTRADA_CSV_DIR,
    OUTPUT_XLSX_PATH,
    PLANILHAS_ANTERIORES_DIR,
    TRIAGEM_DB_PATH,
)
from src.core.database import connect_db, db_transaction, get_cep_db, get_triagem_db
from src.core.normalizer import (
    extrair_cnaes,
    formatar_cep,
    formatar_cnae,
    formatar_cnpj,
    montar_endereco_completo,
    normalizar_bairro,
    normalizar_cep,
    normalizar_cnpj,
)

__all__ = [
    "BASE_DIR",
    "TRIAGEM_DB_PATH",
    "CEP_DB_PATH",
    "OUTPUT_XLSX_PATH",
    "ENTRADA_CSV_DIR",
    "PLANILHAS_ANTERIORES_DIR",
    "connect_db",
    "get_triagem_db",
    "get_cep_db",
    "db_transaction",
    "normalizar_cnpj",
    "formatar_cnpj",
    "normalizar_cep",
    "formatar_cep",
    "normalizar_bairro",
    "extrair_cnaes",
    "formatar_cnae",
    "montar_endereco_completo",
]
