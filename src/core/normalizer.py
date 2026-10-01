"""Módulo de Sanitização e Normalização de Dados.

Implementa algoritmos rigorosos para higienização e formatação de:
- CNPJs (máscaras e dígitos puros com tratamento de float do Excel)
- CEPs (8 dígitos e máscara XXXXX-XXX)
- Bairros (decomposição NFKD, remoção de acentos e expansão de abreviações)
- CNAEs (extração de códigos, fatiamento para 7 dígitos e máscara oficial)
- Endereços completos harmonizados
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Sequence


def normalizar_cnpj(cnpj_bruto: Any) -> str:
    """Higieniza e normaliza CNPJ para uma string numérica de 14 dígitos.

    Retorna string vazia se o valor for nulo ou inválido.
    Suporta floats do Excel, strings com .0 e casos onde zeros à esquerda foram perdidos (11 a 14 dígitos).
    """
    if not cnpj_bruto:
        return ""
    if isinstance(cnpj_bruto, float):
        cnpj_bruto = int(cnpj_bruto)
    s = str(cnpj_bruto).strip()
    if s.endswith(".0"):
        s = s[:-2]
    apenas_digitos = re.sub(r"\D", "", s)
    if len(apenas_digitos) not in range(11, 15):
        return ""
    return apenas_digitos.zfill(14)


def formatar_cnpj(cnpj_limpo: str | None) -> str:
    """Aplica a máscara oficial XX.XXX.XXX/XXXX-XX em um CNPJ limpo de 14 dígitos."""
    if not cnpj_limpo:
        return ""
    digitos = normalizar_cnpj(cnpj_limpo)
    if len(digitos) != 14:
        return cnpj_limpo
    return f"{digitos[0:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:14]}"


def normalizar_cep(cep_bruto: Any) -> str:
    """Higieniza e normaliza CEP para uma string numérica de 8 dígitos.

    Retorna string vazia se o valor for nulo ou inválido (ex: menos de 7 dígitos).
    Suporta floats do Excel e strings com .0.
    """
    if not cep_bruto:
        return ""
    if isinstance(cep_bruto, float):
        cep_bruto = int(cep_bruto)
    s = str(cep_bruto).strip()
    if s.endswith(".0"):
        s = s[:-2]
    apenas_digitos = re.sub(r"\D", "", s)
    # Um CEP válido possui 8 dígitos (ou 7 caso tenha perdido zero à esquerda no Excel)
    if len(apenas_digitos) not in (7, 8):
        return ""
    return apenas_digitos.zfill(8)


def formatar_cep(cep_limpo: str | None) -> str:
    """Aplica a máscara oficial XXXXX-XXX em um CEP limpo de 8 dígitos."""
    if not cep_limpo:
        return ""
    digitos = normalizar_cep(cep_limpo)
    if len(digitos) != 8:
        return cep_limpo
    return f"{digitos[0:5]}-{digitos[5:8]}"


def normalizar_bairro(nome_bairro: str | None) -> str:
    """Normaliza nome de bairro para casamento exato em dicionários e tabelas.

    Passos:
    1. Decomposição Unicode NFKD para remoção de acentos e diacríticos.
    2. Conversão para maiúsculas e substituição de pontuações por espaço.
    3. Expansão de abreviações consagradas de Rio Preto (JD, PQ, VL, RES, etc.).
       NOTA: Letras isoladas como 'V' ou 'J' NÃO são abreviadas para evitar
       corrupção de numerais romanos (ex: 'DAMHA V', 'LIBERDADE V') ou setores.
    4. Remoção de espaços múltiplos e trims.
    """
    if not nome_bairro:
        return ""

    texto_str = str(nome_bairro).strip()
    if not texto_str:
        return ""

    # 1. Decomposição e remoção de acentos
    nfkd = unicodedata.normalize("NFKD", texto_str)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))

    # 2. Maiúsculas e higienização de pontuação
    texto = sem_acento.upper()
    texto = re.sub(r"[\.,\-_\/\\;\(\)]", " ", texto)

    # 3. Dicionário de expansão de abreviações consagradas
    substituicoes = {
        r"\bJD\b": "JARDIM",
        r"\bPQ\b": "PARQUE",
        r"\bPRQ\b": "PARQUE",
        r"\bVL\b": "VILA",
        r"\bRES\b": "RESIDENCIAL",
        r"\bCONJ\b": "CONJUNTO",
        r"\bCJ\b": "CONJUNTO",
        r"\bHAB\b": "HABITACIONAL",
        r"\bCH\b": "CHACARA",
        r"\bCHAC\b": "CHACARA",
        r"\bEST\b": "ESTANCIA",
        r"\bDT\b": "DISTRITO",
        r"\bDIST\b": "DISTRITO",
        r"\bLOT\b": "LOTEAMENTO",
        r"\bRET\b": "RETALHAMENTO",
        r"\bFAZ\b": "FAZENDA",
        r"\bCOND\b": "CONDOMINIO",
        r"\bIND\b": "INDUSTRIAL",
        r"\bREC\b": "RECANTO",
    }

    for padrao, sub in substituicoes.items():
        texto = re.sub(padrao, sub, texto)

    # 4. Colapso de múltiplos espaços em um único
    return re.sub(r"\s+", " ", texto).strip()


def extrair_cnaes(cnaes_str: str | None) -> list[str]:
    """Extrai lista de códigos CNAE fatiados em 7 dígitos limpos.

    No CSV do VRE, os CNAEs vêm em formato de 9 dígitos separados por vírgula.
    Esta função higieniza cada token, extrai os 7 primeiros dígitos e preserva
    a ordem sem duplicatas.
    """
    if not cnaes_str:
        return []

    tokens = re.split(r"[,;]+", str(cnaes_str))
    resultado: list[str] = []
    vistos: set[str] = set()

    for token in tokens:
        digitos = re.sub(r"\D", "", token.strip())
        if not digitos:
            continue
        # Fatiamento dos 7 primeiros dígitos (classe/subclasse base)
        cnae_7 = digitos[:7] if len(digitos) >= 7 else digitos
        if len(cnae_7) == 7 and cnae_7 not in vistos:
            vistos.add(cnae_7)
            resultado.append(cnae_7)

    return resultado


def formatar_cnae(cnae_7: str | None) -> str:
    """Aplica a máscara oficial XXXX-X/XX em um código CNAE de 7 dígitos.

    Exemplo: '4687701' -> '4687-7/01'.
    """
    if not cnae_7:
        return ""
    digitos = re.sub(r"\D", "", str(cnae_7).strip())
    if len(digitos) != 7:
        return cnae_7
    return f"{digitos[0:4]}-{digitos[4:5]}/{digitos[5:7]}"


def montar_endereco_completo(
    tipo: str | None = None,
    logradouro: str | None = None,
    numero: str | None = None,
    complemento: str | None = None,
    bairro: str | None = None,
) -> str:
    """Monta uma string amigável de endereço sem pontuações órfãs ou ':'.

    Exemplo: 'RUA AIMORES, 287 - VILA ANCHIETA'
    """
    partes_logradouro = []
    if tipo and str(tipo).strip():
        partes_logradouro.append(str(tipo).strip())
    if logradouro and str(logradouro).strip():
        partes_logradouro.append(str(logradouro).strip())

    base = " ".join(partes_logradouro).strip()

    num = str(numero).strip() if numero is not None else ""
    if num:
        base = f"{base}, {num}" if base else num

    comp = str(complemento).strip() if complemento is not None else ""
    if comp and comp not in (":", "-", "--"):
        base = f"{base} {comp}" if base else comp

    bai = str(bairro).strip() if bairro is not None else ""
    if bai:
        base = f"{base} - {bai}" if base else bai

    return re.sub(r"\s+", " ", base).strip()
