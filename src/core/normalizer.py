"""Módulo de Sanitização e Normalização de Dados.

Implementa algoritmos determinísticos de alta precisão para higienização,
padronização e formatação de entidades de domínio:
- CNPJs: Suporte a strings puras, máscaras, inteiros, floats do Excel (.0) e recuperação de zeros à esquerda (11 a 14 dígitos).
- CEPs: Higienização para 8 dígitos numéricos, recuperação de zeros à esquerda (7 ou 8 dígitos) e máscara `XXXXX-XXX`.
- Bairros: Decomposição Unicode NFKD, remoção de diacríticos, expansão de abreviações consagradas de São José do Rio Preto e preservação de numerais romanos.
- CNAEs: Extração em lote, fatiamento para 7 dígitos base (classe/subclasse), deduplicação preservando ordem e máscara `XXXX-X/XX`.
- Endereços: Harmonização concatenada de tipo de logradouro, nome, número, complemento e bairro sem pontuações órfãs.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Sequence

# ==============================================================================
# EXPRESSÕES REGULARES PRÉ-COMPILADAS (OTIMIZAÇÃO DE PERFORMANCE)
# ==============================================================================
# A compilação em nível de módulo elimina o overhead de parsing de regex em loops de milhares de linhas.
_RE_DIGITOS = re.compile(r"\D")
_RE_PONTUACAO_BAIRRO = re.compile(r"[\.,\-_\/\\;\(\)]")
_RE_SPACES = re.compile(r"\s+")
_RE_CNAE_SPLIT = re.compile(r"[,;]+")

# Dicionário pré-compilado de abreviações consagradas de bairros em São José do Rio Preto.
# ATENÇÃO: Letras unitárias como 'V' ou 'J' NÃO são expandidas isoladamente para evitar corromper
# numerais romanos (ex: 'DAMHA V' -> 'DAMHA VILA', 'LIBERDADE IV' -> 'LIBERDADE IVILA').
_RE_BAIRRO_ABREVIACOES: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bJD\b"), "JARDIM"),
    (re.compile(r"\bPQ\b"), "PARQUE"),
    (re.compile(r"\bPRQ\b"), "PARQUE"),
    (re.compile(r"\bVL\b"), "VILA"),
    (re.compile(r"\bRES\b"), "RESIDENCIAL"),
    (re.compile(r"\bCONJ\b"), "CONJUNTO"),
    (re.compile(r"\bCJ\b"), "CONJUNTO"),
    (re.compile(r"\bHAB\b"), "HABITACIONAL"),
    (re.compile(r"\bCH\b"), "CHACARA"),
    (re.compile(r"\bCHAC\b"), "CHACARA"),
    (re.compile(r"\bEST\b"), "ESTANCIA"),
    (re.compile(r"\bDT\b"), "DISTRITO"),
    (re.compile(r"\bDIST\b"), "DISTRITO"),
    (re.compile(r"\bLOT\b"), "LOTEAMENTO"),
    (re.compile(r"\bRET\b"), "RETALHAMENTO"),
    (re.compile(r"\bFAZ\b"), "FAZENDA"),
    (re.compile(r"\bCOND\b"), "CONDOMINIO"),
    (re.compile(r"\bIND\b"), "INDUSTRIAL"),
    (re.compile(r"\bREC\b"), "RECANTO"),
)


def normalizar_cnpj(cnpj_bruto: Any) -> str:
    """Higieniza e normaliza CNPJ para uma string estrita de 14 dígitos numéricos.

    Trata casos patológicos típicos do ecossistema Excel/CSV da fiscalização:
    1. Valores nulos, booleanos ou strings vazias retornam `""`.
    2. Valores float (ex: `44444444000144.0`) convertidos para int antes da extração.
    3. Strings terminadas em `.0` (geradas por leitores de planilha) têm o sufixo removido.
    4. Remoção de todos os caracteres não-dígitos.
    5. Se o comprimento resultante estiver entre 11 e 14 dígitos, preenche com zeros à esquerda
       via `zfill(14)` (recuperando zeros descartados pelo Excel). Caso contrário, considera inválido.

    Args:
        cnpj_bruto: Valor bruto recebido do CSV ou Excel (str, int, float ou None).

    Returns:
        str: String de 14 dígitos numéricos ou string vazia `""` se inválido/nulo.

    Complexity:
        Tempo: O(K) onde K é o comprimento do texto de entrada (~14 a 20 caracteres).
        Espaço: O(1) - alocação de pequenas strings transitórias.
    """
    if not cnpj_bruto:
        return ""
    if isinstance(cnpj_bruto, float):
        cnpj_bruto = int(cnpj_bruto)
    s = str(cnpj_bruto).strip()
    if s.endswith(".0"):
        s = s[:-2]
    apenas_digitos = _RE_DIGITOS.sub("", s)
    if len(apenas_digitos) not in range(11, 15):
        return ""
    return apenas_digitos.zfill(14)


def formatar_cnpj(cnpj_limpo: str | None) -> str:
    """Aplica a máscara oficial XX.XXX.XXX/XXXX-XX sobre um CNPJ limpo de 14 dígitos.

    Se a entrada contiver formatação parcial ou não normalizar para 14 dígitos,
    retorna a própria entrada original sem corrompê-la.

    Args:
        cnpj_limpo: String de CNPJ puro de 14 dígitos ou com formatação preexistente.

    Returns:
        str: CNPJ formatado no padrão `00.000.000/0000-00` ou string vazia se nulo.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if not cnpj_limpo:
        return ""
    digitos = normalizar_cnpj(cnpj_limpo)
    if len(digitos) != 14:
        return str(cnpj_limpo).strip()
    return f"{digitos[0:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:14]}"


def normalizar_cep(cep_bruto: Any) -> str:
    """Higieniza e normaliza CEP para uma string estrita de 8 dígitos numéricos.

    Trata casos de perda de zero inicial no Excel (onde CEPs como `01505-090` viram `1505090` de 7 dígitos),
    preenchendo com zeros à esquerda via `zfill(8)`. Suporta floats e strings com sufixo `.0`.

    Args:
        cep_bruto: CEP bruto recebido do CSV ou planilha (str, int, float ou None).

    Returns:
        str: String de 8 dígitos numéricos puros ou string vazia `""` se inválido.

    Complexity:
        Tempo: O(K) onde K é o tamanho da string de entrada (~8 a 12 caracteres).
        Espaço: O(1).
    """
    if not cep_bruto:
        return ""
    if isinstance(cep_bruto, float):
        cep_bruto = int(cep_bruto)
    s = str(cep_bruto).strip()
    if s.endswith(".0"):
        s = s[:-2]
    apenas_digitos = _RE_DIGITOS.sub("", s)
    # Um CEP válido possui 8 dígitos (ou 7 caso tenha perdido zero à esquerda no Excel)
    if len(apenas_digitos) not in (7, 8):
        return ""
    return apenas_digitos.zfill(8)


def formatar_cep(cep_limpo: str | None) -> str:
    """Aplica a máscara oficial XXXXX-XXX em um CEP limpo de 8 dígitos.

    Args:
        cep_limpo: String de CEP de 8 dígitos numéricos ou valor bruto.

    Returns:
        str: CEP com máscara postal (ex: `15050-090`) ou o valor de entrada se inválido.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if not cep_limpo:
        return ""
    digitos = normalizar_cep(cep_limpo)
    if len(digitos) != 8:
        return str(cep_limpo).strip()
    return f"{digitos[0:5]}-{digitos[5:8]}"


def normalizar_bairro(nome_bairro: str | None) -> str:
    """Normaliza o nome de bairro para casamento exato em dicionários e tabelas SQL.

    Pipeline de Processamento:
    1. Verificação de nulidade e sanitização básica de espaços em branco.
    2. Decomposição Unicode NFKD para separar caracteres base de diacríticos combinantes
       (ex: 'SÃO JOSÉ' -> 'SAO JOSE').
    3. Conversão para caixa alta (UPPERCASE) e substituição de pontuações por espaços.
    4. Aplicação do dicionário pré-compilado de expansão de abreviações consagradas
       (ex: 'JD' -> 'JARDIM', 'RES' -> 'RESIDENCIAL', 'VL' -> 'VILA').
    5. Colapso de múltiplos espaços consecutivos em um espaço único e trim.

    Args:
        nome_bairro: Nome original do bairro extraído do cadastro ou planilha.

    Returns:
        str: Nome padronizado em caixa alta, sem acentos e com abreviações expandidas.

    Complexity:
        Tempo: O(M) onde M é o número de caracteres do nome do bairro.
        Espaço: O(M) para as transformações de string intermediárias.
    """
    if not nome_bairro:
        return ""

    texto_str = str(nome_bairro).strip()
    if not texto_str:
        return ""

    # 1. Decomposição e remoção de acentos/diacríticos
    nfkd = unicodedata.normalize("NFKD", texto_str)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))

    # 2. Maiúsculas e higienização de pontuação
    texto = sem_acento.upper()
    texto = _RE_PONTUACAO_BAIRRO.sub(" ", texto)

    # 3. Dicionário de expansão de abreviações consagradas
    for padrao, sub in _RE_BAIRRO_ABREVIACOES:
        texto = padrao.sub(sub, texto)

    # 4. Colapso de múltiplos espaços em um único
    return _RE_SPACES.sub(" ", texto).strip()


def extrair_cnaes(cnaes_str: str | None) -> list[str]:
    """Extrai lista de códigos CNAE fatiados em 7 dígitos limpos.

    No arquivo CSV exportado pelo sistema VRE de São Paulo, os CNAEs vêm em formato de
    9 dígitos numéricos (subclasse completa) separados por vírgula ou ponto-e-vírgula.
    Esta função higieniza cada token, extrai os 7 primeiros dígitos base (classe oficial)
    e preserva a ordem original de inserção, eliminando duplicatas.

    Args:
        cnaes_str: String concatenada com um ou múltiplos CNAEs (ex: '561120101, 471130200').

    Returns:
        list[str]: Lista de códigos CNAE com exatamente 7 dígitos numéricos sem formatação.

    Complexity:
        Tempo: O(T) onde T é o comprimento total da string de CNAEs.
        Espaço: O(N) onde N é o número de códigos CNAE distintos encontrados.
    """
    if not cnaes_str:
        return []

    tokens = _RE_CNAE_SPLIT.split(str(cnaes_str))
    resultado: list[str] = []
    vistos: set[str] = set()

    for token in tokens:
        digitos = _RE_DIGITOS.sub("", token.strip())
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

    Args:
        cnae_7: String de 7 dígitos do código CNAE ou código formatado.

    Returns:
        str: CNAE com pontuação oficial do IBGE ou valor de entrada se não tiver 7 dígitos.

    Complexity:
        Tempo: O(1).
        Espaço: O(1).
    """
    if not cnae_7:
        return ""
    digitos = _RE_DIGITOS.sub("", str(cnae_7).strip())
    if len(digitos) != 7:
        return str(cnae_7).strip()
    return f"{digitos[0:4]}-{digitos[4:5]}/{digitos[5:7]}"


def montar_endereco_completo(
    tipo: str | None = None,
    logradouro: str | None = None,
    numero: str | None = None,
    complemento: str | None = None,
    bairro: str | None = None,
) -> str:
    """Monta uma string amigável de endereço sem pontuações órfãs ou caracteres parasitas.

    Exemplo: 'RUA AIMORES, 287 - VILA ANCHIETA'.
    Elimina símbolos de pontuação isolados como ':', '-', '--' frequentemente encontrados
    nos campos de complemento em cadastros legados da Junta Comercial.

    Args:
        tipo: Tipo do logradouro (ex: 'RUA', 'AV', 'RODOVIA').
        logradouro: Nome do logradouro (ex: 'AIMORES').
        numero: Número predial ou 'S/N' (ex: '287').
        complemento: Informação complementar (ex: 'SALA 10', 'GALPAO A').
        bairro: Nome do bairro (ex: 'VILA ANCHIETA').

    Returns:
        str: String única e harmonizada pronta para exibição nas planilhas de trabalho.

    Complexity:
        Tempo: O(L) onde L é o comprimento total das parcelas de endereço.
        Espaço: O(L).
    """
    partes_logradouro: list[str] = []
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

    return _RE_SPACES.sub(" ", base).strip()
