"""Módulo de Gravação e Otimização de Planilhas Excel com OpenPyXL.

Gera o arquivo oficial `Planilha_Controle_VRE_Ativa.xlsx` reproduzindo fielmente:
- Visual da planilha de trabalho do exemplo oficial da Inspetoria Fiscal
- Ordem e dimensões exatas das 31 colunas (A a AE)
- Injeção de Plain Values (texto puro estático, sem fórmulas lentas)
- As 12 regras de formatação condicional na exata ordem institucional
- Congelamento de painel em A3 e auto-filtro A2:AE
- Cópia de backup preventivo na pasta `/Planilhas anteriores`
"""

from __future__ import annotations

import datetime
import json
import shutil
from copy import copy
from pathlib import Path
from typing import Any, Sequence

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
from src.core.normalizer import montar_endereco_completo

# Definição oficial das 32 colunas da planilha de trabalho ativa (31 originais + nova coluna única de Restrição)
COLUNAS_CABECALHO = [
    ("Número ordem ", 17.43),
    ("Status", 15.57),
    ("Fiscal", 15.57),
    ("Tipo de Serviço", 18.29),
    ("Restrição", 16.0),          # Nova coluna única de classificação: 'Sem Restrição', 'Com Restrição', 'Noturno'
    ("Data", 15.57),
    ("IPTU", 12.43),
    ("ÁREA", 20.0),
    ("Endereço", 85.14),
    ("Protocolo Existente", 55.0),
    ("ProtocoloSigm / OS", 21.0),
    ("ProtocoloSil", 11.29),
    ("ProtocoloLicenca", 16.57),
    ("Data", 10.71),
    ("GrauRiscoGeral", 7.0),
    ("SituacaoPostura", 13.14),
    ("InscricaoMunicipal", 21.0),
    ("Cnpj", 18.43),
    ("NomeEmpresarial", 48.0),
    ("Telefone", 11.57),
    ("E-mail", 17.43),
    ("TipoEndereco", 5.71),
    ("Logradouro", 27.86),
    ("Numero", 8.0),
    ("Complemento", 12.43),
    ("Bairro", 56.14),
    ("Cep", 11.57),
    ("FinalidadeEndereco", 17.86),
    ("CnaesLicenciadas", 50.71),
    ("CnaesNaoLicenciadas", 11.57),
    ("Iptu", 13.0),
    ("CEP4", 13.0),
]

# Regras de Formatação Condicional na EXATA ordem e prioridade institucional
REGRAS_CONDICIONAIS_ORDEM = [
    ('$D3="Sem Restrição Livre"', COR_SEM_RESTRICAO),
    ('$E3="Sem Restrição"', COR_SEM_RESTRICAO),
    ('$P3="INDEFERIDO"', COR_INDEFERIDO),
    ('$P3="CANCELADO"', COR_INDEFERIDO),
    ('$D3="Noturno entregue"', COR_NOTURNO_ENTREGUE),
    ('$D3="Noturno Livre"', COR_NOTURNO_LIVRE),
    ('$E3="Noturno"', COR_NOTURNO_LIVRE),
    ('$B3="Em Andamento"', COR_EM_ANDAMENTO),
    ('$B3="Indeferido"', COR_FINALIZADO_STATUS),
    ('$B3="Cancelado"', COR_FINALIZADO_STATUS),
    ('$B3="Cassado"', COR_FINALIZADO_STATUS),
    ('$B3="Deferido"', COR_FINALIZADO_STATUS),
    ('$D3="Com Restrição Livre"', COR_RESTRICAO_LIVRE),
    ('$E3="Com Restrição"', COR_RESTRICAO_LIVRE),
    ('$C3<>""', COR_FISCAL_ATRIBUIDO),
]


def aplicar_regras_condicionais(ws, sqref: str) -> None:
    """Aplica as regras de formatação condicional na exata ordem institucional do exemplo."""
    for formula, cor in REGRAS_CONDICIONAIS_ORDEM:
        fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
        rule = FormulaRule(formula=[formula], fill=fill)
        ws.conditional_formatting.add(sqref, rule)


def criar_template_padrao(caminho_template: Path | str | None = None, forcar: bool = False) -> Path:
    """Cria e formata o arquivo template_triagem.xlsx caso ele não exista."""
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

    Garante velocidade máxima e zero travamento injetando apenas Plain Values
    e aplicando as 12 regras de formatação condicional na exata ordem do exemplo.
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

        # Classificação da nova coluna única 'Restrição'
        flag_r = int(reg.get("flag_restricao") or 0)
        flag_n = int(reg.get("flag_noturno") or 0)
        if flag_n:
            classificacao_restricao = "Noturno"
        elif flag_r:
            classificacao_restricao = "Com Restrição"
        else:
            classificacao_restricao = "Sem Restrição"

        # Tipo de Serviço sugerido (mantém valor existente ou deduz)
        tipo_servico = reg.get("tipo_servico")
        if not tipo_servico:
            if flag_n:
                tipo_servico = "Noturno Livre"
            elif flag_r:
                tipo_servico = "Com Restrição Livre"
            else:
                tipo_servico = "Sem Restrição Livre"

        # Endereço estático limpo (Plain Values)
        tipo_end = extras.get("TipoEndereco", "")
        logr = extras.get("Logradouro", "")
        num = extras.get("Numero", "")
        comp = extras.get("Complemento", "")
        bairro = reg.get("bairro_original", "") or extras.get("Bairro", "")
        endereco_calculado = montar_endereco_completo(tipo_end, logr, num, comp, bairro)

        # CEP e CEP4 (5 dígitos)
        cep_bruto = str(extras.get("Cep", "") or "")
        cep4 = cep_bruto[:5] if len(cep_bruto) >= 5 else cep_bruto

        # Higienização de Status: não deve ser escrito não distribuído e/ou distribuído, apenas deixe vazio
        status_raw = str(reg.get("status") or "").strip()
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

        # Data do VRE (usar data do VRE não a colocada)
        data_vre = extras.get("Data", "") or reg.get("data_importacao", "")

        # Mapeamento oficial das 32 colunas
        valores_linha = [
            seq_num,                                                  # Col 1 (A): Número de ordem
            status_exibicao,                                          # Col 2 (B): Status (vazio para pendente/distribuído)
            reg.get("fiscal_atribuido") or "",                        # Col 3 (C): Fiscal
            tipo_servico,                                             # Col 4 (D): Tipo de Serviço
            classificacao_restricao,                                  # Col 5 (E): Restrição (NOVA COLUNA ÚNICA)
            data_vre,                                                 # Col 6 (F): Data do VRE
            extras.get("Iptu", "") or "",                             # Col 7 (G): IPTU
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
            extras.get("Iptu", ""),                                   # Col 31 (AE): Iptu
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

    # Sincroniza saída e template com a pasta fbuild se existir
    try:
        fbuild_dir = BASE_DIR / "fbuild"
        if fbuild_dir.exists():
            fbuild_saida = fbuild_dir / "Planilha_Controle_VRE_Ativa.xlsx"
            if dest_path.resolve() != fbuild_saida.resolve():
                shutil.copy2(str(dest_path), str(fbuild_saida))
            fbuild_tmpl = fbuild_dir / "Configurações" / "templates" / "template_triagem.xlsx"
            if template_path.resolve() != fbuild_tmpl.resolve() and template_path.exists():
                fbuild_tmpl.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(template_path), str(fbuild_tmpl))
    except Exception:
        pass

    return dest_path

