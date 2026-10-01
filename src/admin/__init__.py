"""Subpacote Administrativo com interface gráfica Tkinter e serviços de gestão."""

from src.admin.archive_service import (
    alternar_status_competencia,
    aplicar_reorganizacao_planilha,
    arquivar_competencia,
    buscar_historico_fiscais,
    listar_competencias_status,
    obter_estatisticas_gerais,
    reativar_competencia,
)

__all__ = [
    "alternar_status_competencia",
    "aplicar_reorganizacao_planilha",
    "arquivar_competencia",
    "buscar_historico_fiscais",
    "listar_competencias_status",
    "obter_estatisticas_gerais",
    "reativar_competencia",
]
