"""Subpacote Administrativo e de Manutenção do Sistema VRE.

Contém serviços de backoffice e a interface gráfica do usuário:
- `archive_service`: Gestão de visibilidade/arquivamento de meses de competência e busca multi-critério de fiscais.
- `main_admin`: Aplicação executável gráfica nativa em Tkinter (`config_vre.exe`).
"""

from __future__ import annotations

from src.admin.archive_service import (
    alternar_status_competencia,
    aplicar_reorganizacao_planilha,
    arquivar_competencia,
    buscar_historico_fiscais,
    definir_status_competencia,
    fazer_backup_planilha_ativa,
    listar_competencias_status,
    listar_fiscais_cadastrados,
    obter_estatisticas_gerais,
    obter_ultimos_atendimentos,
    reativar_competencia,
    sincronizar_competencias,
)

__all__ = [
    "sincronizar_competencias",
    "obter_estatisticas_gerais",
    "listar_competencias_status",
    "definir_status_competencia",
    "alternar_status_competencia",
    "arquivar_competencia",
    "reativar_competencia",
    "fazer_backup_planilha_ativa",
    "aplicar_reorganizacao_planilha",
    "listar_fiscais_cadastrados",
    "obter_ultimos_atendimentos",
    "buscar_historico_fiscais",
]
