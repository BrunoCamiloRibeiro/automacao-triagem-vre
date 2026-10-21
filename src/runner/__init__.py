"""Subpacote Runner com executáveis operacionais CLI e rotinas de pipeline.

Módulos disponíveis:
- `main_seco`: Pipeline headless principal de triagem e processamento diário de CSVs.
- `importar_antigas`: Utilitário de linha de comando para carga pontual de planilhas históricas.
"""

from __future__ import annotations

from src.runner.main_seco import processar_fluxo_seco

__all__ = ["processar_fluxo_seco"]
