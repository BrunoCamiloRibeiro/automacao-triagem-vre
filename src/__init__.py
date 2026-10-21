"""Pacote Principal do Sistema de Automação e Triagem VRE - SMAURB.

Sistema corporativo de alta performance para saneamento, triagem regulatória,
mapeamento fiscal e geração de planilhas de trabalho para a Secretaria Municipal
de Habitação e Urbanismo (SMAURB) / Inspetoria Fiscal de São José do Rio Preto - SP.

Módulos Arquiteturais:
- `src.core`: Serviços de baixo nível, persistência SQLite, normalização e manipulação do Excel.
- `src.runner`: Executáveis operacionais de linha de comando (CLI) e rotinas headless.
- `src.admin`: Serviços de arquivamento, métricas e interface gráfica administrativa (GUI).
"""

from __future__ import annotations

__version__ = "2.2.0"
__author__ = "SMAURB / Inspetoria Fiscal - São José do Rio Preto"
__license__ = "Proprietário / Uso Institucional"
