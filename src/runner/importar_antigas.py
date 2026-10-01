"""Script dedicado para Carga e Importação Pontual de Planilhas Históricas (SMAURB).

Lê planilhas de trabalho anteriores (meses ou anos passados) depositadas
em 'Planilhas anteriores/' ou na raiz, higieniza linhas desorganizadas,
descarta duplicatas e alimenta:
1. `historico_fiscais` (triagem_vre.db)
2. `cep_mapeamento` (cep_areas.db)

Executado sob demanda via 'Configurações/Importar_Planilhas_Antigas.bat'.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

# Adiciona raiz do projeto ao path de módulos
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.config import BASE_DIR, PLANILHAS_ANTERIORES_DIR
from src.core.historico_importer import (
    escanear_e_importar_planilhas_antigas,
    importar_planilha_historico,
)


def main() -> None:
    print("=" * 78)
    print(" SMAURB / INSPETORIA FISCAL - CARGA DE PLANILHAS HISTÓRICAS")
    print(" Mapeamento Automático de Fiscais e CEPs x Áreas")
    print("=" * 78)
    print()

    print(f"[*] Diretório base: {BASE_DIR}")
    print(f"[*] Pasta de planilhas anteriores: {PLANILHAS_ANTERIORES_DIR}")
    print("[*] Procurando planilhas Excel (.xlsx, .xls)...")
    print("-" * 78)

    # Permite passar caminho de arquivo específico por argumento (ex: drag & drop sobre o .bat)
    caminho_especifico = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else None

    if caminho_especifico:
        p = Path(caminho_especifico).resolve()
        if not p.exists():
            print(f"[ERRO] Arquivo especificado não encontrado: {caminho_especifico}")
            sys.exit(1)
        arquivos = [p]
    else:
        arquivos = []
        if PLANILHAS_ANTERIORES_DIR.exists():
            for ext in ["*.xlsx", "*.xls"]:
                arquivos.extend(sorted(PLANILHAS_ANTERIORES_DIR.glob(ext)))
        if BASE_DIR.exists():
            for ext in ["*.xlsx", "*.xls"]:
                for f in sorted(BASE_DIR.glob(ext)):
                    nome_up = f.name.upper()
                    if "ATIVA" not in nome_up and "TEMPLATE" not in nome_up and f not in arquivos:
                        arquivos.append(f)

    if not arquivos:
        print()
        print(" [!] NENHUMA PLANILHA ANTIGA FOI ENCONTRADA.")
        print()
        print(" Instruções de uso:")
        print(" 1. Coloque seus arquivos .xlsx de anos ou meses anteriores na pasta:")
        print(f"    -> {PLANILHAS_ANTERIORES_DIR}")
        print(" 2. Ou arraste um arquivo de planilha por cima deste atalho.")
        print(" 3. Execute este arquivo novamente.")
        print()
        print("=" * 78)
        sys.exit(0)

    print(f"[*] Localizada(s) {len(arquivos)} planilha(s) para importação.")
    print()

    total_linhas = 0
    total_duplicatas = 0
    total_fiscais = 0
    total_ceps = 0

    t_inicio = time.time()

    for idx, arq in enumerate(arquivos, start=1):
        print(f"[{idx}/{len(arquivos)}] Lendo: {arq.name}...")
        try:
            res = importar_planilha_historico(arq, forcar=True)
            if res.get("status") == "ERRO_LEITURA":
                print(f"    [!] Erro ao abrir planilha: {res.get('mensagem')}")
                continue

            lidas = res.get("linhas_lidas", 0)
            dup = res.get("duplicatas_descartadas", 0)
            fisc = res.get("fiscais_inseridos", 0)
            ceps = res.get("ceps_inseridos", 0)

            total_linhas += lidas
            total_duplicatas += dup
            total_fiscais += fisc
            total_ceps += ceps

            print(f"    [✓] Linhas analisadas: {lidas}")
            print(f"    [✓] Duplicatas/inválidas eliminadas: {dup}")
            print(f"    [✓] Atendimentos de fiscais consolidados: +{fisc}")
            print(f"    [✓] Correlações de CEP mapeadas: +{ceps}")
            print()
        except Exception as e:
            print(f"    [ERRO] Falha ao processar {arq.name}: {e}")
            print()

    tempo_gasto = time.time() - t_inicio

    print("=" * 78)
    print(" [SUCESSO] Carga de histórico concluída com êxito!")
    print(f" Tempo de processamento: {tempo_gasto:.1f} segundos")
    print(f" - Planilhas processadas:              {len(arquivos)}")
    print(f" - Total de linhas brutas lidas:       {total_linhas}")
    print(f" - Linhas duplicadas/descartadas:      {total_duplicatas}")
    print(f" - Atendimentos únicos de fiscais:     {total_fiscais}")
    print(f" - Correlações de CEP x Área salvas:   {total_ceps}")
    print("=" * 78)
    print(" O banco de dados agora reconhece esses fiscais na triagem automática.")
    print("=" * 78)


if __name__ == "__main__":
    main()
