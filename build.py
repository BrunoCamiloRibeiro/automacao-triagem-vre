"""Script de Build e Empacotamento Multiplataforma (Windows e Linux).

Executa as 5 etapas de preparação e compilação do sistema VRE:
1. Instalação e verificação de dependências (pip)
2. Inicialização e seeding dos bancos SQLite (triagem_vre.db e cep_areas.db)
3. Compilação do executável operacional headless: triagem_vre.exe
4. Compilação do executável administrativo GUI: config_vre.exe
5. Organização da entrega final na pasta 'fbuild' conforme a estrutura requisitada:
   - fbuild/Planilha_Controle_VRE_Ativa.xlsx
   - fbuild/triagem_vre.exe
   - fbuild/config_vre.exe
   - fbuild/Planilhas anteriores/
   - fbuild/Configurações/ (bancos, templates, schemas, entrada_csv, src)
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

# Fixa o diretório de execução na raiz do projeto
ROOT_DIR = Path(__file__).resolve().parent
os.chdir(ROOT_DIR)


def log_etapa(numero: str, mensagem: str) -> None:
    print(f"\n[{numero}] {mensagem}")


def executar_comando(cmd: list[str], descricao: str) -> None:
    print(f"    -> Executando: {' '.join(cmd)}")
    resultado = subprocess.run(cmd, cwd=ROOT_DIR)
    if resultado.returncode != 0:
        print(f"\n[ERRO] Falha ao executar: {descricao} (código de saída: {resultado.returncode})")
        sys.exit(resultado.returncode)


def main() -> None:
    print("=" * 78)
    print(" SMAURB / INSPETORIA FISCAL - BUILD & EMPACOTAMENTO DA TRIAGEM VRE")
    print("=" * 78)

    req_file = ROOT_DIR / "requirements.txt"

    # 1. Dependências Python
    log_etapa("1/5", "Verificando dependências Python...")
    modulos_necessarios = ["openpyxl", "xlrd", "PyInstaller"]
    faltando = []
    for mod in modulos_necessarios:
        try:
            __import__(mod)
        except ImportError:
            faltando.append(mod)

    if faltando:
        print(f"    -> Instalando dependências pendentes: {faltando}...")
        if req_file.exists():
            res = subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(req_file)], cwd=ROOT_DIR)
            if res.returncode != 0 and sys.platform.startswith("linux"):
                subprocess.run([sys.executable, "-m", "pip", "install", "--break-system-packages", "-r", str(req_file)], cwd=ROOT_DIR)
    else:
        print("    [✓] Todas as dependências (openpyxl, xlrd, PyInstaller) já estão disponíveis.")

    # 2. Inicialização e Seeding dos Bancos
    log_etapa("2/5", "Inicializando bancos SQLite e gerando templates padrão...")
    init_db_script = ROOT_DIR / "database" / "init_db.py"
    if not init_db_script.exists():
        print(f"[ERRO FATAL] Arquivo não encontrado: {init_db_script}")
        print(f"Diretório atual: {ROOT_DIR}")
        sys.exit(1)

    executar_comando(
        [sys.executable, str(init_db_script), "--seed"],
        "Inicialização dos bancos de dados SQLite",
    )

    # Separador de dados para o PyInstaller (';' no Windows, ':' no Linux/macOS)
    sep = ";" if sys.platform.startswith("win") else ":"

    # 3. Compilação do Executável Seco: triagem_vre
    log_etapa("3/5", "Compilando executável operacional: triagem_vre...")
    cmd_triagem = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--noconsole",
        "--name",
        "triagem_vre",
        "--add-data",
        f"templates{sep}templates",
        "--add-data",
        f"database{os.sep}schema.sql{sep}database",
        "--add-data",
        f"database{os.sep}schema_cep.sql{sep}database",
        str(ROOT_DIR / "src" / "runner" / "main_seco.py"),
    ]
    executar_comando(cmd_triagem, "Compilação do triagem_vre")

    # 4. Compilação do Executável Admin: config_vre
    log_etapa("4/5", "Compilando executável administrativo: config_vre...")
    cmd_config = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "config_vre",
        "--add-data",
        f"database{os.sep}schema.sql{sep}database",
        "--add-data",
        f"database{os.sep}schema_cep.sql{sep}database",
        str(ROOT_DIR / "src" / "admin" / "main_admin.py"),
    ]
    executar_comando(cmd_config, "Compilação do config_vre")

    # 5. Organização da Entrega Final em 'fbuild/'
    log_etapa("5/5", "Organizando a versão final na pasta 'fbuild'...")

    fbuild_dir = ROOT_DIR / "fbuild"
    fbuild_dir.mkdir(parents=True, exist_ok=True)

    # Subpastas requeridas pela arquitetura
    planilhas_anteriores_dir = fbuild_dir / "Planilhas anteriores"
    planilhas_anteriores_dir.mkdir(parents=True, exist_ok=True)

    configuracoes_dir = fbuild_dir / "Configurações"
    configuracoes_dir.mkdir(parents=True, exist_ok=True)

    # Remove pasta duplicada sem acento se porventura existir de builds anteriores
    pasta_duplicada = fbuild_dir / "Configuracoes"
    if pasta_duplicada.exists():
        shutil.rmtree(pasta_duplicada, ignore_errors=True)

    # Remove lançadores legados obsoletos e binários ELF se existirem
    itens_obsoletos = [
        "Executar_Triagem_VRE.bat",
        "Abrir_Painel_Configuracao.bat",
        "assinar_executaveis.ps1",
        "triagem_vre",
        "config_vre",
    ]
    for item in itens_obsoletos:
        p_obs = fbuild_dir / item
        if p_obs.exists():
            p_obs.unlink()

    (configuracoes_dir / "database").mkdir(parents=True, exist_ok=True)
    (configuracoes_dir / "templates").mkdir(parents=True, exist_ok=True)
    (configuracoes_dir / "entrada_csv" / "processados").mkdir(parents=True, exist_ok=True)

    # Copiar executáveis Windows compilados (se estiver no Windows)
    if sys.platform.startswith("win"):
        dist_dir = ROOT_DIR / "dist"
        triagem_bin = dist_dir / "triagem_vre.exe"
        if triagem_bin.exists():
            shutil.copy2(triagem_bin, fbuild_dir / "triagem_vre.exe")
            shutil.copy2(triagem_bin, ROOT_DIR / "triagem_vre.exe")
            print("    [✓] triagem_vre.exe copiado para fbuild/ e raiz.")

        config_bin = dist_dir / "config_vre.exe"
        if config_bin.exists():
            shutil.copy2(config_bin, fbuild_dir / "config_vre.exe")
            shutil.copy2(config_bin, ROOT_DIR / "config_vre.exe")
            print("    [✓] config_vre.exe copiado para fbuild/ e raiz.")

    # Copiar planilha de trabalho ativa atual para a raiz do fbuild/
    candidatos_planilha = [
        ROOT_DIR / "Planilha_Controle_VRE_Ativa.xlsx",
        ROOT_DIR / "planilhas_trabalho" / "Planilha_Controle_VRE_Ativa.xlsx",
        ROOT_DIR / "templates" / "template_triagem.xlsx",
    ]
    for p in candidatos_planilha:
        if p.exists():
            shutil.copy2(p, fbuild_dir / "Planilha_Controle_VRE_Ativa.xlsx")
            print("    [✓] Planilha_Controle_VRE_Ativa.xlsx copiada para a raiz do fbuild/.")
            break

    # Copiar bancos e schemas para Configurações/database
    arquivos_db = [
        ("database/triagem_vre.db", "triagem_vre.db"),
        ("database/cep_areas.db", "cep_areas.db"),
        ("database/schema.sql", "schema.sql"),
        ("database/schema_cep.sql", "schema_cep.sql"),
    ]
    for src_rel, dest_name in arquivos_db:
        orig = ROOT_DIR / src_rel
        if orig.exists():
            shutil.copy2(orig, configuracoes_dir / "database" / dest_name)

    # Copiar template para Configurações/templates
    tmpl_orig = ROOT_DIR / "templates" / "template_triagem.xlsx"
    if tmpl_orig.exists():
        shutil.copy2(tmpl_orig, configuracoes_dir / "templates" / "template_triagem.xlsx")

    # Copiar código-fonte e documentação para Configurações
    src_orig = ROOT_DIR / "src"
    if src_orig.exists():
        dest_src = configuracoes_dir / "src"
        if dest_src.exists():
            shutil.rmtree(dest_src)
        shutil.copytree(src_orig, dest_src, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

    if req_file.exists():
        shutil.copy2(req_file, configuracoes_dir / "requirements.txt")

    readme = ROOT_DIR / "README.md"
    if readme.exists():
        shutil.copy2(readme, configuracoes_dir / "README.md")

    # Copiar script de assinatura digital e importador de histórico para Configurações
    script_assinar = ROOT_DIR / "assinar_executaveis.ps1"
    if script_assinar.exists():
        shutil.copy2(script_assinar, configuracoes_dir / "assinar_executaveis.ps1")

    bat_importar = ROOT_DIR / "Importar_Planilhas_Antigas.bat"
    if bat_importar.exists():
        shutil.copy2(bat_importar, configuracoes_dir / "Importar_Planilhas_Antigas.bat")
        print("    [✓] Importar_Planilhas_Antigas.bat copiado para Configurações/.")

    # Copiar lançadores oficiais para a raiz do fbuild/
    lancadores = [
        "Executar_Triagem.bat",
        "Executar_Triagem_Silencioso.vbs",
        "Painel_Administrativo.bat",
    ]
    for lanc in lancadores:
        arq_lanc = ROOT_DIR / lanc
        if arq_lanc.exists():
            shutil.copy2(arq_lanc, fbuild_dir / lanc)
            print(f"    [✓] {lanc} copiado para fbuild/.")

    # 6. Geração dos Pacotes de Distribuição (.zip)
    import zipfile

    zip_portatil = ROOT_DIR / "Triagem_VRE_Portatil.zip"
    print("    -> Gerando Triagem_VRE_Portatil.zip a partir de 'fbuild/'...")
    with zipfile.ZipFile(zip_portatil, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for item in fbuild_dir.rglob("*"):
            if "__pycache__" in item.parts or item.suffix in (".pyc", ".log"):
                continue
            rel_path = item.relative_to(fbuild_dir)
            zf.write(item, str(rel_path))
    print(f"    [✓] Triagem_VRE_Portatil.zip gerado ({zip_portatil.stat().st_size // (1024*1024)} MB).")

    zip_deploy = ROOT_DIR / "PROJETO_VRE_DEPLOY.zip"
    print("    -> Gerando PROJETO_VRE_DEPLOY.zip (código-fonte completo)...")
    ignore_dirs = {".git", "__pycache__", ".pytest_cache"}
    ignore_suffixes = {".pyc", ".log"}
    with zipfile.ZipFile(zip_deploy, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for item in ROOT_DIR.rglob("*"):
            if item.name in ("PROJETO_VRE_DEPLOY.zip", "Triagem_VRE_Portatil.zip", ".cache_python_amd64.zip"):
                continue
            if any(part in ignore_dirs for part in item.parts):
                continue
            if item.suffix in ignore_suffixes:
                continue
            rel_path = item.relative_to(ROOT_DIR)
            zf.write(item, str(rel_path))
    print(f"    [✓] PROJETO_VRE_DEPLOY.zip gerado ({zip_deploy.stat().st_size // (1024*1024)} MB).")

    print("\n" + "=" * 78)
    print(" [SUCESSO] Build concluído com êxito!")
    print(" A versão final funcional está organizada na pasta 'fbuild/':")
    print(f"   -> {fbuild_dir / 'Planilha_Controle_VRE_Ativa.xlsx'}")
    print(f"   -> {fbuild_dir / 'Executar_Triagem.bat'}")
    print(f"   -> {fbuild_dir / 'Executar_Triagem_Silencioso.vbs'}")
    print(f"   -> {fbuild_dir / 'Painel_Administrativo.bat'}")
    print(f"   -> {fbuild_dir / 'Planilhas anteriores'}")
    print(f"   -> {fbuild_dir / 'Configurações'}")
    print("=" * 78)


if __name__ == "__main__":
    main()
