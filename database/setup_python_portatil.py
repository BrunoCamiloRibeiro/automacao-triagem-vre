"""Script de Preparação do Python Portátil Oficial (Windows 64-bit).

Baixa o runtime oficial do Python da python.org com assinatura digital Authenticode
da Python Software Foundation e instala as dependências (openpyxl, xlrd, et-xmlfile)
dentro de Lib/site-packages.

Isso elimina 100% dos falsos positivos de Antivírus Corporativos, pois:
1. python.exe e pythonw.exe são binários oficiais e confiáveis globalmente.
2. Não há descompactador PyInstaller atuando como 'dropper' na pasta %TEMP%.
3. Não requer instalação nem privilégios de Administrador na máquina do usuário final.
"""

from __future__ import annotations

import io
import os
import shutil
import urllib.request
import zipfile
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT_DIR / "Configurações"
FBUILD_CONFIG_DIR = ROOT_DIR / "fbuild" / "Configurações"

PYTHON_ZIP_URL = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.zip"
CACHE_ZIP_PATH = ROOT_DIR / ".cache_python_amd64.zip"


def baixar_ou_obter_zip() -> bytes:
    """Obtém os bytes do runtime oficial do Python em cache ou via download."""
    if CACHE_ZIP_PATH.exists() and CACHE_ZIP_PATH.stat().st_size > 10_000_000:
        print(f"[*] Usando pacote Python em cache local: {CACHE_ZIP_PATH}")
        return CACHE_ZIP_PATH.read_bytes()

    print(f"[*] Baixando runtime oficial do Python ({PYTHON_ZIP_URL})...")
    req = urllib.request.Request(PYTHON_ZIP_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp:
        data = resp.read()

    print(f"[✓] Download concluído: {len(data) / (1024*1024):.2f} MB")
    try:
        CACHE_ZIP_PATH.write_bytes(data)
    except Exception:
        pass
    return data


def instalar_python_portatil(target_python_dir: Path, zip_bytes: bytes) -> None:
    """Extrai os binários limpos e copia dependências Python necessárias."""
    print(f"[*] Instalando Python portátil em: {target_python_dir}...")
    target_python_dir.mkdir(parents=True, exist_ok=True)

    pastas_ignoradas = ("doc/", "doc\\", "include/", "include\\")
    extensoes_ignoradas = (".pdb", ".chm")

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for member in z.infolist():
            nome = member.filename
            if any(nome.lower().startswith(p) for p in pastas_ignoradas):
                continue
            if any(nome.lower().endswith(ext) for ext in extensoes_ignoradas):
                continue
            z.extract(member, target_python_dir)

    print("    [✓] Arquivos base e DLLs do Python extraídos.")

    # Cria pasta Lib/site-packages
    site_packages = target_python_dir / "Lib" / "site-packages"
    site_packages.mkdir(parents=True, exist_ok=True)

    # Copia openpyxl, xlrd e et_xmlfile do ambiente Python local
    modulos = ["openpyxl", "xlrd", "et_xmlfile"]
    for mod_name in modulos:
        try:
            mod = __import__(mod_name)
            mod_file = Path(mod.__file__).resolve()
            if mod_file.name == "__init__.py":
                orig_pkg = mod_file.parent
                dest_pkg = site_packages / mod_name
                if dest_pkg.exists():
                    shutil.rmtree(dest_pkg)
                shutil.copytree(orig_pkg, dest_pkg, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
                print(f"    [✓] Pacote '{mod_name}' copiado para Lib/site-packages.")
            else:
                shutil.copy2(mod_file, site_packages / mod_file.name)
                print(f"    [✓] Módulo '{mod_name}' copiado para Lib/site-packages.")
        except Exception as e:
            print(f"    [AVISO] Não foi possível copiar '{mod_name}': {e}")


def main() -> None:
    print("=" * 78)
    print(" SMAURB - PREPARAÇÃO DO PYTHON PORTÁTIL EMBUTIDO (WINDOWS)")
    print("=" * 78)

    data = baixar_ou_obter_zip()

    destinos = [
        FBUILD_CONFIG_DIR / "python",
        CONFIG_DIR / "python",
    ]

    for d in destinos:
        instalar_python_portatil(d, data)

    print("\n[✓] Configuração do interpretador portátil concluída com sucesso!")


if __name__ == "__main__":
    main()
