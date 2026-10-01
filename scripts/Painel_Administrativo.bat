@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 1. Localiza o Python portatil oficial embutido
set "PYW_BIN="
if exist "Configurações\python\pythonw.exe" set "PYW_BIN=Configurações\python\pythonw.exe"
if not defined PYW_BIN if exist "Configuracoes\python\pythonw.exe" set "PYW_BIN=Configuracoes\python\pythonw.exe"
if not defined PYW_BIN if exist "python\pythonw.exe" set "PYW_BIN=python\pythonw.exe"
if not defined PYW_BIN set "PYW_BIN=pythonw"

:: 2. Localiza o script do painel
set "TARGET_SCRIPT="
if exist "Configurações\src\admin\main_admin.py" set "TARGET_SCRIPT=Configurações\src\admin\main_admin.py"
if not defined TARGET_SCRIPT if exist "Configuracoes\src\admin\main_admin.py" set "TARGET_SCRIPT=Configuracoes\src\admin\main_admin.py"
if not defined TARGET_SCRIPT if exist "src\admin\main_admin.py" set "TARGET_SCRIPT=src\admin\main_admin.py"

if not defined TARGET_SCRIPT (
    echo [ERRO] Script do painel administrativo nao encontrado na pasta Configuracoes\src.
    pause
    exit /b 1
)

start "" "%PYW_BIN%" "%TARGET_SCRIPT%" %*
