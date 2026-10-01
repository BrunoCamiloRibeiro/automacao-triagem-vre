@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: Se estiver executando de dentro da pasta Configurações, posiciona na raiz da Triagem
if exist "python\python.exe" (
    set "CFG_DIR=%~dp0"
    cd ..
) else if exist "Configurações\python\python.exe" (
    set "CFG_DIR=%~dp0Configurações\"
) else (
    set "CFG_DIR=%~dp0"
)

echo ==============================================================================
echo  SMAURB / INSPETORIA FISCAL - CARGA DE PLANILHAS HISTORICAS
echo ==============================================================================
echo.

:: 1. Localiza interpretador Python portatil embutido
set "PY_BIN="
if exist "%CFG_DIR%python\python.exe" set "PY_BIN=%CFG_DIR%python\python.exe"
if not defined PY_BIN if exist "Configurações\python\python.exe" set "PY_BIN=Configurações\python\python.exe"
if not defined PY_BIN if exist "python\python.exe" set "PY_BIN=python\python.exe"
if not defined PY_BIN set "PY_BIN=python"

:: 2. Localiza o script de importação
set "TARGET_SCRIPT="
if exist "%CFG_DIR%src\runner\importar_antigas.py" set "TARGET_SCRIPT=%CFG_DIR%src\runner\importar_antigas.py"
if not defined TARGET_SCRIPT if exist "Configurações\src\runner\importar_antigas.py" set "TARGET_SCRIPT=Configurações\src\runner\importar_antigas.py"
if not defined TARGET_SCRIPT if exist "src\runner\importar_antigas.py" set "TARGET_SCRIPT=src\runner\importar_antigas.py"

if not defined TARGET_SCRIPT (
    echo [ERRO] Script de importacao nao localizado na pasta src\runner.
    echo Pressione qualquer tecla para sair...
    pause >nul
    exit /b 1
)

:: 3. Executa a importação das planilhas antigas
"%PY_BIN%" "%TARGET_SCRIPT%" %*

echo.
echo Pressione qualquer tecla para fechar esta janela...
pause >nul
