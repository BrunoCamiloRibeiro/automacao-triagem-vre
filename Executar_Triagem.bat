@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==============================================================================
echo  SMAURB / INSPETORIA FISCAL - TRIAGEM AUTOMATICA DO VRE
echo ==============================================================================
echo.

:: 1. Localiza o Python portatil oficial embutido
set "PY_BIN="
if exist "Configurações\python\python.exe" set "PY_BIN=Configurações\python\python.exe"
if not defined PY_BIN if exist "Configuracoes\python\python.exe" set "PY_BIN=Configuracoes\python\python.exe"
if not defined PY_BIN if exist "python\python.exe" set "PY_BIN=python\python.exe"
if not defined PY_BIN set "PY_BIN=python"

:: 2. Localiza o script de processamento
set "TARGET_SCRIPT="
if exist "Configurações\src\runner\main_seco.py" set "TARGET_SCRIPT=Configurações\src\runner\main_seco.py"
if not defined TARGET_SCRIPT if exist "Configuracoes\src\runner\main_seco.py" set "TARGET_SCRIPT=Configuracoes\src\runner\main_seco.py"
if not defined TARGET_SCRIPT if exist "src\runner\main_seco.py" set "TARGET_SCRIPT=src\runner\main_seco.py"

if not defined TARGET_SCRIPT (
    echo [ERRO] Codigo-fonte da triagem nao localizado na pasta Configuracoes\src.
    echo Pressione qualquer tecla para sair...
    pause >nul
    exit /b 1
)

:: 3. Executa a triagem (aceita arrastar arquivo CSV por cima do bat)
echo [*] Iniciando processamento...
"%PY_BIN%" "%TARGET_SCRIPT%" %*

if %errorlevel% equ 0 (
    echo.
    echo ==============================================================================
    echo  [SUCESSO] Operação concluída com êxito!
    echo  Planilha de trabalho atualizada: Planilha_Controle_VRE_Ativa.xlsx
    echo ==============================================================================
    echo.
    echo Pressione qualquer tecla para fechar ou aguarde 8 segundos...
    timeout /t 8 >nul
) else (
    echo.
    echo ==============================================================================
    echo  [ERRO] Ocorreu uma falha durante o processamento.
    echo  Verifique as mensagens acima ou o arquivo de log: triagem_erro.log
    echo ==============================================================================
    pause
)
