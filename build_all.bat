@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==============================================================================
echo  SMAURB - Disparando Build e Empacotamento
echo ==============================================================================

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRO] O Python nao foi encontrado no PATH do Windows.
    echo Por favor, instale o Python ou adicione-o ao PATH.
    pause
    exit /b 1
)

python build.py
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Ocorreu uma falha durante o processo de build.
    pause
    exit /b %errorlevel%
)

if exist "%~dp0assinar_executaveis.ps1" (
    echo.
    echo [*] Aplicando assinatura digital nos executaveis via PowerShell...
    powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0assinar_executaveis.ps1"
)

