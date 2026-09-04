@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   CONFIGURACAO AUTOMATICA DO SISTEMA DE PRECOS
echo ====================================================================
echo.
echo Verificando instalacao do Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] O Python nao foi encontrado neste computador!
    echo Por favor, instale o Python em https://www.python.org/
    echo (Lembre-se de marcar a opcao "Add Python to PATH" durante a instalacao).
    echo.
    pause
    exit /b
)

echo [1/3] Criando ambiente virtual isolado (.venv)...
python -m venv .venv

echo [2/3] Instalando todas as dependencias necessarias...
"%~dp0.venv\Scripts\pip.exe" install --upgrade pip
"%~dp0.venv\Scripts\pip.exe" install -r requirements.txt

echo.
echo ====================================================================
echo   CONFIGURACAO CONCLUIDA COM SUCESSO!
echo   Agora voce ja pode abrir o INICIAR_PAINEL_WEB.bat
echo ====================================================================
echo.
pause
