@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   RESETAR MEMORIA / HISTORICO DO DIA
echo ====================================================================
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" gerar_relatorio_pdf.py --modo reset
) else (
    python gerar_relatorio_pdf.py --modo reset
)
echo.
pause
