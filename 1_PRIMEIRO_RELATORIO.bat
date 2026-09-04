@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   1. PRIMEIRO RELATORIO DO DIA
echo   Gera todas as alteracoes feitas ate este momento e salva a memoria
echo ====================================================================
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx 1_RELATORIO_PRECOS.pdf --modo primeiro
) else (
    python gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx 1_RELATORIO_PRECOS.pdf --modo primeiro
)
echo.
pause
