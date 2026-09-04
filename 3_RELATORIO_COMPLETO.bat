@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   RELATORIO CONSOLIDADO DO DIA (TODAS AS ALTERACOES)
echo ====================================================================
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx RELATORIO_ALTERACAO_PRECO.pdf --modo todos
) else (
    python gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx RELATORIO_ALTERACAO_PRECO.pdf --modo todos
)
echo.
pause
