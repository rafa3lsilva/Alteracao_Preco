@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   2. SEGUNDO RELATORIO DO DIA
echo   Gera APENAS as novas alteracoes feitas depois do 1o relatorio
echo ====================================================================
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx 2_RELATORIO_PRECOS_NOVOS.pdf --modo segundo
) else (
    python gerar_relatorio_pdf.py ALTERACAO_PRECO.xlsx 2_RELATORIO_PRECOS_NOVOS.pdf --modo segundo
)
echo.
pause
