@echo off
setlocal
cd /d "%~dp0"
cls
echo ====================================================================
echo   INICIANDO PAINEL WEB DE ALTERACAO DE PRECOS (STREAMLIT)
echo ====================================================================
echo.
echo Abrindo painel no seu navegador...
echo Para fechar o painel, feche esta janela do terminal.
echo.
if exist "%~dp0.venv\Scripts\streamlit.exe" (
    "%~dp0.venv\Scripts\streamlit.exe" run app.py
) else (
    streamlit run app.py
)
pause
