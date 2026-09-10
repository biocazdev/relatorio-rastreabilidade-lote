@echo off
setlocal

cd /d "%~dp0"

set TABELA=%~1
if "%TABELA%"=="" (
    echo.
    set /p TABELA="Digite o nome da tabela (ex: SB8010) e pressione Enter: "
)
if "%TABELA%"=="" (
    echo Nenhum nome informado. Encerrando.
    pause
    exit /b 1
)

if not exist .env (
    echo.
    echo [ERRO] Arquivo .env nao encontrado nesta pasta.
    pause
    exit /b 1
)

if not exist .venv (
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

python listar_colunas.py %TABELA%

echo.
pause
