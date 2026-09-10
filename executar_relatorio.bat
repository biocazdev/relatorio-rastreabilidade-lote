@echo off
setlocal

cd /d "%~dp0"

if not exist .env (
    echo.
    echo [ERRO] Arquivo .env nao encontrado nesta pasta.
    echo Copie .env.example para .env e preencha os dados de conexao antes de continuar.
    echo.
    pause
    exit /b 1
)

if not exist .venv (
    echo Criando ambiente virtual Python em .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERRO] Nao foi possivel criar o ambiente virtual. Verifique se o Python esta instalado e no PATH.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

echo Instalando/atualizando dependencias...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar as dependencias. Veja a mensagem acima.
    pause
    exit /b 1
)

echo.
echo Iniciando o relatorio no navegador...
echo (Para encerrar, feche esta janela ou pressione Ctrl+C)
echo.
streamlit run app.py

pause
