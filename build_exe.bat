@echo off
setlocal
cd /d "%~dp0"

if not exist .venv (
    echo.
    echo [ERRO] Ambiente virtual .venv nao encontrado nesta pasta.
    echo Rode primeiro "executar_relatorio.bat" uma vez - ele cria o .venv
    echo e instala as dependencias do app - depois rode este build_exe.bat
    echo de novo.
    echo.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo Instalando PyInstaller, cryptography e pywebview no ambiente virtual
echo (se necessario)...
pip install --quiet pyinstaller cryptography pywebview
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias. Veja a mensagem acima.
    pause
    exit /b 1
)

echo.
echo Gerando conn.enc (string de conexao do .env, criptografada) para
echo embutir no instalador...
python gerar_conn_criptografada.py
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel gerar conn.enc. Veja a mensagem acima -
    echo geralmente e porque o .env nao existe nesta pasta ainda.
    pause
    exit /b 1
)

echo.
echo Limpando builds anteriores (pastas build\ e dist\)...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Gerando o executavel do app (RelatorioMVPed.exe) - isso pode demorar
echo alguns minutos na primeira vez...
echo.
echo AVISO: este build inclui conn.enc - a string de conexao do .env,
echo criptografada - dentro do .exe, para o app conectar sozinho sem pedir
echo nada no primeiro uso. A criptografia impede que alguem abra a pasta
echo instalada num bloco de notas e leia a senha direto, mas nao e um
echo cofre inviolavel: a chave para decifrar tambem viaja dentro do
echo proprio app (senao ele nao conseguiria se conectar sozinho), entao
echo uma pessoa com conhecimento tecnico ainda consegue extrair a senha
echo com esforco. Trate o instalador com cuidado mesmo assim - envie so
echo para quem for usar, nunca em canal publico.
echo.
echo OBS: app.py e carregado pelo Streamlit em tempo de execucao (nao e
echo "importado" normalmente), entao o PyInstaller nao enxerga sozinho
echo tudo que app.py usa - por isso listamos --collect-all para CADA
echo biblioteca de requirements.txt explicitamente abaixo.
echo.
echo OBS: --windowed remove a janela de terminal - o app abre direto numa
echo janela propria (pywebview), sem aba de navegador. Sem terminal, erros
echo do processo servidor vao parar em %%APPDATA%%\RelatorioMVPed\erro.log
echo em vez de aparecer na tela - se o app nao conectar, olhe esse arquivo.

pyinstaller --noconfirm --clean --name RelatorioMVPed --onedir --windowed --icon icon.ico ^
    --add-data "app.py;." ^
    --add-data "conn.enc;." ^
    --add-data "logo.png;." ^
    --add-data "logo_original_fundo_branco.png;." ^
    --add-data ".streamlit\config.toml;.streamlit" ^
    --collect-all streamlit ^
    --copy-metadata streamlit ^
    --collect-all altair ^
    --collect-all pandas ^
    --collect-all pyodbc ^
    --collect-all reportlab ^
    --collect-all sqlalchemy ^
    --collect-all dotenv ^
    --collect-all xlsxwriter ^
    --collect-all cryptography ^
    --collect-all webview ^
    --collect-all clr_loader ^
    --hidden-import sqlalchemy.dialects.mssql.pyodbc ^
    run_app.py

if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao gerar o executavel. Veja a mensagem acima.
    pause
    exit /b 1
)

echo.
echo ==========================================================
echo  Pronto! App empacotado em:
echo    dist\RelatorioMVPed\RelatorioMVPed.exe
echo.
echo  IMPORTANTE - antes de gerar o instalador:
echo  1. Abra esse .exe manualmente (duplo clique) e confirme que abre
echo     uma janela propria (sem terminal, sem navegador) ja conectada
echo     no banco, direto na tela principal do relatorio.
echo  2. Teste tambem os botoes de "Baixar Excel" e "Baixar PDF" nas
echo     abas, ja que essas bibliotecas sao as que mais costumam
echo     faltar num primeiro build (reportlab, xlsxwriter).
echo  3. Se a janela nao abrir ou ficar em branco, veja o arquivo
echo     %%APPDATA%%\RelatorioMVPed\erro.log (nao ha terminal para
echo     mostrar o erro na tela desta vez).
echo  4. So depois disso, compile o instalador.iss no Inno Setup
echo     (veja GUIA_EMPACOTAMENTO.md).
echo ==========================================================
echo.
pause
