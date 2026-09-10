; Script do Inno Setup para gerar o instalador (Setup.exe) do Relatorio de
; Rastreabilidade de Lote.
;
; PRE-REQUISITO: antes de compilar este script, rode build_exe.bat nesta
; mesma pasta e confirme que dist\RelatorioMVPed\RelatorioMVPed.exe abre
; certinho (veja GUIA_EMPACOTAMENTO.md para o passo a passo completo).
;
; Como compilar:
;   1. Instale o Inno Setup (gratuito): https://jrsoftware.org/isinfo.php
;   2. Abra este arquivo (installer.iss) no Inno Setup Compiler e clique
;      em "Compile" (ou rode pela linha de comando:
;      "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss)
;   3. O instalador final fica em: instalador\RelatorioMVPed_Setup.exe
;      Esse e o UNICO arquivo que precisa ser enviado para quem for usar
;      o app.

#define MyAppName "Relatorio de Rastreabilidade de Lote"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "BioCAZ"
#define MyAppExeName "RelatorioMVPed.exe"

[Setup]
AppId={{7C6F1B2A-9E3D-4A6B-8C2F-5B1E7D4A9F10}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Instala na pasta do usuario (AppData\Local\Programs) - NAO precisa de
; administrador, entao qualquer pessoa consegue instalar sozinha so
; recebendo o .exe, sem pedir senha de admin do Windows.
DefaultDirName={localappdata}\Programs\RelatorioMVPed
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=instalador
OutputBaseFilename=RelatorioMVPed_Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; Se a compilacao reclamar que nao encontra "BrazilianPortuguese.isl",
; comente as duas linhas da secao [Languages] abaixo (o instalador so
; fica em ingles) OU reabra o instalador do Inno Setup e marque a opcao
; de instalar "arquivos de idioma adicionais" (Portuguese Brazilian).
[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Area de Trabalho"; GroupDescription: "Atalhos adicionais:"

[Files]
; Copia TUDO que o build_exe.bat gerou em dist\RelatorioMVPed. ATENCAO:
; por decisao explicita do dono deste projeto, a string de conexao do
; SQL Server (servidor, usuario, senha) esta embutida dentro deste pacote,
; como conn.enc - um arquivo CRIPTOGRAFADO (nao e o .env em texto puro),
; gerado automaticamente pelo build_exe.bat. A chave para decifrar viaja
; junto no proprio app (senao ele nao conseguiria se conectar sozinho),
; entao isso impede a leitura casual (abrir num bloco de notas) mas nao
; e inviolavel para alguem tecnicamente capaz. Distribua
; RelatorioMVPed_Setup.exe so para pessoas de confianca dentro da empresa.
Source: "dist\RelatorioMVPed\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName} agora"; Flags: nowait postinstall skipifsilent

[Code]
// Verifica se o "ODBC Driver 18 for SQL Server" da Microsoft ja esta
// instalado no Windows - esse driver e usado pelo app para conversar com
// o SQL Server do Protheus e NAO vem junto com o Python/pip, precisa ser
// instalado a parte no sistema operacional (uma vez so por computador).
function OdbcDriver18Instalado: Boolean;
begin
  Result :=
    RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server') or
    RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ErrorCode: Integer;
begin
  if (CurStep = ssPostInstall) and (not OdbcDriver18Instalado) then
  begin
    if MsgBox(
      'Este app precisa do "ODBC Driver 18 for SQL Server" (da Microsoft) ' + #13#10 +
      'instalado no Windows para conseguir conectar ao banco de dados - ' + #13#10 +
      'ele nao foi encontrado neste computador.' + #13#10 + #13#10 +
      'Clique em Sim para abrir a pagina de download da Microsoft agora ' + #13#10 +
      '(e gratuito - baixe e instale o "msodbcsql18" de 64 bits antes de ' + #13#10 +
      'usar o app).',
      mbConfirmation, MB_YESNO) = IDYES then
      ShellExec('open',
        'https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server',
        '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  end;
end;
