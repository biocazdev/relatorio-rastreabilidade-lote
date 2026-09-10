# Empacotando o app como instalador (.exe) com Inno Setup

Este guia explica como transformar este app Streamlit num instalador
Windows único (`RelatorioMVPed_Setup.exe`) que qualquer pessoa na empresa
pode instalar sozinha, sem precisar ter Python instalado, sem precisar
saber o que é Streamlit, sem precisar de permissão de administrador — e
já conectando direto no banco, sem preencher nada na primeira vez.

Foram adicionados a este projeto (além dos arquivos que já existiam):

- `run_app.py` — um "lançador" que o PyInstaller usa para empacotar o
  Streamlit como um `.exe` de verdade: ele sobe o Streamlit escondido nos
  bastidores e abre uma **janela própria do programa** (via `pywebview`),
  sem aba de navegador e sem janela de terminal — para quem usa, é
  indistinguível de qualquer outro programa instalado no Windows.
- `gerar_conn_criptografada.py` — gera `conn.enc` a partir do seu `.env`
  real, criptografado, para ser embutido no instalador (ver seção "Sobre
  a senha do banco" abaixo — é a parte mais importante deste guia).
- `build_exe.bat` — gera o `.exe` do app a partir do código-fonte (chama
  o script acima sozinho, não precisa rodar na mão).
- `installer.iss` — script do Inno Setup que empacota esse `.exe` num
  instalador único.
- `icon.ico` — ícone gerado a partir de `logo_original_fundo_branco.png`,
  usado no `.exe`, no instalador e nos atalhos.
- Uma camada de **config embutida e criptografada** foi adicionada ao
  `app.py`, além da tela de configuração manual que já existia como
  reserva.

## Visão geral do processo

```
código-fonte (app.py + .env)  →  build_exe.bat (PyInstaller)  →  dist\RelatorioMVPed\RelatorioMVPed.exe
                                                                          │
                                                                          ▼
                                                          installer.iss (Inno Setup)
                                                                          │
                                                                          ▼
                                                      instalador\RelatorioMVPed_Setup.exe
                                                        (ESTE é o arquivo que você distribui)
```

Duas ferramentas, duas etapas. O PyInstaller pega o app Python/Streamlit
e o transforma num programa Windows independente (não precisa de Python
instalado na máquina de quem for usar). O Inno Setup pega esse programa e
o embrulha num instalador com tela de boas-vindas, atalho no menu
iniciar, atalho na área de trabalho (opcional) e desinstalador.

## Sobre a senha do banco (leia antes de distribuir)

**Decisão deste projeto: a string de conexão real (servidor, usuário e
senha do SQL Server, a mesma do seu `.env`) fica embutida dentro do
`.exe`/instalador, de forma criptografada.** Isso foi decidido porque o
banco só é alcançável de qualquer forma estando na rede da empresa ou
conectado na VPN — então o app instalado só funciona mesmo para quem já
tem esse acesso de rede, e faz sentido cada notebook já abrir conectado,
sem pedir nada.

Como funciona: a cada `build_exe.bat`, o script `gerar_conn_criptografada.py`
lê o `.env` real desta pasta e gera `conn.enc` — um arquivo criptografado
(usando a biblioteca `cryptography`, cifra simétrica Fernet) — que é
empacotado dentro do `.exe` no lugar do `.env` puro. Em tempo de
execução, o app decifra esse arquivo em memória para conectar; nada é
salvo em disco em texto puro.

**Seja honesto consigo mesmo sobre o nível de proteção disso:** a chave
para decifrar viaja junto dentro do próprio código do app (`app.py`),
porque sem ela o app não conseguiria se conectar sozinho na máquina de
quem instalou. Isso significa que:

- ✅ Alguém que abrir a pasta instalada num bloco de notas comum **não**
  vai ver a senha em texto puro (diferente de um `.env` cru).
- ✅ Protege contra curiosidade casual e contra ferramentas simples de
  busca de texto (`strings`, Ctrl+F num editor, etc).
- ❌ **Não é inviolável.** Alguém com conhecimento técnico (que saiba
  abrir o `.exe` com uma ferramenta de engenharia reversa e entenda
  Python) consegue, com algum esforço, extrair a chave e decifrar a
  senha. Isso é uma limitação de qualquer app que precisa se conectar
  sozinho sem pedir credenciais — não existe forma 100% seguindo de
  embutir um segredo permanente num programa que roda na máquina de
  outra pessoa.

Recomendações práticas com essa decisão:

- Envie o instalador só para quem realmente vai usar o app, por um canal
  que você confia (nunca num grupo público, link aberto, ou pasta
  compartilhada sem controle de quem acessa).
- Se esse arquivo algum dia for parar em mãos erradas de forma que te
  preocupe (perdido, enviado para a pessoa errada, notebook roubado com
  ele salvo), **troque a senha do usuário `DESENVLEITURA` no SQL Server**
  e gere um novo instalador — é a única forma de "revogar" o acesso de
  verdade.
- Ao atualizar o app no futuro, sempre rode `build_exe.bat` de novo antes
  de distribuir (ele regenera o `conn.enc` a partir do `.env` atual, então
  se a senha mudar, a nova versão do instalador já sai com a senha nova).

A tela de "configuração inicial" que existe no `app.py` continua lá como
reserva (só aparece se, por algum motivo, nem `conn.enc` nem `.env` forem
encontrados em tempo de execução), mas em uso normal ela nunca deve
aparecer — o app já abre direto conectado.

## Passo 1 — Gerar o executável (PyInstaller)

Pré-requisito: Python instalado (você já tem, é o mesmo usado pelo
`executar_relatorio.bat`), e um `.env` válido nesta pasta (o mesmo que o
app já usa para rodar localmente).

1. Dê duplo clique em **`build_exe.bat`** (ou rode pelo terminal, dentro
   da pasta do projeto).
2. Aguarde — a primeira vez demora alguns minutos (o PyInstaller baixa e
   empacota o Streamlit e todas as dependências).
3. Ao final, o app empacotado estará em:
   `dist\RelatorioMVPed\RelatorioMVPed.exe`
4. **Teste esse `.exe` manualmente antes de continuar**: dê duplo clique
   nele. Não deve abrir nem janela de terminal nem aba de navegador — o
   app abre direto numa janela própria (título "Relatório de
   Rastreabilidade de Lote"), já conectada no banco, direto na tela
   principal do relatório (sem pedir nada, já que a conexão vem
   embutida). Teste também os botões "Baixar Excel" e "Baixar PDF" nas
   abas. Para fechar o app, feche essa janela normalmente (no X do
   canto, como qualquer programa).

Como o build agora usa `--windowed`, não existe mais console para
mostrar mensagem de erro na tela. Se a janela não abrir, ficar em branco,
ou fechar sozinha, olhe o arquivo:

```
%APPDATA%\RelatorioMVPed\erro.log
```

Esse arquivo é criado automaticamente na primeira vez que algo dá
errado (por exemplo: o Streamlit demorou demais pra subir, ou faltou
alguma dependência no build). Ele acumula um registro por tentativa,
então o erro mais recente fica no final do arquivo.

## Passo 2 — Instalar o Inno Setup (só na primeira vez)

Baixe gratuitamente em <https://jrsoftware.org/isinfo.php> (link
"Download") e instale normalmente. Na tela de seleção de componentes do
instalador do Inno Setup, se aparecer uma opção de **idiomas
adicionais**, marque "Brazilian Portuguese" (não é obrigatório, mas deixa
o instalador final em português).

## Passo 3 — Gerar o instalador (Inno Setup)

1. Abra o **Inno Setup Compiler**.
2. Abra o arquivo `installer.iss` (dentro da pasta do projeto).
3. Clique em **Build → Compile** (ou aperte Ctrl+F9).
4. Se aparecer um erro dizendo que não encontra
   `BrazilianPortuguese.isl`, é porque o idioma adicional do Passo 2 não
   foi instalado — abra `installer.iss` num editor de texto e comente
   (coloque `;` na frente) as duas linhas da seção `[Languages]`, depois
   compile de novo (o instalador só fica em inglês, sem afetar o app em
   si, que continua em português).
5. Ao terminar, o instalador final estará em:
   `instalador\RelatorioMVPed_Setup.exe`

**Esse arquivo (`RelatorioMVPed_Setup.exe`) é o único que você precisa
enviar** para quem for usar o app (e-mail, pendrive, link de download
interno, etc.) — trate-o com o mesmo cuidado que trataria a senha do
banco (ver seção acima).

## O que acontece quando alguém recebe e roda o `RelatorioMVPed_Setup.exe`

1. Instala em `%LOCALAPPDATA%\Programs\RelatorioMVPed` — **não pede senha
   de administrador do Windows**, porque instala só para aquele usuário.
2. Cria atalho no Menu Iniciar (e na Área de Trabalho, se a pessoa marcar
   essa opção durante a instalação).
3. Verifica se o **"ODBC Driver 18 for SQL Server"** (da Microsoft) está
   instalado no Windows daquele computador. Esse driver é o que permite
   ao app conversar com o SQL Server — ele não vem junto com o instalador
   (por tamanho e licenciamento) e precisa ser instalado uma vez por
   computador. Se não encontrar, o instalador oferece abrir a página
   oficial de download da Microsoft.
4. Ao abrir o app pela primeira vez, já conecta direto no banco — a
   string de conexão (servidor, usuário, senha) veio embutida e
   criptografada dentro do próprio instalador (ver seção "Sobre a senha
   do banco" acima).
5. O computador também precisa conseguir alcançar o servidor do SQL
   Server pela rede (mesma VPN/rede da empresa que já é necessária hoje
   para quem roda o app localmente).
6. A janela do app usa o **"WebView2 Runtime"** da Microsoft para
   desenhar a tela (é a mesma engine do navegador Edge, só que sem a
   moldura de navegador). Isso já vem instalado de fábrica em qualquer
   Windows 10 atualizado (versão 2004 ou mais nova) e em todo Windows 11
   — ou seja, na prática quase ninguém vai precisar instalar nada extra.
   Só em um Windows 10 bem desatualizado isso poderia faltar; se
   acontecer, o Windows Update resolve, ou dá pra baixar em
   <https://developer.microsoft.com/microsoft-edge/webview2/>.

## Atualizando o app depois (nova versão)

Sempre que `app.py` ou o `.env` mudarem:

1. Rode `build_exe.bat` de novo (regenera `conn.enc` a partir do `.env`
   atual e gera um novo `dist\RelatorioMVPed`).
2. Abra `installer.iss`, aumente o número em `MyAppVersion` (ex:
   `"1.0.0"` → `"1.1.0"`), e compile de novo.
3. Distribua o novo `RelatorioMVPed_Setup.exe` — quem já tinha instalado
   pode rodar o novo instalador por cima (ele substitui os arquivos do
   app).

## Arquivos que entram (e que NÃO entram) no instalador

O `build_exe.bat` empacota explicitamente: `app.py`, `conn.enc` (a
conexão real, criptografada — gerada na hora a partir do seu `.env`), os
dois logos, e `.streamlit\config.toml` (só o tema visual). O `.env` em
texto puro, `.env.example`, `.streamlit\secrets.toml` e
`.streamlit\secrets.toml.example` **não** são copiados — só o `conn.enc`
criptografado carrega a informação de conexão para dentro do pacote.

Se um dia você quiser voltar ao modelo anterior (instalador sem nenhuma
credencial embutida, pedindo os dados no primeiro uso de cada máquina),
basta remover as linhas relacionadas a `gerar_conn_criptografada.py` e
`conn.enc` do `build_exe.bat` — a tela de configuração inicial do
`app.py` já está pronta para assumir nesse cenário.
