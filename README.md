# Relatório de Rastreabilidade de Lote (Protheus/TOTVS)

App em Streamlit que consulta o SQL Server do Protheus (SD1, SD2, SB1, SD5,
SB6, SB8, SA1, SA2, SBM) para rastreabilidade de lote e controle de
terceiros, em 4 abas.

> **Escopo: só Produto Acabado (PA)**. Todas as abas filtram os produtos
> por `SB1.B1_TIPO = 'PA'` — matéria-prima, embalagem e outros tipos não
> aparecem em nenhum relatório, mesmo tendo controle de lote no Protheus.
> Isso vale também para a **Busca por Lote** (procurar o número de um
> lote de matéria-prima não vai encontrar nada) e o **Controle de
> Terceiros** (remessas de matéria-prima para terceirização não
> aparecem). Se algum desses dois casos for necessário no futuro, é só
> pedir para eu tirar a trava só ali.

1. **Relatório de Movimentos** — a consulta original, com filtros de
   filial, período, produto, pedido, tipo de movimento e status de
   rastreabilidade (agora incluindo detecção de **cobertura parcial de
   lote** — quando só parte da quantidade do documento está vinculada a um
   lote na SD5), gráfico clicável, tabela e exportação para Excel ou PDF.
2. **Consolidado por Lote** — agrupa tudo por PRODUTO + LOTE: saldo atual
   e validade (via SB8), total comprado/vendido e quantos clientes
   distintos receberam aquele lote, com alerta visual de lotes vencidos ou
   vencendo em até 30 dias.
3. **Busca por Lote** — digite um número de lote e veja todas as entradas
   (de quem veio) e saídas (para quem foi vendido) — útil num cenário de
   recall, para saber rapidamente a origem e o destino de um lote
   específico. Cada movimento mostra também a descrição do armazém (não só
   o código), via cadastro de Armazéns (NNR).
4. **Controle de Terceiros** — remessas e devoluções de terceiros (SB6):
   o que foi enviado a clientes/fornecedores ou recebido deles, com o
   saldo em aberto de cada documento. Filtros de período (emissão),
   produto, cliente/fornecedor (código ou nome), tipo (cliente/
   fornecedor), direção (de/em terceiros) e situação. A situação de cada
   documento (**Em aberto** / **Parcial** / **Atendido** / **Devolução**)
   já vem calculada na consulta a partir do saldo restante, com destaque
   visual na tabela (amarelo/laranja/azul) e um indicador de **saldo
   pendente estimado** (soma de saldo × preço unitário dos documentos em
   aberto ou parciais — um número aproximado, não um valor fiscal).

Nas quatro abas, ao lado do botão "Baixar Excel" tem um "Baixar PDF": gera
um PDF em paisagem com a logo e as cores da BioCAZ no cabeçalho, um resumo
de KPIs (os mesmos números que aparecem na tela, em formato de cartões)
logo abaixo do título, a tabela com o mesmo destaque visual usado na tela
(colunas numéricas alinhadas à direita), cabeçalho da tabela repetido em
cada página e um rodapé com o nome do relatório, "Documento de uso
interno", aviso de confidencialidade e numeração de página. Como alguns
relatórios têm muitas colunas, os nomes das colunas aparecem na vertical
no cabeçalho da tabela — assim cabem inteiros, sem quebrar a palavra no
meio, mesmo com relatórios largos como o de Movimentos.

As abas 2, 3 e 4 usam consultas separadas da aba 1 de propósito: se algum
nome de campo (por exemplo do SB8, saldo/validade do lote) for diferente
no seu ambiente, só aquela aba mostra um aviso de erro — a aba 1 (que já
está validada em produção) continua funcionando normalmente.

Nos filtros de produto (abas 1, 2 e 4), além de digitar o código exato do
Protheus, dá para abrir o "🔍 Não sabe o código? Buscar produto por nome"
e digitar parte da descrição — o app lista os produtos (só Produto
Acabado) que batem com o nome e, ao escolher um, preenche o código
automaticamente. Não é mais necessário consultar o Protheus antes só para
descobrir um código.

Assim que uma tabela é exibida (abas 1 a 4), aparece um aviso "🕒 Dados de
HH:MM" acima dela, com o horário em que aquela consulta foi de fato
executada no banco — como o resultado fica em cache por 5 minutos, esse
aviso ajuda a saber se o que está na tela é recente ou se vale clicar em
"Executar consulta"/"Consultar"/"Buscar" de novo para atualizar.

> **Nota sobre a versão do pandas**: o `requirements.txt` trava
> `pandas<3.0`. A partir da versão 3.0 o pandas passou a inferir
> automaticamente um novo tipo de dado para colunas de texto, o que faz
> datas em branco aparecerem como `nan` em vez de célula vazia. Manter a
> versão 2.x evita esse problema — não é necessário fazer nada, o `.bat`
> já instala a versão correta.

## Arquivos

- `app.py` — o app Streamlit (conexão, filtros, KPIs, tabela, exportação, 4 abas).
- `test_conexao.py` — script simples para testar a conexão com o banco antes de abrir o app.
- `executar_relatorio.bat` — **clique duplo** para instalar dependências (se preciso) e abrir o relatório no navegador.
- `testar_conexao.bat` — **clique duplo** para só testar a conexão com o banco (usa o `.env`).
- `requirements.txt` — dependências Python.
- `packages.txt` — pacotes de sistema (drivers) usados no Streamlit Community Cloud.
- `.env.example` — modelo de conexão para **teste local rápido** (ex: homologação).
- `.streamlit/secrets.toml.example` — modelo de conexão para **produção / Streamlit Cloud**.
- `.streamlit/config.toml` — tema visual com a paleta da marca BioCAZ.
- `logo.png` — logomarca exibida no topo da barra lateral e como ícone da aba do navegador.

## Identidade visual

O app usa a logo da BioCAZ (arquivo `logo.png`, na raiz do projeto) e a
paleta oficial da marca, definida em `.streamlit/config.toml`:

- Verde `#43AA8A` — cor primária (botões, itens selecionados, gráfico de status)
- Azul-marinho `#0C1B7D` — cor de links
- Preto `#000000` — cor do texto

Para trocar a logo, basta substituir o arquivo `logo.png` por outro PNG de
mesmo nome (o app carrega o arquivo direto da pasta do projeto).

O app decide sozinho de onde ler a conexão: primeiro tenta
`.streamlit/secrets.toml`, e se essa seção não existir, usa as variáveis do
`.env`. A barra lateral mostra qual das duas está sendo usada.

## 0. Jeito mais rápido de testar (Windows, sem digitar comando nenhum)

1. Copie `.env.example` para `.env` (mesma pasta) e preencha com os dados
   do banco de homologação — **confira se o Explorer não salvou como
   `.env.txt`** (ative "Extensões de nome de arquivo" na aba Exibir).
2. Dê **duplo clique em `testar_conexao.bat`** para confirmar que a conexão
   funciona (ele mostra a versão do SQL Server e a contagem de linhas de
   cada tabela usada na consulta).
3. Se o teste passar, dê **duplo clique em `executar_relatorio.bat`** para
   abrir o relatório no navegador.

Os dois `.bat` criam sozinhos um ambiente virtual (`.venv`) na primeira
execução e instalam as dependências — só exigem que o Python esteja
instalado e no PATH do Windows.

## 1. Rodando localmente (linha de comando, alternativa aos .bat)

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux/Mac

pip install -r requirements.txt
```

### Opção A — teste rápido com `.env` (recomendado para homologação)

Copie `.env.example` para `.env`, na mesma pasta do `app.py`, e preencha
com os dados do banco de **homologação**:

```
SQLSERVER_SERVIDOR=servidor-homologacao-ou-ip
SQLSERVER_PORTA=1433
SQLSERVER_DATABASE=PROTHEUS_HOMOLOGACAO
SQLSERVER_USUARIO=usuario_leitura
SQLSERVER_SENHA=sua_senha_aqui
SQLSERVER_DRIVER=ODBC Driver 18 for SQL Server
SQLSERVER_ODBC_EXTRA=TrustServerCertificate=yes;Encrypt=yes
```

O arquivo `.env` **não** deve ir para o Git (já está no `.gitignore`). Ele
só é lido se não existir `.streamlit/secrets.toml` com uma seção `[sqlserver]`
— então, para testar com `.env`, não crie o `secrets.toml` ainda.

### Opção B — `.streamlit/secrets.toml` (formato usado em produção)

Copie `.streamlit/secrets.toml.example` para `.streamlit/secrets.toml` e
preencha com os dados reais de conexão. Esse arquivo também **não** deve
ir para o Git.

```bash
streamlit run app.py
```

O app abre no navegador (normalmente `http://localhost:8501`). Na barra
lateral, confira se está escrito "Conexão via: **.env**" (ou
"**secrets.toml**") para saber qual configuração está em uso, ajuste os
filtros e clique em "Executar consulta".

## 2. Driver ODBC — atenção especial

- **Localmente / servidor Windows da empresa**: normalmente já existe o
  "ODBC Driver 17" ou "18 for SQL Server" da Microsoft instalado. Use esse
  nome no campo `driver` do `secrets.toml`.
- **Streamlit Community Cloud (Linux)**: o driver ODBC da Microsoft **não**
  vem instalado e não pode ser adicionado via `packages.txt` (exigiria
  configurar o repositório da Microsoft, que o Cloud não permite rodar).
  Este projeto já inclui um `packages.txt` com o driver **FreeTDS**, que
  funciona nativamente no Cloud. Nesse caso, no secrets do Cloud, use:

  ```toml
  driver     = "FreeTDS"
  odbc_extra = "TDS_Version=7.4;"
  ```

  Se preferir não lidar com ODBC no Cloud, a alternativa mais simples é
  trocar `pyodbc` por `pymssql` (`pip install pymssql`, sem dependência de
  driver do sistema) e ajustar a string de conexão em `get_engine()` para
  `mssql+pymssql://usuario:senha@servidor:porta/database`.

## 3. Publicando no Streamlit Community Cloud

1. Suba este projeto para um repositório no GitHub (**sem** o arquivo
   `secrets.toml` real — só o `.example`).
2. Em [share.streamlit.io](https://share.streamlit.io), clique em
   "New app", selecione o repositório, branch e o arquivo `app.py`.
3. Em **Settings → Secrets**, cole o conteúdo do seu `secrets.toml` real
   (com os dados de conexão verdadeiros).
4. Deploy.

### ⚠️ Importante: o banco precisa ser alcançável pela internet

O Streamlit Community Cloud roda em servidores públicos da Streamlit/Snowflake,
fora da rede da empresa. Para o app conseguir se conectar ao SQL Server do
Protheus, uma das opções abaixo é necessária:

- **Não recomendado**: abrir a porta do SQL Server (1433) diretamente para
  a internet — risco alto de segurança, mesmo com usuário restrito.
- **Recomendado**: criar um túnel seguro entre o Cloud e o servidor interno,
  por exemplo com [Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/)
  ou [Tailscale Funnel](https://tailscale.com/kb/1223/funnel), expondo só
  a porta do banco para um endpoint autenticado.
- **Alternativa mais simples**: hospedar o app em um servidor dentro da
  própria rede da empresa (Docker, IIS + reverse proxy, ou uma VM), em vez
  do Community Cloud — assim o app já enxerga o SQL Server sem exposição
  externa. Se quiser seguir por esse caminho depois, um `Dockerfile` pode
  ser adicionado facilmente a este projeto.

Em qualquer cenário, use um usuário de banco **somente leitura** e com
acesso restrito às tabelas necessárias (`SD1010`, `SD2010`, `SB1010`,
`SD5010`, `SB8010`, `SB6010`, `SBM010`, `SA1010`, `SA2010`, `NNR010`),
nunca o usuário administrativo do Protheus.

## 4. Filtros disponíveis no app

- Filial (texto, ex: `010101`)
- Período (data inicial/final, opcional)
- Produto (código exato, opcional)
- Pedido (número exato, opcional)
- Tipo de movimento (Entrada/Saída)
- Status de rastreabilidade (OK / Alerta / Sem controle / Verificar)

Os resultados ficam em cache por 5 minutos (`st.cache_data(ttl=300)`) para
evitar sobrecarregar o banco a cada interação; o botão "Executar consulta"
força uma nova busca.
