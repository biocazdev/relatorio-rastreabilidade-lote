"""
Relatorio de Rastreabilidade de Lote (SD1/SD2/SB1/SD5/SB8/SB6) - Protheus/TOTVS
App Streamlit: conecta no SQL Server do Protheus, roda consultas com filtros
dinamicos e permite exportar cada resultado para Excel ou PDF (com o
logotipo da BioCAZ).

O app tem 4 abas independentes (cada uma com sua propria consulta SQL,
guardada num session_state proprio para nao se perder ao trocar de aba):
    1. Relatório de Movimentos  -> aba_relatorio  (BASE_QUERY)
    2. Consolidado por Lote     -> aba_consolidado (CONSOLIDADO_QUERY + SALDO_LOTE_QUERY)
    3. Busca por Lote           -> aba_busca       (BUSCA_LOTE_QUERY)
    4. Controle de Terceiros    -> aba_terceiros   (TERCEIROS_QUERY)

Escopo importante: TODAS as consultas filtram `B1_TIPO = 'PA'` (Produto
Acabado) direto no SQL - matéria-prima e outros tipos de produto nunca
aparecem em nenhum relatório, mesmo tendo controle de lote no Protheus.
Essa trava é permanente e proposital (decisão do usuário do app).

Como rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py

Configuracao da conexao:
    - Teste local rapido: copie .env.example para .env e preencha (ver README).
    - Producao / Streamlit Cloud: use .streamlit/secrets.toml (ver secrets.toml.example).
    O app tenta st.secrets primeiro e cai para variaveis de ambiente (.env) se
    a secao [sqlserver] nao existir em secrets.toml.
"""

import io
import os
from datetime import date, datetime, timedelta  # datas dos filtros e dos nomes de arquivo exportado

import altair as alt  # gráfico interativo (clicável) da aba 1
import pandas as pd  # DataFrames: resultado das consultas SQL e manipulação em memória
import streamlit as st  # framework da interface web (sidebar, abas, tabelas, botões, cache)
import streamlit.components.v1 as components  # usado só pelo botão "Sair do aplicativo" (roda um JS pequeno)
from dotenv import load_dotenv, dotenv_values  # le o arquivo .env (conexão local) para os.environ
from pathlib import Path  # caminho do config local salvo no computador de quem instalou o app
# --- ReportLab: biblioteca usada para montar os PDFs exportáveis ---
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT  # alinhamento de texto dentro de um Paragraph (cabeçalho e colunas numéricas do PDF)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth  # mede a largura de um texto numa fonte/tamanho (usado no cabeçalho vertical)
from reportlab.platypus import Flowable  # classe-base para desenhar elementos customizados no PDF (usada no cabeçalho vertical)
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
# --- SQLAlchemy: conexão com o SQL Server e execução das consultas parametrizadas ---
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine, URL

load_dotenv()  # le variaveis do arquivo .env (se existir) para os.environ


# --------------------------------------------------------------------------
# Config local (usada quando o app roda empacotado/instalado em outro PC,
# sem .env nem secrets.toml - ver tela de configuracao inicial mais abaixo)
# --------------------------------------------------------------------------
def _caminho_config_local() -> Path:
    """Pasta de dados do usuario no Windows (%APPDATA%), fora da pasta de
    instalacao - assim funciona mesmo instalado em Program Files (sem
    permissao de escrita) e sobrevive a uma reinstalacao/atualizacao do
    app. Fora do Windows, cai no diretorio home (uso em dev)."""
    base = os.getenv("APPDATA") or os.getenv("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "RelatorioMVPed" / "config.env"


def _config_local_valores() -> dict:
    caminho = _caminho_config_local()
    if not caminho.exists():
        return {}
    return dict(dotenv_values(caminho))


def _salvar_config_local(dados: dict) -> None:
    caminho = _caminho_config_local()
    caminho.parent.mkdir(parents=True, exist_ok=True)
    conteudo = "\n".join(f"{chave}={valor}" for chave, valor in dados.items())
    caminho.write_text(conteudo + "\n", encoding="utf-8")


# --------------------------------------------------------------------------
# Config embutida e criptografada (gerada por gerar_conn_criptografada.py,
# chamado automaticamente pelo build_exe.bat a partir do .env real, no
# momento de empacotar o instalador - ver GUIA_EMPACOTAMENTO.md). Evita
# deixar a senha em texto puro dentro da pasta instalada; nao e um cofre
# inviolavel (a chave viaja junto no codigo, senao o app nao conseguiria
# se conectar sozinho), mas impede a leitura casual num bloco de notas.
# --------------------------------------------------------------------------
_CHAVE_CONFIG_EMBUTIDA = b"CZ1ZqsuHru8OvAEQ9B0MPMd77M3PFZAQSgtxWEbDc54="


def _caminho_config_embutida() -> Path:
    """conn.enc fica ao lado de app.py dentro do pacote (mesma logica do
    LOGO_PATH) - tanto em dev (pasta do projeto, se alguem gerar um
    conn.enc manualmente) quanto empacotado (PyInstaller copia para o
    mesmo diretorio de app.py via --add-data "conn.enc;.")."""
    return Path(os.path.dirname(os.path.abspath(__file__))) / "conn.enc"


def _config_embutida_valores() -> dict:
    caminho = _caminho_config_embutida()
    if not caminho.exists():
        return {}
    try:
        from cryptography.fernet import Fernet

        conteudo = Fernet(_CHAVE_CONFIG_EMBUTIDA).decrypt(caminho.read_bytes())
    except Exception:
        # conn.enc corrompido, chave errada, ou biblioteca ausente - trata
        # como "sem config embutida" em vez de derrubar o app inteiro.
        return {}

    valores: dict = {}
    for linha in conteudo.decode("utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, _, valor = linha.partition("=")
        valores[chave.strip()] = valor.strip()
    return valores


# --------------------------------------------------------------------------
# Configuracao da pagina
# --------------------------------------------------------------------------
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

st.set_page_config(
    page_title="Rastreabilidade de Lote - BioCAZ",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "📦",
    layout="wide",
)

# Deixa todos os botoes do app (sidebar e abas) menores e com o mesmo
# tamanho visual, em vez de cada um esticar para a largura da coluna onde
# foi criado (o que fazia alguns ficarem bem maiores que outros).
st.markdown(
    """
    <style>
    div.stButton > button {
        padding: 0.25rem 0.9rem;
        font-size: 0.85rem;
        min-height: 2.1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if os.path.exists(LOGO_PATH):
    # st.logo() tem tamanho maximo fixo ("large" ~32px) - para uma logo maior
    # de verdade, exibimos a imagem direto no topo da barra lateral, com
    # largura customizada.
    st.sidebar.image(LOGO_PATH, width=260)

# Os 5 status possiveis calculados pela BASE_QUERY (coluna STATUS_RASTREABILIDADE)
# - usados no filtro da sidebar, no grafico da aba 1 e na legenda de cores:
#   OK - PRODUTO RASTREADO                       -> lote informado, tudo certo
#   ALERTA - PRODUTO RASTREADO SEM LOTE NO MOVIMENTO -> produto tem controle de
#                                                    lote mas o movimento nao
#                                                    tem lote nenhum vinculado
#   ALERTA - COBERTURA DE LOTE PARCIAL           -> só parte da quantidade do
#                                                    movimento tem lote (via SD5)
#   SEM CONTROLE DE LOTE                          -> produto nem tem rastreio (B1_RASTRO)
#   VERIFICAR                                     -> caso nao previsto nas regras acima
STATUS_OPCOES = [
    "OK - PRODUTO RASTREADO",
    "ALERTA - PRODUTO RASTREADO SEM LOTE NO MOVIMENTO",
    "ALERTA - COBERTURA DE LOTE PARCIAL",
    "SEM CONTROLE DE LOTE",
    "VERIFICAR",
]
TIPO_MOV_OPCOES = ["ENTRADA", "SAIDA"]

# Paleta da marca BioCAZ, aplicada em ordem alternada a cada categoria do
# gráfico (repete do início se houver mais categorias do que cores).
PALETA_MARCA = ["#43AA8A", "#0C1B7D", "#000000"]
CORES_STATUS = [PALETA_MARCA[i % len(PALETA_MARCA)] for i in range(len(STATUS_OPCOES))]


# --------------------------------------------------------------------------
# Conexao com o banco (SQL Server / Protheus)
# --------------------------------------------------------------------------
def _tem_secrets_sqlserver() -> bool:
    """
    Verifica com seguranca se existe secrets.toml com secao [sqlserver].
    Streamlit lanca StreamlitSecretNotFoundError ao usar `in st.secrets`
    quando NENHUM arquivo secrets.toml existe no sistema - por isso o try/except.
    """
    try:
        return "sqlserver" in st.secrets
    except Exception:
        return False


def config_disponivel() -> bool:
    """True se houver configuracao de conexao via secrets.toml, .env,
    config embutida e criptografada (conn.enc, empacotada pelo
    build_exe.bat) OU config local salvo pela tela de configuracao
    inicial (fallback, so entra em jogo se nenhuma das outras existir)."""
    tem_env = bool(os.getenv("SQLSERVER_SERVIDOR"))
    tem_config_embutida = bool(_config_embutida_valores().get("SQLSERVER_SERVIDOR"))
    tem_config_local = bool(_config_local_valores().get("SQLSERVER_SERVIDOR"))
    return _tem_secrets_sqlserver() or tem_env or tem_config_embutida or tem_config_local


def obter_config_conexao() -> tuple[dict, str]:
    """
    Retorna (config, origem). Prioriza st.secrets["sqlserver"]; se a secao
    nao existir (ou nao houver secrets.toml algum), usa variaveis de
    ambiente (carregadas do .env por load_dotenv()).
    """
    if _tem_secrets_sqlserver():
        cfg = dict(st.secrets["sqlserver"])
        return cfg, "secrets.toml"

    tem_env = bool(os.getenv("SQLSERVER_SERVIDOR"))
    if tem_env:
        fonte, origem = os.environ, ".env"
    else:
        embutida = _config_embutida_valores()
        if embutida.get("SQLSERVER_SERVIDOR"):
            fonte, origem = embutida, "instalador (criptografado)"
        else:
            fonte, origem = _config_local_valores(), "config local (1º uso)"

    cfg = {
        "servidor": fonte.get("SQLSERVER_SERVIDOR"),
        "porta": fonte.get("SQLSERVER_PORTA", "1433"),
        "database": fonte.get("SQLSERVER_DATABASE"),
        "usuario": fonte.get("SQLSERVER_USUARIO"),
        "senha": fonte.get("SQLSERVER_SENHA"),
        "driver": fonte.get("SQLSERVER_DRIVER", "ODBC Driver 18 for SQL Server"),
        "odbc_extra": fonte.get("SQLSERVER_ODBC_EXTRA", "TrustServerCertificate=yes;Encrypt=yes"),
    }
    faltando = [k for k in ("servidor", "database", "usuario", "senha") if not cfg.get(k)]
    if faltando:
        raise RuntimeError(
            "Configuracao de conexao incompleta (" + origem + "): faltam "
            + ", ".join(f"SQLSERVER_{k.upper()}" for k in faltando)
        )
    return cfg, origem


def _tela_configuracao_inicial() -> None:
    """Formulario de 1o uso: pede os dados de conexao com o SQL Server e
    salva em config local (fora da pasta do app - ver _caminho_config_local),
    para quem recebeu o app ja empacotado (instalador), sem .env nem
    secrets.toml. Encerra a execucao do script com st.stop() em qualquer
    caminho (formulario ainda nao enviado, ou enviado com sucesso -> rerun)."""
    st.title("📦 Relatório de Rastreabilidade de Lote")
    st.subheader("Configuração inicial da conexão com o banco")
    st.write(
        "Primeiro uso deste app neste computador. Preencha os dados de "
        "conexão com o SQL Server do Protheus (peça ao TI/administrador se "
        "não souber). Isso só precisa ser feito uma vez: os dados ficam "
        "salvos apenas neste computador (nunca dentro do instalador)."
    )
    with st.form("form_config_inicial"):
        servidor = st.text_input("Servidor", placeholder="ex: 181.41.163.164")
        porta = st.text_input("Porta", value="1433")
        database = st.text_input("Banco de dados (database)")
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        with st.expander("Avançado (normalmente não precisa mudar)"):
            driver = st.text_input("Driver ODBC", value="ODBC Driver 18 for SQL Server")
            odbc_extra = st.text_input(
                "Parâmetros extra do ODBC", value="TrustServerCertificate=yes;Encrypt=yes"
            )
        enviado = st.form_submit_button("Salvar e conectar", type="primary")

    if enviado:
        obrigatorios = {
            "Servidor": servidor,
            "Banco de dados": database,
            "Usuário": usuario,
            "Senha": senha,
        }
        faltando = [nome for nome, valor in obrigatorios.items() if not valor.strip()]
        if faltando:
            st.error("Preencha: " + ", ".join(faltando))
        else:
            _salvar_config_local(
                {
                    "SQLSERVER_SERVIDOR": servidor.strip(),
                    "SQLSERVER_PORTA": (porta.strip() or "1433"),
                    "SQLSERVER_DATABASE": database.strip(),
                    "SQLSERVER_USUARIO": usuario.strip(),
                    "SQLSERVER_SENHA": senha,
                    "SQLSERVER_DRIVER": (driver.strip() or "ODBC Driver 18 for SQL Server"),
                    "SQLSERVER_ODBC_EXTRA": odbc_extra.strip(),
                }
            )
            st.success("Configuração salva! Conectando...")
            st.rerun()
    st.stop()


@st.cache_resource(show_spinner=False)
def get_engine() -> Engine:
    """Cria (uma unica vez por sessao, gracas a @st.cache_resource) o Engine
    do SQLAlchemy usado por todas as consultas do app. Diferente de
    @st.cache_data, @st.cache_resource nao tenta serializar/copiar o
    retorno - certo para objetos como conexao de banco, que devem ser
    reaproveitados, nao recriados a cada rerun da pagina."""
    cfg, origem = obter_config_conexao()
    driver = cfg.get("driver") or "ODBC Driver 18 for SQL Server"
    odbc_extra = cfg.get("odbc_extra") or "TrustServerCertificate=yes;Encrypt=yes"

    # Monta os parametros extra (ex: "TrustServerCertificate=yes;Encrypt=yes")
    # como um dicionario, em vez de concatenar a string manualmente - isso
    # evita erros de parsing no driver ODBC e cuida de caracteres especiais
    # na senha automaticamente.
    query = {"driver": driver}
    for par in odbc_extra.split(";"):
        par = par.strip()
        if not par:
            continue
        chave, _, valor = par.partition("=")
        if chave:
            query[chave] = valor

    url = URL.create(
        "mssql+pyodbc",
        username=cfg["usuario"],
        password=cfg["senha"],
        host=cfg["servidor"],
        port=int(cfg.get("porta") or 1433),
        database=cfg["database"],
        query=query,
    )
    # pool_pre_ping=True: testa a conexao antes de reusa-la (evita erro apos
    # o SQL Server derrubar conexoes ociosas). fast_executemany=True: envia
    # varios parametros de uma vez ao driver ODBC (mais rapido, embora aqui
    # so façamos SELECTs).
    engine = create_engine(url, pool_pre_ping=True, fast_executemany=True)
    st.session_state["_conn_origem"] = origem  # exibido na sidebar ("Conexão via: ...")
    return engine


# --------------------------------------------------------------------------
# Consulta SQL (baseada na consulta original, com WHERE parametrizado)
# --------------------------------------------------------------------------
# Visao geral da BASE_QUERY (aba 1 - Relatório de Movimentos):
#   1. CTE "MOVIMENTOS" unifica entradas (SD1010, compras) e saidas (SD2010,
#      vendas) numa unica lista de movimentos, com as mesmas colunas.
#   2. O SELECT externo busca, para cada movimento, dados complementares via
#      OUTER APPLY (equivalente a um "LEFT JOIN LATERAL"):
#        - B1: cadastro do produto (SB1010) - descricao, tipo de rastreio e
#          tipo do produto (B1_TIPO), escolhendo a filial mais especifica.
#        - CF: nome do parceiro (cliente na SA1010 se for SAIDA, fornecedor
#          na SA2010 se for ENTRADA).
#        - D5: agregados da SD5010 (movimentacao de lotes) para o mesmo
#          documento - quantas linhas tem, quantas tem lote e a soma da
#          quantidade com lote (usado para detectar cobertura parcial).
#        - B8: contagem de registros/lotes na SB8010 (saldo de lote) do
#          mesmo produto+local - so um indicador auxiliar.
#   3. A coluna STATUS_RASTREABILIDADE cruza essas informacoes numa unica
#      classificacao (ver comentario nos STATUS_OPCOES acima para o
#      significado de cada valor).
#   4. O WHERE fixa a filial (:filial) e trava B1.B1_TIPO = 'PA' (só
#      Produto Acabado); {filtros_internos} e {filtros_status} sao
#      preenchidos dinamicamente por montar_query() com base nos filtros
#      escolhidos na sidebar.
BASE_QUERY = """
;WITH MOVIMENTOS AS
(
    ------------------------------------------------------------
    -- ENTRADAS - SD1
    ------------------------------------------------------------
    SELECT
        'ENTRADA'           AS TIPO_MOV,
        'COMPRA'            AS TIPO_PEDIDO,
        D1.D1_FILIAL        AS FILIAL,
        D1.D1_DOC           AS DOC,
        D1.D1_SERIE         AS SERIE,
        D1.D1_ITEM          AS ITEM,
        D1.D1_PEDIDO        AS PEDIDO,
        D1.D1_ITEMPC        AS ITEM_PEDIDO,
        D1.D1_COD           AS PRODUTO,
        D1.D1_LOCAL         AS LOCAL,
        D1.D1_FORNECE       AS CLIFOR,
        D1.D1_LOJA          AS LOJA,
        D1.D1_EMISSAO       AS DATA_MOV,
        D1.D1_QUANT         AS QUANTIDADE,
        D1.D1_LOTECTL       AS LOTE_DOCUMENTO,
        D1.D1_NUMLOTE       AS SUBLOTE_DOCUMENTO,
        D1.D1_DTVALID       AS VALIDADE_DOCUMENTO
    FROM SD1010 D1
    WHERE D1.D_E_L_E_T_ = ' '

    UNION ALL

    ------------------------------------------------------------
    -- SAIDAS - SD2
    ------------------------------------------------------------
    SELECT
        'SAIDA'             AS TIPO_MOV,
        'VENDA'             AS TIPO_PEDIDO,
        D2.D2_FILIAL        AS FILIAL,
        D2.D2_DOC           AS DOC,
        D2.D2_SERIE         AS SERIE,
        D2.D2_ITEM          AS ITEM,
        D2.D2_PEDIDO        AS PEDIDO,
        D2.D2_ITEMPV        AS ITEM_PEDIDO,
        D2.D2_COD           AS PRODUTO,
        D2.D2_LOCAL         AS LOCAL,
        D2.D2_CLIENTE       AS CLIFOR,
        D2.D2_LOJA          AS LOJA,
        D2.D2_EMISSAO       AS DATA_MOV,
        D2.D2_QUANT         AS QUANTIDADE,
        D2.D2_LOTECTL       AS LOTE_DOCUMENTO,
        D2.D2_NUMLOTE       AS SUBLOTE_DOCUMENTO,
        D2.D2_DTVALID       AS VALIDADE_DOCUMENTO
    FROM SD2010 D2
    WHERE D2.D_E_L_E_T_ = ' '
)
SELECT * FROM (
    SELECT
        M.FILIAL, M.TIPO_MOV, M.DATA_MOV, M.DOC, M.SERIE, M.ITEM,
        M.TIPO_PEDIDO, M.PEDIDO, M.ITEM_PEDIDO,
        M.PRODUTO, B1.B1_DESC AS DESCRICAO,
        M.LOCAL, M.CLIFOR, CF.NOME AS CLIFOR_DESCRICAO, CF.CGC AS CLIFOR_CNPJ, M.LOJA, M.QUANTIDADE,

        ISNULL(B1.B1_RASTRO, 'N') AS B1_RASTRO,
        CASE
            WHEN B1.B1_RASTRO = 'L' THEN 'LOTE'
            WHEN B1.B1_RASTRO = 'S' THEN 'SUBLOTE'
            ELSE 'NAO RASTREADO'
        END AS TIPO_RASTREABILIDADE,

        NULLIF(LTRIM(RTRIM(M.LOTE_DOCUMENTO)), '')    AS LOTE_DOCUMENTO,
        NULLIF(LTRIM(RTRIM(M.SUBLOTE_DOCUMENTO)), '') AS SUBLOTE_DOCUMENTO,
        M.VALIDADE_DOCUMENTO,

        ISNULL(D5.QTD_MOV_SD5, 0)      AS QTD_REGISTROS_SD5,
        ISNULL(D5.QTD_MOV_COM_LOTE, 0) AS QTD_REGISTROS_SD5_COM_LOTE,
        D5.LOTE_SD5,
        D5.SUBLOTE_SD5,

        ISNULL(B8.QTD_REGISTROS_SB8, 0) AS QTD_REGISTROS_SB8,
        ISNULL(B8.QTD_LOTES_SB8, 0)     AS QTD_LOTES_SB8,

        CASE
            WHEN NULLIF(LTRIM(RTRIM(M.LOTE_DOCUMENTO)), '') IS NOT NULL
                 OR ISNULL(D5.QTD_MOV_COM_LOTE, 0) > 0
            THEN 'SIM' ELSE 'NAO'
        END AS HOUVE_MOVIMENTACAO_LOTE,

        CASE
            WHEN ISNULL(B1.B1_RASTRO, 'N') NOT IN ('L', 'S')
            THEN 'SEM CONTROLE DE LOTE'

            -- lote informado direto no documento: assume cobertura total do movimento
            WHEN NULLIF(LTRIM(RTRIM(M.LOTE_DOCUMENTO)), '') IS NOT NULL
            THEN 'OK - PRODUTO RASTREADO'

            -- sem lote no documento, mas a soma das quantidades com lote na SD5
            -- cobre (ou excede) a quantidade total do movimento
            WHEN ISNULL(D5.QTD_COM_LOTE_SOMA, 0) > 0
                 AND ISNULL(D5.QTD_COM_LOTE_SOMA, 0) >= M.QUANTIDADE
            THEN 'OK - PRODUTO RASTREADO'

            -- tem alguma quantidade com lote na SD5, mas nao cobre o total do movimento
            WHEN ISNULL(D5.QTD_COM_LOTE_SOMA, 0) > 0
            THEN 'ALERTA - COBERTURA DE LOTE PARCIAL'

            -- nenhuma quantidade com lote em lugar nenhum
            WHEN ISNULL(D5.QTD_MOV_COM_LOTE, 0) = 0
            THEN 'ALERTA - PRODUTO RASTREADO SEM LOTE NO MOVIMENTO'

            ELSE 'VERIFICAR'
        END AS STATUS_RASTREABILIDADE

    FROM MOVIMENTOS M

    OUTER APPLY
    (
        SELECT TOP 1 SB1.B1_COD, SB1.B1_DESC, SB1.B1_RASTRO, SB1.B1_TIPO
        FROM SB1010 SB1
        WHERE SB1.D_E_L_E_T_ = ' ' AND SB1.B1_COD = M.PRODUTO
        ORDER BY
            CASE
                WHEN SB1.B1_FILIAL = LEFT(M.FILIAL, 4) THEN 1
                WHEN SB1.B1_FILIAL = '' THEN 2
                ELSE 3
            END
    ) B1

    ------------------------------------------------------------
    -- NOME DO CLIENTE (SA1, para SAIDA/VENDA) OU FORNECEDOR (SA2, para ENTRADA/COMPRA)
    ------------------------------------------------------------
    OUTER APPLY
    (
        SELECT TOP 1 X.NOME, X.CGC
        FROM
        (
            SELECT A2_NOME AS NOME, A2_CGC AS CGC, A2_FILIAL AS FILIAL_PARC
            FROM SA2010
            WHERE D_E_L_E_T_ = ' '
              AND M.TIPO_MOV = 'ENTRADA'
              AND A2_COD  = M.CLIFOR
              AND A2_LOJA = M.LOJA

            UNION ALL

            SELECT A1_NOME AS NOME, A1_CGC AS CGC, A1_FILIAL AS FILIAL_PARC
            FROM SA1010
            WHERE D_E_L_E_T_ = ' '
              AND M.TIPO_MOV = 'SAIDA'
              AND A1_COD  = M.CLIFOR
              AND A1_LOJA = M.LOJA
        ) X
        ORDER BY
            CASE
                WHEN X.FILIAL_PARC = LEFT(M.FILIAL, 4) THEN 1
                WHEN X.FILIAL_PARC = '' THEN 2
                ELSE 3
            END
    ) CF

    OUTER APPLY
    (
        SELECT
            COUNT(*) AS QTD_MOV_SD5,
            SUM(CASE WHEN NULLIF(LTRIM(RTRIM(D5.D5_LOTECTL)), '') IS NOT NULL THEN 1 ELSE 0 END) AS QTD_MOV_COM_LOTE,
            SUM(CASE WHEN NULLIF(LTRIM(RTRIM(D5.D5_LOTECTL)), '') IS NOT NULL THEN ABS(D5.D5_QUANT) ELSE 0 END) AS QTD_COM_LOTE_SOMA,
            MAX(NULLIF(LTRIM(RTRIM(D5.D5_LOTECTL)), '')) AS LOTE_SD5,
            MAX(NULLIF(LTRIM(RTRIM(D5.D5_NUMLOTE)), '')) AS SUBLOTE_SD5
        FROM SD5010 D5
        WHERE D5.D_E_L_E_T_ = ' '
          AND D5.D5_FILIAL  = M.FILIAL
          AND D5.D5_PRODUTO = M.PRODUTO
          AND D5.D5_DOC     = M.DOC
          AND D5.D5_SERIE   = M.SERIE
          AND D5.D5_CLIFOR  = M.CLIFOR
          AND D5.D5_LOJA    = M.LOJA
    ) D5

    OUTER APPLY
    (
        SELECT
            COUNT(*) AS QTD_REGISTROS_SB8,
            COUNT(DISTINCT NULLIF(LTRIM(RTRIM(B8.B8_LOTECTL)), '')) AS QTD_LOTES_SB8
        FROM SB8010 B8
        WHERE B8.D_E_L_E_T_ = ' '
          AND B8.B8_FILIAL  = M.FILIAL
          AND B8.B8_PRODUTO = M.PRODUTO
          AND B8.B8_LOCAL   = M.LOCAL
    ) B8

    WHERE M.FILIAL = :filial
      AND B1.B1_TIPO = 'PA'
    {filtros_internos}
) REL
WHERE 1 = 1
{filtros_status}
ORDER BY
    REL.DATA_MOV DESC,
    REL.PEDIDO,
    REL.PRODUTO,
    REL.DOC,
    REL.ITEM
"""


def montar_clausula_in(coluna: str, prefixo: str, valores: list[str]) -> tuple[str, dict]:
    """Monta 'coluna IN (:p0, :p1, ...)' com parametros nomeados (evita SQL injection)."""
    nomes = [f"{prefixo}_{i}" for i in range(len(valores))]
    clausula = f"AND {coluna} IN ({', '.join(':' + n for n in nomes)})"
    params = {n: v for n, v in zip(nomes, valores)}
    return clausula, params


def montar_query(filtros: dict) -> tuple[str, dict]:
    """Monta a BASE_QUERY final substituindo {filtros_internos} (filtros que
    entram DENTRO da subquery, antes de calcular STATUS_RASTREABILIDADE) e
    {filtros_status} (filtro pelo status ja calculado, aplicado por fora).
    So adiciona uma clausula de filtro quando o usuario de fato restringiu
    aquele campo - se nao, a consulta roda sem essa condicao."""
    params = {"filial": filtros["filial"]}
    filtros_internos = []
    filtros_status = []

    if filtros.get("dt_ini") and filtros.get("dt_fim"):
        filtros_internos.append("AND M.DATA_MOV BETWEEN :dt_ini AND :dt_fim")
        params["dt_ini"] = filtros["dt_ini"].strftime("%Y%m%d")
        params["dt_fim"] = filtros["dt_fim"].strftime("%Y%m%d")

    if filtros.get("produto"):
        filtros_internos.append("AND M.PRODUTO = :produto")
        params["produto"] = filtros["produto"].strip()

    if filtros.get("pedido"):
        filtros_internos.append("AND M.PEDIDO = :pedido")
        params["pedido"] = filtros["pedido"].strip()

    if filtros.get("tipo_mov") and len(filtros["tipo_mov"]) < len(TIPO_MOV_OPCOES):
        clausula, p = montar_clausula_in("M.TIPO_MOV", "tipomov", filtros["tipo_mov"])
        filtros_internos.append(clausula)
        params.update(p)

    if filtros.get("status") and len(filtros["status"]) < len(STATUS_OPCOES):
        clausula, p = montar_clausula_in("REL.STATUS_RASTREABILIDADE", "status", filtros["status"])
        filtros_status.append(clausula)
        params.update(p)

    sql = BASE_QUERY.format(
        filtros_internos="\n    ".join(filtros_internos),
        filtros_status="\n".join(filtros_status),
    )
    return sql, params


@st.cache_data(ttl=300, show_spinner="Consultando o banco de dados...")
def carregar_dados(_engine: Engine, filtros: dict) -> tuple[pd.DataFrame, datetime]:
    """Executa a consulta e devolve (DataFrame, horário da consulta).
    @st.cache_data guarda o resultado por 5 minutos (ttl=300) usando os
    argumentos como chave do cache - por isso o parametro se chama
    "_engine" (o underscore diz ao Streamlit para NAO tentar usar o engine
    como parte da chave, já que um objeto de conexão não é "hasheável" da
    forma que o cache precisa; quem varia a chave de fato é o dict
    `filtros`). O botão "Executar consulta" força uma nova busca limpando
    esse cache. O datetime.now() é capturado AQUI DENTRO, ou seja, só muda
    quando a consulta realmente roda no banco (cache miss) - é o que
    alimenta o aviso "Dados de HH:MM" na tela, para o usuário saber se o
    que está vendo é recente ou já passou dos 5 minutos de cache."""
    sql, params = montar_query(filtros)
    with _engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)
    return df, datetime.now()


# --------------------------------------------------------------------------
# Consulta: Visao consolidada por lote (agrupa SD1+SD2 por PRODUTO+LOTE)
# --------------------------------------------------------------------------
# Usada na aba 2 (Consolidado por Lote). Mesma ideia da BASE_QUERY (unifica
# entradas e saidas numa CTE), mas aqui o resultado e AGRUPADO por
# PRODUTO + LOTE + LOCAL em vez de listar cada movimento - o objetivo e
# responder "quanto entrou, quanto saiu e para quantos clientes diferentes
# esse lote foi vendido", nao mostrar cada documento individualmente. O
# saldo/validade "reais" (SB8) sao buscados a parte, em SALDO_LOTE_QUERY,
# e depois unidos (merge) em Python - ver comentario logo abaixo.
CONSOLIDADO_QUERY = """
;WITH MOVIMENTOS AS
(
    SELECT
        'ENTRADA' AS TIPO_MOV, D1.D1_FILIAL AS FILIAL, D1.D1_COD AS PRODUTO,
        D1.D1_LOCAL AS LOCAL, D1.D1_EMISSAO AS DATA_MOV, D1.D1_QUANT AS QUANTIDADE,
        D1.D1_DOC AS DOC, D1.D1_FORNECE AS CLIFOR,
        NULLIF(LTRIM(RTRIM(D1.D1_LOTECTL)), '') AS LOTE
    FROM SD1010 D1
    WHERE D1.D_E_L_E_T_ = ' '

    UNION ALL

    SELECT
        'SAIDA' AS TIPO_MOV, D2.D2_FILIAL AS FILIAL, D2.D2_COD AS PRODUTO,
        D2.D2_LOCAL AS LOCAL, D2.D2_EMISSAO AS DATA_MOV, D2.D2_QUANT AS QUANTIDADE,
        D2.D2_DOC AS DOC, D2.D2_CLIENTE AS CLIFOR,
        NULLIF(LTRIM(RTRIM(D2.D2_LOTECTL)), '') AS LOTE
    FROM SD2010 D2
    WHERE D2.D_E_L_E_T_ = ' '
)
SELECT
    M.PRODUTO,
    B1.B1_DESC AS DESCRICAO,
    M.LOTE,
    M.LOCAL,
    SUM(CASE WHEN M.TIPO_MOV = 'ENTRADA' THEN M.QUANTIDADE ELSE 0 END) AS TOTAL_ENTRADA,
    SUM(CASE WHEN M.TIPO_MOV = 'SAIDA'   THEN M.QUANTIDADE ELSE 0 END) AS TOTAL_SAIDA,
    COUNT(DISTINCT M.DOC) AS QTD_DOCUMENTOS,
    COUNT(DISTINCT CASE WHEN M.TIPO_MOV = 'SAIDA' THEN M.CLIFOR END) AS QTD_CLIENTES_DISTINTOS,
    MIN(M.DATA_MOV) AS PRIMEIRA_MOVIMENTACAO,
    MAX(M.DATA_MOV) AS ULTIMA_MOVIMENTACAO
FROM MOVIMENTOS M
OUTER APPLY
(
    SELECT TOP 1 SB1.B1_DESC, SB1.B1_TIPO
    FROM SB1010 SB1
    WHERE SB1.D_E_L_E_T_ = ' ' AND SB1.B1_COD = M.PRODUTO
    ORDER BY
        CASE
            WHEN SB1.B1_FILIAL = LEFT(M.FILIAL, 4) THEN 1
            WHEN SB1.B1_FILIAL = '' THEN 2
            ELSE 3
        END
) B1
WHERE M.FILIAL = :filial
  AND M.LOTE IS NOT NULL
  AND B1.B1_TIPO = 'PA'
  {filtro_produto}
GROUP BY M.PRODUTO, B1.B1_DESC, M.LOTE, M.LOCAL
ORDER BY MAX(M.DATA_MOV) DESC
"""

# Consulta isolada de saldo/validade por lote (SB8). Fica separada de proposito:
# se B8_SALDO/B8_DTVALID nao baterem com o ambiente, so essa parte falha, sem
# afetar o resto do app.
# Campos confirmados via listar_colunas.py SB8010: B8_SALDO (float) e
# B8_DTVALID (varchar, formato 'YYYYMMDD', ex: '20270702').
SALDO_LOTE_QUERY = """
SELECT
    B8_PRODUTO                           AS PRODUTO,
    B8_LOCAL                             AS LOCAL,
    NULLIF(LTRIM(RTRIM(B8_LOTECTL)), '') AS LOTE,
    B8_SALDO                             AS SALDO_ATUAL,
    B8_DTVALID                           AS VALIDADE
FROM SB8010
WHERE D_E_L_E_T_ = ' '
  AND B8_FILIAL = :filial
  AND NULLIF(LTRIM(RTRIM(B8_LOTECTL)), '') IS NOT NULL
"""


def montar_query_consolidado(filial: str, produto: str = "") -> tuple[str, dict]:
    params = {"filial": filial}
    filtro_produto = ""
    if produto:
        filtro_produto = "AND M.PRODUTO = :produto"
        params["produto"] = produto.strip()
    sql = CONSOLIDADO_QUERY.format(filtro_produto=filtro_produto)
    return sql, params


@st.cache_data(ttl=300, show_spinner="Consultando saldos por lote...")
def carregar_consolidado(_engine: Engine, filial: str, produto: str = "") -> tuple[pd.DataFrame, datetime]:
    """Devolve (DataFrame, horário da consulta) - ver docstring de
    carregar_dados() para o porquê do datetime.now() ficar dentro da
    função cacheada."""
    sql, params = montar_query_consolidado(filial, produto)
    with _engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)
    return df, datetime.now()


@st.cache_data(ttl=300, show_spinner=False)
def carregar_saldo_lotes(_engine: Engine, filial: str) -> pd.DataFrame:
    """Saldo atual e validade por lote (SB8), buscado a parte de proposito -
    ver comentario acima de SALDO_LOTE_QUERY."""
    with _engine.connect() as conn:
        return pd.read_sql(text(SALDO_LOTE_QUERY), conn, params={"filial": filial})


def status_validade(validade) -> str:
    """Classifica a validade de um lote (B8_DTVALID, varchar 'YYYYMMDD') em
    SEM VALIDADE / VENCIDO / VENCE EM ATE 30 DIAS / OK."""
    if validade is None or (isinstance(validade, float) and pd.isna(validade)):
        return "SEM VALIDADE"
    if not hasattr(validade, "strftime"):
        s = str(validade).strip()
        if not s or s in ("0", "00000000"):
            return "SEM VALIDADE"
    try:
        venc = pd.to_datetime(validade)
        if pd.isna(venc):
            return "SEM VALIDADE"
        venc = venc.date()
    except (ValueError, TypeError):
        return "SEM VALIDADE"
    if venc.year <= 1901:
        return "SEM VALIDADE"  # data "zerada" tipica do Protheus
    hoje = date.today()
    if venc < hoje:
        return "VENCIDO"
    if venc <= hoje + timedelta(days=30):
        return "VENCE EM ATE 30 DIAS"
    return "OK"


# --------------------------------------------------------------------------
# Consulta: Busca reversa por lote (de onde veio / para onde foi)
# --------------------------------------------------------------------------
# Usada na aba 3 (Busca por Lote). Recebe um numero de lote especifico
# (:lote) e devolve TODAS as entradas e saidas registradas para ele, com o
# nome do fornecedor (entrada) ou cliente (saida) de cada movimento - o
# cenario de uso e um recall: dado um lote, descobrir rapidamente de onde
# ele veio e para quem foi vendido.
BUSCA_LOTE_QUERY = """
;WITH MOVIMENTOS AS
(
    SELECT
        'ENTRADA' AS TIPO_MOV, 'COMPRA' AS TIPO_PEDIDO,
        D1.D1_FILIAL AS FILIAL, D1.D1_DOC AS DOC, D1.D1_SERIE AS SERIE,
        D1.D1_COD AS PRODUTO, D1.D1_LOCAL AS LOCAL,
        D1.D1_FORNECE AS CLIFOR, D1.D1_LOJA AS LOJA,
        D1.D1_EMISSAO AS DATA_MOV, D1.D1_QUANT AS QUANTIDADE,
        NULLIF(LTRIM(RTRIM(D1.D1_LOTECTL)), '') AS LOTE,
        NULLIF(LTRIM(RTRIM(D1.D1_NUMLOTE)), '') AS SUBLOTE
    FROM SD1010 D1
    WHERE D1.D_E_L_E_T_ = ' '

    UNION ALL

    SELECT
        'SAIDA' AS TIPO_MOV, 'VENDA' AS TIPO_PEDIDO,
        D2.D2_FILIAL AS FILIAL, D2.D2_DOC AS DOC, D2.D2_SERIE AS SERIE,
        D2.D2_COD AS PRODUTO, D2.D2_LOCAL AS LOCAL,
        D2.D2_CLIENTE AS CLIFOR, D2.D2_LOJA AS LOJA,
        D2.D2_EMISSAO AS DATA_MOV, D2.D2_QUANT AS QUANTIDADE,
        NULLIF(LTRIM(RTRIM(D2.D2_LOTECTL)), '') AS LOTE,
        NULLIF(LTRIM(RTRIM(D2.D2_NUMLOTE)), '') AS SUBLOTE
    FROM SD2010 D2
    WHERE D2.D_E_L_E_T_ = ' '
)
SELECT
    M.TIPO_MOV, M.TIPO_PEDIDO, M.DATA_MOV, M.DOC, M.SERIE,
    M.PRODUTO, B1.B1_DESC AS DESCRICAO, M.LOCAL, ARM.NNR_DESCRI AS ARMAZEM_DESCRICAO, M.QUANTIDADE,
    M.LOTE, M.SUBLOTE, CF.NOME AS CLIFOR_DESCRICAO, CF.CGC AS CLIFOR_CNPJ

FROM MOVIMENTOS M

OUTER APPLY
(
    SELECT TOP 1 SB1.B1_DESC, SB1.B1_TIPO
    FROM SB1010 SB1
    WHERE SB1.D_E_L_E_T_ = ' ' AND SB1.B1_COD = M.PRODUTO
    ORDER BY
        CASE
            WHEN SB1.B1_FILIAL = LEFT(M.FILIAL, 4) THEN 1
            WHEN SB1.B1_FILIAL = '' THEN 2
            ELSE 3
        END
) B1

-- Descrição do armazém (cadastro NNR010) para a coluna LOCAL - mesmo
-- padrão de filial "mais específica primeiro" usado no B1 acima. NNR_DESCRI
-- é o nome padrão desse campo no Protheus; se o ambiente usar outro nome,
-- só essa coluna (e a aba de Busca por Lote) mostra erro, sem afetar as
-- outras abas.
OUTER APPLY
(
    SELECT TOP 1 NNR.NNR_DESCRI
    FROM NNR010 NNR
    WHERE NNR.D_E_L_E_T_ = ' ' AND NNR.NNR_CODIGO = M.LOCAL
    ORDER BY
        CASE
            WHEN NNR.NNR_FILIAL = LEFT(M.FILIAL, 4) THEN 1
            WHEN NNR.NNR_FILIAL = '' THEN 2
            ELSE 3
        END
) ARM

OUTER APPLY
(
    SELECT TOP 1 X.NOME, X.CGC
    FROM
    (
        SELECT A2_NOME AS NOME, A2_CGC AS CGC, A2_FILIAL AS FILIAL_PARC
        FROM SA2010
        WHERE D_E_L_E_T_ = ' ' AND M.TIPO_MOV = 'ENTRADA'
          AND A2_COD = M.CLIFOR AND A2_LOJA = M.LOJA

        UNION ALL

        SELECT A1_NOME AS NOME, A1_CGC AS CGC, A1_FILIAL AS FILIAL_PARC
        FROM SA1010
        WHERE D_E_L_E_T_ = ' ' AND M.TIPO_MOV = 'SAIDA'
          AND A1_COD = M.CLIFOR AND A1_LOJA = M.LOJA
    ) X
    ORDER BY
        CASE
            WHEN X.FILIAL_PARC = LEFT(M.FILIAL, 4) THEN 1
            WHEN X.FILIAL_PARC = '' THEN 2
            ELSE 3
        END
) CF

WHERE M.FILIAL = :filial
  AND M.LOTE = :lote
  AND B1.B1_TIPO = 'PA'
ORDER BY M.DATA_MOV, M.TIPO_MOV
"""


@st.cache_data(ttl=300, show_spinner="Buscando movimentações do lote...")
def carregar_busca_lote(_engine: Engine, filial: str, lote: str) -> tuple[pd.DataFrame, datetime]:
    """Devolve (DataFrame, horário da consulta) - ver docstring de
    carregar_dados() para o porquê do datetime.now() ficar dentro da
    função cacheada."""
    with _engine.connect() as conn:
        df = pd.read_sql(text(BUSCA_LOTE_QUERY), conn, params={"filial": filial, "lote": lote.strip()})
    return df, datetime.now()


# --------------------------------------------------------------------------
# Consulta: Controle de Terceiros (remessas e devolucoes - SB6)
# --------------------------------------------------------------------------
# Usada na aba 4 (Controle de Terceiros). Baseada na SB6010 (remessas para
# terceiros / de terceiros), diferente das outras abas que partem de
# SD1/SD2. Cada linha e um documento de remessa ou devolucao; os CASE
# abaixo traduzem os codigos do Protheus para texto legivel:
#   B6_TPCF  ('C'/'F')      -> TIPO_CLI_FOR / RAZAO_SOCIAL (busca nome em SA1 ou SA2)
#   B6_TIPO  ('E'/'D')      -> DE_EM_TERCEIROS ("EM TERCEIROS" = saiu daqui / "DE TERCEIROS" = veio de fora)
#   B6_PODER3 ('R'/'D')     -> MOVIMENTO ("REMESSA"/"DEVOLUCAO")
#   SITUACAO (calculada)    -> DEVOLUCAO (se o proprio movimento for devolucao),
#                              senão ATENDIDO (saldo zerado) / PARCIAL (saldo <
#                              quantidade original) / EM ABERTO (saldo = quantidade)
TERCEIROS_TIPO_CLIFOR_OPCOES = ["Cliente", "Fornecedor"]
TERCEIROS_DE_EM_OPCOES = ["DE TERCEIROS", "EM TERCEIROS"]
TERCEIROS_SITUACAO_OPCOES = ["EM ABERTO", "PARCIAL", "ATENDIDO", "DEVOLUCAO"]

TERCEIROS_QUERY = """
SELECT
    R.B6_FILIAL AS FILIAL,
    RTRIM(R.B6_PRODUTO) AS PRODUTO,
    RTRIM(PROD.B1_DESC) AS DESCRICAO,
    NULLIF(LTRIM(RTRIM(LOTEDOC.LOTE)), '') AS LOTE,
    RTRIM(R.B6_UM) AS UNID_MEDIDA,
    RTRIM(PROD.B1_GRUPO) AS GRUPO,
    RTRIM(GRP.BM_DESC) AS DESC_GRUPO,
    RTRIM(R.B6_LOCAL) AS ARMAZEM,
    RTRIM(R.B6_CLIFOR) AS CLIENTE_FORN,
    RTRIM(R.B6_LOJA) AS LOJA,
    CASE
        WHEN R.B6_TPCF = 'C' THEN RTRIM(CLI.A1_NOME)
        WHEN R.B6_TPCF = 'F' THEN RTRIM(FORN.A2_NOME)
        ELSE ''
    END AS RAZAO_SOCIAL,
    CASE
        WHEN R.B6_TPCF = 'C' THEN RTRIM(CLI.A1_CGC)
        WHEN R.B6_TPCF = 'F' THEN RTRIM(FORN.A2_CGC)
        ELSE ''
    END AS CNPJ,
    CASE
        WHEN R.B6_TPCF = 'C' THEN RTRIM(CLI.A1_MUN)
        WHEN R.B6_TPCF = 'F' THEN RTRIM(FORN.A2_MUN)
        ELSE ''
    END AS CIDADE,
    CASE
        WHEN R.B6_TPCF = 'C' THEN RTRIM(CLI.A1_EST)
        WHEN R.B6_TPCF = 'F' THEN RTRIM(FORN.A2_EST)
        ELSE ''
    END AS ESTADO,
    RTRIM(R.B6_TPCF) AS TIPO_CLI_FOR,
    RTRIM(R.B6_DOC) AS DOC_ORIGINAL,
    RTRIM(R.B6_SERIE) AS SERIE,
    TRY_CONVERT(DATE, NULLIF(RTRIM(R.B6_EMISSAO), ''), 112) AS DT_EMISSAO,
    R.B6_QUANT AS QUANTIDADE,
    R.B6_SALDO AS SALDO,
    R.B6_PRUNIT AS PRECO_UNITARIO,
    CAST(ROUND(R.B6_QUANT * R.B6_PRUNIT, 2) AS DECIMAL(18, 2)) AS TOTAL_NF,
    RTRIM(R.B6_TES) AS TES,
    RTRIM(R.B6_IDENT) AS IDENTIFICADOR,
    CASE R.B6_TIPO
        WHEN 'E' THEN 'EM TERCEIROS'
        WHEN 'D' THEN 'DE TERCEIROS'
        ELSE RTRIM(R.B6_TIPO)
    END AS DE_EM_TERCEIROS,
    CASE R.B6_PODER3
        WHEN 'R' THEN 'REMESSA'
        WHEN 'D' THEN 'DEVOLUCAO'
        ELSE RTRIM(R.B6_PODER3)
    END AS MOVIMENTO,
    CASE
        WHEN R.B6_PODER3 = 'D' THEN 'DEVOLUCAO'
        WHEN R.B6_SALDO = 0 THEN 'ATENDIDO'
        WHEN R.B6_SALDO < R.B6_QUANT THEN 'PARCIAL'
        ELSE 'EM ABERTO'
    END AS SITUACAO

FROM SB6010 R

LEFT JOIN SB1010 PROD
       ON PROD.B1_FILIAL = LEFT(R.B6_FILIAL, 4)
      AND PROD.B1_COD    = R.B6_PRODUTO
      AND PROD.D_E_L_E_T_ = ' '

LEFT JOIN SA1010 CLI
       ON CLI.A1_FILIAL = LEFT(R.B6_FILIAL, 4)
      AND CLI.A1_COD    = R.B6_CLIFOR
      AND CLI.A1_LOJA   = R.B6_LOJA
      AND CLI.D_E_L_E_T_ = ' '

LEFT JOIN SA2010 FORN
       ON FORN.A2_FILIAL = LEFT(R.B6_FILIAL, 4)
      AND FORN.A2_COD    = R.B6_CLIFOR
      AND FORN.A2_LOJA   = R.B6_LOJA
      AND FORN.D_E_L_E_T_ = ' '

-- A SB6010 (remessa/devolucao para terceiros) nao guarda o lote diretamente
-- (confirmado via listar_colunas.py SB6010 - nao existe B6_LOTECTL nem
-- parecido nessa tabela, e tambem nao existe um numero de item na SB6010
-- que ligue direto a um item especifico da nota fiscal). O lote fica no
-- documento de origem (a nota fiscal que gerou essa remessa/devolucao),
-- entao buscamos em SD1010 (documentos de entrada) ou SD2010 (documentos
-- de saida) casando pelo mesmo documento/serie/produto/parceiro/loja/local.
-- Quando a mesma nota tem mais de um item do mesmo produto (cada item com
-- um lote diferente - ja vimos isso acontecer aqui: TESTELOTE5 x
-- TESTELOTE6 do mesmo produto/documento), so documento+produto+parceiro
-- nao e suficiente pra saber qual item e qual: por isso a quantidade
-- tambem entra no casamento (cada item costuma ter uma quantidade
-- diferente), o que resolve a maioria dos casos - mas se dois itens do
-- mesmo produto, no mesmo documento, tiverem a MESMA quantidade, ainda
-- pode dar ambiguidade (sem um numero de item na SB6010 nao tem como
-- garantir 100% - se acontecer, o LOTE dessa linha pode sair errado).
OUTER APPLY
(
    SELECT TOP 1 X.LOTE
    FROM
    (
        SELECT D1.D1_LOTECTL AS LOTE
        FROM SD1010 D1
        WHERE D1.D_E_L_E_T_ = ' '
          AND D1.D1_FILIAL  = R.B6_FILIAL
          AND D1.D1_DOC     = R.B6_DOC
          AND D1.D1_SERIE   = R.B6_SERIE
          AND D1.D1_COD     = R.B6_PRODUTO
          AND D1.D1_FORNECE = R.B6_CLIFOR
          AND D1.D1_LOJA    = R.B6_LOJA
          AND D1.D1_LOCAL   = R.B6_LOCAL
          AND D1.D1_QUANT   = R.B6_QUANT

        UNION ALL

        SELECT D2.D2_LOTECTL AS LOTE
        FROM SD2010 D2
        WHERE D2.D_E_L_E_T_ = ' '
          AND D2.D2_FILIAL  = R.B6_FILIAL
          AND D2.D2_DOC     = R.B6_DOC
          AND D2.D2_SERIE   = R.B6_SERIE
          AND D2.D2_COD     = R.B6_PRODUTO
          AND D2.D2_CLIENTE = R.B6_CLIFOR
          AND D2.D2_LOJA    = R.B6_LOJA
          AND D2.D2_LOCAL   = R.B6_LOCAL
          AND D2.D2_QUANT   = R.B6_QUANT
    ) X
) LOTEDOC

OUTER APPLY
(
    SELECT TOP 1
        BM.BM_DESC
    FROM SBM010 BM
    WHERE BM.D_E_L_E_T_ = ' '
      AND BM.BM_GRUPO = PROD.B1_GRUPO
    ORDER BY
        CASE
            WHEN BM.BM_FILIAL = LEFT(R.B6_FILIAL, 4) THEN 1
            WHEN RTRIM(BM.BM_FILIAL) = '' THEN 2
            ELSE 3
        END,
        BM.R_E_C_N_O_
) GRP

WHERE
    R.D_E_L_E_T_ = ' '
    AND R.B6_FILIAL = :filial
    AND PROD.B1_TIPO = 'PA'
    {filtros_terceiros}

ORDER BY
    R.B6_PRODUTO,
    R.B6_IDENT,
    CASE R.B6_PODER3
        WHEN 'R' THEN 1
        WHEN 'D' THEN 2
        ELSE 3
    END,
    R.B6_DOC
"""


def montar_query_terceiros(filtros: dict) -> tuple[str, dict]:
    params = {"filial": filtros["filial"]}
    filtros_sql = []

    if filtros.get("dt_ini") and filtros.get("dt_fim"):
        filtros_sql.append("AND R.B6_EMISSAO BETWEEN :dt_ini AND :dt_fim")
        params["dt_ini"] = filtros["dt_ini"].strftime("%Y%m%d")
        params["dt_fim"] = filtros["dt_fim"].strftime("%Y%m%d")

    if filtros.get("produto"):
        filtros_sql.append("AND R.B6_PRODUTO = :produto")
        params["produto"] = filtros["produto"].strip()

    sql = TERCEIROS_QUERY.format(filtros_terceiros="\n    ".join(filtros_sql))
    return sql, params


@st.cache_data(ttl=300, show_spinner="Consultando controle de terceiros...")
def carregar_terceiros(_engine: Engine, filtros: dict) -> tuple[pd.DataFrame, datetime]:
    """Devolve (DataFrame, horário da consulta) - ver docstring de
    carregar_dados() para o porquê do datetime.now() ficar dentro da
    função cacheada."""
    sql, params = montar_query_terceiros(filtros)
    with _engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)
    return df, datetime.now()


# --------------------------------------------------------------------------
# Consulta extra: vendas/compras (SD2/SD1) dos MESMOS lotes que aparecem no
# Controle de Terceiros. Um lote remetido para um terceiro tambem pode ter
# sido vendido/comprado normalmente, fora do fluxo de terceiros - sem isso,
# esse outro movimento ficava invisivel nessa tela. Reaproveita o mesmo
# padrao de OUTER APPLY (nome/CNPJ/cidade/estado) ja usado em BASE_QUERY e
# BUSCA_LOTE_QUERY. O "{lotes_in}" e preenchido em
# montar_query_movimentos_lote() com uma lista de parametros nomeados
# (:lote0, :lote1, ...), ja que o pyodbc/SQLAlchemy usados aqui nao tem um
# jeito direto de passar uma lista Python como um unico parametro "IN".
TERCEIROS_MOVIMENTOS_QUERY = """
;WITH MOVIMENTOS AS
(
    SELECT
        'COMPRA' AS MOVIMENTO,
        D1.D1_FILIAL AS FILIAL, D1.D1_DOC AS DOC, D1.D1_SERIE AS SERIE,
        D1.D1_COD AS PRODUTO, D1.D1_FORNECE AS CLIFOR, D1.D1_LOJA AS LOJA,
        'F' AS TIPO_CLI_FOR,
        D1.D1_EMISSAO AS DATA_MOV, D1.D1_QUANT AS QUANTIDADE,
        D1.D1_LOCAL AS LOCAL, D1.D1_UM AS UNID_MEDIDA, D1.D1_TES AS TES,
        D1.D1_VUNIT AS PRECO_UNITARIO, D1.D1_TOTAL AS TOTAL_NF,
        NULLIF(LTRIM(RTRIM(D1.D1_LOTECTL)), '') AS LOTE
    FROM SD1010 D1
    WHERE D1.D_E_L_E_T_ = ' '
      AND D1.D1_FILIAL = :filial
      AND NULLIF(LTRIM(RTRIM(D1.D1_LOTECTL)), '') IN ({lotes_in})

    UNION ALL

    SELECT
        'VENDA' AS MOVIMENTO,
        D2.D2_FILIAL AS FILIAL, D2.D2_DOC AS DOC, D2.D2_SERIE AS SERIE,
        D2.D2_COD AS PRODUTO, D2.D2_CLIENTE AS CLIFOR, D2.D2_LOJA AS LOJA,
        'C' AS TIPO_CLI_FOR,
        D2.D2_EMISSAO AS DATA_MOV, D2.D2_QUANT AS QUANTIDADE,
        D2.D2_LOCAL AS LOCAL, D2.D2_UM AS UNID_MEDIDA, D2.D2_TES AS TES,
        D2.D2_PRCVEN AS PRECO_UNITARIO, D2.D2_TOTAL AS TOTAL_NF,
        NULLIF(LTRIM(RTRIM(D2.D2_LOTECTL)), '') AS LOTE
    FROM SD2010 D2
    WHERE D2.D_E_L_E_T_ = ' '
      AND D2.D2_FILIAL = :filial
      AND NULLIF(LTRIM(RTRIM(D2.D2_LOTECTL)), '') IN ({lotes_in})
)
SELECT
    M.MOVIMENTO, M.FILIAL, M.PRODUTO, RTRIM(PROD.B1_DESC) AS DESCRICAO, M.LOTE,
    RTRIM(M.UNID_MEDIDA) AS UNID_MEDIDA, RTRIM(PROD.B1_GRUPO) AS GRUPO, RTRIM(GRP.BM_DESC) AS DESC_GRUPO,
    RTRIM(M.LOCAL) AS ARMAZEM, RTRIM(M.CLIFOR) AS CLIENTE_FORN, RTRIM(M.LOJA) AS LOJA, M.TIPO_CLI_FOR,
    CF.NOME AS RAZAO_SOCIAL, CF.CGC AS CNPJ, CF.MUN AS CIDADE, CF.EST AS ESTADO,
    RTRIM(M.DOC) AS DOC_ORIGINAL, RTRIM(M.SERIE) AS SERIE,
    TRY_CONVERT(DATE, NULLIF(RTRIM(M.DATA_MOV), ''), 112) AS DT_EMISSAO,
    M.QUANTIDADE, RTRIM(M.TES) AS TES, M.PRECO_UNITARIO, M.TOTAL_NF

FROM MOVIMENTOS M

LEFT JOIN SB1010 PROD
       ON PROD.B1_FILIAL = LEFT(M.FILIAL, 4)
      AND PROD.B1_COD    = M.PRODUTO
      AND PROD.D_E_L_E_T_ = ' '

OUTER APPLY
(
    SELECT TOP 1
        BM.BM_DESC
    FROM SBM010 BM
    WHERE BM.D_E_L_E_T_ = ' '
      AND BM.BM_GRUPO = PROD.B1_GRUPO
    ORDER BY
        CASE
            WHEN BM.BM_FILIAL = LEFT(M.FILIAL, 4) THEN 1
            WHEN RTRIM(BM.BM_FILIAL) = '' THEN 2
            ELSE 3
        END,
        BM.R_E_C_N_O_
) GRP

OUTER APPLY
(
    SELECT TOP 1 X.NOME, X.CGC, X.MUN, X.EST
    FROM
    (
        SELECT A2_NOME AS NOME, A2_CGC AS CGC, A2_MUN AS MUN, A2_EST AS EST, A2_FILIAL AS FILIAL_PARC
        FROM SA2010
        WHERE D_E_L_E_T_ = ' ' AND M.TIPO_CLI_FOR = 'F'
          AND A2_COD = M.CLIFOR AND A2_LOJA = M.LOJA

        UNION ALL

        SELECT A1_NOME AS NOME, A1_CGC AS CGC, A1_MUN AS MUN, A1_EST AS EST, A1_FILIAL AS FILIAL_PARC
        FROM SA1010
        WHERE D_E_L_E_T_ = ' ' AND M.TIPO_CLI_FOR = 'C'
          AND A1_COD = M.CLIFOR AND A1_LOJA = M.LOJA
    ) X
    ORDER BY
        CASE
            WHEN X.FILIAL_PARC = LEFT(M.FILIAL, 4) THEN 1
            WHEN X.FILIAL_PARC = '' THEN 2
            ELSE 3
        END
) CF

WHERE PROD.B1_TIPO = 'PA'
ORDER BY M.DATA_MOV
"""


def montar_query_movimentos_lote(filial: str, lotes: list[str]) -> tuple[str, dict]:
    """Monta a query TERCEIROS_MOVIMENTOS_QUERY com um parametro nomeado
    para cada lote (:lote0, :lote1, ...) no lugar do "{lotes_in}"."""
    lotes_unicos = [l for l in dict.fromkeys(lotes) if l]
    params = {"filial": filial}
    nomes_params = []
    for i, lote in enumerate(lotes_unicos):
        chave = f"lote{i}"
        params[chave] = lote
        nomes_params.append(f":{chave}")
    sql = TERCEIROS_MOVIMENTOS_QUERY.format(lotes_in=", ".join(nomes_params))
    return sql, params


@st.cache_data(ttl=300, show_spinner="Buscando vendas/compras dos lotes...")
def carregar_movimentos_lote(_engine: Engine, filial: str, lotes: tuple[str, ...]) -> pd.DataFrame:
    """Vendas/compras (fora do fluxo de terceiros) dos lotes informados.
    `lotes` precisa ser uma tuple (e nao uma list) para o Streamlit
    conseguir usar como parte da chave do cache."""
    if not lotes:
        return pd.DataFrame()
    sql, params = montar_query_movimentos_lote(filial, list(lotes))
    with _engine.connect() as conn:
        return pd.read_sql(text(sql), conn, params=params)


def gerar_excel(df: pd.DataFrame) -> bytes:
    """Gera um .xlsx em memoria (sem gravar em disco) a partir do DataFrame,
    com a largura de cada coluna ajustada automaticamente pelo maior valor
    entre o nome da coluna e o conteudo mais longo dela (limitada a 40).
    Usado pelos 4 botões "Baixar Excel"."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Rastreabilidade")
        ws = writer.sheets["Rastreabilidade"]
        for i, col in enumerate(df.columns):
            largura = min(max(len(str(col)), df[col].astype(str).str.len().max() if len(df) else 10) + 2, 40)
            ws.set_column(i, i, largura)
    return buffer.getvalue()


class _CabecalhoVerticalPDF(Flowable):
    """Celula de cabecalho com o texto na vertical (rotacionado 90°) - usada
    na tabela do PDF para caber nomes de coluna longos sem quebrar palavra
    no meio, mesmo com muitas colunas na mesma pagina."""

    def __init__(self, texto: str, tamanho: float = 7):
        Flowable.__init__(self)
        self.texto = texto
        self.tamanho = tamanho
        self._w = 0.0
        self._h = 0.0

    def wrap(self, largura_disponivel, altura_disponivel):
        self._w = largura_disponivel
        self._h = altura_disponivel
        return largura_disponivel, altura_disponivel

    def draw(self):
        canvas_obj = self.canv
        canvas_obj.saveState()
        canvas_obj.setFont("Helvetica-Bold", self.tamanho)
        canvas_obj.setFillColor(colors.white)
        largura_texto = stringWidth(self.texto, "Helvetica-Bold", self.tamanho)
        x = self._w / 2 + self.tamanho * 0.35
        y = max(2, (self._h - largura_texto) / 2)
        canvas_obj.translate(x, y)
        canvas_obj.rotate(90)
        canvas_obj.drawString(0, 0, self.texto)
        canvas_obj.restoreState()


def gerar_pdf(
    df: pd.DataFrame, titulo: str, subtitulo: str = "", kpis: list[tuple[str, object]] | None = None
) -> bytes:
    """Gera um PDF em paisagem com a logo da BioCAZ, titulo/subtitulo, um
    resumo de KPIs (opcional - mesmos números mostrados na tela) e a
    tabela de dados (cabecalho repetido em cada pagina, colunas ajustadas
    para caber na largura da folha, colunas numéricas alinhadas à
    direita)."""
    buffer = io.BytesIO()
    largura_pagina, altura_pagina = landscape(A4)
    margem = 12 * mm
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=margem,
        rightMargin=margem,
        topMargin=margem,
        bottomMargin=16 * mm,  # 2mm a mais que antes, para caber a 2ª linha do rodapé
    )

    verde_marca, azul_marca = PALETA_MARCA[0], PALETA_MARCA[1]
    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloBioCAZ",
        parent=estilos["Heading1"],
        fontSize=15,
        leading=18,
        textColor=colors.HexColor(azul_marca),
        spaceAfter=2,
    )
    estilo_subtitulo = ParagraphStyle(
        "SubtituloBioCAZ", parent=estilos["Normal"], fontSize=8.5, leading=11, textColor=colors.HexColor("#444444")
    )
    estilo_celula = ParagraphStyle("CelulaPDF", parent=estilos["Normal"], fontSize=6.5, leading=8)
    # Mesmo estilo da célula normal, só que alinhado à direita - usado nas
    # colunas numéricas (quantidade, saldo, preço, total). O comando
    # "ALIGN" do TableStyle não alinha o TEXTO de dentro de um Paragraph
    # (só posiciona o Paragraph inteiro dentro da célula) - o alinhamento
    # do texto em si precisa vir do próprio estilo do Paragraph.
    estilo_celula_num = ParagraphStyle("CelulaPDFNum", parent=estilo_celula, alignment=TA_RIGHT)
    # Estilos do "cartão" de KPI (rótulo pequeno em cima, valor grande em
    # baixo, igual ao st.metric() da tela).
    estilo_kpi_rotulo = ParagraphStyle(
        "KpiRotulo",
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#5a5a5a"),
        alignment=TA_CENTER,
    )
    estilo_kpi_valor = ParagraphStyle(
        "KpiValor",
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=17,
        textColor=colors.HexColor(azul_marca),
        alignment=TA_CENTER,
    )
    tamanho_cabecalho_tabela = 7

    elementos = []

    bloco_titulo = [Paragraph(titulo, estilo_titulo)]
    if subtitulo:
        bloco_titulo.append(Paragraph(subtitulo, estilo_subtitulo))
    bloco_titulo.append(Paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_subtitulo))

    if os.path.exists(LOGO_PATH):
        try:
            logo = RLImage(LOGO_PATH, width=34 * mm, height=17 * mm, kind="proportional")
            cabecalho = Table([[logo, bloco_titulo]], colWidths=[38 * mm, None])
            cabecalho.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (0, 0), 0),
                        ("LEFTPADDING", (1, 0), (1, 0), 6),
                    ]
                )
            )
            elementos.append(cabecalho)
        except Exception:
            elementos.extend(bloco_titulo)
    else:
        elementos.extend(bloco_titulo)

    elementos.append(Spacer(1, 4 * mm))

    # Resumo de KPIs (opcional) - os mesmos números que aparecem nos
    # st.metric() da tela, em formato de "cartões" lado a lado, logo abaixo
    # do cabeçalho. Assim o PDF fica autossuficiente: quem recebe o
    # arquivo por e-mail entende o panorama geral sem precisar abrir o
    # app. Cada item de `kpis` é uma tupla (rótulo, valor já formatado).
    if kpis:
        largura_disponivel_kpi = largura_pagina - 2 * margem
        largura_cartao = largura_disponivel_kpi / len(kpis)
        linha_kpi = [
            [Paragraph(str(rotulo), estilo_kpi_rotulo), Paragraph(str(valor), estilo_kpi_valor)]
            for rotulo, valor in kpis
        ]
        tabela_kpi = Table([linha_kpi], colWidths=[largura_cartao] * len(kpis))
        tabela_kpi.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F3")),
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D8DEDC")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#D8DEDC")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elementos.append(tabela_kpi)
        elementos.append(Spacer(1, 5 * mm))

    if df.empty:
        elementos.append(Paragraph("Nenhum registro encontrado.", estilo_subtitulo))
    else:
        # Colunas numéricas do DataFrame ORIGINAL (antes do fillna(""), que
        # converteria tudo para string e apagaria essa informação) - usadas
        # para alinhar essas colunas à direita na tabela, em vez de à
        # esquerda junto com o texto.
        colunas_numericas = {i for i, col in enumerate(df.columns) if pd.api.types.is_numeric_dtype(df[col])}

        df_pdf = df.fillna("")
        cabecalho_tabela = [
            _CabecalhoVerticalPDF(str(c), tamanho_cabecalho_tabela) for c in df_pdf.columns
        ]
        linhas = [
            [
                Paragraph(str(v), estilo_celula_num if i in colunas_numericas else estilo_celula)
                for i, v in enumerate(linha)
            ]
            for linha in df_pdf.itertuples(index=False, name=None)
        ]
        dados_tabela = [cabecalho_tabela] + linhas

        # Altura do cabecalho: cabe o nome de coluna mais longo na vertical,
        # sem depender da largura da coluna (evita quebrar palavra no meio
        # quando ha muitas colunas na mesma pagina).
        altura_cabecalho = max(
            stringWidth(str(c), "Helvetica-Bold", tamanho_cabecalho_tabela) for c in df_pdf.columns
        ) + 8
        altura_cabecalho = min(max(altura_cabecalho, 20 * mm), 65 * mm)

        # Largura das colunas: baseada so no conteudo dos dados (o cabecalho
        # agora e vertical e nao precisa de largura extra).
        largura_disponivel = largura_pagina - 2 * margem
        pesos = []
        for col in df_pdf.columns:
            tam_col = int(df_pdf[col].astype(str).str.len().mean()) if len(df_pdf) else 8
            pesos.append(max(min(tam_col, 35), 6))
        soma_pesos = sum(pesos) or 1
        larguras = [max(11 * mm, min(45 * mm, largura_disponivel * (p / soma_pesos))) for p in pesos]
        fator_ajuste = largura_disponivel / sum(larguras)
        larguras = [w * fator_ajuste for w in larguras]

        tabela = Table(
            dados_tabela,
            colWidths=larguras,
            rowHeights=[altura_cabecalho] + [None] * len(linhas),
            repeatRows=1,
        )

        comandos_estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(verde_marca)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#DDDDDD")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F5F3")]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]

        # Mesmo destaque visual da tela (vermelho/amarelo) para status de
        # rastreabilidade ou de validade, quando a coluna existir.
        col_destaque = next(
            (c for c in ("STATUS_VALIDADE", "STATUS_RASTREABILIDADE", "SITUACAO") if c in df_pdf.columns), None
        )
        if col_destaque:
            for idx_linha, valor in enumerate(df_pdf[col_destaque].astype(str), start=1):
                if valor == "VENCIDO" or valor.startswith("ALERTA"):
                    comandos_estilo.append(("BACKGROUND", (0, idx_linha), (-1, idx_linha), colors.HexColor("#ffe1e1")))
                elif valor in ("VENCE EM ATE 30 DIAS", "VERIFICAR", "EM ABERTO"):
                    comandos_estilo.append(("BACKGROUND", (0, idx_linha), (-1, idx_linha), colors.HexColor("#fff3cd")))
                elif valor == "PARCIAL":
                    comandos_estilo.append(("BACKGROUND", (0, idx_linha), (-1, idx_linha), colors.HexColor("#ffe8cc")))
                elif valor == "DEVOLUCAO":
                    comandos_estilo.append(("BACKGROUND", (0, idx_linha), (-1, idx_linha), colors.HexColor("#e7f0ff")))

        tabela.setStyle(TableStyle(comandos_estilo))
        elementos.append(tabela)

    def _rodape(canvas_obj, documento):
        # Rodapé em 2 linhas: nome do relatório + página (linha de cima) e
        # um aviso de confidencialidade centralizado (linha de baixo) -
        # deixa o documento mais formal para quando ele circula fora do
        # time (por e-mail, por exemplo).
        canvas_obj.saveState()
        canvas_obj.setFont("Helvetica", 7)
        canvas_obj.setFillColor(colors.HexColor("#666666"))
        canvas_obj.drawString(margem, 10 * mm, "BioCAZ — Relatório de Rastreabilidade de Lote  |  Documento de uso interno")
        canvas_obj.drawRightString(largura_pagina - margem, 10 * mm, f"Página {documento.page}")
        canvas_obj.setFont("Helvetica", 6)
        canvas_obj.setFillColor(colors.HexColor("#999999"))
        canvas_obj.drawCentredString(
            largura_pagina / 2, 6 * mm, "Confidencial — não distribuir fora da empresa sem autorização"
        )
        canvas_obj.restoreState()

    doc.build(elementos, onFirstPage=_rodape, onLaterPages=_rodape)
    return buffer.getvalue()


COLUNAS_DATA = ("DATA_MOV", "VALIDADE_DOCUMENTO", "DT_EMISSAO")


def _formatar_valor_data(v):
    """Converte um valor de data (datetime, 'YYYYMMDD' ou outro formato) para DD/MM/YYYY."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return None
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%d/%m/%Y")
        except ValueError:
            return None  # datas "zeradas" tipo 0001-01-01 do Protheus
    s = str(v).strip()
    if not s or s in ("0", "00000000"):
        return None
    if len(s) == 8 and s.isdigit():
        try:
            return pd.to_datetime(s, format="%Y%m%d").strftime("%d/%m/%Y")
        except ValueError:
            return s
    try:
        return pd.to_datetime(s).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return s


def formatar_datas_para_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna uma copia do df com as colunas de data no formato DD/MM/YYYY (para tela e Excel)."""
    df2 = df.copy()
    for col in COLUNAS_DATA:
        if col in df2.columns:
            df2[col] = df2[col].apply(_formatar_valor_data)
    return df2


def destacar_status(row: pd.Series) -> list[str]:
    cor = ""
    if row.get("STATUS_RASTREABILIDADE", "").startswith("ALERTA"):
        cor = "background-color: #ffe1e1"
    elif row.get("STATUS_RASTREABILIDADE", "") == "VERIFICAR":
        cor = "background-color: #fff3cd"
    return [cor] * len(row)


def formatar_numero_br(v, casas: int = 2) -> str:
    """Formata numero no padrao BR (milhar com ponto, decimal com virgula)."""
    try:
        return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(v)


# --------------------------------------------------------------------------
# Seletor de produto por nome (abas 1, 2 e 4) - complementa o campo de
# código exato para quem não decora os códigos do Protheus.
# --------------------------------------------------------------------------
def _obter_engine_seguro():
    """Tenta obter o engine sem lançar exceção. Usado só pelo seletor de
    produto da SIDEBAR, que é renderizado ANTES da checagem de conexão do
    corpo principal (então a conexão pode ainda não estar configurada, ou
    pode falhar) - nesses casos devolve None e a busca por nome fica
    desabilitada silenciosamente, sem derrubar a sidebar inteira."""
    if not config_disponivel():
        return None
    try:
        return get_engine()
    except Exception:
        return None


@st.cache_data(ttl=300, show_spinner=False)
def buscar_produtos_por_descricao(_engine: Engine, termo: str, limite: int = 30) -> pd.DataFrame:
    """Busca produtos (só Produto Acabado, mesma trava B1_TIPO='PA' do
    resto do app) cuja descrição contenha o termo digitado. Só busca a
    partir de 2 caracteres, para não trazer resultado nenhum (ou a tabela
    inteira) a cada tecla digitada. Resultado limitado a `limite` linhas."""
    termo = (termo or "").strip()
    if len(termo) < 2:
        return pd.DataFrame(columns=["CODIGO", "DESCRICAO"])
    sql = """
        SELECT DISTINCT TOP (:limite) B1_COD AS CODIGO, B1_DESC AS DESCRICAO
        FROM SB1010
        WHERE D_E_L_E_T_ = ' '
          AND B1_TIPO = 'PA'
          AND B1_DESC LIKE :termo
        ORDER BY B1_DESC
    """
    with _engine.connect() as conn:
        return pd.read_sql(text(sql), conn, params={"termo": f"%{termo}%", "limite": limite})


def seletor_produto(
    engine, key_prefix: str, container=None, label_codigo: str = "Produto (código exato, opcional)"
) -> str:
    """Campo de filtro de produto com duas formas de preencher: digitar o
    código exato do Protheus (como sempre funcionou) OU, dentro do
    expander "buscar por nome", digitar parte da descrição e escolher da
    lista de resultados - a escolha preenche o código automaticamente.
    `container` é onde os widgets são desenhados (st, st.sidebar ou uma
    coluna/st.columns) - assim a mesma função serve para a sidebar e para
    dentro de uma aba. Devolve o código final a ser usado no filtro SQL."""
    alvo = container if container is not None else st
    produto_codigo = alvo.text_input(label_codigo, value="", key=f"{key_prefix}_produto_codigo")

    if engine is None:
        return produto_codigo

    with alvo.expander("🔍 Não sabe o código? Buscar produto por nome"):
        termo_busca = st.text_input(
            "Digite parte do nome do produto (mín. 2 letras)", value="", key=f"{key_prefix}_produto_busca_nome"
        )
        if termo_busca.strip():
            try:
                resultados = buscar_produtos_por_descricao(engine, termo_busca)
            except Exception as e:
                st.warning(f"Não foi possível buscar por nome agora: {e}")
                resultados = pd.DataFrame(columns=["CODIGO", "DESCRICAO"])

            if resultados.empty:
                st.caption("Nenhum produto encontrado com esse nome.")
            else:
                opcoes = ["(selecione um produto)"] + [
                    f"{linha.CODIGO} - {linha.DESCRICAO}" for linha in resultados.itertuples()
                ]
                escolha = st.selectbox("Resultados encontrados", opcoes, key=f"{key_prefix}_produto_escolha")
                if escolha != "(selecione um produto)":
                    produto_codigo = escolha.split(" - ", 1)[0]
                    st.caption(f"✅ Produto selecionado: **{produto_codigo}**")

    return produto_codigo


# --------------------------------------------------------------------------
# UI - Sidebar (filtros)
# --------------------------------------------------------------------------
# Estes filtros pertencem só à aba 1 (Relatório de Movimentos) - as outras
# 3 abas têm seus próprios filtros, definidos dentro de cada `with aba_...:`
# mais abaixo. Nenhum widget aqui dispara consulta sozinho: os valores só
# são usados quando o botão "Executar consulta" é clicado (ou no primeiro
# carregamento da página).
st.sidebar.header("Filtros")

filial = st.sidebar.text_input("Filial", value="010101")

usar_periodo = st.sidebar.checkbox("Filtrar por período", value=False)
if usar_periodo:
    hoje = date.today()
    col_ini, col_fim = st.sidebar.columns(2)
    dt_ini = col_ini.date_input("De", value=hoje - timedelta(days=30), format="DD/MM/YYYY")
    dt_fim = col_fim.date_input("Até", value=hoje, format="DD/MM/YYYY")
else:
    dt_ini = dt_fim = None  # sem filtro de período -> monta_query() nao adiciona a clausula de data

_engine_sidebar = _obter_engine_seguro()
produto = seletor_produto(_engine_sidebar, "sidebar", container=st.sidebar, label_codigo="Produto (código exato)")
pedido = st.sidebar.text_input("Pedido (número exato)", value="")

tipo_mov = st.sidebar.multiselect("Tipo de movimento", TIPO_MOV_OPCOES, default=[])
status_sel = st.sidebar.multiselect("Status de rastreabilidade", STATUS_OPCOES, default=[])

# Tanto vazio (nada marcado, o padrao ao abrir) quanto TODAS as opcoes
# marcadas contam como "sem filtro" para montar_query() (ver
# "len(filtros['tipo_mov']) < len(TIPO_MOV_OPCOES)") - a consulta so fica
# mais restrita quando o usuario marca ALGUMAS opcoes, mas nao todas.
executar = st.sidebar.button("Executar consulta", type="primary")

st.sidebar.caption(
    "Os dados ficam em cache por 5 minutos. Use 'Executar consulta' após "
    "trocar filtros para forçar uma nova busca imediata."
)

# --------------------------------------------------------------------------
# Corpo principal
# --------------------------------------------------------------------------
st.title("📦 Relatório de Rastreabilidade de Lote")
st.caption("Baseado em SD1, SD2, SB1, SD5, SB8, SA1 e SA2 do Protheus/TOTVS")

if not config_disponivel():
    _tela_configuracao_inicial()

_, _origem_conf = obter_config_conexao()
st.sidebar.divider()
with st.sidebar.expander("⚙️ Conexão"):
    st.caption(f"🔌 Conexão via: **{_origem_conf}**")
    if _origem_conf == ".env":
        st.caption("⚠️ Confirme se o `.env` aponta para o banco correto (ex: homologação).")
    if _origem_conf.startswith("config local"):
        if st.button("🔄 Reconfigurar conexão"):
            _caminho_config_local().unlink(missing_ok=True)
            st.cache_resource.clear()
            st.rerun()

if st.sidebar.button("🚪 Sair do aplicativo"):
    # So funciona de verdade dentro do .exe empacotado (run_app.py registra
    # essa funcao "sair" na janela do pywebview). Rodando via
    # `streamlit run app.py` no navegador comum, o botao nao faz nada -
    # ai e so fechar a aba mesmo.
    components.html(
        """
        <script>
        (function () {
            var api = null;
            if (window.pywebview && window.pywebview.api) {
                api = window.pywebview.api;
            } else if (window.parent && window.parent.pywebview && window.parent.pywebview.api) {
                api = window.parent.pywebview.api;
            }
            if (api && api.sair) {
                api.sair();
            }
        })();
        </script>
        """,
        height=0,
    )
    st.sidebar.caption("Fechando o aplicativo...")

try:
    engine = get_engine()
except Exception as e:
    st.error(f"Erro ao conectar no banco de dados: {e}")
    st.stop()

aba_relatorio, aba_consolidado, aba_busca, aba_terceiros, aba_atraso = st.tabs(
    [
        "📊 Relatório de Movimentos", "📦 Consolidado por Lote", "🔍 Busca por Lote",
        "🔄 Controle de Terceiros", "⏳ Atraso de Terceiros",
    ]
)

# ==========================================================================
# ABA 1 - Relatório de movimentos (usa os filtros da barra lateral)
# ==========================================================================
# Padrão usado em TODAS as 4 abas: o resultado da consulta é guardado no
# st.session_state (com uma chave por aba, ex.: "df", "df_consolidado") em
# vez de ser recalculado a cada interação da tela. Assim, clicar num
# gráfico, expandir uma seção ou trocar o filtro de outra aba não dispara
# uma nova consulta ao banco - só o botão correspondente ("Executar
# consulta", "Consultar" ou "Buscar") faz isso.
with aba_relatorio:
    if executar:
        try:
            filtros = {
                "filial": filial,
                "dt_ini": dt_ini,
                "dt_fim": dt_fim,
                "produto": produto,
                "pedido": pedido,
                "tipo_mov": tipo_mov,
                "status": status_sel,
            }
            st.session_state["df"], st.session_state["df_timestamp"] = carregar_dados(engine, filtros)
        except Exception as e:
            st.error(f"Erro ao consultar o banco de dados: {e}")
            st.stop()

    df = st.session_state.get("df", pd.DataFrame())

    if df.empty:
        if "df" not in st.session_state:
            st.info(
                "👈 Escolha os filtros na barra lateral (pelo menos a Filial) e "
                "clique em **Executar consulta** para carregar o relatório."
            )
        else:
            st.info("Nenhum registro encontrado para os filtros selecionados.")
    else:
        _ts = st.session_state.get("df_timestamp")
        if _ts:
            st.caption(
                f"🕒 Dados de {_ts.strftime('%d/%m/%Y %H:%M')} "
                "(cache de 5 min — clique em 'Executar consulta' para atualizar)"
            )

        # Gráfico interativo por status — clique numa barra para filtrar os
        # KPIs, a tabela e a exportação abaixo. Clique em outra barra para
        # trocar o filtro, ou use o botão "Limpar filtro" para ver tudo.
        resumo_status = df["STATUS_RASTREABILIDADE"].value_counts().reset_index()
        resumo_status.columns = ["Status", "Quantidade"]

        # "chart_reset_counter" muda o "key" do gráfico (mais abaixo) sempre
        # que o botão "Limpar filtro" é clicado - trocar o key força o
        # Streamlit/Altair a recriar o widget do zero, o que "desmarca" a
        # barra que estava selecionada (não existe um jeito direto de
        # limpar uma seleção do Altair sem recriar o componente).
        if "chart_reset_counter" not in st.session_state:
            st.session_state["chart_reset_counter"] = 0

        # selection_point com empty="all": quando nada foi clicado ainda,
        # a seleção "vale" para todas as barras (ninguém fica esmaecido).
        selecao_status = alt.selection_point(name="sel_status", fields=["Status"], on="click", empty="all")

        grafico_status = (
            alt.Chart(resumo_status)
            .mark_bar()
            .encode(
                y=alt.Y("Status:N", sort="-x", title=None, axis=alt.Axis(labelLimit=1000)),
                x=alt.X("Quantidade:Q", title="Quantidade"),
                color=alt.Color(
                    "Status:N",
                    scale=alt.Scale(domain=STATUS_OPCOES, range=CORES_STATUS),
                    legend=None,
                ),
                opacity=alt.condition(selecao_status, alt.value(1.0), alt.value(0.35)),
                tooltip=[alt.Tooltip("Status:N", title="Status"), alt.Tooltip("Quantidade:Q", title="Quantidade")],
            )
            .add_params(selecao_status)
            .properties(height=42 * len(resumo_status) + 30)
        )

        evento_grafico = st.altair_chart(
            grafico_status,
            use_container_width=True,
            on_select="rerun",
            key=f"grafico_status_{st.session_state['chart_reset_counter']}",
        )

        # Le a barra clicada no gráfico (se houver) a partir do evento de
        # seleção devolvido por st.altair_chart(on_select="rerun") - cada
        # clique dispara um rerun da página e o Streamlit devolve, aqui,
        # qual ponto ("Status") estava selecionado no momento do clique.
        pontos_selecionados = evento_grafico.selection.get("sel_status", []) if evento_grafico else []
        status_clicado = pontos_selecionados[0]["Status"] if pontos_selecionados else None

        if status_clicado:
            col_msg, col_btn = st.columns([5, 1])
            col_msg.caption(f"🔎 Filtrando por clique no gráfico: **{status_clicado}**")
            if col_btn.button("✖️ Limpar filtro"):
                st.session_state["chart_reset_counter"] += 1
                st.rerun()
            df_visao = df[df["STATUS_RASTREABILIDADE"] == status_clicado]
        else:
            df_visao = df  # sem clique no gráfico -> mostra tudo que veio do banco

        st.divider()

        # KPIs (refletem o filtro do clique no gráfico, quando ativo)
        total = len(df_visao)
        alertas_sem_lote = int(
            (df_visao["STATUS_RASTREABILIDADE"] == "ALERTA - PRODUTO RASTREADO SEM LOTE NO MOVIMENTO").sum()
        )
        alertas_parcial = int((df_visao["STATUS_RASTREABILIDADE"] == "ALERTA - COBERTURA DE LOTE PARCIAL").sum())
        ok_rastreado = int((df_visao["STATUS_RASTREABILIDADE"] == "OK - PRODUTO RASTREADO").sum())
        produtos_distintos = df_visao["PRODUTO"].nunique()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total de registros", f"{total:,}".replace(",", "."))
        c2.metric("OK - rastreado", f"{ok_rastreado:,}".replace(",", "."))
        c3.metric("⚠️ Sem lote", f"{alertas_sem_lote:,}".replace(",", "."))
        c4.metric("⚠️ Cobertura parcial", f"{alertas_parcial:,}".replace(",", "."))
        c5.metric("Produtos distintos", f"{produtos_distintos:,}".replace(",", "."))

        st.divider()

        # Tabela (formatada para exibição: datas em DD/MM/YYYY) — reflete o mesmo filtro
        df_exibicao = formatar_datas_para_exibicao(df_visao)

        with st.expander("📋 Detalhamento", expanded=False):
            st.dataframe(
                df_exibicao.style.apply(destacar_status, axis=1),
                use_container_width=True,
                height=520,
            )

        # Subtítulo do PDF resume os filtros ativos (filial, período e o
        # status clicado no gráfico, se houver) para quem abrir o PDF depois
        # saber exatamente o que está vendo, sem precisar do app aberto.
        subtitulo_pdf = f"Filial: {filial}"
        if usar_periodo and dt_ini and dt_fim:
            subtitulo_pdf += f" | Período: {dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}"
        if status_clicado:
            subtitulo_pdf += f" | Status: {status_clicado}"

        col_dl_xlsx1, col_dl_pdf1 = st.columns(2)
        excel_bytes = gerar_excel(df_exibicao)
        col_dl_xlsx1.download_button(
            label="⬇️ Baixar Excel",
            data=excel_bytes,
            file_name=f"rastreabilidade_lote_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_relatorio",
            use_container_width=True,
        )
        kpis_pdf = [
            ("Total de registros", f"{total:,}".replace(",", ".")),
            ("OK - rastreado", f"{ok_rastreado:,}".replace(",", ".")),
            ("Sem lote", f"{alertas_sem_lote:,}".replace(",", ".")),
            ("Cobertura parcial", f"{alertas_parcial:,}".replace(",", ".")),
            ("Produtos distintos", f"{produtos_distintos:,}".replace(",", ".")),
        ]
        pdf_bytes = gerar_pdf(df_exibicao, "Relatório de Rastreabilidade de Lote", subtitulo_pdf, kpis=kpis_pdf)
        col_dl_pdf1.download_button(
            label="📄 Baixar PDF",
            data=pdf_bytes,
            file_name=f"rastreabilidade_lote_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            key="download_relatorio_pdf",
            use_container_width=True,
        )

# ==========================================================================
# ABA 2 - Visão consolidada por lote (saldo atual, validade, quem recebeu)
# ==========================================================================
with aba_consolidado:
    st.caption(
        "Agrupa todas as movimentações por PRODUTO + LOTE: saldo atual e "
        "validade (via SB8), total comprado/vendido e quantos clientes "
        "distintos receberam aquele lote."
    )
    col_prod, col_btn_c = st.columns([3, 1])
    produto_consolidado = seletor_produto(
        engine, "consolidado", container=col_prod, label_codigo="Filtrar por produto (código exato, opcional)"
    )
    buscar_consolidado = col_btn_c.button("Consultar", type="primary", key="btn_consolidado")

    if buscar_consolidado:
        try:
            st.session_state["df_consolidado"], st.session_state["df_consolidado_timestamp"] = carregar_consolidado(
                engine, filial, produto_consolidado
            )
        except Exception as e:
            st.error(f"Erro ao consultar o consolidado por lote: {e}")
            st.session_state.pop("df_consolidado", None)
            st.session_state.pop("df_consolidado_timestamp", None)

    df_consolidado = st.session_state.get("df_consolidado")

    if df_consolidado is None:
        st.info("Defina o filtro (opcional) e clique em 'Consultar'.")
    elif df_consolidado.empty:
        st.info("Nenhum lote encontrado.")
    else:
        _ts_c = st.session_state.get("df_consolidado_timestamp")
        if _ts_c:
            st.caption(f"🕒 Dados de {_ts_c.strftime('%d/%m/%Y %H:%M')} (cache de 5 min — clique em 'Consultar' para atualizar)")

        df_c = df_consolidado.copy()
        # Saldo "calculado" = diferenca entre tudo que entrou e tudo que saiu
        # (baseado em SD1/SD2, historico completo) - diferente do
        # SALDO_ATUAL vindo da SB8 (saldo real e atual do estoque), que so
        # aparece se a busca abaixo funcionar.
        df_c["SALDO_CALCULADO"] = df_c["TOTAL_ENTRADA"] - df_c["TOTAL_SAIDA"]

        # Saldo/validade reais (SB8) — busca isolada; se falhar, so avisa e
        # segue mostrando o resto sem essas colunas (nao derruba a aba).
        try:
            df_saldo = carregar_saldo_lotes(engine, filial)
            df_c = df_c.merge(df_saldo, on=["PRODUTO", "LOCAL", "LOTE"], how="left")
            df_c["STATUS_VALIDADE"] = df_c["VALIDADE"].apply(status_validade)
        except Exception as e:
            st.warning(
                "Não foi possível obter saldo/validade atual do lote (SB8) — "
                f"exibindo só os totais movimentados. Detalhe técnico: {e}"
            )

        for col in ("PRIMEIRA_MOVIMENTACAO", "ULTIMA_MOVIMENTACAO", "VALIDADE"):
            if col in df_c.columns:
                df_c[col] = df_c[col].apply(_formatar_valor_data)

        if "STATUS_VALIDADE" in df_c.columns:
            cv1, cv2, cv3 = st.columns(3)
            cv1.metric("🔴 Lotes vencidos", int((df_c["STATUS_VALIDADE"] == "VENCIDO").sum()))
            cv2.metric("🟡 Vencendo em até 30 dias", int((df_c["STATUS_VALIDADE"] == "VENCE EM ATE 30 DIAS").sum()))
            cv3.metric("Total de lotes", len(df_c))
            st.divider()

        def _destacar_validade(row):
            cor = ""
            status_v = row.get("STATUS_VALIDADE", "")
            if status_v == "VENCIDO":
                cor = "background-color: #ffe1e1"
            elif status_v == "VENCE EM ATE 30 DIAS":
                cor = "background-color: #fff3cd"
            return [cor] * len(row)

        estilo = df_c.style.apply(_destacar_validade, axis=1) if "STATUS_VALIDADE" in df_c.columns else df_c.style
        st.dataframe(estilo, use_container_width=True, height=520)

        subtitulo_pdf_c = f"Filial: {filial}"
        if produto_consolidado:
            subtitulo_pdf_c += f" | Produto: {produto_consolidado}"

        col_dl_xlsx2, col_dl_pdf2 = st.columns(2)
        excel_consolidado = gerar_excel(df_c)
        col_dl_xlsx2.download_button(
            label="⬇️ Baixar Excel",
            data=excel_consolidado,
            file_name=f"consolidado_por_lote_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_consolidado",
            use_container_width=True,
        )
        kpis_pdf_c = [("Total de lotes", f"{len(df_c):,}".replace(",", "."))]
        if "STATUS_VALIDADE" in df_c.columns:
            kpis_pdf_c = [
                ("Lotes vencidos", f"{int((df_c['STATUS_VALIDADE'] == 'VENCIDO').sum()):,}".replace(",", ".")),
                (
                    "Vencendo em até 30 dias",
                    f"{int((df_c['STATUS_VALIDADE'] == 'VENCE EM ATE 30 DIAS').sum()):,}".replace(",", "."),
                ),
                ("Total de lotes", f"{len(df_c):,}".replace(",", ".")),
            ]
        pdf_consolidado = gerar_pdf(df_c, "Consolidado por Lote", subtitulo_pdf_c, kpis=kpis_pdf_c)
        col_dl_pdf2.download_button(
            label="📄 Baixar PDF",
            data=pdf_consolidado,
            file_name=f"consolidado_por_lote_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            key="download_consolidado_pdf",
            use_container_width=True,
        )

# ==========================================================================
# ABA 3 - Busca reversa por lote (de onde veio / para onde foi)
# ==========================================================================
with aba_busca:
    st.caption(
        "Digite um número de lote para ver todas as entradas (de quem "
        "vendeu/forneceu) e saídas (para quem foi vendido) registradas para ele."
    )
    col_lote, col_btn_b = st.columns([3, 1])
    lote_busca = col_lote.text_input("Número do lote", value="", key="lote_busca")
    buscar_lote = col_btn_b.button("Buscar", type="primary", key="btn_busca_lote")

    if buscar_lote:
        if not lote_busca.strip():
            st.warning("Digite um número de lote para buscar.")
        else:
            try:
                st.session_state["df_busca_lote"], st.session_state["df_busca_lote_timestamp"] = carregar_busca_lote(
                    engine, filial, lote_busca
                )
                st.session_state["lote_busca_atual"] = lote_busca.strip()
            except Exception as e:
                st.error(f"Erro ao buscar o lote: {e}")
                st.session_state.pop("df_busca_lote", None)
                st.session_state.pop("df_busca_lote_timestamp", None)

    df_busca = st.session_state.get("df_busca_lote")

    if df_busca is None:
        st.info("Informe um lote e clique em 'Buscar'.")
    elif df_busca.empty:
        st.info(f"Nenhuma movimentação encontrada para o lote '{st.session_state.get('lote_busca_atual', '')}'.")
    else:
        _ts_b = st.session_state.get("df_busca_lote_timestamp")
        if _ts_b:
            st.caption(f"🕒 Dados de {_ts_b.strftime('%d/%m/%Y %H:%M')} (cache de 5 min — clique em 'Buscar' para atualizar)")

        lote_atual = st.session_state.get("lote_busca_atual", "")
        total_entrada = df_busca.loc[df_busca["TIPO_MOV"] == "ENTRADA", "QUANTIDADE"].sum()
        total_saida = df_busca.loc[df_busca["TIPO_MOV"] == "SAIDA", "QUANTIDADE"].sum()
        produtos_encontrados = df_busca["PRODUTO"].nunique()

        if produtos_encontrados > 1:
            st.subheader(f"Lote {lote_atual} (usado em {produtos_encontrados} produtos diferentes)")
        else:
            st.subheader(f"Lote {lote_atual} — {df_busca['PRODUTO'].iloc[0]} - {df_busca['DESCRICAO'].iloc[0]}")

        cb1, cb2, cb3 = st.columns(3)
        cb1.metric("Total entrada", formatar_numero_br(total_entrada))
        cb2.metric("Total saída", formatar_numero_br(total_saida))
        cb3.metric("Saldo calculado", formatar_numero_br(total_entrada - total_saida))

        # Saldo/validade reais do lote (SB8) — best-effort, isolado: essa
        # tabela e so um complemento visual (mostra o saldo atual em
        # estoque daquele lote especifico), entao qualquer erro aqui e
        # silenciado (except: pass) para nao atrapalhar a busca principal,
        # que ja funcionou.
        try:
            df_saldo_lote = carregar_saldo_lotes(engine, filial)
            df_saldo_lote = df_saldo_lote[df_saldo_lote["LOTE"] == lote_atual]
            if not df_saldo_lote.empty:
                st.caption("Saldo atual (SB8) por local:")
                df_saldo_exib = df_saldo_lote.copy()
                df_saldo_exib["VALIDADE"] = df_saldo_exib["VALIDADE"].apply(_formatar_valor_data)
                st.dataframe(df_saldo_exib, use_container_width=True, hide_index=True)
        except Exception:
            pass  # saldo real e so um complemento - a busca principal segue sem ele

        st.divider()

        df_busca_exib = formatar_datas_para_exibicao(df_busca)

        col_e, col_s = st.columns(2)
        with col_e:
            st.markdown("**⬅️ Entradas (de onde veio)**")
            st.dataframe(
                df_busca_exib[df_busca_exib["TIPO_MOV"] == "ENTRADA"].drop(columns=["TIPO_MOV"]),
                use_container_width=True,
                hide_index=True,
            )
        with col_s:
            st.markdown("**➡️ Saídas (para onde foi)**")
            st.dataframe(
                df_busca_exib[df_busca_exib["TIPO_MOV"] == "SAIDA"].drop(columns=["TIPO_MOV"]),
                use_container_width=True,
                hide_index=True,
            )

        subtitulo_pdf_b = f"Filial: {filial} | Lote: {lote_atual}"
        if produtos_encontrados == 1:
            subtitulo_pdf_b += f" | Produto: {df_busca['PRODUTO'].iloc[0]} - {df_busca['DESCRICAO'].iloc[0]}"

        col_dl_xlsx3, col_dl_pdf3 = st.columns(2)
        excel_busca = gerar_excel(df_busca_exib)
        col_dl_xlsx3.download_button(
            label="⬇️ Baixar Excel",
            data=excel_busca,
            file_name=f"lote_{lote_atual}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_busca_lote",
            use_container_width=True,
        )
        kpis_pdf_b = [
            ("Total entrada", formatar_numero_br(total_entrada)),
            ("Total saída", formatar_numero_br(total_saida)),
            ("Saldo calculado", formatar_numero_br(total_entrada - total_saida)),
        ]
        pdf_busca = gerar_pdf(df_busca_exib, f"Busca por Lote — {lote_atual}", subtitulo_pdf_b, kpis=kpis_pdf_b)
        col_dl_pdf3.download_button(
            label="📄 Baixar PDF",
            data=pdf_busca,
            file_name=f"lote_{lote_atual}_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            key="download_busca_lote_pdf",
            use_container_width=True,
        )

# ==========================================================================
# ABA 4 - Controle de Terceiros (remessas e devolucoes - SB6)
# ==========================================================================
with aba_terceiros:
    st.caption(
        "Remessas e devoluções de terceiros (SB6): o que foi enviado para "
        "clientes/fornecedores ou recebido deles, com o saldo em aberto de "
        "cada documento — útil para cobrar devoluções pendentes e saber o "
        "que ainda está fora da empresa."
    )

    with st.expander("🔍 Filtros", expanded=False):
        col_t1, col_t2, col_t3 = st.columns(3)
        usar_periodo_terceiros = col_t1.checkbox(
            "Filtrar por período (emissão)", value=False, key="usar_periodo_terceiros"
        )
        if usar_periodo_terceiros:
            col_ini_t, col_fim_t = col_t1.columns(2)
            dt_ini_terceiros = col_ini_t.date_input(
                "De", value=date.today() - timedelta(days=90), format="DD/MM/YYYY", key="dt_ini_terceiros"
            )
            dt_fim_terceiros = col_fim_t.date_input(
                "Até", value=date.today(), format="DD/MM/YYYY", key="dt_fim_terceiros"
            )
        else:
            dt_ini_terceiros = dt_fim_terceiros = None
        produto_terceiros = seletor_produto(
            engine, "terceiros", container=col_t2, label_codigo="Produto (código exato, opcional)"
        )
        clifor_terceiros = col_t3.text_input(
            "Cliente/Fornecedor (código ou nome, opcional)", value="", key="clifor_terceiros"
        )

        col_t4, col_t5, col_t6 = st.columns(3)
        tipo_clifor_sel = col_t4.multiselect(
            "Tipo", TERCEIROS_TIPO_CLIFOR_OPCOES, default=TERCEIROS_TIPO_CLIFOR_OPCOES, key="tipo_clifor_terceiros"
        )
        de_em_sel = col_t5.multiselect(
            "De/Em terceiros", TERCEIROS_DE_EM_OPCOES, default=TERCEIROS_DE_EM_OPCOES, key="de_em_terceiros_sel"
        )
        situacao_sel = col_t6.multiselect(
            "Situação", TERCEIROS_SITUACAO_OPCOES, default=TERCEIROS_SITUACAO_OPCOES, key="situacao_terceiros_sel"
        )

        buscar_terceiros = st.button("Consultar", type="primary", key="btn_terceiros")

    if buscar_terceiros:
        try:
            filtros_terceiros = {
                "filial": filial,
                "dt_ini": dt_ini_terceiros,
                "dt_fim": dt_fim_terceiros,
                "produto": produto_terceiros,
            }
            st.session_state["df_terceiros"], st.session_state["df_terceiros_timestamp"] = carregar_terceiros(
                engine, filtros_terceiros
            )
        except Exception as e:
            st.error(f"Erro ao consultar o controle de terceiros: {e}")
            st.session_state.pop("df_terceiros", None)
            st.session_state.pop("df_terceiros_timestamp", None)

    df_terceiros_raw = st.session_state.get("df_terceiros")

    if df_terceiros_raw is None:
        st.info("Defina os filtros (opcionais) e clique em 'Consultar'.")
    elif df_terceiros_raw.empty:
        st.info("Nenhum registro encontrado.")
    else:
        _ts_t = st.session_state.get("df_terceiros_timestamp")
        if _ts_t:
            st.caption(f"🕒 Dados de {_ts_t.strftime('%d/%m/%Y %H:%M')} (cache de 5 min — clique em 'Consultar' para atualizar)")

        df_t = df_terceiros_raw.copy()

        # Filtros de Tipo / De-Em terceiros / Situação e a busca por
        # cliente/fornecedor NÃO vão para o SQL: como essas colunas já são
        # calculadas por CASE dentro da própria TERCEIROS_QUERY, é mais
        # simples (e evita duplicar a lógica do CASE num WHERE) filtrar o
        # DataFrame já carregado, aqui em Python, do que re-executar a
        # consulta a cada mudança desses filtros. Só período e produto vão
        # de fato para o SQL (montar_query_terceiros), pois esses reduzem o
        # volume de dados trazido do banco.
        if tipo_clifor_sel and len(tipo_clifor_sel) < len(TERCEIROS_TIPO_CLIFOR_OPCOES):
            mapa_tipo = {"Cliente": "C", "Fornecedor": "F"}
            codigos_tipo = [mapa_tipo[t] for t in tipo_clifor_sel]
            df_t = df_t[df_t["TIPO_CLI_FOR"].isin(codigos_tipo)]
        if de_em_sel and len(de_em_sel) < len(TERCEIROS_DE_EM_OPCOES):
            df_t = df_t[df_t["DE_EM_TERCEIROS"].isin(de_em_sel)]
        if situacao_sel and len(situacao_sel) < len(TERCEIROS_SITUACAO_OPCOES):
            df_t = df_t[df_t["SITUACAO"].isin(situacao_sel)]
        if clifor_terceiros.strip():
            termo = clifor_terceiros.strip().upper()
            df_t = df_t[
                df_t["CLIENTE_FORN"].astype(str).str.upper().str.contains(termo, na=False)
                | df_t["RAZAO_SOCIAL"].astype(str).str.upper().str.contains(termo, na=False)
            ]

        if df_t.empty:
            st.info("Nenhum registro para os filtros selecionados.")
        else:
            # KPIs (contagens de documentos por situação + saldo pendente
            # estimado - ver o st.caption logo abaixo com a fórmula usada)
            docs_aberto = int((df_t["SITUACAO"] == "EM ABERTO").sum())
            docs_parcial = int((df_t["SITUACAO"] == "PARCIAL").sum())
            docs_devolucao = int((df_t["SITUACAO"] == "DEVOLUCAO").sum())
            valor_pendente = (
                df_t.loc[df_t["SITUACAO"].isin(["EM ABERTO", "PARCIAL"]), "SALDO"]
                * df_t.loc[df_t["SITUACAO"].isin(["EM ABERTO", "PARCIAL"]), "PRECO_UNITARIO"]
            ).sum()

            ct1, ct2, ct3, ct4, ct5 = st.columns(5)
            ct1.metric("Documentos", f"{len(df_t):,}".replace(",", "."))
            ct2.metric("🟡 Em aberto", f"{docs_aberto:,}".replace(",", "."))
            ct3.metric("🟠 Parcial", f"{docs_parcial:,}".replace(",", "."))
            ct4.metric("🔵 Devolução", f"{docs_devolucao:,}".replace(",", "."))
            ct5.metric("Saldo pendente (valor est.)", f"R$ {formatar_numero_br(valor_pendente)}")
            st.caption(
                "Saldo pendente estimado = SALDO × PREÇO_UNITARIO dos documentos em aberto ou parciais "
                "(indicador aproximado, não é um valor fiscal)."
            )

            st.divider()

            # Resumo por lote: quanto ainda esta em poder do terceiro, somando
            # o SALDO das remessas (a devolucao ja e descontada automaticamente
            # do saldo da remessa original pelo Protheus, entao as linhas de
            # DEVOLUCAO nao entram nessa soma - senao contaria duas vezes).
            st.subheader("📦 Saldo por lote — ainda em poder do terceiro")
            df_remessas_lote = df_t[df_t["MOVIMENTO"] == "REMESSA"].copy()
            if df_remessas_lote.empty:
                st.info("Nenhuma remessa em aberto para os filtros selecionados.")
            else:
                df_resumo_lote = (
                    df_remessas_lote
                    .groupby(
                        [
                            "LOTE", "PRODUTO", "DESCRICAO", "CLIENTE_FORN", "RAZAO_SOCIAL", "CIDADE", "ESTADO",
                            "DOC_ORIGINAL", "SERIE",
                        ],
                        dropna=False,
                    )
                    .agg(
                        QTD_REMETIDA=("QUANTIDADE", "sum"),
                        saldoemterceiro=("SALDO", "sum"),
                    )
                    .reset_index()
                    .rename(columns={"DOC_ORIGINAL": "NOTA_SAIDA"})
                )
                df_resumo_lote["LOTE"] = df_resumo_lote["LOTE"].fillna("(sem lote)")
                df_resumo_lote = df_resumo_lote[
                    [
                        "LOTE", "PRODUTO", "DESCRICAO", "NOTA_SAIDA", "SERIE",
                        "CLIENTE_FORN", "RAZAO_SOCIAL", "CIDADE", "ESTADO",
                        "QTD_REMETIDA", "saldoemterceiro",
                    ]
                ]
                df_resumo_lote = df_resumo_lote[df_resumo_lote["saldoemterceiro"] > 0]
                if df_resumo_lote.empty:
                    st.info("Nenhum saldo em aberto para os filtros selecionados (tudo já devolvido).")
                else:
                    df_resumo_lote = df_resumo_lote.sort_values("saldoemterceiro", ascending=False)
                    st.dataframe(df_resumo_lote, use_container_width=True, hide_index=True)
                    excel_resumo_lote = gerar_excel(df_resumo_lote)
                    st.download_button(
                        label="⬇️ Baixar Excel (saldo por lote)",
                        data=excel_resumo_lote,
                        file_name=f"saldo_por_lote_terceiros_{date.today().isoformat()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_resumo_lote_terceiros",
                    )
            st.caption(
                "Se o mesmo lote/produto foi remetido para mais de um cliente/fornecedor, "
                "aparece uma linha para cada um."
            )

            st.divider()

            # Alem de remessa/devolucao (SB6), tambem busca venda/compra
            # (SD1/SD2) dos MESMOS lotes que aparecem aqui - assim, se um
            # lote foi remetido para um terceiro E TAMBEM vendido/comprado
            # normalmente (fora do fluxo de terceiros), esse outro
            # movimento aparece junto na coluna MOVIMENTO, em vez de ficar
            # invisivel nessa tela. So busca pelos lotes realmente
            # presentes no resultado filtrado (df_t), nao pelo lote inteiro
            # da base.
            lotes_presentes = tuple(sorted(df_t["LOTE"].dropna().unique().tolist()))
            df_movimentos_extra = carregar_movimentos_lote(engine, filial, lotes_presentes)
            if not df_movimentos_extra.empty:
                for col in df_t.columns:
                    if col not in df_movimentos_extra.columns:
                        # Numerico vira NaN (aparece em branco no st.dataframe);
                        # texto vira "" (idem) - "pd.NA" puro aparecia como o
                        # texto literal "None" na tela, o que parecia um erro.
                        if pd.api.types.is_numeric_dtype(df_t[col]):
                            df_movimentos_extra[col] = float("nan")
                        else:
                            df_movimentos_extra[col] = ""
                df_movimentos_extra["SITUACAO"] = ""
                df_movimentos_extra["DE_EM_TERCEIROS"] = ""
                df_t_completo = pd.concat(
                    [df_t, df_movimentos_extra[df_t.columns]], ignore_index=True
                ).sort_values(["LOTE", "DT_EMISSAO"], na_position="last")
            else:
                df_t_completo = df_t
            st.caption(
                "As linhas com MOVIMENTO = VENDA ou COMPRA são vendas/compras normais (fora do fluxo de "
                "terceiros) dos mesmos lotes acima — aparecem aqui só para rastreabilidade, sem SALDO/SITUAÇÃO "
                "(esses conceitos não existem numa venda/compra comum). IDENTIFICADOR também fica em branco "
                "nessas linhas, pois não existe um equivalente direto numa venda/compra comum."
            )

            def _destacar_situacao(row):
                cor = ""
                situacao_linha = row.get("SITUACAO", "")
                movimento_linha = row.get("MOVIMENTO", "")
                if situacao_linha == "EM ABERTO":
                    cor = "background-color: #fff3cd"
                elif situacao_linha == "PARCIAL":
                    cor = "background-color: #ffe8cc"
                elif situacao_linha == "DEVOLUCAO":
                    cor = "background-color: #e7f0ff"
                elif movimento_linha in ("VENDA", "COMPRA"):
                    cor = "background-color: #e9ecef"
                return [cor] * len(row)

            df_t_exib = formatar_datas_para_exibicao(df_t_completo)
            estilo_t = df_t_exib.style.apply(_destacar_situacao, axis=1)
            st.dataframe(estilo_t, use_container_width=True, height=520)

            excel_terceiros = gerar_excel(df_t_exib)
            st.download_button(
                label="⬇️ Baixar Excel",
                data=excel_terceiros,
                file_name=f"controle_terceiros_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_terceiros",
            )

# ==========================================================================
# ABA 5 - Atraso de Terceiros (aging do saldo em poder do terceiro)
# ==========================================================================
# Reaproveita a mesma TERCEIROS_QUERY/carregar_terceiros da aba "Controle de
# Terceiros" (mesmo cache de 5 min - se os filtros baterem exatamente, nem
# repete a consulta no banco). Cada linha de REMESSA na SB6010 ja carrega o
# proprio SALDO em aberto (o que ainda falta devolver) e a propria data de
# emissao, entao o "atraso" de cada remessa e simplesmente dias corridos
# desde essa emissao - nao precisa de nenhum calculo de FIFO entre remessas
# e devolucoes (o Protheus ja desconta a devolucao do saldo da remessa
# original).
with aba_atraso:
    st.caption(
        "Há quanto tempo cada remessa em aberto está em poder do terceiro sem devolução — "
        "útil para cobrar devoluções pendentes antes que virem passivo esquecido."
    )

    with st.expander("🔍 Filtros", expanded=False):
        col_a1, col_a2, col_a3 = st.columns(3)
        produto_atraso = seletor_produto(
            engine, "atraso", container=col_a1, label_codigo="Produto (código exato, opcional)"
        )
        clifor_atraso = col_a2.text_input(
            "Cliente/Fornecedor (código ou nome, opcional)", value="", key="clifor_atraso"
        )
        dias_alerta_atraso = col_a3.number_input(
            "Alertar (🟠) a partir de quantos dias", min_value=1, value=60, step=1, key="dias_alerta_atraso"
        )
        tipo_clifor_atraso_sel = st.multiselect(
            "Tipo", TERCEIROS_TIPO_CLIFOR_OPCOES, default=TERCEIROS_TIPO_CLIFOR_OPCOES, key="tipo_clifor_atraso"
        )
        buscar_atraso = st.button("Consultar", type="primary", key="btn_atraso")

    if buscar_atraso:
        try:
            filtros_atraso = {"filial": filial, "dt_ini": None, "dt_fim": None, "produto": produto_atraso}
            st.session_state["df_atraso"], st.session_state["df_atraso_timestamp"] = carregar_terceiros(
                engine, filtros_atraso
            )
        except Exception as e:
            st.error(f"Erro ao consultar o atraso de terceiros: {e}")
            st.session_state.pop("df_atraso", None)
            st.session_state.pop("df_atraso_timestamp", None)

    df_atraso_raw = st.session_state.get("df_atraso")

    if df_atraso_raw is None:
        st.info("Defina os filtros (opcionais) e clique em 'Consultar'.")
    elif df_atraso_raw.empty:
        st.info("Nenhum registro encontrado.")
    else:
        _ts_a = st.session_state.get("df_atraso_timestamp")
        if _ts_a:
            st.caption(f"🕒 Dados de {_ts_a.strftime('%d/%m/%Y %H:%M')} (cache de 5 min — clique em 'Consultar' para atualizar)")

        df_a = df_atraso_raw.copy()

        # So interessa remessa com saldo em aberto (devolucao ja fechada nao
        # "atrasa" nada, e a devolucao ja desconta do saldo da propria
        # remessa - por isso so olhamos MOVIMENTO == 'REMESSA').
        df_a = df_a[(df_a["MOVIMENTO"] == "REMESSA") & (df_a["SALDO"] > 0)]

        if tipo_clifor_atraso_sel and len(tipo_clifor_atraso_sel) < len(TERCEIROS_TIPO_CLIFOR_OPCOES):
            mapa_tipo_a = {"Cliente": "C", "Fornecedor": "F"}
            codigos_tipo_a = [mapa_tipo_a[t] for t in tipo_clifor_atraso_sel]
            df_a = df_a[df_a["TIPO_CLI_FOR"].isin(codigos_tipo_a)]
        if clifor_atraso.strip():
            termo_a = clifor_atraso.strip().upper()
            df_a = df_a[
                df_a["CLIENTE_FORN"].astype(str).str.upper().str.contains(termo_a, na=False)
                | df_a["RAZAO_SOCIAL"].astype(str).str.upper().str.contains(termo_a, na=False)
            ]

        if df_a.empty:
            st.info("Nenhuma remessa em aberto para os filtros selecionados.")
        else:
            hoje_ts = pd.Timestamp(date.today())
            df_a["DIAS_EM_ABERTO"] = (hoje_ts - pd.to_datetime(df_a["DT_EMISSAO"])).dt.days

            def _status_atraso(dias) -> str:
                if pd.isna(dias):
                    return "SEM DATA"
                if dias > 90:
                    return "🔴 CRÍTICO (>90 dias)"
                if dias >= int(dias_alerta_atraso):
                    return "🟠 ATENÇÃO"
                return "🟢 RECENTE"

            df_a["STATUS_ATRASO"] = df_a["DIAS_EM_ABERTO"].apply(_status_atraso)

            qtd_critico = int((df_a["STATUS_ATRASO"] == "🔴 CRÍTICO (>90 dias)").sum())
            qtd_atencao = int((df_a["STATUS_ATRASO"] == "🟠 ATENÇÃO").sum())
            maior_atraso = int(df_a["DIAS_EM_ABERTO"].max())
            valor_parado = (df_a["SALDO"] * df_a["PRECO_UNITARIO"]).sum()

            ca1, ca2, ca3, ca4 = st.columns(4)
            ca1.metric("🔴 Críticos (>90 dias)", qtd_critico)
            ca2.metric(f"🟠 Atenção (≥{int(dias_alerta_atraso)} dias)", qtd_atencao)
            ca3.metric("Maior atraso", f"{maior_atraso} dias")
            ca4.metric("Saldo parado (valor est.)", f"R$ {formatar_numero_br(valor_parado)}")
            st.caption(
                "Saldo parado estimado = SALDO × PREÇO_UNITARIO das remessas em aberto (indicador "
                "aproximado, não é um valor fiscal). 'Dias em aberto' conta a partir da emissão da "
                "remessa até hoje — como não dá pra saber a data exata de cada devolução parcial, o "
                "prazo é sempre contado a partir do envio original."
            )

            st.divider()

            df_resumo_atraso = (
                df_a
                .groupby(
                    [
                        "LOTE", "PRODUTO", "DESCRICAO", "CLIENTE_FORN", "RAZAO_SOCIAL", "CIDADE", "ESTADO",
                        "DOC_ORIGINAL", "SERIE",
                    ],
                    dropna=False,
                )
                .agg(
                    DT_EMISSAO=("DT_EMISSAO", "min"),
                    DIAS_EM_ABERTO=("DIAS_EM_ABERTO", "max"),
                    saldoemterceiro=("SALDO", "sum"),
                )
                .reset_index()
                .rename(columns={"DOC_ORIGINAL": "NOTA_SAIDA"})
            )
            df_resumo_atraso["LOTE"] = df_resumo_atraso["LOTE"].fillna("(sem lote)")
            df_resumo_atraso["STATUS_ATRASO"] = df_resumo_atraso["DIAS_EM_ABERTO"].apply(_status_atraso)
            df_resumo_atraso = df_resumo_atraso.sort_values("DIAS_EM_ABERTO", ascending=False)

            df_resumo_atraso_exib = formatar_datas_para_exibicao(df_resumo_atraso)
            df_resumo_atraso_exib = df_resumo_atraso_exib[
                [
                    "STATUS_ATRASO", "DIAS_EM_ABERTO", "LOTE", "PRODUTO", "DESCRICAO", "NOTA_SAIDA", "SERIE",
                    "DT_EMISSAO", "CLIENTE_FORN", "RAZAO_SOCIAL", "CIDADE", "ESTADO", "saldoemterceiro",
                ]
            ]

            def _destacar_atraso(row):
                cor = ""
                status_linha = row.get("STATUS_ATRASO", "")
                if status_linha == "🔴 CRÍTICO (>90 dias)":
                    cor = "background-color: #f8d7da"
                elif status_linha == "🟠 ATENÇÃO":
                    cor = "background-color: #ffe8cc"
                return [cor] * len(row)

            estilo_atraso = df_resumo_atraso_exib.style.apply(_destacar_atraso, axis=1)
            st.dataframe(estilo_atraso, use_container_width=True, hide_index=True, height=480)

            excel_atraso = gerar_excel(df_resumo_atraso_exib)
            st.download_button(
                label="⬇️ Baixar Excel",
                data=excel_atraso,
                file_name=f"atraso_terceiros_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_atraso_terceiros",
            )
