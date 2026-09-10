"""
Verifica se um numero de documento (nota fiscal) aparece na SB6010
(controle de terceiros) e/ou nas SD1010/SD2010 (documentos de
entrada/saida). Usa o mesmo .env do app.py.

Serve para descobrir, quando um documento nao aparece no relatorio, se ele
simplesmente nao existe na tabela esperada (por exemplo: a NF foi lancada
com uma TES que nao atualiza a SB6010, ou esta em outra filial) - em vez
de ficar no chute.

Uso:
    python verificar_nf.py 43
    python verificar_nf.py 000000043
"""

import os
import sys

from dotenv import load_dotenv

load_dotenv()

if len(sys.argv) < 2:
    print("Uso: python verificar_nf.py NUMERO_DOCUMENTO   (ex: 43 ou 000000043)")
    sys.exit(1)

doc_raw = sys.argv[1].strip()
doc_padded = doc_raw.zfill(9)  # Protheus geralmente grava o doc com 9 digitos, zero a esquerda

servidor = os.getenv("SQLSERVER_SERVIDOR")
porta = os.getenv("SQLSERVER_PORTA", "1433")
database = os.getenv("SQLSERVER_DATABASE")
usuario = os.getenv("SQLSERVER_USUARIO")
senha = os.getenv("SQLSERVER_SENHA")
driver = os.getenv("SQLSERVER_DRIVER", "ODBC Driver 18 for SQL Server")
odbc_extra = os.getenv("SQLSERVER_ODBC_EXTRA", "TrustServerCertificate=yes;Encrypt=yes")

faltando = [k for k, v in {
    "SQLSERVER_SERVIDOR": servidor, "SQLSERVER_DATABASE": database,
    "SQLSERVER_USUARIO": usuario, "SQLSERVER_SENHA": senha,
}.items() if not v]
if faltando:
    print(f"Faltam variaveis no .env: {', '.join(faltando)}")
    sys.exit(1)

from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

query = {"driver": driver}
for par in odbc_extra.split(";"):
    par = par.strip()
    if not par:
        continue
    chave, _, valor = par.partition("=")
    if chave:
        query[chave] = valor

url = URL.create(
    "mssql+pyodbc", username=usuario, password=senha, host=servidor,
    port=int(porta), database=database, query=query,
)
engine = create_engine(url)

print(f"Procurando documento '{doc_raw}' (tambem tentando com zeros: '{doc_padded}')...\n")


def mostrar(titulo, resultado):
    print(f"--- {titulo} ---")
    linhas = resultado.fetchall()
    colunas = resultado.keys()
    if not linhas:
        print("  (nenhum registro encontrado)")
    for l in linhas:
        print(" ", dict(zip(colunas, l)))
    print()


with engine.connect() as conn:
    r = conn.execute(
        text(
            "SELECT B6_FILIAL, B6_DOC, B6_SERIE, B6_PRODUTO, B6_CLIFOR, B6_LOJA, "
            "B6_LOCAL, B6_QUANT, B6_SALDO, B6_TPCF, B6_TIPO, B6_PODER3, B6_EMISSAO, D_E_L_E_T_ "
            "FROM SB6010 WHERE RTRIM(B6_DOC) IN (:d1, :d2) ORDER BY R_E_C_N_O_"
        ),
        {"d1": doc_raw, "d2": doc_padded},
    )
    mostrar("SB6010 (controle de terceiros)", r)

    r = conn.execute(
        text(
            "SELECT D1_FILIAL, D1_DOC, D1_SERIE, D1_COD, D1_FORNECE, D1_LOJA, D1_LOCAL, "
            "D1_QUANT, D1_LOTECTL, D1_EMISSAO, D_E_L_E_T_ "
            "FROM SD1010 WHERE RTRIM(D1_DOC) IN (:d1, :d2) ORDER BY R_E_C_N_O_"
        ),
        {"d1": doc_raw, "d2": doc_padded},
    )
    mostrar("SD1010 (documentos de entrada)", r)

    r = conn.execute(
        text(
            "SELECT D2_FILIAL, D2_DOC, D2_SERIE, D2_COD, D2_CLIENTE, D2_LOJA, D2_LOCAL, "
            "D2_QUANT, D2_LOTECTL, D2_EMISSAO, D_E_L_E_T_ "
            "FROM SD2010 WHERE RTRIM(D2_DOC) IN (:d1, :d2) ORDER BY R_E_C_N_O_"
        ),
        {"d1": doc_raw, "d2": doc_padded},
    )
    mostrar("SD2010 (documentos de saida)", r)
