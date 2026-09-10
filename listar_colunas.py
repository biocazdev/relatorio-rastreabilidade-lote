"""
Lista as colunas reais de uma tabela do banco (usa o mesmo .env do app.py).
Serve para conferir o nome exato de campos antes de usa-los numa consulta -
por exemplo, descobrir se o saldo do lote na SB8010 chama B8_QUANT, B8_SALDO,
ou outro nome.

Uso:
    python listar_colunas.py SB8010
    python listar_colunas.py SD5010
"""

import os
import sys

from dotenv import load_dotenv

load_dotenv()

if len(sys.argv) < 2:
    print("Uso: python listar_colunas.py NOME_DA_TABELA   (ex: SB8010)")
    sys.exit(1)

tabela = sys.argv[1]

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
    print(f"❌ Faltam variáveis no .env: {', '.join(faltando)}")
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

with engine.connect() as conn:
    print(f"Colunas de {tabela}:\n")
    resultado = conn.execute(
        text(
            "SELECT COLUMN_NAME, DATA_TYPE "
            "FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_NAME = :tabela "
            "ORDER BY ORDINAL_POSITION"
        ),
        {"tabela": tabela},
    )
    linhas = resultado.fetchall()
    if not linhas:
        print(f"  (nenhuma coluna encontrada — confira se '{tabela}' é o nome certo da tabela)")
    for nome, tipo in linhas:
        print(f"  {nome:<20} {tipo}")

    # Dá uma dica visual: mostra 1 linha de exemplo, se houver dados
    try:
        exemplo = conn.execute(text(f"SELECT TOP 1 * FROM {tabela}")).mappings().first()
        if exemplo:
            print("\nExemplo de 1 linha:")
            for k, v in exemplo.items():
                print(f"  {k:<20} = {v}")
    except Exception as e:
        print(f"\n(não consegui puxar uma linha de exemplo: {e})")
