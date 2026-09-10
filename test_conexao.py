"""
Teste rápido de conexão com o SQL Server, usando o mesmo .env do app.py.
Uso:
    python test_conexao.py
"""

import os
import sys

from dotenv import load_dotenv

load_dotenv()

obrigatorias = ["SQLSERVER_SERVIDOR", "SQLSERVER_DATABASE", "SQLSERVER_USUARIO", "SQLSERVER_SENHA"]
faltando = [v for v in obrigatorias if not os.getenv(v)]
if faltando:
    print(f"❌ Faltam variáveis no .env: {', '.join(faltando)}")
    sys.exit(1)

servidor = os.getenv("SQLSERVER_SERVIDOR")
porta = os.getenv("SQLSERVER_PORTA", "1433")
database = os.getenv("SQLSERVER_DATABASE")
usuario = os.getenv("SQLSERVER_USUARIO")
senha = os.getenv("SQLSERVER_SENHA")
driver = os.getenv("SQLSERVER_DRIVER", "ODBC Driver 18 for SQL Server")
odbc_extra = os.getenv("SQLSERVER_ODBC_EXTRA", "TrustServerCertificate=yes;Encrypt=yes")

print(f"Servidor : {servidor}:{porta}")
print(f"Database : {database}")
print(f"Usuário  : {usuario}")
print(f"Driver   : {driver}")
print("Conectando...")

try:
    from sqlalchemy import create_engine, text
except ImportError:
    print("❌ SQLAlchemy não instalado. Rode: pip install -r requirements.txt")
    sys.exit(1)

conn_str = (
    f"mssql+pyodbc://{usuario}:{senha}@{servidor}:{porta}/{database}"
    f"?driver={driver.replace(' ', '+')}&{odbc_extra}"
)

try:
    engine = create_engine(conn_str)
    with engine.connect() as conn:
        versao = conn.execute(text("SELECT @@VERSION AS v")).scalar()
        print("✅ Conexão OK!")
        print(f"SQL Server: {versao.splitlines()[0]}")

        for tabela in ("SD1010", "SD2010", "SB1010", "SD5010", "SB8010"):
            try:
                qtd = conn.execute(text(f"SELECT COUNT(*) FROM {tabela}")).scalar()
                print(f"  {tabela}: {qtd} linhas (usuário tem acesso de leitura)")
            except Exception as e:
                print(f"  ⚠️  {tabela}: erro ao ler -> {e}")

except Exception as e:
    print(f"❌ Falha na conexão: {e}")
    sys.exit(1)
