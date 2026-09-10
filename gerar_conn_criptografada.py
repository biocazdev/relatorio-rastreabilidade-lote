"""
Gera conn.enc a partir do .env real desta pasta - roda automaticamente
dentro do build_exe.bat, antes do PyInstaller, para embutir a string de
conexao (criptografada) dentro do instalador.

Nao precisa rodar isso manualmente no dia a dia (o build_exe.bat ja
chama este script sozinho a cada build). So use direto se quiser gerar
um conn.enc "na mao" pra testar.

IMPORTANTE: a chave abaixo (_CHAVE) tem que ser EXATAMENTE a mesma que
esta em app.py (funcao _config_embutida_valores) - senao o app nao
consegue decifrar o conn.enc gerado aqui. Se um dia trocar a chave,
troque nos dois arquivos.
"""
from pathlib import Path

from cryptography.fernet import Fernet

_CHAVE = b"CZ1ZqsuHru8OvAEQ9B0MPMd77M3PFZAQSgtxWEbDc54="


def main() -> None:
    env_path = Path(".env")
    if not env_path.exists():
        raise SystemExit(
            "[ERRO] .env nao encontrado nesta pasta - crie a partir de "
            ".env.example e preencha com os dados reais de conexao antes "
            "de gerar o build."
        )

    conteudo = env_path.read_bytes()
    cifrado = Fernet(_CHAVE).encrypt(conteudo)
    Path("conn.enc").write_bytes(cifrado)
    print("conn.enc gerado com sucesso a partir do .env atual.")


if __name__ == "__main__":
    main()
