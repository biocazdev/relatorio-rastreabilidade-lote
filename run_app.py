"""
Launcher usado pelo PyInstaller para empacotar o app Streamlit como um
programa de verdade: janela propria (via pywebview), sem aba de navegador
e sem janela de terminal visivel.

Como o executavel gerado pelo PyInstaller nao inclui um "python.exe"
separado para rodar `streamlit run app.py` como um subprocesso comum,
este script se reinvoca a si mesmo (o proprio RelatorioMVPed.exe) com um
argumento especial para servir de "processo servidor" escondido - e essa
segunda instancia que roda o Streamlit de verdade. A primeira instancia
(a que a pessoa realmente abre, pelo atalho) so sobe esse processo filho
e abre a janela do pywebview apontando para o servidor local.

Nao precisa mexer aqui no dia a dia - isso so entra em cena quando o app
ja foi empacotado pelo build_exe.bat. Para desenvolvimento, continue
usando `streamlit run app.py` normalmente.
"""
import os
import subprocess
import sys
import time
import traceback
import urllib.request
from pathlib import Path

_FLAG_SERVIDOR = "--modo-servidor-streamlit"
_PORTA = "8765"


def _base_path() -> Path:
    """Pasta com os arquivos empacotados (app.py, conn.enc, logos, ...)."""
    return Path(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))))


def _pasta_dados_usuario() -> Path:
    base = os.getenv("APPDATA") or os.getenv("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "RelatorioMVPed"


def _preparar_credenciais_streamlit() -> None:
    """Evita o prompt de 'digite seu e-mail' que o Streamlit mostra no
    primeiro uso de `streamlit run` em qualquer maquina - sem isso, o
    processo servidor (sem console de verdade) travaria esperando uma
    resposta que nunca vem."""
    caminho = Path.home() / ".streamlit" / "credentials.toml"
    if caminho.exists():
        return
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text('[general]\nemail = ""\n', encoding="utf-8")


def _registrar_erro(origem: str, exc: BaseException) -> None:
    """Como nao ha janela de terminal para mostrar erros, qualquer
    exception vira do processo servidor cai aqui - facilita suporte."""
    try:
        pasta = _pasta_dados_usuario()
        pasta.mkdir(parents=True, exist_ok=True)
        with open(pasta / "erro.log", "a", encoding="utf-8") as f:
            f.write(f"\n--- {origem} ---\n")
            f.write(traceback.format_exc())
    except Exception:
        pass


def _rodar_como_servidor() -> None:
    """Processo filho, escondido: sobe o Streamlit de verdade (bloqueante,
    sem abrir navegador sozinho - a janela do pywebview e quem mostra a
    tela, no processo pai)."""
    base = _base_path()
    os.chdir(base)
    _preparar_credenciais_streamlit()
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

    try:
        from streamlit.web import cli as stcli

        sys.argv = [
            "streamlit",
            "run",
            str(base / "app.py"),
            "--server.headless=true",
            f"--server.port={_PORTA}",
            "--global.developmentMode=false",
            "--browser.gatherUsageStats=false",
        ]
        sys.exit(stcli.main())
    except SystemExit:
        raise
    except BaseException as exc:  # nao deixa o processo servidor morrer silencioso
        _registrar_erro("processo servidor (streamlit)", exc)
        raise


def _rodar_como_janela() -> None:
    """Processo principal (o que a pessoa realmente abre): sobe o processo
    servidor escondido e abre a janela nativa do pywebview apontando pra
    ele."""
    exe = sys.executable
    args = [exe, _FLAG_SERVIDOR]
    if not getattr(sys, "frozen", False):
        # em desenvolvimento (nao empacotado), sys.executable e o python.exe
        # "puro" - precisa apontar de volta pra este arquivo .py
        args = [exe, os.path.abspath(__file__), _FLAG_SERVIDOR]

    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    try:
        processo = subprocess.Popen(args, creationflags=creationflags)
    except Exception as exc:
        _registrar_erro("iniciar processo servidor", exc)
        raise

    url = f"http://localhost:{_PORTA}"
    servidor_respondeu = False
    for _ in range(120):  # ate 60s esperando o Streamlit subir
        try:
            urllib.request.urlopen(url, timeout=1)
            servidor_respondeu = True
            break
        except Exception:
            if processo.poll() is not None:
                # o processo servidor morreu sozinho - nao adianta esperar mais
                break
            time.sleep(0.5)

    if not servidor_respondeu:
        _registrar_erro(
            "servidor nao respondeu a tempo",
            RuntimeError("Streamlit nao subiu em 60s - veja se ha outro erro logo acima neste log."),
        )

    import webview

    class _ApiJanela:
        """Exposta como `window.pywebview.api` dentro da pagina - e o que o
        botao "Sair do aplicativo" (na sidebar do app.py) chama para fechar
        a janela de verdade, ao inves de so a aba/pagina."""

        def sair(self) -> None:
            for janela in list(webview.windows):
                janela.destroy()

    # Sem isso, os links de download que o Streamlit gera ("Baixar Excel",
    # "Baixar PDF") ficam mudos dentro da janela do pywebview - por padrao,
    # o pywebview bloqueia downloads dentro da propria janela por seguranca.
    webview.settings['ALLOW_DOWNLOADS'] = True

    webview.create_window(
        "Relatório de Rastreabilidade de Lote",
        url,
        width=1300,
        height=850,
        min_size=(900, 600),
        # Libera o Ctrl + scroll do mouse e o Ctrl com +/- para quem usa
        # o app poder aumentar ou diminuir o zoom da tela, igual num
        # navegador comum. Sem isso o pywebview trava o zoom fixo em 100%.
        zoomable=True,
        # Liga o botao "Sair do aplicativo" da sidebar (ver _ApiJanela acima).
        js_api=_ApiJanela(),
    )
    webview.start()

    processo.terminate()
    try:
        processo.wait(timeout=5)
    except Exception:
        processo.kill()


def main() -> None:
    if _FLAG_SERVIDOR in sys.argv:
        _rodar_como_servidor()
    else:
        _rodar_como_janela()


if __name__ == "__main__":
    main()
