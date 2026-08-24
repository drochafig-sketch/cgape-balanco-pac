# -*- coding: utf-8 -*-
"""
CGAPE / BALANCO PAC - Diagnostico de ambiente.

Roda sozinho, sem dependencia externa. Confere, nesta ordem:
  1. Versao e arquitetura do Python
  2. Bibliotecas de terceiros usadas pelo relatorio, com versao
  3. Comparacao com o requirements-travado.txt da maquina antiga
  4. Arquivos de apoio que o script procura na propria pasta
  5. Itens especificos do Windows (tkinter, WebView2, sockets)

Uso:
    python verificar_ambiente.py
    python verificar_ambiente.py --pasta "C:\\caminho\\do\\projeto"
"""

import os
import sys
import platform
import struct
import socket
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

OK = "[ OK ]"
ERRO = "[FALHA]"
ALERTA = "[AVISO]"

PACOTES = {
    "pandas": "pandas",
    "numpy": "numpy",
    "reportlab": "reportlab",
    "openpyxl": "openpyxl",
    "webview": "pywebview",
}

ARQUIVOS_APOIO = [
    ("PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026.xlsx", True),
    ("PAC.png", True),
    ("GOVERNO.PNG", True),
    ("municipios_bahia.geojson", False),
    ("calibri.ttf", False),
    ("calibrib.ttf", False),
    ("DIN.ttf", False),
    ("DIN-Bold.ttf", False),
    ("bahnschrift.ttf", False),
]

problemas = []
avisos = []


def titulo(texto):
    print()
    print("=" * 62)
    print(" " + texto)
    print("=" * 62)


def verificar_python():
    titulo("1. PYTHON")
    print(f"  Versao      : {platform.python_version()}")
    print(f"  Arquitetura : {struct.calcsize('P') * 8} bits")
    print(f"  Executavel  : {sys.executable}")
    print(f"  Sistema     : {platform.platform()}")
    dentro_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    print(f"  Ambiente    : {'.venv ativo' if dentro_venv else 'Python global'}")
    if not dentro_venv:
        avisos.append(
            "Voce esta no Python global, nao no .venv. "
            "Selecione .venv\\Scripts\\python.exe no VS Code."
        )
    if struct.calcsize("P") * 8 != 64:
        problemas.append("Python 32 bits. Use a versao 64 bits.")


def versao_de(modulo, nome_pip):
    try:
        import importlib.metadata as md
        return md.version(nome_pip)
    except Exception:
        return getattr(modulo, "__version__", "?")


def verificar_bibliotecas():
    titulo("2. BIBLIOTECAS DO RELATORIO")
    encontradas = {}
    for nome_import, nome_pip in PACOTES.items():
        try:
            mod = __import__(nome_import)
            v = versao_de(mod, nome_pip)
            encontradas[nome_pip.lower()] = v
            print(f"  {OK} {nome_pip:<12} {v}")
        except Exception as e:
            print(f"  {ERRO} {nome_pip:<12} nao instalado ({e.__class__.__name__})")
            problemas.append(f"Biblioteca ausente: {nome_pip}")
    return encontradas


def verificar_divergencias(encontradas, pasta):
    caminho = os.path.join(pasta, "_EXPORTADO", "requirements-travado.txt")
    if not os.path.exists(caminho):
        caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "_EXPORTADO", "requirements-travado.txt")
    titulo("3. COMPARACAO COM A MAQUINA ANTIGA")
    if not os.path.exists(caminho):
        print("  requirements-travado.txt nao encontrado. Comparacao pulada.")
        return
    esperadas = {}
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "==" not in linha:
                continue
            nome, _, versao = linha.partition("==")
            esperadas[nome.strip().lower().replace("_", "-")] = versao.strip()
    if not esperadas:
        print("  Arquivo sem versoes travadas (==). Comparacao pulada.")
        return
    houve = False
    for nome_pip in [p.lower() for p in PACOTES.values()]:
        atual = encontradas.get(nome_pip)
        antiga = esperadas.get(nome_pip)
        if antiga and atual and atual != antiga:
            houve = True
            print(f"  {ALERTA} {nome_pip}: antiga {antiga} / atual {atual}")
            avisos.append(f"{nome_pip} mudou de versao ({antiga} para {atual}).")
        elif antiga and atual:
            print(f"  {OK} {nome_pip}: {atual} (identica)")
    if not houve:
        print("  Nenhuma divergencia nas bibliotecas criticas.")


def verificar_arquivos(pasta):
    titulo("4. ARQUIVOS DE APOIO EM " + pasta)
    for nome, obrigatorio in ARQUIVOS_APOIO:
        caminho = os.path.join(pasta, nome)
        if os.path.exists(caminho):
            tamanho = os.path.getsize(caminho)
            print(f"  {OK} {nome}  ({tamanho:,} bytes)".replace(",", "."))
        elif obrigatorio:
            print(f"  {ERRO} {nome}  AUSENTE (o programa nao roda sem ele)")
            problemas.append(f"Arquivo obrigatorio ausente: {nome}")
        else:
            print(f"  {ALERTA} {nome}  ausente (o programa roda, com fonte/mapa de reserva)")
            avisos.append(f"Recurso opcional ausente: {nome}")


def verificar_windows(pasta):
    titulo("5. ITENS DO WINDOWS")

    try:
        import tkinter  # noqa: F401
        print(f"  {OK} tkinter disponivel")
    except Exception:
        print(f"  {ERRO} tkinter ausente. Reinstale o Python marcando 'tcl/tk and IDLE'.")
        problemas.append("tkinter nao esta disponivel.")

    # Teste de socket: reproduz o cenario do WinError 10022 do pywebview.
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(("127.0.0.1", 0))
        porta = s.getsockname()[1]
        s.close()
        print(f"  {OK} Sockets locais funcionando (porta livre testada: {porta})")
    except OSError as e:
        codigo = getattr(e, "winerror", None) or e.errno
        print(f"  {ERRO} Falha ao abrir socket local (erro {codigo}).")
        print("         Catalogo Winsock provavelmente corrompido.")
        print("         Correcao: abrir CMD como administrador e rodar")
        print("           netsh winsock reset")
        print("         depois REINICIAR o computador.")
        problemas.append(f"Socket local falhou (erro {codigo}). Rode 'netsh winsock reset'.")

    if platform.system() == "Windows":
        achou = False
        try:
            import winreg
            chaves = [
                (winreg.HKEY_LOCAL_MACHINE,
                 r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
                (winreg.HKEY_LOCAL_MACHINE,
                 r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
            ]
            for raiz, caminho_reg in chaves:
                try:
                    with winreg.OpenKey(raiz, caminho_reg) as k:
                        v, _ = winreg.QueryValueEx(k, "pv")
                        print(f"  {OK} WebView2 Runtime instalado (versao {v})")
                        achou = True
                        break
                except FileNotFoundError:
                    continue
        except Exception:
            pass
        if not achou:
            print(f"  {ALERTA} WebView2 Runtime nao detectado no registro.")
            print("         O painel do pywebview pode nao abrir.")
            print("         Baixe 'Evergreen Standalone Installer' em:")
            print("         https://developer.microsoft.com/microsoft-edge/webview2/")
            avisos.append("WebView2 Runtime nao detectado.")
    else:
        print(f"  {ALERTA} Sistema nao e Windows. Checagens especificas puladas.")


def teste_reportlab(pasta):
    titulo("6. TESTE PRATICO DO REPORTLAB")
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        import tempfile

        fonte_usada = "Helvetica"
        alvo = os.path.join(pasta, "calibri.ttf")
        if os.path.exists(alvo):
            try:
                pdfmetrics.registerFont(TTFont("Calibri", alvo))
                fonte_usada = "Calibri"
            except Exception as e:
                print(f"  {ALERTA} calibri.ttf existe mas nao registrou: {e}")

        destino = os.path.join(tempfile.gettempdir(), "teste_cgape.pdf")
        estilo = getSampleStyleSheet()["Normal"]
        estilo.fontName = fonte_usada
        doc = SimpleDocTemplate(destino)
        doc.build([Paragraph("Teste de geracao CGAPE", estilo)])
        print(f"  {OK} PDF de teste gerado com a fonte {fonte_usada}")
        print(f"       {destino}")
        os.remove(destino)
    except Exception as e:
        print(f"  {ERRO} Nao consegui gerar PDF de teste: {e}")
        problemas.append(f"ReportLab falhou no teste pratico: {e}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--pasta", default=None,
                   help="Pasta do projeto (onde estao o .py, a planilha e as fontes)")
    args = p.parse_args()

    pasta = args.pasta or os.path.dirname(os.path.abspath(__file__))
    pasta = os.path.abspath(pasta)

    print()
    print("#" * 62)
    print("#  CGAPE / BALANCO PAC - DIAGNOSTICO DE AMBIENTE")
    print("#" * 62)

    verificar_python()
    encontradas = verificar_bibliotecas()
    verificar_divergencias(encontradas, pasta)
    verificar_arquivos(pasta)
    verificar_windows(pasta)
    teste_reportlab(pasta)

    titulo("RESUMO")
    if problemas:
        print(f"  {len(problemas)} problema(s) que impedem ou quebram a execucao:")
        for i, x in enumerate(problemas, 1):
            print(f"    {i}. {x}")
    else:
        print("  Nenhum problema bloqueante encontrado.")
    if avisos:
        print()
        print(f"  {len(avisos)} aviso(s):")
        for i, x in enumerate(avisos, 1):
            print(f"    {i}. {x}")
    print()
    if not problemas:
        print("  Ambiente pronto para rodar o relatorio.")
    print()
    return 1 if problemas else 0


if __name__ == "__main__":
    sys.exit(main())
