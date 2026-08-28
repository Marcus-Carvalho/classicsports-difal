# -*- coding: utf-8 -*-
"""
Classic Sports - Guias DIFAL
Instalador Windows (empacotado com PyInstaller).

Cria a pasta do programa, baixa os arquivos do GitHub, detecta/instala o
Python, instala as dependencias, cria o .bat, o atalho e o registro em
"Adicionar ou Remover Programas".

Gerar o .exe (use o build_instalador.bat, que ja embute os arquivos):
    python -m PyInstaller --onefile --windowed --icon=icon.ico ^
        --add-data "difal_core.py;." --add-data "difal_painel.py;." ^
        --add-data "atualizador.py;." --add-data "modelo_c6.xlsx;." ^
        --add-data "icon.ico;." --add-data "LEIAME.txt;." ^
        --add-data "ClassicSportsDIFAL.bat;." --add-data "versao.json;." ^
        --name=ClassicSports_DIFAL_Instalador instalador_gui.py
"""

import os
import sys
import json
import shutil
import subprocess
import tempfile
import threading
import urllib.request
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

REPOSITORIO = "Marcus-Carvalho/classicsports-difal"
URL_RAW_BASE = "https://raw.githubusercontent.com/" + REPOSITORIO + "/main/"
URL_PYTHON = "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"

PASTA_PADRAO = r"C:\ClassicSportsApps\DIFAL"
NOME_APP = "Classic Sports - Guias DIFAL"
CHAVE_DESINSTALAR = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\ClassicSportsDIFAL"

ARQUIVOS = [
    "difal_core.py",
    "difal_painel.py",
    "atualizador.py",
    "modelo_c6.xlsx",
    "icon.ico",
    "LEIAME.txt",
    "ClassicSportsDIFAL.bat",
    "versao.json",
]

DEPENDENCIAS = ["pdfplumber", "openpyxl"]

BG = "#001228"
HEADER = "#00142E"
CARD = "#002850"
CYAN = "#00AAFF"
BRIGHT = "#E0ECFA"
MUTED = "#7A9DBE"
SUCCESS = "#2ECC71"
BTN = "#0055AA"


def baixar(nome_ou_url, destino=None, tempo=60):
    url = nome_ou_url if nome_ou_url.startswith("http") else URL_RAW_BASE + nome_ou_url
    requisicao = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
    with urllib.request.urlopen(requisicao, timeout=tempo) as resposta:
        conteudo = resposta.read()
    if destino:
        Path(destino).parent.mkdir(parents=True, exist_ok=True)
        Path(destino).write_bytes(conteudo)
    return conteudo


def pastas_recursos():
    """Onde procurar os arquivos que vieram junto com o instalador."""
    lugares = []
    embutido = getattr(sys, "_MEIPASS", None)          # dentro do .exe do PyInstaller
    if embutido:
        lugares.append(Path(embutido))
    try:
        lugares.append(Path(sys.argv[0]).resolve().parent)    # ao lado do .exe
    except Exception:
        pass
    lugares.append(Path(__file__).resolve().parent)           # rodando como .py
    return lugares


def obter_arquivo(nome, destino):
    """
    Usa o arquivo que veio junto com o instalador; se nao houver, baixa do
    GitHub. Devolve a origem usada, so para aparecer no log.
    """
    destino = Path(destino)
    for pasta in pastas_recursos():
        origem = pasta / nome
        try:
            if not origem.exists() or origem.stat().st_size == 0:
                continue
            if origem.resolve() == destino.resolve():
                continue
        except Exception:
            continue
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(origem), str(destino))
        return "embutido"
    baixar(nome, destino)
    return "GitHub"


def achar_python():
    """Procura um python.exe utilizavel na maquina."""
    candidatos = []
    for comando in ("python", "python3", "py"):
        caminho = shutil.which(comando)
        if caminho:
            candidatos.append(caminho)
    base = Path(os.path.expanduser("~")) / "AppData" / "Local" / "Programs" / "Python"
    if base.exists():
        for pasta in sorted(base.glob("Python3*"), reverse=True):
            exe = pasta / "python.exe"
            if exe.exists():
                candidatos.append(str(exe))
    for pasta in (r"C:\Python312", r"C:\Python311", r"C:\Python310"):
        exe = Path(pasta) / "python.exe"
        if exe.exists():
            candidatos.append(str(exe))

    for caminho in candidatos:
        try:
            saida = subprocess.run([caminho, "--version"], capture_output=True, text=True, timeout=20)
            if saida.returncode == 0 and "Python 3" in (saida.stdout + saida.stderr):
                return caminho
        except Exception:
            continue
    return None


def criar_atalho(destino_lnk, alvo, pasta_trabalho, icone):
    """Cria atalho no Windows sem depender de biblioteca externa."""
    script = (
        '$s = New-Object -ComObject WScript.Shell;'
        '$a = $s.CreateShortcut("' + str(destino_lnk) + '");'
        '$a.TargetPath = "' + str(alvo) + '";'
        '$a.WorkingDirectory = "' + str(pasta_trabalho) + '";'
        '$a.IconLocation = "' + str(icone) + '";'
        '$a.Save()'
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", script],
                   capture_output=True, timeout=60)


class Instalador(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Instalador - " + NOME_APP)
        self.geometry("720x520")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.pasta = tk.StringVar(value=PASTA_PADRAO)
        self.instalando = False
        self._montar()

    def _montar(self):
        topo = tk.Frame(self, bg=HEADER, height=76)
        topo.pack(fill="x")
        topo.pack_propagate(False)
        tk.Label(topo, text="GUIAS DIFAL", bg=HEADER, fg=BRIGHT,
                 font=("Segoe UI", 18, "bold")).pack(side="left", padx=22, pady=10)
        tk.Label(topo, text="Instalador  ·  Classic Sports", bg=HEADER, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="left")

        corpo = tk.Frame(self, bg=BG)
        corpo.pack(fill="both", expand=True, padx=22, pady=16)

        tk.Label(corpo, text="Pasta de instalação:", bg=BG, fg=BRIGHT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        linha = tk.Frame(corpo, bg=BG)
        linha.pack(fill="x", pady=(4, 14))
        tk.Entry(linha, textvariable=self.pasta, bg=CARD, fg=BRIGHT, relief="flat",
                 insertbackground=BRIGHT, font=("Segoe UI", 10)).pack(side="left", fill="x",
                                                                     expand=True, ipady=5)
        tk.Button(linha, text="Procurar", command=self.escolher, bg=BTN, fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), cursor="hand2",
                  padx=14).pack(side="left", padx=8)

        self.log = tk.Text(corpo, bg=CARD, fg=BRIGHT, relief="flat", height=15,
                           font=("Consolas", 9), wrap="word")
        self.log.pack(fill="both", expand=True)

        self.barra = ttk.Progressbar(corpo, mode="determinate", maximum=100)
        self.barra.pack(fill="x", pady=10)

        rodape = tk.Frame(self, bg=HEADER, height=68)
        rodape.pack(fill="x")
        rodape.pack_propagate(False)
        self.btn = tk.Button(rodape, text="INSTALAR", command=self.instalar, bg=SUCCESS,
                             fg="white", relief="flat", font=("Segoe UI", 11, "bold"),
                             cursor="hand2", width=18, pady=8)
        self.btn.pack(side="right", padx=20, pady=14)
        tk.Button(rodape, text="Sair", command=self.destroy, bg=BTN, fg="white",
                  relief="flat", font=("Segoe UI", 10), cursor="hand2",
                  width=10, pady=8).pack(side="right", pady=14)

        self.escrever("Este instalador vai:")
        self.escrever("  1. criar a pasta do programa")
        self.escrever("  2. instalar os arquivos que vieram dentro dele")
        self.escrever("     (se faltar algum, baixa de " + REPOSITORIO + ")")
        self.escrever("  3. localizar o Python (instala o 3.12 se não houver)")
        self.escrever("  4. instalar as bibliotecas " + ", ".join(DEPENDENCIAS))
        self.escrever("  5. criar o atalho na Área de Trabalho")
        self.escrever("")

    def escrever(self, texto):
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.update_idletasks()

    def passo(self, valor):
        self.barra["value"] = valor
        self.update_idletasks()

    def escolher(self):
        pasta = filedialog.askdirectory(title="Pasta de instalação")
        if pasta:
            self.pasta.set(pasta)

    def instalar(self):
        if self.instalando:
            return
        self.instalando = True
        self.btn.configure(state="disabled", text="INSTALANDO...")
        threading.Thread(target=self._instalar, daemon=True).start()

    def _instalar(self):
        try:
            destino = Path(self.pasta.get())
            destino.mkdir(parents=True, exist_ok=True)
            self.escrever("Pasta: " + str(destino))
            self.passo(5)

            # 1) arquivos do programa (embutidos no instalador; GitHub como reserva)
            temporaria = Path(tempfile.mkdtemp(prefix="difal_inst_"))
            for indice, nome in enumerate(ARQUIVOS, 1):
                self.escrever("Instalando " + nome + "...")
                origem = obter_arquivo(nome, temporaria / nome)
                if origem == "GitHub":
                    self.escrever("   (baixado do GitHub)")
                self.passo(5 + int(35 * indice / len(ARQUIVOS)))
            for nome in ARQUIVOS:
                shutil.copy2(str(temporaria / nome), str(destino / nome))
            shutil.rmtree(temporaria, ignore_errors=True)
            self.escrever("Arquivos instalados.")
            self.passo(45)

            # 2) Python
            python = achar_python()
            if not python:
                self.escrever("Python não encontrado. Baixando o Python 3.12...")
                instalador_py = Path(tempfile.gettempdir()) / "python_setup.exe"
                baixar(URL_PYTHON, instalador_py, tempo=600)
                self.escrever("Instalando o Python (pode demorar alguns minutos)...")
                subprocess.run([str(instalador_py), "/passive", "InstallAllUsers=0",
                                "PrependPath=1", "Include_pip=1"], timeout=1800)
                python = achar_python()
            if not python:
                raise RuntimeError("Não foi possível localizar o Python após a instalação.")
            self.escrever("Python: " + python)
            self.passo(60)

            # 3) dependencias
            self.escrever("Instalando " + ", ".join(DEPENDENCIAS) + "...")
            subprocess.run([python, "-m", "pip", "install", "--upgrade", "pip"],
                           capture_output=True, timeout=900)
            resultado = subprocess.run([python, "-m", "pip", "install"] + DEPENDENCIAS,
                                       capture_output=True, text=True, timeout=1800)
            if resultado.returncode != 0:
                self.escrever("Aviso do pip: " + (resultado.stderr or "")[-400:])
            self.passo(80)

            # 4) .bat com o caminho real do pythonw
            pythonw = str(Path(python).with_name("pythonw.exe"))
            if not Path(pythonw).exists():
                pythonw = python
            (destino / "pythonw_path.txt").write_text(pythonw, encoding="utf-8")
            bat = ('@echo off\r\n'
                   'cd /d "%~dp0"\r\n'
                   'start "" "' + pythonw + '" difal_painel.py\r\n')
            (destino / "ClassicSportsDIFAL.bat").write_text(bat, encoding="utf-8")

            desinstalar = ('@echo off\r\n'
                           'echo Removendo ' + NOME_APP + '...\r\n'
                           'reg delete "HKCU\\' + CHAVE_DESINSTALAR + '" /f >nul 2>&1\r\n'
                           'del "%USERPROFILE%\\Desktop\\' + NOME_APP + '.lnk" >nul 2>&1\r\n'
                           'echo Apague a pasta "' + str(destino) + '" para concluir.\r\n'
                           'pause\r\n')
            (destino / "desinstalar.bat").write_text(desinstalar, encoding="utf-8")
            self.passo(88)

            # 5) atalho + registro
            area_trabalho = Path(os.path.expanduser("~")) / "Desktop"
            criar_atalho(area_trabalho / (NOME_APP + ".lnk"),
                         destino / "ClassicSportsDIFAL.bat", destino, destino / "icon.ico")
            self.escrever("Atalho criado na Área de Trabalho.")

            try:
                versao = json.loads((destino / "versao.json").read_text(encoding="utf-8")).get("versao", "1.0.0")
            except Exception:
                versao = "1.0.0"
            try:
                import winreg
                chave = winreg.CreateKey(winreg.HKEY_CURRENT_USER, CHAVE_DESINSTALAR)
                winreg.SetValueEx(chave, "DisplayName", 0, winreg.REG_SZ, NOME_APP)
                winreg.SetValueEx(chave, "DisplayVersion", 0, winreg.REG_SZ, versao)
                winreg.SetValueEx(chave, "Publisher", 0, winreg.REG_SZ, "Classic Sports")
                winreg.SetValueEx(chave, "InstallLocation", 0, winreg.REG_SZ, str(destino))
                winreg.SetValueEx(chave, "DisplayIcon", 0, winreg.REG_SZ, str(destino / "icon.ico"))
                winreg.SetValueEx(chave, "UninstallString", 0, winreg.REG_SZ,
                                  str(destino / "desinstalar.bat"))
                winreg.CloseKey(chave)
            except Exception as erro:
                self.escrever("Aviso: não foi possível registrar em Programas (" + str(erro) + ")")

            self.passo(100)
            self.escrever("")
            self.escrever("INSTALAÇÃO CONCLUÍDA - versão " + versao)
            self.after(0, self._fim, destino)
        except Exception as erro:
            self.escrever("")
            self.escrever("ERRO: " + str(erro))
            self.after(0, self._erro, str(erro))

    def _fim(self, destino):
        self.instalando = False
        self.btn.configure(state="normal", text="INSTALAR")
        messagebox.showinfo("Instalador",
                            "Instalado em:\n" + str(destino) +
                            "\n\nUse o atalho na Área de Trabalho para abrir o programa.")
        self.destroy()

    def _erro(self, mensagem):
        self.instalando = False
        self.btn.configure(state="normal", text="INSTALAR")
        messagebox.showerror("Instalador", "Falha na instalação:\n\n" + mensagem)


if __name__ == "__main__":
    lock = Path(tempfile.gettempdir()) / "difal_instalador.lock"
    if lock.exists():
        try:
            lock.unlink()
        except Exception:
            pass
    lock.write_text("1", encoding="utf-8")
    try:
        Instalador().mainloop()
    finally:
        try:
            lock.unlink()
        except Exception:
            pass
