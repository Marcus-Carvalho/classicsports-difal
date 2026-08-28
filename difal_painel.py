# -*- coding: utf-8 -*-
"""
Classic Sports - Guias DIFAL
Painel (interface grafica).

Fluxo: adicionar os ZIPs -> processar -> planilha de conferencia
       -> exportar os arquivos no layout do banco C6 (100 por arquivo).
"""

import os
import sys
import json
import threading
import subprocess
import traceback
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

PASTA = Path(__file__).resolve().parent
sys.path.insert(0, str(PASTA))


def _verificar_deps():
    faltando = []
    for modulo, pacote in (("pdfplumber", "pdfplumber"), ("openpyxl", "openpyxl")):
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(pacote)
    if faltando:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install"] + faltando,
                           check=False, capture_output=True)
        except Exception:
            pass


_verificar_deps()
import difal_core as core          # noqa: E402
import atualizador                 # noqa: E402

VERSAO = atualizador.ler_versao_local()
ARQ_CONFIG = PASTA / "config_difal.json"
MODELO_BANCO = PASTA / "modelo_c6.xlsx"

# ------------------------------------------------------- cores da marca

BG = "#001228"
HEADER = "#00142E"
CARD = "#002850"
BLUE = "#0070CC"
CYAN = "#00AAFF"
BRIGHT = "#E0ECFA"
MUTED = "#7A9DBE"
SUCCESS = "#2ECC71"
WARNING = "#F39C12"
DANGER = "#E74C3C"
BTN = "#0055AA"
BTN_TXT = "white"


def carregar_config():
    if ARQ_CONFIG.exists():
        try:
            return json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def salvar_config(dados):
    try:
        ARQ_CONFIG.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def abrir_no_sistema(caminho):
    caminho = str(caminho)
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
    except Exception:
        pass


def brl(valor):
    return "R$ " + "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Classic Sports - Guias DIFAL  v" + VERSAO)
        self.geometry("1120x760")
        self.minsize(980, 660)
        self.configure(bg=BG)
        try:
            if (PASTA / "icon.ico").exists():
                self.iconbitmap(str(PASTA / "icon.ico"))
        except Exception:
            pass

        self.cfg = carregar_config()
        self.zips = []               # lista de Path
        self.resultado = None
        self.planilha = None
        self.pasta_banco = None
        self.carimbo = datetime.now().strftime("%Y-%m-%d_%H%M")
        self.processando = False

        self._estilos()
        self._montar()
        self._pintar_lista()

    # ------------------------------------------------ estilo

    def _estilos(self):
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except Exception:
            pass
        estilo.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=BRIGHT,
                         rowheight=24, borderwidth=0, font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", background=HEADER, foreground=CYAN,
                         font=("Segoe UI", 9, "bold"), borderwidth=0)
        estilo.map("Treeview", background=[("selected", BLUE)])
        estilo.configure("TNotebook", background=BG, borderwidth=0)
        estilo.configure("TNotebook.Tab", background=HEADER, foreground=MUTED,
                         padding=(18, 8), font=("Segoe UI", 9, "bold"))
        estilo.map("TNotebook.Tab", background=[("selected", CARD)],
                   foreground=[("selected", BRIGHT)])
        estilo.configure("Barra.Horizontal.TProgressbar", background=CYAN,
                         troughcolor=HEADER, borderwidth=0)

    def _botao(self, pai, texto, comando, cor=BTN, largura=18):
        return tk.Button(pai, text=texto, command=comando, bg=cor, fg=BTN_TXT,
                         activebackground=CYAN, activeforeground="white",
                         font=("Segoe UI", 10, "bold"), relief="flat",
                         cursor="hand2", width=largura, pady=7, bd=0)

    # ------------------------------------------------ layout

    def _montar(self):
        topo = tk.Frame(self, bg=HEADER, height=64)
        topo.pack(fill="x")
        topo.pack_propagate(False)
        tk.Label(topo, text="GUIAS DIFAL", bg=HEADER, fg=BRIGHT,
                 font=("Segoe UI", 17, "bold")).pack(side="left", padx=20)
        tk.Label(topo, text="Classic Sports  ·  GNRE e DARE-SP  ·  pagamento em massa C6",
                 bg=HEADER, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
        self.var_versao = tk.StringVar(value="v" + VERSAO)
        tk.Label(topo, textvariable=self.var_versao, bg=HEADER, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="right", padx=20)

        corpo = ttk.Notebook(self)
        corpo.pack(fill="both", expand=True, padx=18, pady=(14, 8))
        self.aba_lote = tk.Frame(corpo, bg=BG)
        self.aba_resumo = tk.Frame(corpo, bg=BG)
        self.aba_banco = tk.Frame(corpo, bg=BG)
        self.aba_relatorio = tk.Frame(corpo, bg=BG)
        self.aba_atualizacao = tk.Frame(corpo, bg=BG)
        corpo.add(self.aba_lote, text="LOTE")
        corpo.add(self.aba_resumo, text="RESUMO")
        corpo.add(self.aba_banco, text="BANCO C6")
        corpo.add(self.aba_relatorio, text="CONFERÊNCIA")
        corpo.add(self.aba_atualizacao, text="ATUALIZAÇÕES")
        self.abas = corpo

        self._aba_lote()
        self._aba_resumo()
        self._aba_banco()
        self._aba_relatorio()
        self._aba_atualizacao()

        rodape = tk.Frame(self, bg=HEADER, height=70)
        rodape.pack(fill="x")
        rodape.pack_propagate(False)
        self.barra = ttk.Progressbar(rodape, style="Barra.Horizontal.TProgressbar",
                                     mode="determinate", length=380)
        self.barra.pack(side="left", padx=20)
        self.var_status = tk.StringVar(value="Adicione os ZIPs de guias para começar.")
        tk.Label(rodape, textvariable=self.var_status, bg=HEADER, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=6)
        self.btn_processar = self._botao(rodape, "PROCESSAR LOTE", self.processar, SUCCESS, 18)
        self.btn_processar.pack(side="right", padx=18, pady=14)
        self.btn_planilha = self._botao(rodape, "Abrir planilha", self.abrir_planilha, BTN, 14)
        self.btn_planilha.pack(side="right", pady=14)
        self.btn_planilha.configure(state="disabled")

    # -------------------------------------------------- aba LOTE

    def _aba_lote(self):
        acoes = tk.Frame(self.aba_lote, bg=BG)
        acoes.pack(fill="x", pady=(12, 8))
        self._botao(acoes, "+ Adicionar ZIPs", self.adicionar_zips, BLUE, 17).pack(side="left", padx=(4, 6))
        self._botao(acoes, "+ Pasta inteira", self.adicionar_pasta, BTN, 15).pack(side="left", padx=6)
        self._botao(acoes, "Remover", self.remover_selecionados, BTN, 11).pack(side="left", padx=6)
        self._botao(acoes, "Limpar", self.limpar_lista, DANGER, 10).pack(side="left", padx=6)

        info = tk.Frame(self.aba_lote, bg=BG)
        info.pack(fill="x")
        self.var_info = tk.StringVar(value="Nenhum ZIP na lista.")
        tk.Label(info, textvariable=self.var_info, bg=BG, fg=CYAN,
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=4, pady=(0, 6))

        self.var_rapido = tk.BooleanVar(value=bool(self.cfg.get("modo_rapido", True)))
        tk.Checkbutton(info, text="Modo rápido (vários núcleos)", variable=self.var_rapido,
                       bg=BG, fg=MUTED, selectcolor=CARD, activebackground=BG,
                       activeforeground=BRIGHT, font=("Segoe UI", 9)).pack(side="right", padx=6)

        colunas = ("arquivo", "empresa", "pdfs", "situacao")
        self.tabela = ttk.Treeview(self.aba_lote, columns=colunas, show="headings",
                                   height=16, selectmode="extended")
        for coluna, titulo, largura in (("arquivo", "Arquivo ZIP", 430),
                                        ("empresa", "Empresa", 270),
                                        ("pdfs", "PDFs", 70),
                                        ("situacao", "Situação", 200)):
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor="w")
        self.tabela.pack(fill="both", expand=True, pady=4)
        self.tabela.tag_configure("ok", foreground=BRIGHT)
        self.tabela.tag_configure("alerta", foreground=WARNING)

        tk.Label(self.aba_lote,
                 text="Dica: selecione vários arquivos de uma vez (Ctrl ou Shift). "
                      "Arquivos repetidos são ignorados automaticamente.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=4, pady=(0, 6))

    # -------------------------------------------------- aba RESUMO

    def _aba_resumo(self):
        colunas = ("empresa", "guias", "total")
        self.tabela_resumo = ttk.Treeview(self.aba_resumo, columns=colunas,
                                          show="headings", height=18)
        for coluna, titulo, largura, alinha in (("empresa", "Empresa", 460, "w"),
                                                ("guias", "Guias", 100, "center"),
                                                ("total", "Total", 170, "e")):
            self.tabela_resumo.heading(coluna, text=titulo)
            self.tabela_resumo.column(coluna, width=largura, anchor=alinha)
        self.tabela_resumo.pack(fill="both", expand=True, pady=12)
        self.tabela_resumo.tag_configure("total", foreground=CYAN, font=("Segoe UI", 10, "bold"))

    # -------------------------------------------------- aba BANCO

    def _aba_banco(self):
        painel = tk.Frame(self.aba_banco, bg=CARD)
        painel.pack(fill="x", pady=(12, 8), padx=2)

        linha = tk.Frame(painel, bg=CARD)
        linha.pack(fill="x", padx=14, pady=12)

        tk.Label(linha, text="Data de pagamento:", bg=CARD, fg=BRIGHT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.var_data = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        tk.Entry(linha, textvariable=self.var_data, bg=HEADER, fg=BRIGHT, width=12,
                 insertbackground=BRIGHT, relief="flat", justify="center",
                 font=("Segoe UI", 10)).pack(side="left", padx=10, ipady=4)

        self.var_por_empresa = tk.BooleanVar(value=bool(self.cfg.get("por_empresa", True)))
        tk.Checkbutton(linha, text="Um arquivo por empresa (não mistura CNPJs)",
                       variable=self.var_por_empresa, bg=CARD, fg=BRIGHT, selectcolor=HEADER,
                       activebackground=CARD, activeforeground=BRIGHT,
                       font=("Segoe UI", 9)).pack(side="left", padx=18)

        self.btn_banco = self._botao(linha, "GERAR ARQUIVOS C6", self.exportar_banco, SUCCESS, 20)
        self.btn_banco.pack(side="right")
        self.btn_banco.configure(state="disabled")

        tk.Label(self.aba_banco,
                 text="O C6 aceita no máximo " + str(core.LIMITE_BANCO) +
                      " pagamentos por arquivo. O programa quebra o lote automaticamente.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=4)

        colunas = ("arquivo", "empresa", "guias", "total")
        self.tabela_banco = ttk.Treeview(self.aba_banco, columns=colunas,
                                         show="headings", height=15)
        for coluna, titulo, largura, alinha in (("arquivo", "Arquivo gerado", 400, "w"),
                                                ("empresa", "Empresa", 250, "w"),
                                                ("guias", "Guias", 80, "center"),
                                                ("total", "Total", 140, "e")):
            self.tabela_banco.heading(coluna, text=titulo)
            self.tabela_banco.column(coluna, width=largura, anchor=alinha)
        self.tabela_banco.pack(fill="both", expand=True, pady=8)
        self.tabela_banco.tag_configure("total", foreground=CYAN, font=("Segoe UI", 10, "bold"))
        self.tabela_banco.bind("<Double-1>", self._abrir_arquivo_banco)

        rodape = tk.Frame(self.aba_banco, bg=BG)
        rodape.pack(fill="x", pady=(0, 8))
        self._botao(rodape, "Abrir pasta dos arquivos", self.abrir_pasta_banco, BTN, 22).pack(side="left", padx=4)
        tk.Label(rodape, text="Duplo clique em um arquivo para abri-lo.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(side="left", padx=10)

    # -------------------------------------------------- aba CONFERENCIA

    def _aba_relatorio(self):
        quadro = tk.Frame(self.aba_relatorio, bg=BG)
        quadro.pack(fill="both", expand=True, pady=10)
        self.texto = tk.Text(quadro, bg=CARD, fg=BRIGHT, relief="flat",
                             font=("Consolas", 9), insertbackground=BRIGHT, wrap="none")
        rolagem = ttk.Scrollbar(quadro, orient="vertical", command=self.texto.yview)
        self.texto.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.texto.pack(side="left", fill="both", expand=True)

    # -------------------------------------------------- aba ATUALIZACOES

    def _aba_atualizacao(self):
        quadro = tk.Frame(self.aba_atualizacao, bg=CARD)
        quadro.pack(fill="x", pady=16, padx=2)

        tk.Label(quadro, text="Versão instalada", bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(14, 0))
        tk.Label(quadro, text="v" + VERSAO, bg=CARD, fg=CYAN,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=16)

        self.var_update = tk.StringVar(value="Clique em Verificar para consultar o GitHub.")
        tk.Label(quadro, textvariable=self.var_update, bg=CARD, fg=BRIGHT,
                 font=("Segoe UI", 10), justify="left",
                 wraplength=880).pack(anchor="w", padx=16, pady=12)

        botoes = tk.Frame(quadro, bg=CARD)
        botoes.pack(anchor="w", padx=16, pady=(0, 16))
        self.btn_verificar = self._botao(botoes, "Verificar atualizações", self.verificar_update, BLUE, 20)
        self.btn_verificar.pack(side="left")
        self.btn_atualizar = self._botao(botoes, "Atualizar agora", self.aplicar_update, SUCCESS, 16)
        self.btn_atualizar.pack(side="left", padx=10)
        self.btn_atualizar.configure(state="disabled")

        tk.Label(self.aba_atualizacao,
                 text="Repositório: " + atualizador.REPOSITORIO +
                      "\nA atualização só é instalada depois que todos os arquivos baixam e passam "
                      "na verificação de sintaxe. Planilhas e configurações não são tocadas.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 9), justify="left").pack(anchor="w", padx=6)

    # ------------------------------------------------ lista de ZIPs

    def adicionar_zips(self):
        inicial = self.cfg.get("ultima_pasta", str(Path.home()))
        arquivos = filedialog.askopenfilenames(
            title="Selecione os ZIPs de guias (pode marcar vários)",
            initialdir=inicial,
            filetypes=[("Arquivos ZIP", "*.zip"), ("Todos os arquivos", "*.*")])
        if not arquivos:
            return
        self.cfg["ultima_pasta"] = str(Path(arquivos[0]).parent)
        salvar_config(self.cfg)
        self._incluir([Path(a) for a in arquivos])

    def adicionar_pasta(self):
        inicial = self.cfg.get("ultima_pasta", str(Path.home()))
        pasta = filedialog.askdirectory(title="Pasta com os ZIPs de guias", initialdir=inicial)
        if not pasta:
            return
        self.cfg["ultima_pasta"] = pasta
        salvar_config(self.cfg)
        self._incluir(sorted(Path(pasta).glob("*.zip")))

    def _incluir(self, caminhos):
        existentes = set(str(z).lower() for z in self.zips)
        novos = 0
        for caminho in caminhos:
            if caminho.suffix.lower() != ".zip":
                continue
            if str(caminho).lower() in existentes:
                continue
            self.zips.append(caminho)
            existentes.add(str(caminho).lower())
            novos += 1
        self.zips.sort(key=lambda z: z.name.lower())
        self._pintar_lista()
        if novos == 0 and caminhos:
            self.var_status.set("Esses arquivos já estavam na lista.")

    def remover_selecionados(self):
        selecionados = self.tabela.selection()
        if not selecionados:
            self.var_status.set("Selecione na lista o que deseja remover.")
            return
        indices = sorted((int(self.tabela.item(i)["tags"][-1]) for i in selecionados), reverse=True)
        for indice in indices:
            if 0 <= indice < len(self.zips):
                self.zips.pop(indice)
        self._pintar_lista()

    def limpar_lista(self):
        self.zips = []
        self._pintar_lista()

    def _pintar_lista(self):
        import zipfile
        for item in self.tabela.get_children():
            self.tabela.delete(item)

        reconhecidas = set()
        for indice, arquivo in enumerate(self.zips):
            codigo, empresa, ok = core.empresa_do_zip(arquivo.name)
            try:
                with zipfile.ZipFile(arquivo) as z:
                    qtd = sum(1 for n in z.namelist() if n.lower().endswith(".pdf"))
            except Exception:
                qtd, ok = 0, False
            situacao = "OK" if ok else "empresa não reconhecida"
            if qtd == 0:
                situacao = "ZIP sem PDFs ou ilegível"
            self.tabela.insert("", "end", values=(arquivo.name, empresa, qtd, situacao),
                               tags=(("ok" if (ok and qtd) else "alerta"), str(indice)))
            if ok:
                reconhecidas.add(codigo)

        if not self.zips:
            self.var_info.set("Nenhum ZIP na lista.")
            self.var_status.set("Adicione os ZIPs de guias para começar.")
            return

        faltam = core.TOTAL_EMPRESAS - len(reconhecidas)
        texto = (str(len(self.zips)) + " ZIP(s) na lista  ·  " + str(len(reconhecidas)) +
                 " de " + str(core.TOTAL_EMPRESAS) + " empresas")
        if faltam > 0:
            texto = texto + "  ·  faltam " + str(faltam)
        self.var_info.set(texto)
        self.var_status.set("Pronto para processar.")

    # ------------------------------------------------ processamento

    def processar(self):
        if self.processando:
            return
        if not self.zips:
            messagebox.showwarning("Guias DIFAL", "Nenhum ZIP na lista.")
            return

        reconhecidas = set(core.empresa_do_zip(z.name)[0] for z in self.zips
                           if core.empresa_do_zip(z.name)[2])
        faltam = core.TOTAL_EMPRESAS - len(reconhecidas)
        if faltam > 0:
            pergunta = ("Este lote tem guias de " + str(len(reconhecidas)) + " das " +
                        str(core.TOTAL_EMPRESAS) + " empresas.\n\nProcessar assim mesmo?")
            if not messagebox.askyesno("Guias DIFAL", pergunta):
                return

        self.cfg["modo_rapido"] = bool(self.var_rapido.get())
        salvar_config(self.cfg)

        self.processando = True
        self.btn_processar.configure(state="disabled", text="PROCESSANDO...")
        self.btn_planilha.configure(state="disabled")
        self.btn_banco.configure(state="disabled")
        self.barra["value"] = 0
        threading.Thread(target=self._trabalhar, daemon=True).start()

    def _trabalhar(self):
        try:
            def progresso(feitos, total, texto):
                self.after(0, self._progresso, feitos, total, texto)

            resultado = core.processar_lote(
                [str(z) for z in self.zips],
                pasta_trabalho=PASTA / "tmp_lote",
                progresso=progresso,
                paralelo=self.var_rapido.get(),
                trabalhadores=max(2, (os.cpu_count() or 4) - 1),
            )

            saida = PASTA / "planilhas"
            saida.mkdir(exist_ok=True)
            carimbo = datetime.now().strftime("%Y-%m-%d_%H%M")
            planilha = saida / ("Guias_DIFAL_" + carimbo + ".xlsx")
            core.gerar_planilha(resultado, planilha)
            core.salvar_json(resultado, saida / ("lote_" + carimbo + ".json"))
            relatorio = core.texto_relatorio(resultado)
            (saida / ("relatorio_" + carimbo + ".txt")).write_text(relatorio, encoding="utf-8")

            self.after(0, self._concluir, resultado, planilha, relatorio, carimbo)
        except Exception:
            self.after(0, self._falhar, traceback.format_exc())

    def _progresso(self, feitos, total, texto):
        self.barra["maximum"] = max(1, total)
        self.barra["value"] = feitos
        self.var_status.set(texto + "  " + str(feitos) + "/" + str(total))

    def _concluir(self, resultado, planilha, relatorio, carimbo):
        self.processando = False
        self.resultado = resultado
        self.planilha = planilha
        self.carimbo = carimbo
        self.btn_processar.configure(state="normal", text="PROCESSAR LOTE")
        self.btn_planilha.configure(state="normal")
        self.btn_banco.configure(state="normal" if resultado.guias else "disabled")

        for item in self.tabela_resumo.get_children():
            self.tabela_resumo.delete(item)
        for empresa, dados in resultado.por_empresa().items():
            self.tabela_resumo.insert("", "end", values=(empresa, dados["guias"], brl(dados["total"])))
        self.tabela_resumo.insert("", "end",
                                  values=("TOTAL GERAL", len(resultado.guias), brl(resultado.total)),
                                  tags=("total",))

        self.texto.delete("1.0", "end")
        self.texto.insert("1.0", relatorio)

        pendencias = len(resultado.sem_codigo) + len(resultado.nao_autorizadas)
        self.var_status.set(str(len(resultado.guias)) + " guias  ·  " + brl(resultado.total) +
                            ("  ·  sem pendências" if pendencias == 0
                             else "  ·  " + str(pendencias) + " PDF(s) para conferir"))
        self.abas.select(self.aba_resumo)

        aviso = ""
        if pendencias:
            aviso = ("\n\nAtenção: " + str(pendencias) +
                     " PDF(s) não entraram na planilha. Veja a aba CONFERÊNCIA.")
        messagebox.showinfo("Guias DIFAL",
                            "Planilha gerada:\n" + planilha.name + "\n\n" +
                            str(len(resultado.guias)) + " guias  |  " + brl(resultado.total) + aviso)

    def _falhar(self, erro):
        self.processando = False
        self.btn_processar.configure(state="normal", text="PROCESSAR LOTE")
        self.var_status.set("Erro no processamento.")
        self.texto.delete("1.0", "end")
        self.texto.insert("1.0", erro)
        self.abas.select(self.aba_relatorio)
        messagebox.showerror("Guias DIFAL", "Falha no processamento. Veja a aba CONFERÊNCIA.")

    # ------------------------------------------------ banco

    def exportar_banco(self):
        if not self.resultado or not self.resultado.guias:
            messagebox.showwarning("Guias DIFAL", "Processe um lote antes de gerar os arquivos do banco.")
            return
        if not MODELO_BANCO.exists():
            messagebox.showerror("Guias DIFAL",
                                 "O modelo do banco não foi encontrado:\n" + MODELO_BANCO.name)
            return

        data = self.var_data.get().strip()
        try:
            datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Guias DIFAL", "Data de pagamento inválida. Use dd/mm/aaaa.")
            return

        self.cfg["por_empresa"] = bool(self.var_por_empresa.get())
        salvar_config(self.cfg)

        try:
            pasta = PASTA / "planilhas" / ("banco_c6_" + self.carimbo)
            gerados = core.gerar_arquivos_banco(
                self.resultado, MODELO_BANCO, pasta, data,
                por_empresa=self.var_por_empresa.get())
            relatorio = core.texto_relatorio_banco(gerados, data)
            (pasta / "conferencia.txt").write_text(relatorio, encoding="utf-8")
        except Exception:
            self.texto.delete("1.0", "end")
            self.texto.insert("1.0", traceback.format_exc())
            self.abas.select(self.aba_relatorio)
            messagebox.showerror("Guias DIFAL",
                                 "Falha ao gerar os arquivos do banco. Veja a aba CONFERÊNCIA.")
            return

        self.pasta_banco = pasta
        for item in self.tabela_banco.get_children():
            self.tabela_banco.delete(item)
        total_geral = 0.0
        total_guias = 0
        for caminho, empresa, quantidade, total in gerados:
            self.tabela_banco.insert("", "end",
                                     values=(Path(caminho).name, empresa, quantidade, brl(total)),
                                     tags=(str(caminho),))
            total_geral += total
            total_guias += quantidade
        self.tabela_banco.insert("", "end",
                                 values=(str(len(gerados)) + " arquivo(s)", "TOTAL GERAL",
                                         total_guias, brl(total_geral)),
                                 tags=("total",))

        self.abas.select(self.aba_banco)
        self.var_status.set(str(len(gerados)) + " arquivo(s) para o C6  ·  " + brl(total_geral))
        messagebox.showinfo("Guias DIFAL",
                            str(len(gerados)) + " arquivo(s) gerado(s) em:\n" + pasta.name +
                            "\n\nConfira o primeiro arquivo no Excel antes de subir no banco.")

    def _abrir_arquivo_banco(self, evento=None):
        selecionado = self.tabela_banco.selection()
        if not selecionado:
            return
        tags = self.tabela_banco.item(selecionado[0])["tags"]
        if tags and tags[0] != "total" and Path(tags[0]).exists():
            abrir_no_sistema(tags[0])

    def abrir_pasta_banco(self):
        if self.pasta_banco and Path(self.pasta_banco).exists():
            abrir_no_sistema(self.pasta_banco)
        else:
            messagebox.showinfo("Guias DIFAL", "Gere os arquivos do banco primeiro.")

    def abrir_planilha(self):
        if self.planilha and Path(self.planilha).exists():
            abrir_no_sistema(self.planilha)

    # ------------------------------------------------ atualizacoes

    def verificar_update(self):
        self.btn_verificar.configure(state="disabled", text="Verificando...")
        self.var_update.set("Consultando o GitHub...")
        threading.Thread(target=self._verificar_thread, daemon=True).start()

    def _verificar_thread(self):
        tem, remota, dados, erro = atualizador.consultar()
        self.after(0, self._verificar_fim, tem, remota, dados, erro)

    def _verificar_fim(self, tem, remota, dados, erro):
        self.btn_verificar.configure(state="normal", text="Verificar atualizações")
        if erro:
            self.var_update.set("Não foi possível consultar: " + str(erro))
            self.btn_atualizar.configure(state="disabled")
            return
        self._dados_update = dados
        if tem:
            notas = dados.get("notas", "")
            texto = "Versão " + str(remota) + " disponível."
            if notas:
                texto = texto + "\n\n" + notas
            self.var_update.set(texto)
            self.btn_atualizar.configure(state="normal")
        else:
            self.var_update.set("Você já está na versão mais recente (v" + VERSAO + ").")
            self.btn_atualizar.configure(state="disabled")

    def aplicar_update(self):
        if not messagebox.askyesno("Guias DIFAL",
                                   "Baixar e instalar a atualização agora?\n\n"
                                   "Suas planilhas e configurações não serão alteradas."):
            return
        self.btn_atualizar.configure(state="disabled", text="Atualizando...")
        threading.Thread(target=self._atualizar_thread, daemon=True).start()

    def _atualizar_thread(self):
        def progresso(texto):
            self.after(0, self.var_update.set, texto)

        ok, mensagem = atualizador.atualizar(getattr(self, "_dados_update", None), progresso)
        self.after(0, self._atualizar_fim, ok, mensagem)

    def _atualizar_fim(self, ok, mensagem):
        self.btn_atualizar.configure(text="Atualizar agora",
                                     state="disabled" if ok else "normal")
        self.var_update.set(mensagem)
        if ok:
            messagebox.showinfo("Guias DIFAL", mensagem)
        else:
            messagebox.showerror("Guias DIFAL", mensagem)


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    App().mainloop()
