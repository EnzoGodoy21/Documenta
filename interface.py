"""
interface.py — Interface gráfica para o Gerador de Documentos Word

Executa gera_word.py em background e exibe:
  - Configuração de pastas e opções
  - Log em tempo real durante a geração
  - Painel de resultados por tabela (ok / parcial / erro) com detalhes

Dependências:
  python-docx  (instalado automaticamente se ausente)
  tkinter      (já vem com o Python)

Uso:
  python interface.py
"""

# ─────────────────────────────────────────────────────────────
# VERIFICAÇÃO E INSTALAÇÃO AUTOMÁTICA DE DEPENDÊNCIAS
# Roda antes de qualquer outro import do projeto.
# ─────────────────────────────────────────────────────────────
import sys
import subprocess
import importlib

DEPENDENCIAS = [
    ("docx",       "python-docx"),   # import name, pip package name
]

def _verificar_dependencias():
    """
    Verifica se todas as dependências estão instaladas.
    Se alguma faltar, exibe uma janela de confirmação e instala via pip.
    Retorna True se tudo OK, False se o usuário cancelou ou houve erro.
    """
    faltando = []
    for import_name, pip_name in DEPENDENCIAS:
        try:
            importlib.import_module(import_name)
        except ImportError:
            faltando.append(pip_name)

    if not faltando:
        return True

    # tkinter já está disponível para mostrar o diálogo
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()  # janela invisível só pro diálogo funcionar

    pacotes = "\n  • ".join(faltando)
    resposta = messagebox.askyesno(
        title="Dependências não encontradas",
        message=(
            f"Os seguintes pacotes Python são necessários e não estão instalados:\n\n"
            f"  • {pacotes}\n\n"
            f"Deseja instalá-los agora?"
        ),
    )
    root.destroy()

    if not resposta:
        return False

    # Instala via pip
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", *faltando],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
        )
    except subprocess.CalledProcessError:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Erro na instalação",
            f"Não foi possível instalar os pacotes automaticamente.\n\n"
            f"Execute manualmente no terminal:\n\n"
            f"  pip install {' '.join(faltando)}"
        )
        root.destroy()
        return False

    return True


if not _verificar_dependencias():
    sys.exit(0)


# ─────────────────────────────────────────────────────────────
# IMPORTS PRINCIPAIS (após garantir dependências)
# ─────────────────────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import subprocess
import sys
from pathlib import Path

from gera_word import (
    resolver_template,
    descobrir_chaves,
    descobrir_tabelas,
    processar_tabela,
    gerar_relatorio,
)


# ─────────────────────────────────────────────────────────────
# CORES E ESTILOS
# ─────────────────────────────────────────────────────────────

COR_BG        = "#f7f7f8"
COR_PAINEL    = "#ffffff"
COR_BORDA     = "#e0e0e0"
COR_TITULO    = "#1a1a2e"
COR_SUBTITULO = "#555566"
COR_OK        = "#16a34a"
COR_PARCIAL   = "#d97706"
COR_ERRO      = "#dc2626"
COR_LOG_BG    = "#1e1e2e"
COR_LOG_FG    = "#cdd6f4"
COR_BTN       = "#2563eb"
COR_BTN_FG    = "#ffffff"
FONTE_MONO    = ("Courier New", 10)
FONTE_UI      = ("Segoe UI", 10)
FONTE_TITULO  = ("Segoe UI", 13, "bold")


# ─────────────────────────────────────────────────────────────
# LEITURA DE ARQUIVO DE TABELAS
# ─────────────────────────────────────────────────────────────

def _ler_tabelas_de_arquivo(path: str) -> tuple:
    """
    Lê um arquivo .txt ou .csv de tabelas.

    Retorna (tabelas, dados_por_tabela) onde:
      tabelas         — list[str] com os nomes das tabelas
      dados_por_tabela — dict[str, dict] com colunas extras por tabela
                         ex.: {"VENDAS": {"RESPONSAVEL": "João", "DOMINIO": "Comercial"}}

    .txt — uma tabela por linha; linhas em branco e # ignorados; sem dados extras.
    .csv — coluna 'nome_tabela' obrigatória (ou primeira coluna).
           Colunas extras viram chaves [TXT:*] por tabela.
           Delimitador , ou ; detectado automaticamente.
           Nomes de colunas são normalizados para maiúsculas sem espaços.
    """
    import csv as _csv
    from pathlib import Path as _Path

    path = str(path)
    ext = _Path(path).suffix.lower()
    tabelas = []
    dados = {}

    with open(path, encoding="utf-8-sig") as f:
        if ext == ".csv":
            amostra = f.read(2048)
            f.seek(0)
            try:
                dialeto = _csv.Sniffer().sniff(amostra, delimiters=",;")
            except _csv.Error:
                dialeto = _csv.excel
            reader = _csv.DictReader(f, dialect=dialeto)

            # Normaliza nomes de colunas removendo espaços e BOM
            fieldnames_raw = reader.fieldnames or []
            fieldnames = [c.strip() for c in fieldnames_raw]

            col_nome = next((c for c in fieldnames
                             if c.lower() == "nome_tabela"), None)
            if col_nome is None and fieldnames:
                col_nome = fieldnames[0]

            extras = [c for c in fieldnames if c != col_nome]

            for row in reader:
                # Normaliza chaves da row
                row_norm = {k.strip(): v.strip() for k, v in row.items() if k}
                nome = row_norm.get(col_nome, "").strip()
                if not nome:
                    continue
                tabelas.append(nome)
                if extras:
                    dados[nome] = {c.upper(): row_norm.get(c, "") for c in extras}
        else:
            for linha in f:
                val = linha.strip()
                if val and not val.startswith("#"):
                    tabelas.append(val)

    return tabelas, dados


# ─────────────────────────────────────────────────────────────
# COMPONENTES
# ─────────────────────────────────────────────────────────────

def campo_pasta(parent, label, var, tipo="dir", row=0):
    """Cria uma linha de campo com label, entry e botão de browse."""
    ttk.Label(parent, text=label, font=FONTE_UI).grid(
        row=row, column=0, sticky="w", pady=4, padx=(0, 8))
    entry = ttk.Entry(parent, textvariable=var, width=52, font=FONTE_UI)
    entry.grid(row=row, column=1, sticky="ew", pady=4)

    def browse():
        if tipo == "file":
            path = filedialog.askopenfilename(
                title="Selecionar template",
                filetypes=[("Word Document", "*.docx")])
        else:
            path = filedialog.askdirectory(title=f"Selecionar pasta — {label}")
        if path:
            var.set(path)

    ttk.Button(parent, text="…", width=3, command=browse).grid(
        row=row, column=2, padx=(6, 0), pady=4)
    return entry


# ─────────────────────────────────────────────────────────────
# JANELA PRINCIPAL
# ─────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Gerador de Documentos Word")
        self.geometry("960x720")
        self.minsize(800, 600)
        self.configure(bg=COR_BG)
        self.resizable(True, True)

        self._resultados = []
        self._running = False

        self._build_ui()
        self._aplicar_estilo()

    # ── Estilos ttk ──────────────────────────────────────────

    def _aplicar_estilo(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame",       background=COR_BG)
        style.configure("Card.TFrame",  background=COR_PAINEL, relief="flat")
        style.configure("TLabel",       background=COR_BG,    font=FONTE_UI)
        style.configure("Card.TLabel",  background=COR_PAINEL, font=FONTE_UI)
        style.configure("TEntry",       font=FONTE_UI)
        style.configure("TCheckbutton", background=COR_PAINEL, font=FONTE_UI)
        style.configure("TSpinbox",     font=FONTE_UI)
        style.configure("TLabelframe",  background=COR_PAINEL, font=FONTE_UI)
        style.configure("TLabelframe.Label", background=COR_PAINEL, font=("Segoe UI", 10, "bold"))
        style.configure("TNotebook",    background=COR_BG)
        style.configure("TNotebook.Tab", font=FONTE_UI, padding=[12, 5])

        # Treeview
        style.configure("Resultados.Treeview",
                         font=FONTE_UI, rowheight=26,
                         background=COR_PAINEL, fieldbackground=COR_PAINEL)
        style.configure("Resultados.Treeview.Heading",
                         font=("Segoe UI", 10, "bold"))
        style.map("Resultados.Treeview",
                  background=[("selected", "#dbeafe")])

        # Botão principal
        style.configure("Run.TButton",
                         background=COR_BTN, foreground=COR_BTN_FG,
                         font=("Segoe UI", 11, "bold"), padding=[16, 8])
        style.map("Run.TButton",
                  background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")])

    # ── Layout principal ──────────────────────────────────────

    def _build_ui(self):
        # Cabeçalho
        header = tk.Frame(self, bg="#1a1a2e", pady=12)
        header.pack(fill="x")
        tk.Label(header, text="Gerador de Documentos Word",
                 font=("Segoe UI", 15, "bold"),
                 bg="#1a1a2e", fg="white").pack(side="left", padx=20)
        tk.Label(header, text="por tabela a partir de template",
                 font=("Segoe UI", 10),
                 bg="#1a1a2e", fg="#8899cc").pack(side="left", padx=4)

        tk.Button(header, text="  ?  Ajuda  ", font=("Segoe UI", 10),
                  bg="#2d3561", fg="white", relief="flat", cursor="hand2",
                  activebackground="#3d4571", activeforeground="white",
                  command=self._abrir_ajuda).pack(side="right", padx=16, pady=6)

        # Container central
        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        # ── Painel de configuração ────────────────────────────
        config_frame = ttk.LabelFrame(body, text="Configuração", padding=12)
        config_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        config_frame.columnconfigure(1, weight=1)

        self.v_template = tk.StringVar(value="template/")
        self.v_prints   = tk.StringVar(value="prints/")
        self.v_output   = tk.StringVar(value="output/")
        self.v_logs     = tk.StringVar(value="logs/")

        campo_pasta(config_frame, "Template (.docx):", self.v_template, tipo="file", row=0)
        campo_pasta(config_frame, "Pasta de prints:", self.v_prints, row=1)
        campo_pasta(config_frame, "Pasta de saída:",  self.v_output, row=2)
        campo_pasta(config_frame, "Pasta de logs:",   self.v_logs,   row=3)

        # ── Painel de opções + botão ──────────────────────────
        opts_row = ttk.Frame(body)
        opts_row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        opts_row.columnconfigure(1, weight=1)

        opts_frame = ttk.LabelFrame(opts_row, text="Opções", padding=10)
        opts_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.v_force        = tk.BooleanVar()
        self.v_workers      = tk.IntVar(value=4)
        self.v_limite       = tk.IntVar(value=0)   # 0 = sem limite
        self._tabelas_lista = []   # list[str] carregada do arquivo
        self._tabelas_dados = {}   # dict[str, dict] colunas extras do CSV

        ttk.Checkbutton(opts_frame, text="Forçar reprocessamento (--force)",
                        variable=self.v_force).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(opts_frame, text="Workers:", background=COR_PAINEL).grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Spinbox(opts_frame, from_=1, to=32, textvariable=self.v_workers,
                    width=5).grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))

        ttk.Label(opts_frame, text="Limite (prints):", background=COR_PAINEL).grid(
            row=1, column=2, sticky="w", padx=(16, 0), pady=(6, 0))
        ttk.Spinbox(opts_frame, from_=0, to=99999, textvariable=self.v_limite,
                    width=7).grid(row=1, column=3, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(opts_frame, text="(0 = todas)", background=COR_PAINEL,
                  foreground=COR_SUBTITULO, font=("Segoe UI", 8)).grid(
            row=1, column=4, sticky="w", pady=(6, 0))

        # ── Seleção de tabelas por arquivo ────────────────────
        ttk.Label(opts_frame, text="Filtrar tabelas por arquivo (.txt / .csv):",
                  background=COR_PAINEL).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 2))

        arquivo_row = ttk.Frame(opts_frame, style="Card.TFrame")
        arquivo_row.grid(row=3, column=0, columnspan=3, sticky="ew")

        self.v_arquivo_tabelas = tk.StringVar(value="")
        self._entry_arquivo = ttk.Entry(arquivo_row, textvariable=self.v_arquivo_tabelas,
                                        width=38, state="readonly", font=FONTE_UI)
        self._entry_arquivo.pack(side="left", padx=(0, 6))

        ttk.Button(arquivo_row, text="📂 Selecionar", command=self._carregar_arquivo_tabelas).pack(side="left")
        ttk.Button(arquivo_row, text="✕ Limpar",      command=self._limpar_arquivo_tabelas).pack(side="left", padx=(4, 0))

        self.lbl_tabelas_count = tk.Label(
            opts_frame, text="(sem filtro — processa todas)",
            font=("Segoe UI", 9), fg=COR_SUBTITULO, bg=COR_PAINEL)
        self.lbl_tabelas_count.grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Botão executar
        btn_frame = ttk.Frame(opts_row)
        btn_frame.grid(row=0, column=1, sticky="nsew")
        btn_frame.rowconfigure(0, weight=1)
        btn_frame.columnconfigure(0, weight=1)

        self.btn_run = ttk.Button(
            btn_frame, text="▶  Gerar Documentos",
            style="Run.TButton", command=self._executar)
        self.btn_run.grid(row=0, column=0, sticky="nsew", padx=(0, 0))

        # ── Progresso ─────────────────────────────────────────
        prog_frame = ttk.Frame(body)
        prog_frame.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        prog_frame.columnconfigure(0, weight=1)

        self.v_prog = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(prog_frame, variable=self.v_prog, maximum=100)
        self.progressbar.grid(row=0, column=0, sticky="ew")

        self.lbl_status = ttk.Label(prog_frame, text="Pronto.", foreground=COR_SUBTITULO)
        self.lbl_status.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # ── Notebook: Log + Resultados ────────────────────────
        self.notebook = ttk.Notebook(body)
        self.notebook.grid(row=3, column=0, sticky="nsew", pady=(4, 0))
        body.rowconfigure(3, weight=1)

        self._build_aba_log()
        self._build_aba_resultados()

    # ── Aba Log ───────────────────────────────────────────────

    def _build_aba_log(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Log  ")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            frame, state="disabled",
            bg=COR_LOG_BG, fg=COR_LOG_FG,
            font=FONTE_MONO, wrap="word",
            insertbackground=COR_LOG_FG,
            selectbackground="#44475a",
            relief="flat", padx=10, pady=8)

        scroll_log = ttk.Scrollbar(frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll_log.set)
        scroll_log.grid(row=0, column=1, sticky="ns")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Tags de cor no log
        self.log_text.tag_configure("ok",      foreground=COR_OK)
        self.log_text.tag_configure("parcial",  foreground=COR_PARCIAL)
        self.log_text.tag_configure("erro",     foreground=COR_ERRO)
        self.log_text.tag_configure("info",     foreground="#89b4fa")
        self.log_text.tag_configure("destaque", foreground="#f5c2e7")

    # ── Aba Resultados ────────────────────────────────────────

    def _build_aba_resultados(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Resultados  ")
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        # Barra de resumo
        self.frame_resumo = tk.Frame(frame, bg=COR_PAINEL, pady=8, padx=12)
        self.frame_resumo.grid(row=0, column=0, columnspan=2, sticky="ew")

        self.lbl_resumo_ok      = self._lbl_badge(self.frame_resumo, "OK: —",      COR_OK)
        self.lbl_resumo_parcial = self._lbl_badge(self.frame_resumo, "Parcial: —", COR_PARCIAL)
        self.lbl_resumo_erro    = self._lbl_badge(self.frame_resumo, "Erro: —",    COR_ERRO)

        self.lbl_resumo_ok.pack(side="left", padx=(0, 12))
        self.lbl_resumo_parcial.pack(side="left", padx=(0, 12))
        self.lbl_resumo_erro.pack(side="left")

        # Botão abrir output
        self.btn_abrir = ttk.Button(
            self.frame_resumo, text="📂 Abrir pasta output",
            command=self._abrir_output)
        self.btn_abrir.pack(side="right")

        # Treeview
        colunas = ("tabela", "status", "imgs_ok", "ausentes", "erro")
        self.tree = ttk.Treeview(
            frame, columns=colunas, show="headings",
            style="Resultados.Treeview", selectmode="browse")

        self.tree.heading("tabela",   text="Tabela")
        self.tree.heading("status",   text="Status")
        self.tree.heading("imgs_ok",  text="[IMG] inseridos")
        self.tree.heading("ausentes", text="Ausentes (IMG / TXT)")
        self.tree.heading("erro",     text="Erro")

        self.tree.column("tabela",   width=180, minwidth=100)
        self.tree.column("status",   width=80,  minwidth=70, anchor="center")
        self.tree.column("imgs_ok",  width=180, minwidth=80)
        self.tree.column("ausentes", width=220, minwidth=80)
        self.tree.column("erro",     width=260, minwidth=80)

        self.tree.tag_configure("ok",      foreground=COR_OK)
        self.tree.tag_configure("parcial", foreground=COR_PARCIAL)
        self.tree.tag_configure("erro",    foreground=COR_ERRO)

        scroll_tree = ttk.Scrollbar(frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_tree.set)
        scroll_tree.grid(row=1, column=1, sticky="ns")
        self.tree.grid(row=1, column=0, sticky="nsew")

        # Filtro de status
        filtro_frame = ttk.Frame(frame, padding=(8, 4))
        filtro_frame.grid(row=2, column=0, columnspan=2, sticky="ew")

        ttk.Label(filtro_frame, text="Filtrar:").pack(side="left", padx=(0, 6))
        self.v_filtro = tk.StringVar(value="todos")
        for valor, txt in [("todos", "Todos"), ("ok", "OK"), ("parcial", "Parcial"), ("erro", "Erro")]:
            ttk.Radiobutton(
                filtro_frame, text=txt,
                variable=self.v_filtro, value=valor,
                command=self._aplicar_filtro).pack(side="left", padx=4)

    def _lbl_badge(self, parent, texto, cor):
        return tk.Label(parent, text=texto,
                        font=("Segoe UI", 11, "bold"),
                        fg=cor, bg=COR_PAINEL)

    # ── Carregar arquivo de tabelas ───────────────────────────

    def _carregar_arquivo_tabelas(self):
        path = filedialog.askopenfilename(
            title="Selecionar arquivo de tabelas",
            filetypes=[
                ("Arquivos de texto", "*.txt *.csv"),
                ("Todos os arquivos", "*.*"),
            ]
        )
        if not path:
            return

        tabelas, dados = _ler_tabelas_de_arquivo(path)

        if not tabelas:
            messagebox.showwarning(
                "Arquivo vazio",
                "Nenhum nome de tabela encontrado no arquivo.\n\n"
                "Verifique se o arquivo tem um nome por linha (ou coluna 'nome_tabela' no CSV)."
            )
            return

        self._tabelas_lista = tabelas
        self._tabelas_dados = dados
        self.v_arquivo_tabelas.set(path)

        extras = len(dados.get(tabelas[0], {})) if dados else 0
        info = f"✔ {len(tabelas)} tabela(s)"
        if extras:
            info += f" · {extras} coluna(s) extra(s) como [TXT:*]"
        self.lbl_tabelas_count.config(text=info, fg=COR_OK)

    def _limpar_arquivo_tabelas(self):
        self._tabelas_lista = []
        self._tabelas_dados = {}
        self.v_arquivo_tabelas.set("")
        self.lbl_tabelas_count.config(
            text="(sem filtro — processa todas)",
            fg=COR_SUBTITULO)

    # ── Helpers de UI ─────────────────────────────────────────

    def _log(self, msg, tag=None):
        """Adiciona linha ao log (thread-safe)."""
        def _do():
            self.log_text.config(state="normal")
            if tag:
                self.log_text.insert("end", msg + "\n", tag)
            else:
                self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, _do)

    def _set_status(self, msg, cor=COR_SUBTITULO):
        self.after(0, lambda: self.lbl_status.config(text=msg, foreground=cor))

    def _set_progresso(self, pct):
        self.after(0, lambda: self.v_prog.set(pct))

    def _adicionar_resultado(self, r):
        def _do():
            partes_ausentes = []
            if r["prints_ausentes"]:
                partes_ausentes.append(f"IMG: {', '.join(r['prints_ausentes'])}")
            if r["textos_ausentes"]:
                partes_ausentes.append(f"TXT: {', '.join(r['textos_ausentes'])}")

            self.tree.insert("", "end",
                values=(
                    r["tabela"],
                    r["status"].upper(),
                    ", ".join(r["prints_ok"]) or "—",
                    " | ".join(partes_ausentes) or "—",
                    r["erro"] or "—",
                ),
                tags=(r["status"],))
            self._atualizar_resumo()
        self.after(0, _do)

    def _atualizar_resumo(self):
        ok      = sum(1 for r in self._resultados if r["status"] == "ok")
        parcial = sum(1 for r in self._resultados if r["status"] == "parcial")
        erro    = sum(1 for r in self._resultados if r["status"] == "erro")
        self.lbl_resumo_ok.config(text=f"OK: {ok}")
        self.lbl_resumo_parcial.config(text=f"Parcial: {parcial}")
        self.lbl_resumo_erro.config(text=f"Erro: {erro}")

    def _aplicar_filtro(self):
        filtro = self.v_filtro.get()
        for item in self.tree.get_children():
            tags = self.tree.item(item, "tags")
            if filtro == "todos" or filtro in tags:
                # Re-inserir na posição certa (tkinter não tem "show/hide")
                pass
        # Simples: limpar e re-inserir só os que batem
        for item in self.tree.get_children():
            self.tree.delete(item)
        filtro_atual = self.v_filtro.get()
        for r in self._resultados:
            if filtro_atual == "todos" or r["status"] == filtro_atual:
                partes_ausentes = []
                if r["prints_ausentes"]:
                    partes_ausentes.append(f"IMG: {', '.join(r['prints_ausentes'])}")
                if r["textos_ausentes"]:
                    partes_ausentes.append(f"TXT: {', '.join(r['textos_ausentes'])}")
                self.tree.insert("", "end",
                    values=(
                        r["tabela"],
                        r["status"].upper(),
                        ", ".join(r["prints_ok"]) or "—",
                        " | ".join(partes_ausentes) or "—",
                        r["erro"] or "—",
                    ),
                    tags=(r["status"],))

    def _abrir_ajuda(self):
        """Abre janela de ajuda com duas abas: Como usar e Como criar o template."""
        win = tk.Toplevel(self)
        win.title("Ajuda — Gerador de Documentos Word")
        win.geometry("720x560")
        win.resizable(True, True)
        win.configure(bg=COR_BG)
        win.grab_set()  # modal

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=12, pady=12)

        def aba_texto(titulo, conteudo):
            frame = ttk.Frame(nb)
            nb.add(frame, text=f"  {titulo}  ")
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)
            txt = tk.Text(frame, wrap="word", font=("Segoe UI", 10),
                          bg=COR_PAINEL, fg=COR_TITULO, relief="flat",
                          padx=16, pady=12, state="normal",
                          selectbackground="#dbeafe", cursor="arrow")
            scroll = ttk.Scrollbar(frame, command=txt.yview)
            txt.configure(yscrollcommand=scroll.set)
            scroll.grid(row=0, column=1, sticky="ns")
            txt.grid(row=0, column=0, sticky="nsew")

            # Tags de estilo
            txt.tag_configure("h1",    font=("Segoe UI", 12, "bold"), foreground=COR_TITULO,
                               spacing1=10, spacing3=4)
            txt.tag_configure("h2",    font=("Segoe UI", 10, "bold"), foreground="#2563eb",
                               spacing1=8, spacing3=2)
            txt.tag_configure("code",  font=("Courier New", 10), foreground="#7c3aed",
                               background="#f3f0ff")
            txt.tag_configure("ok",    foreground=COR_OK,      font=("Segoe UI", 10, "bold"))
            txt.tag_configure("warn",  foreground=COR_PARCIAL, font=("Segoe UI", 10, "bold"))
            txt.tag_configure("err",   foreground=COR_ERRO,    font=("Segoe UI", 10, "bold"))
            txt.tag_configure("muted", foreground=COR_SUBTITULO)
            txt.tag_configure("standby", foreground="#92400e", background="#fef3c7",
                               font=("Segoe UI", 9))

            for bloco in conteudo:
                tag, texto = bloco
                txt.insert("end", texto, tag)

            # Somente leitura mas selecionável — bloqueia edição, permite Ctrl+C / Ctrl+A
            def _bloquear_edicao(e):
                if e.state & 0x4 and e.keysym.lower() in ("c", "a"):
                    return  # permite Ctrl+C e Ctrl+A
                return "break"
            txt.bind("<Key>", _bloquear_edicao)
            txt.bind("<Button-2>", lambda e: "break")  # bloqueia colar com botão do meio

            return frame

        # ── Aba 1: Como usar ──────────────────────────────────
        aba_texto("Como usar", [
            ("h1",   "Como usar a interface\n"),
            ("h2",   "1. Configure as pastas\n"),
            ("",     "  Template (.docx)   Arquivo Word com os placeholders\n"),
            ("",     "  Pasta de prints    Onde ficam as imagens geradas\n"),
            ("",     "  Pasta de saída     Onde os .docx serão salvos\n"),
            ("",     "  Pasta de logs      Onde ficam os relatórios de execução\n\n"),
            ("h2",   "2. Arquivo de tabelas (opcional)\n"),
            ("",     "  Deixe em branco para processar todas as tabelas encontradas.\n"),
            ("",     "  Carregue um .txt ou .csv para filtrar e/ou passar dados [TXT:*].\n\n"),
            ("h2",   "3. Opções\n"),
            ("code", "  Forçar reprocessamento"),
            ("",     "  reprocessa mesmo que o .docx já exista em output/\n"),
            ("code", "  Workers"),
            ("",     "             número de documentos gerados em paralelo\n\n"),
            ("h2",   "4. Clique em ▶ Gerar Documentos\n"),
            ("",     "  Acompanhe o progresso na aba "),
            ("code", "Log"),
            ("",     ".\n"),
            ("",     "  Ao final, a aba "),
            ("code", "Resultados"),
            ("",     " mostra o status de cada tabela.\n\n"),
            ("h2",   "Status dos documentos\n"),
            ("ok",   "  OK       "),
            ("",     "  Todos os placeholders substituídos\n"),
            ("warn", "  PARCIAL  "),
            ("",     "  Gerado, mas algum print ou texto estava ausente\n"),
            ("",     "           (placeholder visível no .docx para identificação)\n"),
            ("err",  "  ERRO     "),
            ("",     "  Falha crítica — documento não foi gerado\n\n"),
            ("h2",   "Arquivo de tabelas (.txt)\n"),
            ("muted","  Uma tabela por linha. Linhas com # são ignoradas.\n\n"),
            ("code", "  VENDAS\n  CLIENTES\n  # comentário ignorado\n  PEDIDOS\n\n"),
            ("h2",   "Arquivo de tabelas (.csv) — filtro + dados [TXT:*]\n"),
            ("muted","  Coluna nome_tabela obrigatória. Colunas extras viram [TXT:*] por tabela.\n"),
            ("muted","  Delimitador , ou ; detectado automaticamente.\n\n"),
            ("code", "  nome_tabela ; RESPONSAVEL  ; DOMINIO\n"
                     "  VENDAS      ; João Silva   ; Comercial\n"
                     "  CLIENTES    ; Maria Santos ; CRM\n"),
        ])

        # ── Aba 2: Como criar o template ──────────────────────
        aba_texto("Criar template", [
            ("h1",   "Como criar o template\n"),
            ("",     "Crie um arquivo .docx no Word e use os placeholders abaixo\n"),
            ("",     "onde quiser inserir valores dinâmicos.\n\n"),
            ("h2",   "[TXT:*] — texto automático (sem configuração)\n"),
            ("code", "  [TXT:NOME_TABELA]  "),
            ("",     " nome da tabela                  ex.: VENDAS\n"),
            ("code", "  [TXT:DATA]         "),
            ("",     " data de geração                 ex.: 04/04/2026\n"),
            ("code", "  [TXT:DATA_HORA]    "),
            ("",     " data e hora                     ex.: 04/04/2026 14:30\n"),
            ("code", "  [TXT:ANO]          "),
            ("",     " ano                             ex.: 2026\n"),
            ("code", "  [TXT:MES]          "),
            ("",     " mês                             ex.: 04\n"),
            ("code", "  [TXT:DIA]          "),
            ("",     " dia                             ex.: 04\n\n"),
            ("h2",   "[TXT:*] — texto via CSV (colunas extras)\n"),
            ("muted","  Qualquer coluna do CSV de tabelas vira um placeholder:\n\n"),
            ("code", "  [TXT:RESPONSAVEL]  "),
            ("",     " coluna RESPONSAVEL do CSV\n"),
            ("code", "  [TXT:DOMINIO]      "),
            ("",     " coluna DOMINIO do CSV\n"),
            ("code", "  [TXT:DESCRICAO]    "),
            ("",     " coluna DESCRICAO do CSV\n\n"),
            ("h2",   "[IMG:*] — imagem\n"),
            ("code", "  [IMG:chave]        "),
            ("",     " prints/chave_TABELA.png (ou .jpg)\n\n"),
            ("h2",   "[LEG:*] — legenda de print\n"),
            ("standby", "  Em desenvolvimento — estilo próprio (fonte menor), vinculado ao [IMG:*]\n\n"),
            ("h2",   "Convenção de nome dos prints\n"),
            ("code", "  chave_NOMETABELA.png\n\n"),
            ("muted","  Exemplos:\n"),
            ("code", "  visao_geral_VENDAS.png\n"),
            ("code", "  distribuicao_CLIENTES.png\n\n"),
            ("h2",   "Exemplo de template\n"),
            ("code", "  Tabela: [TXT:NOME_TABELA]       Gerado em: [TXT:DATA]\n"),
            ("code", "  Responsável: [TXT:RESPONSAVEL]  Domínio: [TXT:DOMINIO]\n\n"),
            ("code", "  Visão Geral\n  [IMG:visao_geral]\n\n"),
            ("code", "  Distribuição\n  [IMG:distribuicao]\n\n"),
            ("h2",   "Dicas\n"),
            ("",     "  • Placeholders funcionam em parágrafos, títulos e células de tabela\n"),
            ("",     "  • Placeholder não substituído fica visível no .docx → status PARCIAL\n"),
            ("",     "  • Tabelas são descobertas automaticamente pelos prints — sem lista manual\n"),
        ])

        ttk.Button(win, text="Fechar", command=win.destroy).pack(pady=(0, 12))

    def _abrir_output(self):
        pasta = self.v_output.get()
        if not os.path.exists(pasta):
            messagebox.showinfo("Pasta não encontrada",
                                f"A pasta ainda não existe:\n{pasta}")
            return
        if sys.platform == "win32":
            os.startfile(pasta)
        elif sys.platform == "darwin":
            subprocess.run(["open", pasta])
        else:
            subprocess.run(["xdg-open", pasta])

    # ── Execução ──────────────────────────────────────────────

    def _executar(self):
        if self._running:
            return

        template_arg = self.v_template.get().strip()
        prints_path  = self.v_prints.get().strip()
        output_path  = self.v_output.get().strip()
        logs_path    = self.v_logs.get().strip()
        force        = self.v_force.get()
        workers      = self.v_workers.get()
        filtro_tab   = list(self._tabelas_lista) if self._tabelas_lista else None
        dados_tab    = dict(self._tabelas_dados)
        limite_prints = self.v_limite.get()

        # Reset UI
        self._resultados = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self._atualizar_resumo()
        self.v_prog.set(0)

        self._running = True
        self.btn_run.config(state="disabled")
        self.notebook.select(0)  # vai pro log

        threading.Thread(target=self._thread_execucao, daemon=True, kwargs=dict(
            template_arg=template_arg,
            prints_path=prints_path,
            output_path=output_path,
            logs_path=logs_path,
            force=force,
            workers=workers,
            filtro_tab=filtro_tab,
            dados_tab=dados_tab,
            limite_prints=limite_prints,
        )).start()

    def _thread_execucao(self, template_arg, prints_path, output_path,
                         logs_path, force, workers, filtro_tab, dados_tab,
                         limite_prints=0):
        try:
            # Resolver template
            try:
                template_path = resolver_template(template_arg)
            except FileNotFoundError as e:
                self._log(f"❌ {e}", "erro")
                self._set_status(f"Erro: {e}", COR_ERRO)
                return

            self._log(f"📄 Template : {template_path}", "info")

            # Auto-descoberta
            chaves  = descobrir_chaves(template_path)

            self._log(f"🖼️  [IMG] : {chaves['img'] if chaves['img'] else '(nenhuma)'}", "info")
            self._log(f"📝 [TXT] : {chaves['txt'] if chaves['txt'] else '(nenhuma)'}", "info")

            if filtro_tab:
                # Lista explícita fornecida — usa diretamente (não depende de prints)
                tabelas = list(filtro_tab)
                self._log(f"📂 Arquivo  : {len(tabelas)} tabela(s) carregada(s) do arquivo", "info")
            else:
                # Sem lista explícita — descobre via prints/
                tabelas = descobrir_tabelas(prints_path, chaves["img"])
                self._log(f"📊 Tabelas  : {len(tabelas)} encontrada(s) via prints/", "info")
                if limite_prints and limite_prints > 0:
                    tabelas = tabelas[:limite_prints]
                    self._log(f"🔢 Limite   : {len(tabelas)} tabela(s) (limite aplicado)", "info")

            if not force:
                pendentes = [t for t in tabelas
                             if not (Path(output_path) / f"{t}.docx").exists()]
                ignoradas = len(tabelas) - len(pendentes)
                if ignoradas:
                    self._log(f"⏭️  {ignoradas} já processada(s) — use 'Forçar' para reprocessar", "info")
                tabelas = pendentes

            if not tabelas:
                self._log("✅ Nada a processar.", "ok")
                self._set_status("Nada a processar.", COR_OK)
                return

            from datetime import datetime as _dt
            from concurrent.futures import ThreadPoolExecutor, as_completed

            # Timestamp fixo para todo o lote — DATA/DATA_HORA consistentes entre docs
            ts = _dt.now()
            self._log(f"\n🚀 Iniciando {len(tabelas)} tabela(s) com {workers} worker(s)...", "destaque")
            self._log(f"📅 Timestamp : {ts.strftime('%d/%m/%Y %H:%M:%S')}\n", "info")
            self._set_status(f"Processando 0 / {len(tabelas)}...")

            total = len(tabelas)
            concluidos = [0]

            def on_done(resultado):
                self._resultados.append(resultado)
                self._adicionar_resultado(resultado)
                concluidos[0] += 1

                status = resultado["status"]
                tag = status
                icon = {"ok": "✅", "parcial": "⚠️ ", "erro": "❌"}.get(status, "?")
                linha = f"{icon} [{status.upper():7}] {resultado['tabela']}"
                if resultado["prints_ausentes"]:
                    linha += f"  — IMG ausentes: {', '.join(resultado['prints_ausentes'])}"
                if resultado["textos_ausentes"]:
                    linha += f"  — TXT ausentes: {', '.join(resultado['textos_ausentes'])}"
                if resultado["erro"]:
                    linha += f"  — {resultado['erro']}"
                self._log(linha, tag)

                pct = concluidos[0] / total * 100
                self._set_progresso(pct)
                self._set_status(f"Processando {concluidos[0]} / {total}...")

            with ThreadPoolExecutor(max_workers=workers) as executor:
                futuros = {
                    executor.submit(processar_tabela, t, template_path, prints_path,
                                    output_path, chaves, dados_tab.get(t), on_done, ts): t
                    for t in tabelas
                }
                for f in as_completed(futuros):
                    pass

            caminho_rel = gerar_relatorio(self._resultados, logs_path)

            ok      = sum(1 for r in self._resultados if r["status"] == "ok")
            parcial = sum(1 for r in self._resultados if r["status"] == "parcial")
            erro    = sum(1 for r in self._resultados if r["status"] == "erro")

            self._log(f"\n{'─' * 48}", "info")
            self._log(f"✅ OK: {ok}   ⚠️  Parcial: {parcial}   ❌ Erro: {erro}", "destaque")
            self._log(f"📋 Relatório salvo em: {caminho_rel}", "info")

            msg_status = f"Concluído — OK: {ok} | Parcial: {parcial} | Erro: {erro}"
            cor_status = COR_OK if erro == 0 and parcial == 0 else (COR_ERRO if erro > 0 else COR_PARCIAL)
            self._set_status(msg_status, cor_status)
            self._set_progresso(100)

            # Vai pra aba de resultados automaticamente
            self.after(800, lambda: self.notebook.select(1))

        except Exception as e:
            self._log(f"❌ Erro fatal: {e}", "erro")
            self._set_status(f"Erro fatal: {e}", COR_ERRO)

        finally:
            self._running = False
            self.after(0, lambda: self.btn_run.config(state="normal"))


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
