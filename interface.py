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
# PALETA DE CORES E TIPOGRAFIA  (refinada — visual mais moderno)
# ─────────────────────────────────────────────────────────────

# Superfícies
COR_BG          = "#f1f5f9"   # fundo geral (slate-100)
COR_PAINEL      = "#ffffff"   # cards
COR_PAINEL_ALT  = "#f8fafc"   # variação suave para destacar áreas internas
COR_BORDA       = "#e2e8f0"   # borda card (slate-200)
COR_BORDA_FORTE = "#cbd5e1"   # borda em foco / hover

# Texto
COR_TITULO     = "#0f172a"    # slate-900
COR_TEXTO      = "#1e293b"    # slate-800
COR_SUBTITULO  = "#64748b"    # slate-500
COR_MUTED      = "#94a3b8"    # slate-400

# Estados
COR_OK         = "#10b981"    # emerald-500
COR_OK_BG      = "#d1fae5"    # emerald-100
COR_PARCIAL    = "#f59e0b"    # amber-500
COR_PARCIAL_BG = "#fef3c7"    # amber-100
COR_ERRO       = "#ef4444"    # red-500
COR_ERRO_BG    = "#fee2e2"    # red-100
COR_INFO       = "#3b82f6"    # blue-500
COR_INFO_BG    = "#dbeafe"    # blue-100

# Botão primário (CTA)
COR_BTN        = "#2563eb"    # blue-600
COR_BTN_HOVER  = "#1d4ed8"    # blue-700
COR_BTN_FG     = "#ffffff"
COR_BTN_OFF    = "#94a3b8"    # quando disabled

# Header (sidebar topo)
COR_HEADER     = "#0f172a"    # slate-900
COR_HEADER_FG  = "#ffffff"
COR_HEADER_SUB = "#94a3b8"    # slate-400
COR_HEADER_BTN = "#1e293b"    # slate-800
COR_HEADER_BTN_HOVER = "#334155"  # slate-700

# Console (log)
COR_LOG_BG     = "#0b1020"
COR_LOG_FG     = "#e2e8f0"

# Fontes
FONTE_MONO    = ("Cascadia Mono", 10) if sys.platform == "win32" else ("Menlo", 10)
try:
    # Cascadia pode não estar instalada — fallback automático tratado pelo tk
    pass
except Exception:
    FONTE_MONO = ("Courier New", 10)

FONTE_UI       = ("Segoe UI", 10)
FONTE_UI_BOLD  = ("Segoe UI", 10, "bold")
FONTE_SMALL    = ("Segoe UI", 9)
FONTE_TITULO   = ("Segoe UI", 16, "bold")
FONTE_SECAO    = ("Segoe UI", 11, "bold")
FONTE_BADGE    = ("Segoe UI", 10, "bold")


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
# COMPONENTES VISUAIS
# ─────────────────────────────────────────────────────────────

class Card(tk.Frame):
    """Painel branco com borda sutil — simula 'card' de UI moderna."""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=COR_PAINEL,
                         highlightbackground=COR_BORDA,
                         highlightthickness=1, bd=0, **kw)


class HoverButton(tk.Button):
    """Botão tk plano com efeito hover."""
    def __init__(self, parent, hover_bg=None, hover_fg=None, **kw):
        super().__init__(parent, **kw)
        self._bg_default = kw.get("bg", self.cget("bg"))
        self._fg_default = kw.get("fg", self.cget("fg"))
        self._hover_bg = hover_bg or self._bg_default
        self._hover_fg = hover_fg or self._fg_default
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _):
        if self["state"] != "disabled":
            self.config(bg=self._hover_bg, fg=self._hover_fg)

    def _on_leave(self, _):
        if self["state"] != "disabled":
            self.config(bg=self._bg_default, fg=self._fg_default)


class CampoPasta(tk.Frame):
    """
    Linha de configuração de pasta/arquivo:
      [ícone status] Label
      [Entry........................] [Procurar]
    Mostra ✔ verde se o caminho existe, ⚠ amarelo se não existe.
    """

    def __init__(self, parent, label, var, tipo="dir", placeholder=""):
        super().__init__(parent, bg=COR_PAINEL)
        self.var = var
        self.tipo = tipo

        # Linha superior: ícone + label
        cabecalho = tk.Frame(self, bg=COR_PAINEL)
        cabecalho.pack(fill="x")

        self.lbl_status = tk.Label(cabecalho, text="○", font=FONTE_UI_BOLD,
                                    fg=COR_MUTED, bg=COR_PAINEL, width=2)
        self.lbl_status.pack(side="left")

        tk.Label(cabecalho, text=label, font=FONTE_UI_BOLD,
                 fg=COR_TITULO, bg=COR_PAINEL).pack(side="left")

        # Linha inferior: entry + botão
        linha = tk.Frame(self, bg=COR_PAINEL)
        linha.pack(fill="x", pady=(4, 0))

        self.entry = tk.Entry(linha, textvariable=var, font=FONTE_UI,
                              relief="flat", bd=0,
                              highlightthickness=1,
                              highlightbackground=COR_BORDA,
                              highlightcolor=COR_BTN,
                              bg=COR_PAINEL_ALT, fg=COR_TEXTO)
        self.entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(20, 6))

        HoverButton(linha, text="Procurar", font=FONTE_UI,
                    bg=COR_PAINEL_ALT, fg=COR_TITULO,
                    hover_bg=COR_BORDA, hover_fg=COR_TITULO,
                    relief="flat", bd=0, cursor="hand2",
                    padx=14, pady=6,
                    command=self._browse).pack(side="left")

        # Reavaliação do status quando o valor mudar
        var.trace_add("write", lambda *_: self._atualizar_status())
        self._atualizar_status()

    def _browse(self):
        if self.tipo == "file":
            path = filedialog.askopenfilename(
                title="Selecionar template",
                filetypes=[("Word Document", "*.docx")])
        else:
            path = filedialog.askdirectory(title="Selecionar pasta")
        if path:
            self.var.set(path)

    def _atualizar_status(self):
        v = self.var.get().strip()
        if not v:
            self.lbl_status.config(text="○", fg=COR_MUTED)
            return
        existe = Path(v).exists()
        if existe:
            self.lbl_status.config(text="✓", fg=COR_OK)
        else:
            self.lbl_status.config(text="!", fg=COR_PARCIAL)


# ─────────────────────────────────────────────────────────────
# JANELA PRINCIPAL
# ─────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Documenta — Gerador de Documentos Word")
        self.geometry("1040x760")
        self.minsize(880, 640)
        self.configure(bg=COR_BG)
        self.resizable(True, True)

        self._resultados = []
        self._running = False

        self._aplicar_estilo()
        self._build_ui()

    # ── Estilos ttk ──────────────────────────────────────────

    def _aplicar_estilo(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Bases
        style.configure(".",            background=COR_BG, foreground=COR_TEXTO,
                                         font=FONTE_UI)
        style.configure("TFrame",       background=COR_BG)
        style.configure("Card.TFrame",  background=COR_PAINEL)
        style.configure("TLabel",       background=COR_BG, font=FONTE_UI)
        style.configure("Card.TLabel",  background=COR_PAINEL, font=FONTE_UI)
        style.configure("TCheckbutton", background=COR_PAINEL, font=FONTE_UI,
                                         foreground=COR_TEXTO,
                                         focuscolor=COR_PAINEL)
        style.map("TCheckbutton",
                  background=[("active", COR_PAINEL)],
                  foreground=[("disabled", COR_MUTED)])

        # Spinbox e Entry mais limpos
        style.configure("TSpinbox", font=FONTE_UI, fieldbackground=COR_PAINEL_ALT,
                                     background=COR_PAINEL_ALT,
                                     bordercolor=COR_BORDA, relief="flat",
                                     arrowsize=14)
        style.configure("TEntry",   font=FONTE_UI, fieldbackground=COR_PAINEL_ALT,
                                     bordercolor=COR_BORDA, relief="flat")

        # Radiobutton flat
        style.configure("TRadiobutton", background=COR_BG, font=FONTE_UI,
                                         foreground=COR_TEXTO)
        style.map("TRadiobutton",
                  background=[("active", COR_BG)])

        # Notebook
        style.configure("TNotebook",      background=COR_BG, borderwidth=0,
                                           tabmargins=[0, 0, 0, 0])
        style.configure("TNotebook.Tab",  font=FONTE_UI, padding=[18, 8],
                                           background=COR_BG,
                                           foreground=COR_SUBTITULO,
                                           borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", COR_PAINEL)],
                  foreground=[("selected", COR_TITULO)],
                  expand=[("selected", [1, 1, 1, 0])])

        # Treeview
        style.configure("Resultados.Treeview",
                         font=FONTE_UI, rowheight=30,
                         background=COR_PAINEL,
                         fieldbackground=COR_PAINEL,
                         foreground=COR_TEXTO,
                         borderwidth=0)
        style.configure("Resultados.Treeview.Heading",
                         font=FONTE_UI_BOLD,
                         background=COR_PAINEL_ALT,
                         foreground=COR_TITULO,
                         relief="flat",
                         padding=[8, 8])
        style.map("Resultados.Treeview.Heading",
                  background=[("active", COR_BORDA)])
        style.map("Resultados.Treeview",
                  background=[("selected", COR_INFO_BG)],
                  foreground=[("selected", COR_TITULO)])

        # Progressbar
        style.configure("Run.Horizontal.TProgressbar",
                         background=COR_BTN,
                         troughcolor=COR_BORDA,
                         borderwidth=0,
                         thickness=8)

        # Botão padrão (ttk) — apenas para botões secundários
        style.configure("TButton",
                         font=FONTE_UI,
                         padding=[12, 6],
                         background=COR_PAINEL_ALT,
                         foreground=COR_TITULO,
                         borderwidth=0,
                         focusthickness=0)
        style.map("TButton",
                  background=[("active", COR_BORDA),
                              ("disabled", COR_PAINEL_ALT)],
                  foreground=[("disabled", COR_MUTED)])

        # LabelFrame — usado apenas em casos específicos
        style.configure("TLabelframe",  background=COR_PAINEL,
                                         bordercolor=COR_BORDA,
                                         borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label",
                         background=COR_PAINEL,
                         foreground=COR_TITULO,
                         font=FONTE_SECAO)

    # ── Layout principal ──────────────────────────────────────

    def _build_ui(self):
        # ╭─────────────────────────────────────────────────────╮
        # │  CABEÇALHO                                          │
        # ╰─────────────────────────────────────────────────────╯
        header = tk.Frame(self, bg=COR_HEADER, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Lado esquerdo: logo + título
        bloco_esq = tk.Frame(header, bg=COR_HEADER)
        bloco_esq.pack(side="left", padx=24, fill="y")

        tk.Label(bloco_esq, text="◆", font=("Segoe UI", 22, "bold"),
                 bg=COR_HEADER, fg=COR_INFO).pack(side="left", pady=12)

        bloco_titulo = tk.Frame(bloco_esq, bg=COR_HEADER)
        bloco_titulo.pack(side="left", padx=(10, 0), pady=12)

        tk.Label(bloco_titulo, text="Documenta",
                 font=FONTE_TITULO,
                 bg=COR_HEADER, fg=COR_HEADER_FG).pack(anchor="w")

        tk.Label(bloco_titulo,
                 text="Gerador de Documentos Word por tabela",
                 font=FONTE_SMALL,
                 bg=COR_HEADER, fg=COR_HEADER_SUB).pack(anchor="w")

        # Lado direito: ajuda
        HoverButton(header, text="?  Ajuda",
                    font=FONTE_UI_BOLD,
                    bg=COR_HEADER_BTN, fg=COR_HEADER_FG,
                    hover_bg=COR_HEADER_BTN_HOVER, hover_fg=COR_HEADER_FG,
                    relief="flat", bd=0, cursor="hand2",
                    padx=16, pady=8,
                    command=self._abrir_ajuda).pack(side="right", padx=20, pady=18)

        # ╭─────────────────────────────────────────────────────╮
        # │  CORPO                                              │
        # ╰─────────────────────────────────────────────────────╯
        body = tk.Frame(self, bg=COR_BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        # ── Card de configuração ──────────────────────────────
        card_config = Card(body)
        card_config.grid(row=0, column=0, sticky="ew", pady=(0, 14))

        cabec_cfg = tk.Frame(card_config, bg=COR_PAINEL)
        cabec_cfg.pack(fill="x", padx=18, pady=(14, 6))
        tk.Label(cabec_cfg, text="Configuração",
                 font=FONTE_SECAO, fg=COR_TITULO, bg=COR_PAINEL).pack(side="left")
        tk.Label(cabec_cfg, text="caminhos do template e das pastas de trabalho",
                 font=FONTE_SMALL, fg=COR_SUBTITULO, bg=COR_PAINEL).pack(side="left", padx=8)

        # Grid 2x2 de campos
        grid_pastas = tk.Frame(card_config, bg=COR_PAINEL)
        grid_pastas.pack(fill="x", padx=18, pady=(6, 16))
        grid_pastas.columnconfigure(0, weight=1, uniform="col")
        grid_pastas.columnconfigure(1, weight=1, uniform="col")

        self.v_template = tk.StringVar(value="template/")
        self.v_prints   = tk.StringVar(value="prints/")
        self.v_output   = tk.StringVar(value="output/")
        self.v_logs     = tk.StringVar(value="logs/")

        CampoPasta(grid_pastas, "Template (.docx)", self.v_template,
                   tipo="file").grid(row=0, column=0, sticky="ew", padx=(0, 10), pady=6)
        CampoPasta(grid_pastas, "Pasta de prints", self.v_prints
                   ).grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=6)
        CampoPasta(grid_pastas, "Pasta de saída", self.v_output
                   ).grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=6)
        CampoPasta(grid_pastas, "Pasta de logs", self.v_logs
                   ).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=6)

        # ── Card de opções + Card CTA ─────────────────────────
        linha_opc = tk.Frame(body, bg=COR_BG)
        linha_opc.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        linha_opc.columnconfigure(0, weight=3)
        linha_opc.columnconfigure(1, weight=1)

        # ── Opções ────────────────────────────────────────────
        card_opc = Card(linha_opc)
        card_opc.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        cabec_opc = tk.Frame(card_opc, bg=COR_PAINEL)
        cabec_opc.pack(fill="x", padx=18, pady=(14, 6))
        tk.Label(cabec_opc, text="Opções",
                 font=FONTE_SECAO, fg=COR_TITULO, bg=COR_PAINEL).pack(side="left")
        tk.Label(cabec_opc, text="controle de execução e filtros",
                 font=FONTE_SMALL, fg=COR_SUBTITULO, bg=COR_PAINEL).pack(side="left", padx=8)

        corpo_opc = tk.Frame(card_opc, bg=COR_PAINEL)
        corpo_opc.pack(fill="both", expand=True, padx=18, pady=(6, 16))

        # Vars
        self.v_force        = tk.BooleanVar()
        self.v_workers      = tk.IntVar(value=4)
        self.v_limite       = tk.StringVar(value="0")
        self.v_prefixo      = tk.StringVar(value="")
        self._tabelas_lista = []
        self._tabelas_dados = {}

        # Linha 1: Workers + Limite + Forçar
        row1 = tk.Frame(corpo_opc, bg=COR_PAINEL)
        row1.pack(fill="x")

        def _label(parent, txt, fg=COR_TEXTO):
            return tk.Label(parent, text=txt, font=FONTE_UI, fg=fg, bg=COR_PAINEL)

        _label(row1, "Workers:").pack(side="left")
        ttk.Spinbox(row1, from_=1, to=32, textvariable=self.v_workers,
                    width=5).pack(side="left", padx=(6, 18))

        _label(row1, "Limite:").pack(side="left")
        ttk.Spinbox(row1, from_=0, to=99999, textvariable=self.v_limite,
                    width=8).pack(side="left", padx=(6, 4))
        _label(row1, "(0 = todas)", fg=COR_MUTED).pack(side="left", padx=(0, 18))

        ttk.Checkbutton(row1, text="Forçar reprocessamento",
                        variable=self.v_force).pack(side="left")

        # Linha 2: Prefixo
        row2 = tk.Frame(corpo_opc, bg=COR_PAINEL)
        row2.pack(fill="x", pady=(12, 0))

        _label(row2, "Prefixo do arquivo:").pack(side="left")
        tk.Entry(row2, textvariable=self.v_prefixo, font=FONTE_UI, width=22,
                 relief="flat", bd=0,
                 highlightthickness=1,
                 highlightbackground=COR_BORDA,
                 highlightcolor=COR_BTN,
                 bg=COR_PAINEL_ALT, fg=COR_TEXTO
                 ).pack(side="left", padx=(8, 8), ipady=4)
        _label(row2, "ex.: DOC_ → DOC_VENDAS.docx", fg=COR_MUTED).pack(side="left")

        # Separador visual
        tk.Frame(corpo_opc, bg=COR_BORDA, height=1).pack(fill="x", pady=(14, 12))

        # Linha 3: Filtro por arquivo (.txt / .csv)
        row3 = tk.Frame(corpo_opc, bg=COR_PAINEL)
        row3.pack(fill="x")

        tk.Label(row3, text="Filtrar tabelas por arquivo (opcional)",
                 font=FONTE_UI_BOLD, fg=COR_TITULO, bg=COR_PAINEL).pack(anchor="w")
        tk.Label(row3, text=".txt (uma tabela por linha) ou .csv (com colunas extras [TXT:*])",
                 font=FONTE_SMALL, fg=COR_SUBTITULO, bg=COR_PAINEL).pack(anchor="w")

        row4 = tk.Frame(corpo_opc, bg=COR_PAINEL)
        row4.pack(fill="x", pady=(8, 0))

        self.v_arquivo_tabelas = tk.StringVar(value="")
        self._entry_arquivo = tk.Entry(row4, textvariable=self.v_arquivo_tabelas,
                                        font=FONTE_UI, state="readonly",
                                        relief="flat", bd=0,
                                        highlightthickness=1,
                                        highlightbackground=COR_BORDA,
                                        readonlybackground=COR_PAINEL_ALT,
                                        fg=COR_TEXTO)
        self._entry_arquivo.pack(side="left", fill="x", expand=True, ipady=5)

        HoverButton(row4, text="Selecionar",
                    font=FONTE_UI,
                    bg=COR_PAINEL_ALT, fg=COR_TITULO,
                    hover_bg=COR_BORDA, hover_fg=COR_TITULO,
                    relief="flat", bd=0, cursor="hand2",
                    padx=14, pady=5,
                    command=self._carregar_arquivo_tabelas
                    ).pack(side="left", padx=(8, 4))
        HoverButton(row4, text="Limpar",
                    font=FONTE_UI,
                    bg=COR_PAINEL_ALT, fg=COR_SUBTITULO,
                    hover_bg=COR_BORDA, hover_fg=COR_TITULO,
                    relief="flat", bd=0, cursor="hand2",
                    padx=14, pady=5,
                    command=self._limpar_arquivo_tabelas).pack(side="left")

        self.lbl_tabelas_count = tk.Label(
            corpo_opc, text="Sem filtro — processa todas as tabelas encontradas",
            font=FONTE_SMALL, fg=COR_SUBTITULO, bg=COR_PAINEL)
        self.lbl_tabelas_count.pack(anchor="w", pady=(6, 0))

        # ── CTA — botão principal ─────────────────────────────
        card_cta = Card(linha_opc)
        card_cta.grid(row=0, column=1, sticky="nsew")

        bloco_cta = tk.Frame(card_cta, bg=COR_PAINEL)
        bloco_cta.pack(expand=True, fill="both", padx=18, pady=18)

        tk.Label(bloco_cta, text="Pronto para gerar",
                 font=FONTE_UI_BOLD, fg=COR_TITULO, bg=COR_PAINEL).pack(pady=(8, 4))
        tk.Label(bloco_cta,
                 text="Confira as pastas\ne clique abaixo",
                 font=FONTE_SMALL, fg=COR_SUBTITULO, bg=COR_PAINEL,
                 justify="center").pack(pady=(0, 12))

        self.btn_run = HoverButton(
            bloco_cta, text="▶  Gerar Documentos",
            font=("Segoe UI", 12, "bold"),
            bg=COR_BTN, fg=COR_BTN_FG,
            hover_bg=COR_BTN_HOVER, hover_fg=COR_BTN_FG,
            relief="flat", bd=0, cursor="hand2",
            padx=20, pady=14,
            command=self._executar)
        self.btn_run.pack(fill="x", expand=False)

        # ── Card de Progresso ─────────────────────────────────
        card_prog = Card(body)
        card_prog.grid(row=2, column=0, sticky="ew", pady=(0, 14))

        prog_inner = tk.Frame(card_prog, bg=COR_PAINEL)
        prog_inner.pack(fill="x", padx=18, pady=12)

        topo_prog = tk.Frame(prog_inner, bg=COR_PAINEL)
        topo_prog.pack(fill="x")

        self.lbl_status = tk.Label(topo_prog, text="Pronto.",
                                    font=FONTE_UI_BOLD,
                                    fg=COR_TEXTO, bg=COR_PAINEL)
        self.lbl_status.pack(side="left")

        self.lbl_pct = tk.Label(topo_prog, text="",
                                 font=FONTE_SMALL,
                                 fg=COR_SUBTITULO, bg=COR_PAINEL)
        self.lbl_pct.pack(side="right")

        self.v_prog = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(
            prog_inner, variable=self.v_prog, maximum=100,
            style="Run.Horizontal.TProgressbar")
        self.progressbar.pack(fill="x", pady=(8, 0))

        # ╭─────────────────────────────────────────────────────╮
        # │  NOTEBOOK: Log + Resultados                         │
        # ╰─────────────────────────────────────────────────────╯
        notebook_wrap = Card(body)
        notebook_wrap.grid(row=3, column=0, sticky="nsew")
        body.rowconfigure(3, weight=1)

        self.notebook = ttk.Notebook(notebook_wrap)
        self.notebook.pack(fill="both", expand=True, padx=2, pady=2)

        self._build_aba_log()
        self._build_aba_resultados()

    # ── Aba Log ───────────────────────────────────────────────

    def _build_aba_log(self):
        frame = tk.Frame(self.notebook, bg=COR_PAINEL)
        self.notebook.add(frame, text="  Log  ")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        wrapper = tk.Frame(frame, bg=COR_LOG_BG)
        wrapper.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        wrapper.rowconfigure(0, weight=1)
        wrapper.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            wrapper, state="disabled",
            bg=COR_LOG_BG, fg=COR_LOG_FG,
            font=FONTE_MONO, wrap="word",
            insertbackground=COR_LOG_FG,
            selectbackground="#1e293b",
            relief="flat", bd=0,
            padx=14, pady=12)

        scroll_log = ttk.Scrollbar(wrapper, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll_log.set)
        scroll_log.grid(row=0, column=1, sticky="ns")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Tags de cor no log
        self.log_text.tag_configure("ok",      foreground="#86efac")
        self.log_text.tag_configure("parcial", foreground="#fcd34d")
        self.log_text.tag_configure("erro",    foreground="#fca5a5")
        self.log_text.tag_configure("info",    foreground="#93c5fd")
        self.log_text.tag_configure("destaque", foreground="#f0abfc",
                                     font=(FONTE_MONO[0], FONTE_MONO[1], "bold"))

        # Mensagem inicial
        self._log("Pronto para iniciar. Configure as pastas acima e clique em \"Gerar Documentos\".", "info")

    # ── Aba Resultados ────────────────────────────────────────

    def _build_aba_resultados(self):
        frame = tk.Frame(self.notebook, bg=COR_PAINEL)
        self.notebook.add(frame, text="  Resultados  ")
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        # Barra superior: badges + abrir output
        topo = tk.Frame(frame, bg=COR_PAINEL_ALT, height=56)
        topo.grid(row=0, column=0, columnspan=2, sticky="ew")
        topo.pack_propagate(False)

        bloco_badges = tk.Frame(topo, bg=COR_PAINEL_ALT)
        bloco_badges.pack(side="left", padx=14, pady=10)

        self.lbl_resumo_ok      = self._badge(bloco_badges, "OK",      "—", COR_OK,      COR_OK_BG)
        self.lbl_resumo_parcial = self._badge(bloco_badges, "Parcial", "—", COR_PARCIAL, COR_PARCIAL_BG)
        self.lbl_resumo_erro    = self._badge(bloco_badges, "Erro",    "—", COR_ERRO,    COR_ERRO_BG)

        self.lbl_resumo_ok.pack(side="left", padx=(0, 8))
        self.lbl_resumo_parcial.pack(side="left", padx=(0, 8))
        self.lbl_resumo_erro.pack(side="left")

        # Botão abrir output
        HoverButton(topo, text="📂 Abrir pasta output",
                    font=FONTE_UI,
                    bg=COR_PAINEL, fg=COR_TITULO,
                    hover_bg=COR_BORDA, hover_fg=COR_TITULO,
                    relief="flat", bd=0, cursor="hand2",
                    padx=14, pady=6,
                    command=self._abrir_output).pack(side="right", padx=14, pady=12)

        # Treeview de resultados
        colunas = ("tabela", "status", "imgs_ok", "ausentes", "erro")
        tree_wrap = tk.Frame(frame, bg=COR_PAINEL)
        tree_wrap.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=(8, 0))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_wrap, columns=colunas, show="headings",
            style="Resultados.Treeview", selectmode="browse")

        self.tree.heading("tabela",   text="Tabela")
        self.tree.heading("status",   text="Status")
        self.tree.heading("imgs_ok",  text="Imagens inseridas")
        self.tree.heading("ausentes", text="Ausentes (IMG / TXT)")
        self.tree.heading("erro",     text="Erro")

        self.tree.column("tabela",   width=200, minwidth=120)
        self.tree.column("status",   width=90,  minwidth=80, anchor="center")
        self.tree.column("imgs_ok",  width=200, minwidth=120)
        self.tree.column("ausentes", width=260, minwidth=120)
        self.tree.column("erro",     width=280, minwidth=120)

        self.tree.tag_configure("ok",      foreground=COR_OK)
        self.tree.tag_configure("parcial", foreground=COR_PARCIAL)
        self.tree.tag_configure("erro",    foreground=COR_ERRO)

        scroll_tree = ttk.Scrollbar(tree_wrap, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_tree.set)
        scroll_tree.grid(row=0, column=1, sticky="ns")
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Empty state
        self.lbl_empty = tk.Label(tree_wrap,
                                   text="Os resultados aparecerão aqui após a execução.",
                                   font=FONTE_UI, fg=COR_MUTED, bg=COR_PAINEL)
        self.lbl_empty.place(relx=0.5, rely=0.5, anchor="center")

        # Filtro de status
        filtro_frame = tk.Frame(frame, bg=COR_PAINEL)
        filtro_frame.grid(row=2, column=0, columnspan=2, sticky="ew",
                          padx=14, pady=(6, 10))

        tk.Label(filtro_frame, text="Filtrar:",
                 font=FONTE_UI, fg=COR_SUBTITULO, bg=COR_PAINEL
                 ).pack(side="left", padx=(0, 8))
        self.v_filtro = tk.StringVar(value="todos")
        for valor, txt in [("todos", "Todos"),
                            ("ok", "OK"),
                            ("parcial", "Parcial"),
                            ("erro", "Erro")]:
            ttk.Radiobutton(
                filtro_frame, text=txt,
                variable=self.v_filtro, value=valor,
                command=self._aplicar_filtro).pack(side="left", padx=4)

    def _badge(self, parent, label, valor, cor_fg, cor_bg):
        """Pill-style badge: 'OK · 12'."""
        fr = tk.Frame(parent, bg=cor_bg, padx=10, pady=4,
                       highlightbackground=cor_fg, highlightthickness=0)
        tk.Label(fr, text=label, font=FONTE_BADGE,
                 fg=cor_fg, bg=cor_bg).pack(side="left")
        tk.Label(fr, text=" · ", font=FONTE_BADGE,
                 fg=cor_fg, bg=cor_bg).pack(side="left")
        val_lbl = tk.Label(fr, text=valor, font=FONTE_BADGE,
                            fg=cor_fg, bg=cor_bg)
        val_lbl.pack(side="left")
        fr._valor_lbl = val_lbl  # referência para atualização
        return fr

    def _badge_set(self, badge, valor):
        badge._valor_lbl.config(text=str(valor))

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

        try:
            tabelas, dados = _ler_tabelas_de_arquivo(path)
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler arquivo",
                f"Não foi possível ler o arquivo selecionado:\n\n{e}"
            )
            return

        if not tabelas:
            messagebox.showwarning(
                "Arquivo vazio",
                "Nenhum nome de tabela encontrado no arquivo.\n\n"
                "Verifique se o arquivo tem um nome por linha "
                "(ou coluna 'nome_tabela' no CSV)."
            )
            return

        self._tabelas_lista = tabelas
        self._tabelas_dados = dados
        self.v_arquivo_tabelas.set(path)

        extras = len(dados.get(tabelas[0], {})) if dados else 0
        info = f"✓  {len(tabelas)} tabela(s) carregada(s)"
        if extras:
            info += f"  ·  {extras} coluna(s) extra(s) como [TXT:*]"
        self.lbl_tabelas_count.config(text=info, fg=COR_OK)

    def _limpar_arquivo_tabelas(self):
        self._tabelas_lista = []
        self._tabelas_dados = {}
        self.v_arquivo_tabelas.set("")
        self.lbl_tabelas_count.config(
            text="Sem filtro — processa todas as tabelas encontradas",
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

    def _set_status(self, msg, cor=COR_TEXTO):
        self.after(0, lambda: self.lbl_status.config(text=msg, fg=cor))

    def _set_progresso(self, pct):
        def _do():
            self.v_prog.set(pct)
            self.lbl_pct.config(text=f"{int(pct)}%")
        self.after(0, _do)

    def _adicionar_resultado(self, r):
        def _do():
            # Esconde o empty state
            if self.lbl_empty.winfo_ismapped():
                self.lbl_empty.place_forget()

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
        self._badge_set(self.lbl_resumo_ok, ok)
        self._badge_set(self.lbl_resumo_parcial, parcial)
        self._badge_set(self.lbl_resumo_erro, erro)

    def _aplicar_filtro(self):
        """Re-popula a Treeview de acordo com o filtro selecionado."""
        # Limpa
        for item in self.tree.get_children():
            self.tree.delete(item)

        filtro_atual = self.v_filtro.get()
        for r in self._resultados:
            if filtro_atual != "todos" and r["status"] != filtro_atual:
                continue
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
        win.title("Ajuda — Documenta")
        win.geometry("780x600")
        win.resizable(True, True)
        win.configure(bg=COR_BG)
        win.transient(self)
        win.grab_set()

        # Header da janela de ajuda
        h = tk.Frame(win, bg=COR_HEADER, height=56)
        h.pack(fill="x")
        h.pack_propagate(False)
        tk.Label(h, text="Central de Ajuda",
                 font=("Segoe UI", 14, "bold"),
                 bg=COR_HEADER, fg=COR_HEADER_FG).pack(side="left", padx=20, pady=14)

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=14, pady=14)

        def aba_texto(titulo, conteudo):
            frame = tk.Frame(nb, bg=COR_PAINEL)
            nb.add(frame, text=f"  {titulo}  ")
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)
            txt = tk.Text(frame, wrap="word", font=("Segoe UI", 10),
                          bg=COR_PAINEL, fg=COR_TEXTO, relief="flat",
                          padx=20, pady=16, state="normal",
                          selectbackground=COR_INFO_BG, cursor="arrow")
            scroll = ttk.Scrollbar(frame, command=txt.yview)
            txt.configure(yscrollcommand=scroll.set)
            scroll.grid(row=0, column=1, sticky="ns")
            txt.grid(row=0, column=0, sticky="nsew")

            txt.tag_configure("h1",    font=("Segoe UI", 14, "bold"),
                               foreground=COR_TITULO, spacing1=10, spacing3=8)
            txt.tag_configure("h2",    font=("Segoe UI", 11, "bold"),
                               foreground=COR_INFO, spacing1=12, spacing3=4)
            txt.tag_configure("code",  font=("Cascadia Mono", 10) if sys.platform == "win32" else ("Menlo", 10),
                               foreground="#7c3aed", background="#f3f0ff")
            txt.tag_configure("ok",    foreground=COR_OK,      font=FONTE_UI_BOLD)
            txt.tag_configure("warn",  foreground=COR_PARCIAL, font=FONTE_UI_BOLD)
            txt.tag_configure("err",   foreground=COR_ERRO,    font=FONTE_UI_BOLD)
            txt.tag_configure("muted", foreground=COR_SUBTITULO)
            txt.tag_configure("standby", foreground="#92400e", background="#fef3c7",
                               font=FONTE_SMALL)

            for bloco in conteudo:
                tag, texto = bloco
                txt.insert("end", texto, tag)

            def _bloquear_edicao(e):
                if e.state & 0x4 and e.keysym.lower() in ("c", "a"):
                    return
                return "break"
            txt.bind("<Key>", _bloquear_edicao)
            txt.bind("<Button-2>", lambda e: "break")
            return frame

        # ── Aba 1: Como usar ──────────────────────────────────
        aba_texto("Como usar", [
            ("h1",   "Como usar a interface\n"),
            ("h2",   "1. Configure as pastas\n"),
            ("",     "  Template (.docx)   Arquivo Word com os placeholders\n"),
            ("",     "  Pasta de prints    Onde ficam as imagens geradas\n"),
            ("",     "  Pasta de saída     Onde os .docx serão salvos\n"),
            ("",     "  Pasta de logs      Onde ficam os relatórios de execução\n\n"),
            ("muted","  Dica: o ícone verde ✓ indica que o caminho existe.\n\n"),
            ("h2",   "2. Arquivo de tabelas (opcional)\n"),
            ("",     "  Deixe em branco para processar todas as tabelas encontradas.\n"),
            ("",     "  Carregue um .txt ou .csv para filtrar e/ou passar dados [TXT:*].\n\n"),
            ("h2",   "3. Opções\n"),
            ("code", "  Forçar reprocessamento"),
            ("",     "  reprocessa mesmo que o .docx já exista em output/\n"),
            ("code", "  Workers"),
            ("",     "                 número de documentos gerados em paralelo\n"),
            ("code", "  Limite"),
            ("",     "                  processa apenas as N primeiras tabelas (0 = todas)\n\n"),
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

        # Botão fechar
        rodape = tk.Frame(win, bg=COR_BG)
        rodape.pack(fill="x", pady=(0, 12), padx=14)
        HoverButton(rodape, text="Fechar",
                    font=FONTE_UI_BOLD,
                    bg=COR_BTN, fg=COR_BTN_FG,
                    hover_bg=COR_BTN_HOVER, hover_fg=COR_BTN_FG,
                    relief="flat", bd=0, cursor="hand2",
                    padx=20, pady=8,
                    command=win.destroy).pack(side="right")

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
        filtro_tab    = list(self._tabelas_lista) if self._tabelas_lista else None
        dados_tab     = dict(self._tabelas_dados)
        try:
            limite_prints = int(self.v_limite.get())
        except (ValueError, TypeError):
            limite_prints = 0
        prefixo = self.v_prefixo.get().strip()

        # Reset UI
        self._resultados = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        # Mostra empty state se a aba existir
        if not self.lbl_empty.winfo_ismapped():
            self.lbl_empty.place(relx=0.5, rely=0.5, anchor="center")
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self._atualizar_resumo()
        self.v_prog.set(0)
        self.lbl_pct.config(text="0%")

        self._running = True
        self.btn_run.config(state="disabled", text="Processando...",
                             bg=COR_BTN_OFF)
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
            prefixo=prefixo,
        )).start()

    def _thread_execucao(self, template_arg, prints_path, output_path,
                         logs_path, force, workers, filtro_tab, dados_tab,
                         limite_prints=0, prefixo=""):
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
                if limite_prints > 0:
                    tabelas = tabelas[:limite_prints]
                    self._log(f"🔢 Limite   : {len(tabelas)} tabela(s) (limite aplicado)", "info")

            if prefixo:
                self._log(f"🏷️  Prefixo  : {prefixo}", "info")

            if not force:
                pendentes = [t for t in tabelas
                             if not (Path(output_path) / f"{prefixo}{t}.docx").exists()]
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
            self._set_status(f"Processando 0 / {len(tabelas)}...", COR_INFO)

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
                self._set_status(f"Processando {concluidos[0]} / {total}...", COR_INFO)

            with ThreadPoolExecutor(max_workers=workers) as executor:
                futuros = {
                    executor.submit(processar_tabela, t, template_path, prints_path,
                                    output_path, chaves, dados_tab.get(t), on_done, ts,
                                    prefixo): t
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
            self.after(0, lambda: self.btn_run.config(
                state="normal",
                text="▶  Gerar Documentos",
                bg=COR_BTN))


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
