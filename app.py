"""Gerador de lista de material - Cadari Engenharia.

Lê o CSV por instância exportado do Revit (tabela LISTA DE MATERIAL GERAL) e
gera a lista de material no padrão Cadari, com a folha de rosto oficial.
"""
import datetime as dt
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
PASTA_APP = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

from lista_material import __version__, config                                   # noqa: E402
from lista_material.gerador import DadosCapa, RelacaoProjetos, carregar_csv, gerar, nova_analise  # noqa: E402
from lista_material.leitura import ErroLeitura                                    # noqa: E402
from lista_material.modelo import COMUM, PRIVATIVA                                # noqa: E402
from lista_material.pacote import ErroModelo                                      # noqa: E402

AZUL, AZUL_2 = "#10475C", "#17546C"
DOURADO, DOURADO_2 = "#BCA985", "#AA9978"
CINZA, BRANCO = "#F1F2F2", "#FFFFFF"
TEXTO_DETALHE, BORDA = "#595959", "#C9CED2"
AZUL_CLARO, PENDENTE = "#DCE5E9", "#FFF2CC"
LISTRA = "#C3D1D6"          # faixa das listras para tópicos (#10475C a 25%) sobre branco
NOME_MODELO = "MODELO_-FOLHA_DE_ROSTO.xlsx"
NOMES_ABA = {COMUM: "Áreas de uso comum", PRIVATIVA: "Áreas privativas"}
ABA_POR_NOME = {v: k for k, v in NOMES_ABA.items()}


def localizar(*partes):
    """Procura primeiro ao lado do programa (permite trocar o modelo sem recompilar) e depois no pacote."""
    for raiz in (PASTA_APP, BASE):
        caminho = os.path.join(raiz, *partes)
        if os.path.exists(caminho):
            return caminho
    return os.path.join(PASTA_APP, *partes)


def pasta_config():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "Cadari", "GeradorListaMaterial")


def ler_preferencias():
    try:
        with open(os.path.join(pasta_config(), "preferencias.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def salvar_preferencias(dados):
    try:
        os.makedirs(pasta_config(), exist_ok=True)
        with open(os.path.join(pasta_config(), "preferencias.json"), "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def abrir_no_sistema(caminho):
    if sys.platform.startswith("win"):
        os.startfile(caminho)                       # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


def numero_br(n):
    return f"{n:,}".replace(",", ".")


class App(tk.Tk):
    def __init__(self):
        if sys.platform.startswith("win"):
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
        super().__init__()
        self.escala = max(1.0, self.winfo_fpixels("1i") / 96.0)
        self.title("Cadari  |  Gerador de lista de material")
        self.configure(bg=BRANCO)
        larg = min(int(1440 * self.escala), self.winfo_screenwidth() - 40)
        alt = min(int(900 * self.escala), self.winfo_screenheight() - 80)
        self.geometry(f"{larg}x{alt}")
        self.minsize(min(int(1240 * self.escala), larg), min(int(760 * self.escala), alt))
        self.prefs = ler_preferencias()
        self.fila = queue.Queue()
        self.leitura = None
        self.analise = None
        self.ultimo_resultado = None
        self.sugestao_doc = ""
        self.ocupado = False
        self.caminho_modelo = self.prefs.get("modelo") if os.path.exists(self.prefs.get("modelo", "")) \
            else localizar("modelos", NOME_MODELO)
        self.relacao = RelacaoProjetos(self.caminho_modelo)
        self._fontes()
        self._estilos()
        self._imagens()
        self._montar()
        self.protocol("WM_DELETE_WINDOW", self._sair)
        self.after(80, self._processar_fila)

    # ------------------------------------------------------------------ visual
    def _fontes(self):
        familias = set(tkfont.families(self))
        for nome in ("Franklin Gothic Book", "Libre Franklin", "Franklin Gothic Medium", "Arial"):
            if nome in familias:
                self.F = nome
                break
        else:
            self.F = tkfont.nametofont("TkDefaultFont").actual("family")
        for nome in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(nome).configure(family=self.F, size=10)
            except tk.TclError:
                pass

    def _estilos(self):
        F = self.F
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(F, 10), background=BRANCO, foreground=AZUL)
        st.configure("TFrame", background=BRANCO)
        st.configure("Cinza.TFrame", background=CINZA)
        st.configure("TLabel", background=BRANCO, foreground=AZUL, font=(F, 10))
        st.configure("Rotulo.TLabel", foreground=AZUL_2, font=(F, 9))
        st.configure("Dica.TLabel", foreground=AZUL_2, font=(F, 8))
        st.configure("Info.TLabel", foreground=AZUL, font=(F, 10))
        st.configure("Arquivo.TLabel", foreground=AZUL, font=(F, 11, "bold"))
        st.configure("Titulo.TLabel", foreground=AZUL, font=(F, 15, "bold italic"))
        st.configure("App.TLabel", foreground=AZUL, font=(F, 19, "bold italic"))
        st.configure("Sub.TLabel", foreground=AZUL_2, font=(F, 11))
        st.configure("Alerta.TLabel", background=PENDENTE, foreground=AZUL, font=(F, 10))
        st.configure("Ok.TLabel", foreground=AZUL_2, font=(F, 10))
        st.configure("Primario.TButton", background=AZUL, foreground=BRANCO, font=(F, 11, "bold"),
                     borderwidth=0, focusthickness=0, focuscolor=AZUL, padding=(18, 8), relief="flat")
        st.map("Primario.TButton",
               background=[("disabled", AZUL_2), ("pressed", AZUL_2), ("active", AZUL_2)],
               foreground=[("disabled", CINZA)])
        st.configure("Secundario.TButton", background=BRANCO, foreground=AZUL, font=(F, 11, "bold"),
                     bordercolor=AZUL, lightcolor=BRANCO, darkcolor=BRANCO, borderwidth=1,
                     focusthickness=0, focuscolor=BRANCO, padding=(14, 7), relief="solid")
        st.map("Secundario.TButton",
               background=[("disabled", BRANCO), ("pressed", CINZA), ("active", CINZA)],
               foreground=[("disabled", "#9AAAB2")],
               bordercolor=[("disabled", BORDA)])
        st.configure("Pequeno.TButton", background=BRANCO, foreground=AZUL, font=(F, 10, "bold"),
                     bordercolor=AZUL, lightcolor=BRANCO, darkcolor=BRANCO, borderwidth=1,
                     focusthickness=0, focuscolor=BRANCO, padding=(10, 4), relief="solid")
        st.map("Pequeno.TButton", background=[("active", CINZA), ("disabled", BRANCO)],
               foreground=[("disabled", "#9AAAB2")], bordercolor=[("disabled", BORDA)])
        campo = dict(fieldbackground=CINZA, foreground=AZUL, bordercolor=CINZA, lightcolor=CINZA,
                     darkcolor=CINZA, insertcolor=AZUL, padding=(7, 5))
        st.configure("TEntry", **campo)
        st.map("TEntry", bordercolor=[("focus", DOURADO)], lightcolor=[("focus", DOURADO)],
               darkcolor=[("focus", DOURADO)])
        st.configure("TCombobox", **campo, background=CINZA, arrowcolor=AZUL, arrowsize=13)
        st.map("TCombobox", fieldbackground=[("readonly", CINZA)], background=[("readonly", CINZA)],
               foreground=[("readonly", AZUL)], bordercolor=[("focus", DOURADO)],
               lightcolor=[("focus", DOURADO)], darkcolor=[("focus", DOURADO)],
               selectbackground=[("readonly", CINZA)], selectforeground=[("readonly", AZUL)])
        self.option_add("*TCombobox*Listbox.background", CINZA)
        self.option_add("*TCombobox*Listbox.foreground", AZUL)
        self.option_add("*TCombobox*Listbox.selectBackground", AZUL_2)
        self.option_add("*TCombobox*Listbox.font", (F, 10))
        st.configure("TSpinbox", **campo, background=CINZA, arrowcolor=AZUL, arrowsize=11)
        st.map("TSpinbox", bordercolor=[("focus", DOURADO)], lightcolor=[("focus", DOURADO)],
               darkcolor=[("focus", DOURADO)])
        st.configure("Treeview", background=BRANCO, fieldbackground=BRANCO, foreground=TEXTO_DETALHE,
                     rowheight=int(24 * self.escala), bordercolor=BORDA, lightcolor=BORDA, darkcolor=BORDA,
                     font=(F, 10))
        st.configure("Treeview.Heading", background=AZUL_CLARO, foreground=AZUL, font=(F, 10, "bold"),
                     relief="flat", bordercolor=BORDA, lightcolor=AZUL_CLARO, darkcolor=AZUL_CLARO, padding=(6, 4))
        st.map("Treeview.Heading", background=[("active", AZUL_CLARO)])
        st.map("Treeview", background=[("selected", AZUL_2)], foreground=[("selected", BRANCO)])
        st.configure("Cadari.Horizontal.TProgressbar", troughcolor=CINZA, background=DOURADO,
                     bordercolor=CINZA, lightcolor=DOURADO, darkcolor=DOURADO, thickness=10)
        st.configure("TCheckbutton", background=BRANCO, foreground=AZUL, font=(F, 10), focuscolor=BRANCO)
        st.map("TCheckbutton", background=[("active", BRANCO)],
               indicatorcolor=[("selected", AZUL), ("!selected", CINZA)],
               indicatorbackground=[("selected", AZUL), ("!selected", CINZA)])
        st.configure("Vertical.TScrollbar", background=CINZA, troughcolor=BRANCO, bordercolor=BRANCO,
                     arrowcolor=AZUL_2, lightcolor=CINZA, darkcolor=CINZA)
        st.configure("Horizontal.TScrollbar", background=CINZA, troughcolor=BRANCO, bordercolor=BRANCO,
                     arrowcolor=AZUL_2, lightcolor=CINZA, darkcolor=CINZA)

    def _img(self, base, alturas):
        alvo = alturas[0] * self.escala
        melhor = min(alturas, key=lambda h: abs(h - alvo))
        caminho = localizar("recursos", f"{base}_{melhor}.png")
        try:
            return tk.PhotoImage(file=caminho)
        except tk.TclError:
            return None

    def _imagens(self):
        self.img_marca = self._img("marca_cadari", (57, 86, 114))
        # títulos de seção: listra para tópicos (check dourado + faixa azul translúcida), pacote "Listras para tópicos"
        self.img_tit_inicio = self._img("titulo_inicio", (38, 57, 76))
        self.img_tit_faixa = self._img("titulo_faixa", (38, 57, 76))
        self.alt_titulo = self.img_tit_inicio.height() if self.img_tit_inicio else self.px(38)
        self.recuo = round(1.2 * self.alt_titulo)       # texto do título e corpo da seção começam depois do check
        # rodapé: assinatura C✓ dourada (Ativo 5 do pacote)
        self.img_icone = self._img("assinatura", (24, 36, 48))
        try:
            self.iconphoto(True, tk.PhotoImage(file=localizar("recursos", "icone.png")))
            if sys.platform.startswith("win"):
                self.iconbitmap(default=localizar("recursos", "icone.ico"))
        except tk.TclError:
            pass

    def px(self, v):
        return int(v * self.escala)

    # ------------------------------------------------------------------ layout
    def _montar(self):
        self._cabecalho()
        tk.Frame(self, bg=DOURADO, height=self.px(4)).pack(fill="x")
        self._rodape()
        self._barra_acoes()
        corpo = ttk.Frame(self, padding=(self.px(24), self.px(14), self.px(24), self.px(4)))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=1, uniform="col")
        corpo.columnconfigure(1, weight=1, uniform="col")
        corpo.rowconfigure(0, weight=1)
        esquerda = ttk.Frame(corpo)
        esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, self.px(28)))
        direita = ttk.Frame(corpo)
        direita.grid(row=0, column=1, sticky="nsew")
        self._secao_arquivo(esquerda)
        self._secao_capa(esquerda)
        self._secao_niveis(direita)
        self._secao_ambientes(direita)
        self._atualizar_estado()

    def _cabecalho(self):
        topo = tk.Frame(self, bg=BRANCO)
        topo.pack(fill="x")
        margem = self.px(24)
        if self.img_marca:
            tk.Label(topo, image=self.img_marca, bg=BRANCO, bd=0).pack(side="left", padx=(margem, margem),
                                                                        pady=(self.px(16), self.px(16)))
        tk.Frame(topo, bg=BORDA, width=1).pack(side="left", fill="y", pady=self.px(20))
        bloco = ttk.Frame(topo, padding=(self.px(22), 0, 0, 0))
        bloco.pack(side="left", fill="y")
        ttk.Label(bloco, text="GERADOR DE LISTA DE MATERIAL", style="App.TLabel").pack(anchor="w", pady=(self.px(20), 0))
        ttk.Label(bloco, text="Monta a lista no padrão Cadari a partir do CSV por instância exportado do Revit.",
                  style="Sub.TLabel").pack(anchor="w")
        ttk.Label(topo, text=f"versão {__version__}", style="Dica.TLabel").pack(side="right", anchor="s",
                                                                              padx=margem, pady=self.px(10))

    def _secao_titulo(self, pai, texto, pady=(0, 10)):
        """Título de seção com a listra para tópicos: check + faixa translúcida, título sobre a faixa."""
        linha = ttk.Frame(pai)
        linha.pack(fill="x", pady=(self.px(pady[0]), self.px(pady[1])))
        h = self.alt_titulo
        fonte = (self.F, 15, "bold italic")
        titulo = texto.upper()
        largura = max(round(5.04 * h), self.recuo + tkfont.Font(self, font=fonte).measure(titulo) + round(0.6 * h))
        cv = tk.Canvas(linha, width=largura, height=h, bg=BRANCO, highlightthickness=0, bd=0)
        cv.pack(side="left")
        if self.img_tit_inicio and self.img_tit_faixa:
            inicio = self.img_tit_inicio.width()
            faixa = self.img_tit_faixa.zoom(max(1, largura - inicio), 1)
            cv.imagens = (faixa,)                       # mantém a referência da imagem
            cv.create_image(inicio, 0, image=faixa, anchor="nw")
            cv.create_image(0, 0, image=self.img_tit_inicio, anchor="nw")
        else:
            cv.create_rectangle(0, 0, largura, h, fill=LISTRA, outline=LISTRA)
        cv.create_text(self.recuo, h // 2 + 1, text=titulo, anchor="w", font=fonte, fill=AZUL)
        return linha

    def _campo(self, pai, linha, coluna, rotulo, var, largura=None, colspan=1, combo=None, dica=None):
        ttk.Label(pai, text=rotulo, style="Rotulo.TLabel").grid(row=linha, column=coluna, columnspan=colspan,
                                                                 sticky="w", pady=(self.px(5), self.px(1)))
        if combo is not None:
            w = ttk.Combobox(pai, textvariable=var, values=combo, state="readonly", width=largura or 6)
        else:
            w = ttk.Entry(pai, textvariable=var, width=largura or 10)
        w.grid(row=linha + 1, column=coluna, columnspan=colspan, sticky="ew", padx=(0, self.px(12)))
        if dica:
            ttk.Label(pai, text=dica, style="Dica.TLabel").grid(row=linha + 2, column=coluna, columnspan=colspan,
                                                               sticky="w")
        return w

    def _secao_arquivo(self, pai):
        self._secao_titulo(pai, "Arquivo CSV")
        caixa = ttk.Frame(pai, padding=(self.recuo, 0, 0, 0))
        caixa.pack(fill="x")
        linha = ttk.Frame(caixa)
        linha.pack(fill="x")
        self.bt_csv = ttk.Button(linha, text="Selecionar CSV", style="Primario.TButton", command=self.selecionar_csv)
        self.bt_csv.pack(side="left")
        info = ttk.Frame(linha, padding=(self.px(14), 0, 0, 0))
        info.pack(side="left", fill="x", expand=True)
        self.lb_arquivo = ttk.Label(info, text="Nenhum arquivo selecionado", style="Arquivo.TLabel",
                                    wraplength=self.px(400))
        self.lb_arquivo.pack(anchor="w")
        self.lb_detalhe = ttk.Label(info, text="Exporte por instância a tabela LISTA DE MATERIAL GERAL do Revit.",
                                    style="Rotulo.TLabel", wraplength=self.px(400))
        self.lb_detalhe.pack(anchor="w")
        self.fr_status = ttk.Frame(caixa)
        self.fr_status.pack(fill="x", pady=(self.px(10), 0))
        self.lb_status_csv = ttk.Label(self.fr_status, text="", style="Ok.TLabel", wraplength=self.px(560))
        self.lb_status_csv.pack(anchor="w", fill="x")
        self.bt_problemas = ttk.Button(self.fr_status, text="Ver problemas", style="Pequeno.TButton",
                                       command=self.mostrar_problemas)
        modelo = ttk.Frame(caixa)
        modelo.pack(fill="x", pady=(self.px(8), 0))
        self.lb_modelo = ttk.Label(modelo, text="", style="Dica.TLabel")
        self.lb_modelo.pack(side="left")
        ttk.Button(modelo, text="Trocar modelo", style="Pequeno.TButton", command=self.trocar_modelo).pack(
            side="left", padx=(self.px(10), 0))
        self._mostrar_modelo()

    def _secao_capa(self, pai):
        self._secao_titulo(pai, "Folha de rosto", pady=(16, 4))
        g = ttk.Frame(pai, padding=(self.recuo, 0, 0, 0))
        g.pack(fill="x")
        for c in range(3):
            g.columnconfigure(c, weight=1, uniform="capa")
        p = self.prefs
        self.v_ca = tk.StringVar(value=p.get("ca", ""))
        self.v_doc = tk.StringVar(value="")
        self.v_sistema = tk.StringVar(value=p.get("sistema", "CLIMATIZAÇÃO"))
        self.v_etapa = tk.StringVar(value=p.get("etapa", "PROJETO EXECUTIVO"))
        self.v_documento = tk.StringVar(value=p.get("documento", "LISTA DE MATERIAIS"))
        self.v_prefixo = tk.StringVar(value=p.get("prefixo", "CLI"))
        self.v_rev = tk.StringVar(value="0")
        self.v_te = tk.StringVar(value=p.get("te", "A"))
        self.v_data = tk.StringVar(value=dt.date.today().strftime("%d/%m/%Y"))
        self.v_desc_rev = tk.StringVar(value="EMISSÃO INICIAL")
        self.v_elab = tk.StringVar(value=p.get("elab", ""))
        self.v_verif = tk.StringVar(value=p.get("verif", ""))
        self.v_resp = tk.StringVar(value=p.get("resp", ""))
        self.v_desenhos = tk.StringVar(value="")
        self._campo(g, 0, 0, "Código da obra (CA)", self.v_ca)
        self._campo(g, 0, 1, "Nº do documento (nome do arquivo)", self.v_doc, colspan=2)
        self.lb_obra = ttk.Label(g, text="", style="Dica.TLabel", wraplength=self.px(520))
        self.lb_obra.grid(row=2, column=0, columnspan=3, sticky="w")
        self._campo(g, 3, 0, "Sistema", self.v_sistema)
        self._campo(g, 3, 1, "Etapa", self.v_etapa)
        self._campo(g, 3, 2, "Documento", self.v_documento)
        self._campo(g, 6, 0, "Revisão", self.v_rev)
        self._campo(g, 6, 1, "Tipo de emissão", self.v_te, combo=config.TIPOS_EMISSAO)
        self._campo(g, 6, 2, "Data (dd/mm/aaaa)", self.v_data)
        ttk.Label(g, text="Tipo de emissão: A preliminar, B para aprovação, C para conhecimento, D para cotação, "
                          "E cancelado.", style="Dica.TLabel", wraplength=self.px(580)).grid(
            row=8, column=0, columnspan=3, sticky="w")
        self._campo(g, 9, 0, "Descrição da revisão", self.v_desc_rev, colspan=3)
        self._campo(g, 12, 0, "Elaborado", self.v_elab)
        self._campo(g, 12, 1, "Verificado", self.v_verif)
        self._campo(g, 12, 2, "Responsável", self.v_resp)
        self._campo(g, 15, 0, "Desenhos de referência (nota 7)", self.v_desenhos, colspan=2,
                    dica="Ex.: A056-0001 @ 0011")
        self._campo(g, 15, 2, "Prefixo das abas", self.v_prefixo, dica="Ex.: CLI")
        self.v_ca.trace_add("write", lambda *_: self._atualizar_obra())
        self._atualizar_obra()

    def _tabela(self, pai, colunas, altura):
        caixa = ttk.Frame(pai)
        caixa.pack(fill="both", expand=True)
        tv = ttk.Treeview(caixa, columns=[c[0] for c in colunas], show="headings", height=altura,
                          selectmode="browse")
        for chave, titulo, larg, ancora in colunas:
            tv.heading(chave, text=titulo, anchor=ancora)
            tv.column(chave, width=self.px(larg), minwidth=self.px(36), anchor=ancora,
                      stretch=chave in ("nivel", "cab", "origem", "coluna", "desc"))
        sb = ttk.Scrollbar(caixa, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tv.pack(side="left", fill="both", expand=True)
        tv.tag_configure("par", background=BRANCO)
        tv.tag_configure("impar", background="#F7F8F8")
        return tv

    def _secao_niveis(self, pai):
        self._secao_titulo(pai, "Níveis")
        caixa = ttk.Frame(pai, padding=(self.recuo, 0, 0, 0))
        caixa.pack(fill="both", expand=True)
        # controles primeiro (side=bottom) para nunca serem cortados quando a janela for baixa
        ttk.Label(caixa, text="Repetições é informativo (linha 13 da lista): os quantitativos já vêm somados do "
                              "modelo.", style="Dica.TLabel", wraplength=self.px(600)).pack(side="bottom", anchor="w")
        bts = ttk.Frame(caixa)
        bts.pack(side="bottom", fill="x", pady=(self.px(6), self.px(2)))
        self.bt_n_descer = ttk.Button(bts, text="Descer", style="Pequeno.TButton", command=lambda: self._mover_nivel(1))
        self.bt_n_descer.pack(side="right")
        self.bt_n_subir = ttk.Button(bts, text="Subir", style="Pequeno.TButton", command=lambda: self._mover_nivel(-1))
        self.bt_n_subir.pack(side="right", padx=(0, self.px(6)))
        self.bt_n_aplicar = ttk.Button(bts, text="Aplicar", style="Pequeno.TButton", command=self._aplicar_nivel)
        self.bt_n_aplicar.pack(side="right", padx=(0, self.px(6)))
        ed = ttk.Frame(caixa)
        ed.pack(side="bottom", fill="x", pady=(self.px(6), 0))
        self.v_n_cab = tk.StringVar()
        self.v_n_aba = tk.StringVar()
        self.v_n_rep = tk.StringVar()
        ttk.Label(ed, text="Cabeçalho", style="Rotulo.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(ed, text="Aba", style="Rotulo.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(ed, text="Repetições", style="Rotulo.TLabel").grid(row=0, column=2, sticky="w")
        self.en_n_cab = ttk.Entry(ed, textvariable=self.v_n_cab)
        self.en_n_cab.grid(row=1, column=0, sticky="ew", padx=(0, self.px(10)))
        self.cb_n_aba = ttk.Combobox(ed, textvariable=self.v_n_aba, values=list(NOMES_ABA.values()),
                                     state="readonly", width=20)
        self.cb_n_aba.grid(row=1, column=1, sticky="ew", padx=(0, self.px(10)))
        self.sp_n_rep = ttk.Spinbox(ed, textvariable=self.v_n_rep, from_=1, to=999, width=7)
        self.sp_n_rep.grid(row=1, column=2, sticky="w")
        ed.columnconfigure(0, weight=1)
        self.tv_niveis = self._tabela(caixa, [
            ("nivel", "Nível no modelo", 140, "w"), ("cab", "Cabeçalho na lista", 140, "w"),
            ("aba", "Aba", 104, "w"), ("rep", "Repetições", 88, "center"), ("inst", "Instâncias", 86, "e")], 6)
        self.tv_niveis.bind("<<TreeviewSelect>>", lambda e: self._nivel_selecionado())

    def _secao_ambientes(self, pai):
        self._secao_titulo(pai, "Ambientes (colunas)", pady=(16, 10))
        caixa = ttk.Frame(pai, padding=(self.recuo, 0, 0, 0))
        caixa.pack(fill="both", expand=True)
        ttk.Label(caixa, text="Locais com vários ambientes (ex.: SAT-MOTORISTA) ficam em colunas vizinhas e o "
                              "quantitativo vai em células mescladas.", style="Dica.TLabel",
                  wraplength=self.px(600)).pack(side="bottom", anchor="w")
        self.v_realocar = tk.BooleanVar(value=self.prefs.get("realocar", True))
        ttk.Checkbutton(caixa, text="Lançar no ambiente atendido as peças com nível de referência de outro pavimento",
                        variable=self.v_realocar, command=self._trocar_realocar).pack(side="bottom", anchor="w",
                                                                                     pady=(self.px(6), 0))
        ed = ttk.Frame(caixa)
        ed.pack(side="bottom", fill="x", pady=(self.px(6), 0))
        ttk.Label(ed, text="Nome da coluna", style="Rotulo.TLabel").grid(row=0, column=0, sticky="w")
        self.v_a_nome = tk.StringVar()
        self.en_a_nome = ttk.Entry(ed, textvariable=self.v_a_nome)
        self.en_a_nome.grid(row=1, column=0, sticky="ew", padx=(0, self.px(10)))
        self.bt_a_aplicar = ttk.Button(ed, text="Aplicar", style="Pequeno.TButton", command=self._aplicar_amb)
        self.bt_a_aplicar.grid(row=1, column=1, padx=(0, self.px(6)))
        self.bt_a_esq = ttk.Button(ed, text="Para a esquerda", style="Pequeno.TButton",
                                   command=lambda: self._mover_amb(-1))
        self.bt_a_esq.grid(row=1, column=2, padx=(0, self.px(6)))
        self.bt_a_dir = ttk.Button(ed, text="Para a direita", style="Pequeno.TButton",
                                   command=lambda: self._mover_amb(1))
        self.bt_a_dir.grid(row=1, column=3)
        ed.columnconfigure(0, weight=1)
        self.tv_amb = self._tabela(caixa, [
            ("aba", "Aba", 66, "w"), ("nivel", "Nível", 120, "w"), ("coluna", "Coluna", 130, "w"),
            ("origem", "Local atendido no modelo", 158, "w"), ("inst", "Instâncias", 86, "e")], 7)
        self.tv_amb.bind("<<TreeviewSelect>>", lambda e: self._amb_selecionado())

    def _barra_acoes(self):
        barra = ttk.Frame(self, padding=(self.px(24), self.px(10), self.px(24), self.px(12)))
        barra.pack(side="bottom", fill="x")
        tk.Frame(self, bg=BORDA, height=1).pack(side="bottom", fill="x")
        esquerda = ttk.Frame(barra)
        esquerda.pack(side="left", fill="x", expand=True, padx=(0, self.px(24)))
        topo = ttk.Frame(esquerda)
        topo.pack(fill="x")
        self.lb_etapa = ttk.Label(topo, text="Selecione o CSV para começar.", style="Rotulo.TLabel")
        self.lb_etapa.pack(side="left")
        self.lb_pct = ttk.Label(topo, text="", style="Rotulo.TLabel")
        self.lb_pct.pack(side="right")
        self.progresso = ttk.Progressbar(esquerda, style="Cadari.Horizontal.TProgressbar", maximum=100)
        self.progresso.pack(fill="x", pady=(self.px(4), self.px(4)))
        self.lb_msg = ttk.Label(esquerda, text="", style="Dica.TLabel", wraplength=self.px(760), justify="left")
        self.lb_msg.pack(anchor="w")
        botoes = ttk.Frame(barra)
        botoes.pack(side="right")
        self.bt_pasta = ttk.Button(botoes, text="Abrir pasta", style="Secundario.TButton", command=self.abrir_pasta)
        self.bt_pasta.pack(side="right")
        self.bt_abrir = ttk.Button(botoes, text="Abrir planilha", style="Secundario.TButton",
                                   command=self.abrir_planilha)
        self.bt_abrir.pack(side="right", padx=(0, self.px(10)))
        self.bt_gerar = ttk.Button(botoes, text="Gerar planilha", style="Primario.TButton", command=self.gerar)
        self.bt_gerar.pack(side="right", padx=(0, self.px(10)))

    def _rodape(self):
        h = self.px(44)
        cv = tk.Canvas(self, height=h, bg=AZUL, highlightthickness=0, bd=0)
        cv.pack(side="bottom", fill="x")

        def desenhar(evento=None):
            cv.delete("all")
            w = cv.winfo_width()
            passo = self.px(7)
            for x in range(-h, w + h, passo):
                cv.create_line(x, h, x + h, 0, fill=AZUL_2, width=1)
            if self.img_icone:
                cv.create_image(w - self.px(22), h // 2, image=self.img_icone, anchor="e")
        cv.bind("<Configure>", desenhar)

    # ------------------------------------------------------------------ estado
    def _atualizar_estado(self):
        tem = self.analise is not None and not self.ocupado
        for b in (self.bt_gerar, self.bt_n_aplicar, self.bt_n_subir, self.bt_n_descer,
                  self.bt_a_aplicar, self.bt_a_esq, self.bt_a_dir):
            b.state(["!disabled"] if tem else ["disabled"])
        self.bt_csv.state(["disabled"] if self.ocupado else ["!disabled"])
        pronto = self.ultimo_resultado is not None and not self.ocupado
        for b in (self.bt_abrir, self.bt_pasta):
            b.state(["!disabled"] if pronto else ["disabled"])

    def _mostrar_modelo(self):
        nome = os.path.basename(self.caminho_modelo)
        ok = os.path.exists(self.caminho_modelo)
        self.lb_modelo.configure(text=f"Modelo da folha de rosto: {nome}" if ok else
                                 "Modelo da folha de rosto não encontrado. Selecione o arquivo do modelo.")

    def _atualizar_obra(self):
        codigo = self.v_ca.get()
        if not codigo.strip():
            self.lb_obra.configure(text="Digite o código da obra como na relação de projetos (ex.: CA-1276).")
            return
        proj = self.relacao.buscar(codigo)
        normal = RelacaoProjetos.normalizar_codigo(codigo)
        if proj:
            partes = [f"Obra: {proj['obra']}"]
            if proj.get("cliente"):
                partes.append(f"cliente: {proj['cliente']}")
            if proj.get("referencia"):
                partes.append(f"referência: {proj['referencia']}")
            self.lb_obra.configure(text=", ".join(partes) + ".")
            prefixo = RelacaoProjetos.prefixo_obra(proj)
            atual = self.v_doc.get().strip()
            if prefixo and (not atual or atual == self.sugestao_doc):
                self.sugestao_doc = config.PADRAO_DOCUMENTO.format(obra=prefixo)
                self.v_doc.set(self.sugestao_doc)
        elif self.relacao.carregada:
            self.lb_obra.configure(text=f"{normal} não está na relação de projetos salva no modelo. A obra e a logo "
                                        "do cliente aparecem quando o Excel atualizar o vínculo com a relação.")
        else:
            self.lb_obra.configure(text="O modelo não tem a relação de projetos salva; a obra e a logo do cliente "
                                        "vêm do vínculo quando o arquivo for aberto no Excel.")

    # ------------------------------------------------------------------ CSV
    def selecionar_csv(self):
        inicial = self.prefs.get("pasta_csv") or os.path.expanduser("~")
        caminho = filedialog.askopenfilename(
            title="Selecionar CSV por instância", initialdir=inicial,
            filetypes=[("Tabela exportada do Revit", "*.csv *.txt"), ("Todos os arquivos", "*.*")])
        if not caminho:
            return
        self.prefs["pasta_csv"] = os.path.dirname(caminho)
        self.ultimo_resultado = None
        self.ocupado = True
        self._atualizar_estado()
        self.lb_etapa.configure(text="Lendo o CSV...")
        self.lb_pct.configure(text="")
        self.progresso.configure(mode="determinate", value=0)
        self.lb_msg.configure(text="")

        def tarefa():
            try:
                leitura = carregar_csv(caminho, lambda x: self.fila.put(("prog", 90 * x)))
                analise = nova_analise(leitura, realocar=self.v_realocar.get())
                self.fila.put(("csv", (leitura, analise)))
            except ErroLeitura as e:
                self.fila.put(("erro", ("Não foi possível ler o CSV", str(e))))
            except Exception as e:      # noqa: BLE001
                self.fila.put(("erro", ("Não foi possível ler o CSV",
                                        f"{e}\n\nConfira se o arquivo é a exportação por instância da tabela "
                                        "LISTA DE MATERIAL GERAL.")))
        threading.Thread(target=tarefa, daemon=True).start()

    def _csv_carregado(self, leitura, analise):
        self.leitura, self.analise = leitura, analise
        nome = os.path.basename(leitura.caminho)
        self.lb_arquivo.configure(text=nome)
        n_itens = len({(i.descricao, i.unidade) for i in leitura.instancias_validas})
        self.lb_detalhe.configure(text=f"{numero_br(len(leitura.instancias))} instâncias, {n_itens} itens, "
                                       f"{len(analise.niveis)} níveis. Linhas no arquivo: "
                                       f"{numero_br(leitura.total_linhas)}.")
        self.progresso.configure(value=100)
        self.lb_pct.configure(text="100%")
        if leitura.problemas:
            self.lb_status_csv.configure(
                text=f"{len(leitura.problemas)} linhas com problema no CSV (NF, descrição vazia, sem local atendido "
                     "ou sem nível). Corrija no modelo ou gere sem essas linhas.", style="Alerta.TLabel")
            self.bt_problemas.pack(anchor="w", pady=(self.px(6), 0))
            self.lb_etapa.configure(text="CSV lido com problemas.")
        else:
            self.lb_status_csv.configure(text="Nenhum problema encontrado no CSV.", style="Ok.TLabel")
            self.bt_problemas.pack_forget()
            self.lb_etapa.configure(text="CSV lido. Confira níveis e ambientes e gere a planilha.")
        self._preencher_niveis()
        self._preencher_ambientes()
        self.ocupado = False
        self._atualizar_estado()
        if leitura.problemas:
            self.mostrar_problemas()

    def mostrar_problemas(self):
        if self.leitura and self.leitura.problemas:
            JanelaProblemas(self, self.leitura)

    def trocar_modelo(self):
        caminho = filedialog.askopenfilename(title="Selecionar modelo da folha de rosto",
                                             initialdir=os.path.dirname(self.caminho_modelo),
                                             filetypes=[("Planilha do Excel", "*.xlsx")])
        if caminho:
            self.caminho_modelo = caminho
            self.prefs["modelo"] = caminho
            self.relacao = RelacaoProjetos(caminho)
            self._mostrar_modelo()
            self._atualizar_obra()

    # ------------------------------------------------------------------ níveis e ambientes
    def _preencher_niveis(self, selecionar=None):
        tv = self.tv_niveis
        tv.delete(*tv.get_children())
        for i, (nome, cfg) in enumerate(self.analise.niveis.items()):
            tv.insert("", "end", iid=nome, values=(nome, cfg.cabecalho, NOMES_ABA[cfg.aba], cfg.repeticoes,
                                                   numero_br(cfg.instancias)),
                      tags=("par" if i % 2 == 0 else "impar",))
        if selecionar and tv.exists(selecionar):
            tv.selection_set(selecionar)
            tv.see(selecionar)

    def _preencher_ambientes(self, selecionar=None):
        tv = self.tv_amb
        tv.delete(*tv.get_children())
        self._mapa_amb = {}
        cols = self.analise.colunas()
        i = 0
        for aba in (COMUM, PRIVATIVA):
            for c in cols[aba]:
                iid = f"{c.chave[0]}||{c.chave[1]}"
                self._mapa_amb[iid] = c.chave
                cab = self.analise.niveis[c.nivel].cabecalho
                tv.insert("", "end", iid=iid, values=("Comum" if aba == COMUM else "Privativa", cab, c.titulo,
                                                      "; ".join(c.origens),
                                                      numero_br(c.instancias)),
                          tags=("par" if i % 2 == 0 else "impar",))
                i += 1
        if selecionar and tv.exists(selecionar):
            tv.selection_set(selecionar)
            tv.see(selecionar)

    def _nivel_selecionado(self):
        sel = self.tv_niveis.selection()
        if not sel or not self.analise:
            return
        cfg = self.analise.niveis[sel[0]]
        self.v_n_cab.set(cfg.cabecalho)
        self.v_n_aba.set(NOMES_ABA[cfg.aba])
        self.v_n_rep.set(str(cfg.repeticoes))

    def _aplicar_nivel(self):
        sel = self.tv_niveis.selection()
        if not sel:
            self._mensagem("Selecione um nível na tabela para editar.")
            return
        cfg = self.analise.niveis[sel[0]]
        cfg.cabecalho = self.v_n_cab.get().strip() or cfg.cabecalho
        cfg.aba = ABA_POR_NOME.get(self.v_n_aba.get(), cfg.aba)
        try:
            cfg.repeticoes = max(1, int(self.v_n_rep.get()))
        except ValueError:
            self._mensagem("Repetições precisa ser um número inteiro.")
        self._preencher_niveis(sel[0])
        self._preencher_ambientes()

    def _mover_nivel(self, passo):
        sel = self.tv_niveis.selection()
        if sel:
            self.analise.mover_nivel(sel[0], passo)
            self._preencher_niveis(sel[0])
            self._preencher_ambientes()

    def _amb_selecionado(self):
        sel = self.tv_amb.selection()
        if sel:
            self.v_a_nome.set(self.tv_amb.item(sel[0], "values")[2])

    def _aplicar_amb(self):
        sel = self.tv_amb.selection()
        if not sel:
            self._mensagem("Selecione uma coluna na tabela de ambientes para renomear.")
            return
        nome = self.v_a_nome.get().strip()
        if nome:
            self.analise.titulos_colunas[self._mapa_amb[sel[0]]] = nome
            self._preencher_ambientes(sel[0])

    def _mover_amb(self, passo):
        sel = self.tv_amb.selection()
        if sel:
            self.analise.mover_ambiente(self._mapa_amb[sel[0]], passo)
            self._preencher_ambientes(sel[0])

    def _trocar_realocar(self):
        self.prefs["realocar"] = self.v_realocar.get()
        if self.analise:
            self.analise.realocar = self.v_realocar.get()
            self._preencher_ambientes()

    # ------------------------------------------------------------------ geração
    def _dados_capa(self):
        try:
            data = dt.datetime.strptime(self.v_data.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            raise ValueError("A data da revisão precisa estar no formato dd/mm/aaaa.")
        ca = RelacaoProjetos.normalizar_codigo(self.v_ca.get()) if self.v_ca.get().strip() else ""
        proj = self.relacao.buscar(ca) if ca else None
        return DadosCapa(
            ca=ca, obra=(proj or {}).get("obra", ""), sistema=self.v_sistema.get().strip(),
            etapa=self.v_etapa.get().strip(), documento=self.v_documento.get().strip(),
            numero_documento=self.v_doc.get().strip(), rev=self.v_rev.get().strip(), te=self.v_te.get().strip(),
            descricao_rev=self.v_desc_rev.get().strip(), elab=self.v_elab.get().strip(),
            verif=self.v_verif.get().strip(), resp=self.v_resp.get().strip(), data=data,
            desenhos=self.v_desenhos.get().strip(), prefixo=self.v_prefixo.get().strip() or "CLI")

    def gerar(self):
        if not self.analise:
            return
        if not os.path.exists(self.caminho_modelo):
            messagebox.showerror("Modelo não encontrado",
                                 "O modelo da folha de rosto não foi encontrado. Clique em Trocar modelo e "
                                 f"selecione o arquivo {NOME_MODELO}.", parent=self)
            return
        try:
            capa = self._dados_capa()
        except ValueError as e:
            messagebox.showerror("Dados da folha de rosto", str(e), parent=self)
            return
        if not capa.numero_documento:
            messagebox.showerror("Dados da folha de rosto",
                                 "Preencha o Nº do documento. Ele também é o nome do arquivo gerado.", parent=self)
            return
        if not capa.ca and not messagebox.askyesno(
                "Código da obra vazio", "Sem o código da obra a folha de rosto fica sem o nome da obra e sem a "
                                        "logo do cliente. Gerar assim mesmo?", parent=self):
            return
        if self.leitura.problemas:
            escolha = PerguntaProblemas(self, len(self.leitura.problemas)).resultado
            if escolha == "ver":
                self.mostrar_problemas()
                return
            if escolha != "gerar":
                return
        pasta = self.prefs.get("pasta_saida") or os.path.dirname(self.leitura.caminho)
        caminho = filedialog.asksaveasfilename(
            title="Salvar lista de material", initialdir=pasta, initialfile=capa.numero_documento + ".xlsx",
            defaultextension=".xlsx", filetypes=[("Planilha do Excel", "*.xlsx")])
        if not caminho:
            return
        nome = os.path.splitext(os.path.basename(caminho))[0]
        if nome != capa.numero_documento:
            capa.numero_documento = nome
            self.v_doc.set(nome)
        self.prefs["pasta_saida"] = os.path.dirname(caminho)
        self._guardar_preferencias()
        self.ocupado = True
        self.ultimo_resultado = None
        self._atualizar_estado()
        self.progresso.configure(value=0)
        self.lb_etapa.configure(text="Gerando a planilha...")
        self.lb_msg.configure(text="")
        leitura, analise, modelo = self.leitura, self.analise, self.caminho_modelo

        def tarefa():
            try:
                res = gerar(leitura, analise, capa, modelo, caminho,
                            progresso=lambda x: self.fila.put(("prog", 100 * x)),
                            pasta_recursos=localizar("recursos"))
                self.fila.put(("ok", res))
            except PermissionError:
                self.fila.put(("erro", ("Não foi possível salvar",
                                        "O arquivo está aberto no Excel ou a pasta não permite gravação. Feche o "
                                        "arquivo (ou escolha outra pasta) e gere de novo.")))
            except ErroModelo as e:
                self.fila.put(("erro", ("Modelo da folha de rosto", str(e))))
            except Exception as e:      # noqa: BLE001
                self.fila.put(("erro", ("Não foi possível gerar a planilha", str(e))))
        threading.Thread(target=tarefa, daemon=True).start()

    def _gerado(self, res):
        self.ultimo_resultado = res
        self.ocupado = False
        self._atualizar_estado()
        self.progresso.configure(value=100)
        self.lb_pct.configure(text="100%")
        self.lb_etapa.configure(text=f"Planilha gerada: {os.path.basename(res.caminho)}")
        partes = [f"{res.itens} itens em {len(res.abas)} abas", f"{res.mesclagens} células mescladas",
                  f"{res.comentarios} comentários"]
        texto = ", ".join(partes) + "."
        if res.diferencas:
            difs = "; ".join(f"{d['subgrupo'] or d['grupo']}: CSV {d['csv']:g}, lido {d['soma']:g}"
                             for d in res.diferencas[:4])
            texto += f" Subtotais do CSV com diferença: {difs}."
        else:
            texto += " Leitura conferida com os subtotais do CSV."
        if res.ignoradas:
            fora = [d for d in res.conferencia if abs(d["fora"]) > 1e-9]
            detalhe = "; ".join(f"{d['subgrupo'] or d['grupo']}: {d['fora']:g}" for d in fora[:4])
            texto += f" {res.ignoradas} linhas com problema ficaram fora da lista"
            texto += f" ({detalhe})." if detalhe else "."
        self.lb_msg.configure(text=texto)

    def _mensagem(self, texto):
        self.lb_msg.configure(text=texto)

    def abrir_planilha(self):
        if self.ultimo_resultado:
            try:
                abrir_no_sistema(self.ultimo_resultado.caminho)
            except Exception as e:      # noqa: BLE001
                messagebox.showerror("Abrir planilha", f"Não foi possível abrir o arquivo: {e}", parent=self)

    def abrir_pasta(self):
        if self.ultimo_resultado:
            try:
                abrir_no_sistema(os.path.dirname(self.ultimo_resultado.caminho))
            except Exception as e:      # noqa: BLE001
                messagebox.showerror("Abrir pasta", f"Não foi possível abrir a pasta: {e}", parent=self)

    # ------------------------------------------------------------------ fila e saída
    def _processar_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "prog":
                    self.progresso.configure(mode="determinate", value=dado)
                    self.lb_pct.configure(text=f"{int(dado)}%")
                elif tipo == "csv":
                    self._csv_carregado(*dado)
                elif tipo == "ok":
                    self._gerado(dado)
                elif tipo == "erro":
                    self.ocupado = False
                    self._atualizar_estado()
                    self.progresso.configure(value=0)
                    self.lb_pct.configure(text="")
                    self.lb_etapa.configure(text=dado[0])
                    messagebox.showerror(dado[0], dado[1], parent=self)
        except queue.Empty:
            pass
        self.after(80, self._processar_fila)

    def _guardar_preferencias(self):
        self.prefs.update(dict(sistema=self.v_sistema.get(), etapa=self.v_etapa.get(),
                               documento=self.v_documento.get(), prefixo=self.v_prefixo.get(),
                               te=self.v_te.get(), elab=self.v_elab.get(), verif=self.v_verif.get(),
                               resp=self.v_resp.get(), ca=self.v_ca.get(), realocar=self.v_realocar.get()))
        salvar_preferencias(self.prefs)

    def _sair(self):
        self._guardar_preferencias()
        self.destroy()


class JanelaProblemas(tk.Toplevel):
    def __init__(self, app, leitura):
        super().__init__(app)
        self.app, self.leitura = app, leitura
        self.title("Problemas no CSV")
        self.configure(bg=BRANCO)
        self.geometry(f"{app.px(1120)}x{app.px(580)}")
        self.minsize(app.px(800), app.px(400))
        self.transient(app)
        frame = ttk.Frame(self, padding=(app.px(24), app.px(18), app.px(24), app.px(16)))
        frame.pack(fill="both", expand=True)
        app._secao_titulo(frame, "Problemas no CSV")
        corpo = ttk.Frame(frame, padding=(app.recuo, 0, 0, 0))
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text=f"{len(leitura.problemas)} linhas precisam de correção. Corrija no modelo e exporte o "
                              "CSV de novo. Se gerar a planilha assim mesmo, essas linhas ficam fora da lista.",
                  wraplength=app.px(900), style="Info.TLabel").pack(anchor="w", pady=(0, app.px(10)))
        tv = app._tabela(corpo, [("linha", "Linha do CSV", 100, "center"), ("prob", "Problema", 230, "w"),
                                 ("desc", "Descrição", 290, "w"), ("local", "CA_Local Atendido", 170, "w"),
                                 ("nivel", "CA_Nível de Referência", 180, "w")], 12)
        tv.tag_configure("pendente", background=PENDENTE)
        for p in leitura.problemas:
            tv.insert("", "end", values=(p.linha, p.texto, p.descricao, p.local, p.nivel), tags=("pendente",))
        botoes = ttk.Frame(corpo)
        botoes.pack(fill="x", pady=(app.px(12), 0))
        ttk.Button(botoes, text="Fechar", style="Primario.TButton", command=self.destroy).pack(side="right")
        ttk.Button(botoes, text="Exportar lista", style="Secundario.TButton", command=self.exportar).pack(
            side="right", padx=(0, app.px(10)))
        self.grab_set()
        self.focus_set()

    def exportar(self):
        base = os.path.splitext(os.path.basename(self.leitura.caminho))[0]
        caminho = filedialog.asksaveasfilename(parent=self, title="Exportar lista de problemas",
                                               initialdir=os.path.dirname(self.leitura.caminho),
                                               initialfile=f"{base} - problemas.csv", defaultextension=".csv",
                                               filetypes=[("CSV", "*.csv")])
        if not caminho:
            return
        import csv
        with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Linha do CSV", "Problema", "Descrição", "CA_Local Atendido", "CA_Nível de Referência"])
            for p in self.leitura.problemas:
                w.writerow([p.linha, p.texto, p.descricao, p.local, p.nivel])
        messagebox.showinfo("Exportar lista", f"Lista salva em:\n{caminho}", parent=self)


class PerguntaProblemas(tk.Toplevel):
    def __init__(self, app, quantidade):
        super().__init__(app)
        self.resultado = None
        self.title("Linhas com problema")
        self.configure(bg=BRANCO)
        self.transient(app)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=(app.px(24), app.px(20), app.px(24), app.px(18)))
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"O CSV tem {quantidade} linhas com problema (NF, descrição vazia, sem local atendido ou "
                          "sem nível).", wraplength=app.px(520), style="Info.TLabel").pack(anchor="w")
        ttk.Label(f, text="Se gerar agora, essas linhas ficam fora da lista e os totais não batem com o modelo.",
                  wraplength=app.px(520), style="Rotulo.TLabel").pack(anchor="w", pady=(app.px(6), app.px(16)))
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="Gerar mesmo assim", style="Primario.TButton",
                   command=lambda: self._fim("gerar")).pack(side="right")
        ttk.Button(b, text="Cancelar", style="Secundario.TButton",
                   command=lambda: self._fim(None)).pack(side="right", padx=(0, app.px(10)))
        ttk.Button(b, text="Ver problemas", style="Secundario.TButton",
                   command=lambda: self._fim("ver")).pack(side="right", padx=(0, app.px(10)))
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.grab_set()
        self.wait_window(self)

    def _fim(self, valor):
        self.resultado = valor
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
