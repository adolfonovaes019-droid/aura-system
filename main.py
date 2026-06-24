import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import unicodedata
import json
import os
import sys
import shutil

try:
    from tkcalendar import DateEntry
except ImportError:
    messagebox.showerror("Erro", "Instale o tkcalendar: pip install tkcalendar")
    sys.exit()

try:
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror("Erro", "Instale o pillow: pip install pillow")
    sys.exit()

def obter_diretorio_app():
    """Retorna a pasta de dados do app (LocalAppData, sem precisar de admin).
    Se houver dados antigos em Program Files (versões anteriores), migra automaticamente."""
    base_dir = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    app_dir = os.path.join(base_dir, "AURA SYSTEM")

    try:
        os.makedirs(app_dir, exist_ok=True)
    except OSError as e:
        messagebox.showerror("Erro", f"Não foi possível criar a pasta de dados.\n\n{app_dir}\n\n{e}")
        sys.exit()

    legado_base = os.environ.get("PROGRAMFILES", "C:\\Program Files")
    legado_dir = os.path.join(legado_base, "AURA SYSTEM")
    if os.path.isdir(legado_dir):
        for nome in ("aura_system.db", "config_impressora.json"):
            origem = os.path.join(legado_dir, nome)
            destino = os.path.join(app_dir, nome)
            if os.path.exists(origem) and not os.path.exists(destino):
                try:
                    shutil.copy2(origem, destino)
                except (OSError, PermissionError):
                    pass

    return app_dir

try:
    import win32print
    IMPRESSAO_DISPONIVEL = True
except ImportError:
    IMPRESSAO_DISPONIVEL = False

class Database:
    def __init__(self):
        pasta_app = obter_diretorio_app()
        caminho_db = os.path.join(pasta_app, "aura_system.db") 
        
        self.conn = sqlite3.connect(caminho_db)
        self.cursor = self.conn.cursor()
        self.criar_tabelas()
        self.corrigir_banco() 
        self.restaurar_cardapio_base()

    def criar_tabelas(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT, 
                nome TEXT, telefone TEXT UNIQUE, endereco TEXT, 
                bairro TEXT, bloco TEXT, apto TEXT,
                taxa REAL DEFAULT 0
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT, 
                nome TEXT, 
                preco REAL,
                categoria TEXT DEFAULT 'LANCHES'
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS vendas (
                id INTEGER PRIMARY KEY AUTOINCREMENT, data_hora TEXT, total REAL,
                cliente TEXT, tipo_pedido TEXT, pagamento TEXT, taxa_entrega REAL,
                itens_json TEXT, endereco_entrega TEXT
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS config_loja (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_loja TEXT,
                senha_finan TEXT,
                endereco TEXT DEFAULT '',
                cnpj TEXT DEFAULT '',
                telefone TEXT DEFAULT '',
                instagram TEXT DEFAULT '',
                slogan TEXT DEFAULT ''
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS adicionais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT UNIQUE,
                preco REAL
            )
        """)
        self.conn.commit()

    def corrigir_banco(self):
        colunas_clientes = [("taxa", "REAL DEFAULT 0"), ("bloco", "TEXT"), ("apto", "TEXT")]
        for col, tipo in colunas_clientes:
            try: self.cursor.execute(f"ALTER TABLE clientes ADD COLUMN {col} {tipo}")
            except sqlite3.OperationalError: pass

        try: self.cursor.execute("ALTER TABLE produtos ADD COLUMN categoria TEXT DEFAULT 'LANCHES'")
        except sqlite3.OperationalError: pass

        for col in ["endereco", "cnpj", "telefone", "instagram", "slogan"]:
            try: self.cursor.execute(f"ALTER TABLE config_loja ADD COLUMN {col} TEXT DEFAULT ''")
            except sqlite3.OperationalError: pass

        adicionais_base = [
            ("OVO", 2.0), ("HAMBURGUER", 10.0), ("BACON", 5.0), ("QUEIJO", 4.0),
            ("CATUPIRY", 3.0), ("CHEDDAR", 3.0), ("MILHO", 2.0), ("BATATA EXTRA", 5.0),
            ("MAIONESE AVULSA", 4.0), ("FILE DE FRANGO", 15.0), ("CONTRA FILE", 15.0),
            ("ACRESCIMO", 10.0), ("ABACAXI", 5.0)
        ]
        try:
            self.cursor.executemany("INSERT OR IGNORE INTO adicionais (nome, preco) VALUES (?, ?)", adicionais_base)
        except sqlite3.OperationalError: pass

        self.conn.commit()
        self.semear_adicionais_base()

    def semear_adicionais_base(self):
        res = self.cursor.execute("SELECT COUNT(*) FROM adicionais").fetchone()
        if res[0] == 0:
            base = [
                ("OVO", 2.0), ("HAMBURGUER", 10.0), ("BACON", 5.0), ("QUEIJO", 4.0),
                ("CATUPIRY", 3.0), ("CHEDDAR", 3.0), ("MILHO", 2.0), ("BATATA EXTRA", 5.0),
                ("MAIONESE AVULSA", 4.0), ("FILE DE FRANGO", 15.0), ("CONTRA FILE", 15.0),
                ("ACRESCIMO", 10.0), ("ABACAXI", 5.0)
            ]
            self.cursor.executemany("INSERT OR IGNORE INTO adicionais (nome, preco) VALUES (?, ?)", base)
            self.conn.commit()

    def get_adicionais(self):
        return self.cursor.execute("SELECT id, nome, preco FROM adicionais ORDER BY nome").fetchall()

    def upsert_adicional(self, nome, preco, ad_id=None):
        nome = nome.strip().upper()
        if ad_id:
            self.cursor.execute("UPDATE adicionais SET nome=?, preco=? WHERE id=?", (nome, preco, ad_id))
        else:
            self.cursor.execute("INSERT OR REPLACE INTO adicionais (nome, preco) VALUES (?, ?)", (nome, preco))
        self.conn.commit()

    def excluir_adicional(self, ad_id):
        self.cursor.execute("DELETE FROM adicionais WHERE id=?", (ad_id,))
        self.conn.commit()

    def restaurar_cardapio_base(self):
        res = self.cursor.execute("SELECT COUNT(*) FROM produtos").fetchone()
        if res[0] <= 5:
            self.cursor.execute("DELETE FROM produtos")
            itens = [
                ('X-FRANGO', 31.00, 'LANCHES'), ('X-FRANGO EGG', 33.00, 'LANCHES'), ('X-FRANGO BACON', 36.00, 'LANCHES'), ('X-FRANGO TUDO', 41.00, 'LANCHES'),
                ('X-CHURRASCO', 33.00, 'LANCHES'), ('X-CHURRASCO EGG', 35.00, 'LANCHES'), ('X-CHURRASCO BACON', 38.00, 'LANCHES'), ('X-CHURRASCO TUDO', 43.00, 'LANCHES'),
                ('X-BURGUER', 25.00, 'LANCHES'), ('X-SALADA', 30.00, 'LANCHES'), ('X-EGG', 32.00, 'LANCHES'), ('X-BACON', 34.00, 'LANCHES'), ('X-TUDO', 37.00, 'LANCHES'),
                ('FRANBURGUER', 53.00, 'LANCHES'), ('X-TUDO ESPECIAL', 53.00, 'LANCHES'), ('X-MONSTRO', 58.00, 'LANCHES'), ('X-SALADA ESPECIAL', 35.00, 'LANCHES'),
                ('FRITAS SIMPLES', 28.00, 'PORÇÕES'), ('FRITAS CATUPIRY/CHEDDAR', 35.00, 'PORÇÕES'), ('FRITAS MUSSARELA', 35.00, 'PORÇÕES'), ('FRITAS BACON/CHEDDAR', 42.00, 'PORÇÕES'),
                ('X-SALADA ABACAXI', 36.00, 'LANCHES'), ('X-EGG ABACAXI', 38.00, 'LANCHES'), ('X-BACON ABACAXI', 40.00, 'LANCHES'), ('X-TUDO ABACAXI', 43.00, 'LANCHES'), ('X-BRAVO ABACAXI', 50.00, 'LANCHES'),
                ('BARCA INDIVIDUAL', 48.00, 'PORÇÕES'), ('BARCA GRANDE', 95.00, 'PORÇÕES'),
                ('COCA COLA', 7.00, 'REFRIGERANTES'), ('COCA ZERO', 7.00, 'REFRIGERANTES'), ('GUARANA LATA', 7.00, 'REFRIGERANTES'), ('HEINEKEN LATA', 10.00, 'REFRIGERANTES'), ('ORIGINAL LATA', 8.00, 'REFRIGERANTES'), ('DEL VALLE', 8.00, 'REFRIGERANTES'), ('AGUA MINERAL', 4.00, 'REFRIGERANTES'),
                ('AGUA COM GAS', 5.00, 'REFRIGERANTES'), ('FANTA UVA', 7.00, 'REFRIGERANTES'), ('SCHWEPPES', 7.00, 'REFRIGERANTES'), ('SPRITE', 7.00, 'REFRIGERANTES')
            ]
            self.cursor.executemany("INSERT INTO produtos (nome, preco, categoria) VALUES (?,?,?)", itens)
            self.conn.commit()

    def get_config(self):
        return self.cursor.execute(
            "SELECT nome_loja, senha_finan, endereco, cnpj, telefone, instagram, slogan FROM config_loja WHERE id=1"
        ).fetchone()

    def set_config(self, nome, senha, endereco="", cnpj="", telefone="", instagram="", slogan=""):
        if self.get_config():
            self.cursor.execute(
                "UPDATE config_loja SET nome_loja=?, senha_finan=?, endereco=?, cnpj=?, telefone=?, instagram=?, slogan=? WHERE id=1",
                (nome, senha, endereco, cnpj, telefone, instagram, slogan)
            )
        else:
            self.cursor.execute(
                "INSERT INTO config_loja (nome_loja, senha_finan, endereco, cnpj, telefone, instagram, slogan) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (nome, senha, endereco, cnpj, telefone, instagram, slogan)
            )
        self.conn.commit()

class AppPDV:
    def __init__(self, root):
        self.db = Database()
        self.root = root
        self.root.withdraw()
        self._resize_timer = None

        self.configuracoes_loja = self.db.get_config()
        if not self.configuracoes_loja:
            self.abrir_primeiro_acesso()
        else:
            (self.nome_loja, self.senha_finan, self.loja_endereco,
             self.loja_cnpj, self.loja_telefone, self.loja_instagram,
             self.loja_slogan) = self.configuracoes_loja
            self.iniciar_app()

    def abrir_primeiro_acesso(self):
        BG = "#0E0E10"
        CARD = "#181820"
        SOFT = "#22222C"
        ACCENT = "#E63946"
        YELLOW = "#F4B400"
        MUTED = "#8A8A99"

        win = tk.Toplevel(self.root)
        win.title("AURA SYSTEM - Configuração Inicial")
        win.attributes("-toolwindow", True)
        self.centralizar_janela_toplevel(win, 560, 540)
        win.configure(bg=BG)
        win.protocol("WM_DELETE_WINDOW", sys.exit)

        f_logo = tk.Frame(win, bg=BG)
        f_logo.pack(pady=(36, 6))
        bloco = tk.Frame(f_logo, bg=ACCENT, padx=18, pady=8)
        bloco.pack()
        tk.Label(bloco, text="A", font=("Segoe UI", 26, "bold"), fg="white", bg=ACCENT).pack()

        tk.Label(win, text="AURA SYSTEM", font=("Segoe UI", 26, "bold"), fg=YELLOW, bg=BG).pack(pady=(14, 2))
        tk.Label(win, text="Bem-vindo. Vamos configurar a sua loja em 30 segundos.",
                 font=("Segoe UI", 10), fg=MUTED, bg=BG).pack(pady=(0, 10))

        f_form = tk.Frame(win, bg=BG)
        f_form.pack(pady=20, fill="x", padx=60)

        tk.Label(f_form, text="NOME DA SUA LOJA", font=("Segoe UI", 9, "bold"), fg=MUTED, bg=BG).pack(anchor="w", pady=(10, 4))
        ent_nome = tk.Entry(f_form, font=("Segoe UI", 13), bg=SOFT, fg="white", relief="flat",
                            highlightthickness=1, highlightbackground=SOFT, highlightcolor=ACCENT,
                            insertbackground="white")
        ent_nome.pack(fill="x", ipady=6)

        tk.Label(f_form, text="SENHA PARA O FINANCEIRO", font=("Segoe UI", 9, "bold"), fg=MUTED, bg=BG).pack(anchor="w", pady=(20, 4))
        ent_senha = tk.Entry(f_form, show="*", font=("Segoe UI", 13), bg=SOFT, fg="white", relief="flat",
                             highlightthickness=1, highlightbackground=SOFT, highlightcolor=ACCENT,
                             insertbackground="white")
        ent_senha.pack(fill="x", ipady=6)

        def salvar_primeiro_acesso():
            n = ent_nome.get().strip().upper()
            s = ent_senha.get().strip()
            if not n or not s:
                messagebox.showwarning("Aviso", "Preencha todos os campos!")
                return
            self.db.set_config(n, s)
            self.nome_loja = n
            self.senha_finan = s
            self.loja_endereco = ""
            self.loja_cnpj = ""
            self.loja_telefone = ""
            self.loja_instagram = ""
            self.loja_slogan = ""
            win.destroy()
            self.iniciar_app()

        tk.Button(win, text="INICIAR SISTEMA", font=("Segoe UI", 12, "bold"), bg="#2ECC71",
                  fg="white", relief="flat", command=salvar_primeiro_acesso, cursor="hand2",
                  height=2, bd=0).pack(pady=18, fill="x", padx=60)
        tk.Label(win, text="Você poderá adicionar endereço, CNPJ e telefone na aba CONFIGURAÇÕES.",
                 font=("Segoe UI", 9, "italic"), fg=MUTED, bg=BG).pack(pady=(0, 10))

    def carregar_configuracoes(self):
        pasta_app = obter_diretorio_app()
        self.config_file = os.path.join(pasta_app, "config_impressora.json")
        
        self.valores_taxas = ["5.00", "7.00", "8.00", "10.00", "13.00", "15.00", "20.00", "25.00"]
        
        self.prefs = {
            "impressora": win32print.GetDefaultPrinter() if IMPRESSAO_DISPONIVEL else "SIMULADOR",
            "largura": "80mm", "tamanho_fonte": "NORMAL", "fonte_num": "0", "tema": "PRETO E VERMELHO", "bg_image": "",
            "tempos": {t: "60" for t in self.valores_taxas},
            "tempo_retirada": "20"
        }
        if os.path.exists(self.config_file):
            try: 
                with open(self.config_file, "r", encoding="utf-8") as f: self.prefs.update(json.load(f))
            except: pass
            
        self.original_bg_img = None
        bg_path = self.prefs.get("bg_image", "")
        if bg_path and os.path.exists(bg_path):
            try: self.original_bg_img = Image.open(bg_path)
            except: pass

    def iniciar_app(self):
        self.root.deiconify() 
        self.root.title(f"AURA SYSTEM - {self.nome_loja}")
        self.root.geometry("1200x850")
        
        self.carregar_configuracoes()
        tema = self.prefs.get("tema", "PRETO E VERMELHO")
        
        if tema == "VERMELHO E AMARELO":
            self.bg_dark, self.bg_card, self.bg_soft, self.accent, self.accent2, self.yellow, self.green, self.text_muted = "#1A0000", "#2A0606", "#3B0F0F", "#E74C3C", "#FF6B5A", "#F1C40F", "#2ECC71", "#9B7B7B"
        elif tema == "AZUL E VERDE":
            self.bg_dark, self.bg_card, self.bg_soft, self.accent, self.accent2, self.yellow, self.green, self.text_muted = "#0A1628", "#13243D", "#1B3251", "#2E86DE", "#54A0FF", "#2ECC71", "#27AE60", "#7F8FA6"
        elif tema == "LARANJA E VERMELHO":
            self.bg_dark, self.bg_card, self.bg_soft, self.accent, self.accent2, self.yellow, self.green, self.text_muted = "#150700", "#2A0E00", "#3D1500", "#FF6B1A", "#FF8C42", "#FFB84D", "#CE2029", "#9C8470"
        else:
            self.bg_dark, self.bg_card, self.bg_soft, self.accent, self.accent2, self.yellow, self.green, self.text_muted = "#0E0E10", "#181820", "#22222C", "#E63946", "#FF5A6B", "#F4B400", "#2ECC71", "#8A8A99"

        self.root.configure(bg=self.bg_dark)

        self.tabela_adicionais = {nome: preco for _id, nome, preco in self.db.get_adicionais()}

        self.carrinho = []
        self.total_atual = 0.0
        self.id_pedido_em_edicao = None 
        self.id_produto_em_edicao = None 
        
        self.cliente_padrao = {"nome": "CONSUMIDOR FINAL", "tel": "", "end": "", "bairro": "", "bloco": "", "apto": "", "taxa": 0}
        self.cliente_atual = self.cliente_padrao.copy()
        
        self.estilizar_ttks()
        self.montar_tela()

        self.root.bind("<Return>", self.atalho_enter)
        self.root.bind("<Button-1>", self.remover_selecoes_tabelas)

    def estilizar_ttks(self):
        self.font_family = "Segoe UI"
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background=self.bg_dark, borderwidth=0, tabmargins=[8, 12, 8, 0])
        style.configure("TNotebook.Tab", background=self.bg_card, foreground=self.text_muted,
                        padding=[26, 12], font=(self.font_family, 11, "bold"),
                        borderwidth=0, focuscolor=self.bg_dark)
        style.map("TNotebook.Tab",
                  background=[("selected", self.accent), ("active", self.bg_soft)],
                  foreground=[("selected", "#FFFFFF"), ("active", "#FFFFFF")])
        style.configure("Treeview", background=self.bg_card, foreground="white",
                        fieldbackground=self.bg_card, rowheight=38, borderwidth=0,
                        font=(self.font_family, 11))
        style.map("Treeview", background=[('selected', self.accent)], foreground=[('selected', '#FFFFFF')])
        style.configure("Treeview.Heading", background=self.bg_soft, foreground=self.yellow,
                        font=(self.font_family, 10, "bold"), borderwidth=0, relief="flat", padding=[6, 8])
        style.map("Treeview.Heading", background=[('active', self.accent)], foreground=[('active', '#FFFFFF')])
        style.configure("Cart.Treeview", background=self.bg_card, foreground="white",
                        fieldbackground=self.bg_card, rowheight=24, borderwidth=0,
                        font=(self.font_family, 10))
        style.map("Cart.Treeview", background=[('selected', self.accent)], foreground=[('selected', '#FFFFFF')])
        style.configure("Cart.Treeview.Heading", background=self.bg_soft, foreground=self.yellow,
                        font=(self.font_family, 9, "bold"), borderwidth=0, relief="flat", padding=[4, 5])
        style.map("Cart.Treeview.Heading", background=[('active', self.accent)], foreground=[('active', '#FFFFFF')])
        style.configure("TCombobox", selectbackground=self.accent, selectforeground="white",
                        borderwidth=0, arrowsize=15, padding=6)
        style.map("TCombobox", fieldbackground=[('readonly', self.bg_soft)],
                  foreground=[('readonly', 'white')], background=[('readonly', self.bg_soft)])
        style.configure("Vertical.TScrollbar", background=self.bg_soft, troughcolor=self.bg_dark,
                        bordercolor=self.bg_dark, arrowcolor=self.text_muted, gripcount=0)
        style.map("Vertical.TScrollbar", background=[('active', self.accent)])

    def remover_selecoes_tabelas(self, event):
        try:
            wid = event.widget
            if not isinstance(wid, tk.Button) and not isinstance(wid, ttk.Treeview) and not isinstance(wid, tk.Radiobutton):
                if hasattr(self, 'tree_cli') and self.tree_cli.selection(): self.tree_cli.selection_remove(self.tree_cli.selection())
                if hasattr(self, 'tree_ped') and self.tree_ped.selection(): self.tree_ped.selection_remove(self.tree_ped.selection())
                if hasattr(self, 'tree_prod') and self.tree_prod.selection(): self.tree_prod.selection_remove(self.tree_prod.selection())
                if hasattr(self, 'tree') and self.tree.selection(): self.tree.selection_remove(self.tree.selection())
        except Exception: pass

    def clique_vazio_tree(self, event, tree):
        if tree.identify_region(event.x, event.y) == "nothing":
            tree.selection_remove(tree.selection())

    def salvar_configuracoes(self):
        with open(self.config_file, "w", encoding="utf-8") as f: 
            json.dump(self.prefs, f, ensure_ascii=False)

    def montar_tela(self):
        top_bar = tk.Frame(self.root, bg=self.bg_card, height=84)
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        f_logo = tk.Frame(top_bar, bg=self.accent, padx=14, pady=8)
        f_logo.pack(side="left", padx=18, pady=10)
        tk.Label(f_logo, text="A", font=(self.font_family, 18, "bold"),
                 fg="white", bg=self.accent).pack()

        f_titulo = tk.Frame(top_bar, bg=self.bg_card)
        f_titulo.pack(side="left", padx=4)
        tk.Label(f_titulo, text="AURA SYSTEM", font=(self.font_family, 13, "bold"),
                 fg=self.yellow, bg=self.bg_card).pack(anchor="w")
        tk.Label(f_titulo, text=self.nome_loja, font=(self.font_family, 9),
                 fg=self.text_muted, bg=self.bg_card).pack(anchor="w")

        f_right = tk.Frame(top_bar, bg=self.bg_card)
        f_right.pack(side="right", padx=24, pady=10)

        f_right_top = tk.Frame(f_right, bg=self.bg_card)
        f_right_top.pack(side="top", anchor="e")
        f_status = tk.Frame(f_right_top, bg=self.bg_card)
        f_status.pack(side="right", padx=12)
        tk.Label(f_status, text="●", font=(self.font_family, 14, "bold"),
                 fg=self.green, bg=self.bg_card).pack(side="left")
        tk.Label(f_status, text="SISTEMA ONLINE", font=(self.font_family, 9, "bold"),
                 fg=self.text_muted, bg=self.bg_card).pack(side="left", padx=(4, 0))
        self.lbl_relogio = tk.Label(f_right_top, font=("Consolas", 13, "bold"),
                                    fg=self.yellow, bg=self.bg_card)
        self.lbl_relogio.pack(side="right", padx=(0, 12))
        self.atualizar_relogio()

        f_num_top = tk.Frame(f_right, bg=self.bg_card)
        f_num_top.pack(side="top", anchor="e", pady=(4, 0))
        tk.Label(f_num_top, text="PEDIDO Nº", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 8, "bold")).pack(side="left", padx=(0, 6))
        self.lbl_num_pedido = tk.Label(f_num_top, text=f"{self.obter_proximo_id()}",
                                       font=(self.font_family, 16, "bold"),
                                       fg=self.yellow, bg=self.bg_card)
        self.lbl_num_pedido.pack(side="left")

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True)
        self.aba_venda = tk.Frame(self.tabs, bg=self.bg_dark)
        self.aba_cad = tk.Frame(self.tabs, bg=self.bg_dark)
        self.aba_pedidos = tk.Frame(self.tabs, bg=self.bg_dark)
        self.aba_produtos = tk.Frame(self.tabs, bg=self.bg_dark)
        self.aba_finan = tk.Frame(self.tabs, bg=self.bg_dark)
        self.aba_config = tk.Frame(self.tabs, bg=self.bg_dark)

        self.tabs.add(self.aba_venda, text="  VENDAS  ")
        self.tabs.add(self.aba_cad, text="  CLIENTES  ")
        self.tabs.add(self.aba_pedidos, text="  PEDIDOS  ")
        self.tabs.add(self.aba_produtos, text="  PRODUTOS  ")
        self.tabs.add(self.aba_finan, text="  FINANCEIRO  ")
        self.tabs.add(self.aba_config, text="  CONFIGURAÇÕES  ")
        
        self.aba_finan.bind("<Visibility>", lambda e: self.bloqueio_aba(self.aba_finan, self.montar_financeiro))
        self.aba_config.bind("<Visibility>", lambda e: self.bloqueio_aba(self.aba_config, self.montar_configuracoes))
        
        self.montar_vendas()
        self.montar_aba_clientes()
        self.montar_aba_pedidos()
        self.montar_aba_produtos()

    def atualizar_relogio(self):
        agora = datetime.now().strftime("%d/%m/%Y - %H:%M:%S")
        if hasattr(self, 'lbl_relogio') and self.lbl_relogio.winfo_exists():
            self.lbl_relogio.config(text=agora)
        self.root.after(1000, self.atualizar_relogio)

    def bloqueio_aba(self, aba, metodo_montagem):
        for w in aba.winfo_children(): w.destroy()
        f_centro = tk.Frame(aba, bg=self.bg_dark)
        f_centro.pack(expand=True)

        cartao = tk.Frame(f_centro, bg=self.bg_card, padx=44, pady=34,
                          highlightthickness=1, highlightbackground=self.bg_soft)
        cartao.pack()
        tk.Frame(cartao, bg=self.accent, height=3).pack(fill="x", pady=(0, 16))

        tk.Label(cartao, text="ACESSO RESTRITO", font=(self.font_family, 18, "bold"),
                 fg=self.yellow, bg=self.bg_card).pack(pady=(0, 6))
        tk.Label(cartao, text="Digite a senha para acessar esta área.",
                 font=(self.font_family, 10), fg=self.text_muted, bg=self.bg_card).pack(pady=(0, 18))

        ent_senha = tk.Entry(cartao, show="*", font=(self.font_family, 16), width=22, justify="center",
                             bg=self.bg_soft, fg="white", relief="flat", highlightthickness=1,
                             highlightbackground=self.bg_soft, highlightcolor=self.accent,
                             insertbackground="white")
        ent_senha.pack(pady=(0, 18), ipady=8)
        ent_senha.focus_set()

        def verificar_senha(e=None):
            if ent_senha.get() == self.senha_finan:
                for w in aba.winfo_children(): w.destroy()
                metodo_montagem()
            else:
                messagebox.showerror("Acesso Negado", "Senha incorreta!")
                self.tabs.select(0)

        tk.Button(cartao, text="ENTRAR", bg=self.green, fg="white",
                  font=(self.font_family, 11, "bold"), command=verificar_senha,
                  relief="flat", cursor="hand2", bd=0, width=18, height=2).pack(pady=(0, 4))
        ent_senha.bind("<Return>", verificar_senha)

        tk.Label(aba, text="Developed by Adolfo Rodrigues Novaes",
                 font=(self.font_family, 10, "italic"), fg=self.text_muted,
                 bg=self.bg_dark).pack(side="bottom", pady=20)

    def montar_vendas(self):
        f_topo = tk.Frame(self.aba_venda, bg=self.bg_dark, pady=12)
        f_topo.pack(fill="x", padx=14)

        self.lbl_busca_venda = tk.Label(f_topo, text="NOME", bg=self.bg_dark, fg=self.text_muted,
                                        font=(self.font_family, 9, "bold"))
        self.lbl_busca_venda.pack(side="left", padx=(0, 8))

        self.ent_busca_venda = tk.Entry(f_topo, font=(self.font_family, 13), width=22,
                                        bg=self.bg_soft, fg="white", insertbackground="white",
                                        relief="flat", highlightthickness=1,
                                        highlightbackground=self.bg_soft, highlightcolor=self.accent)
        self.ent_busca_venda.pack(side="left", padx=4, ipady=7)
        tk.Button(f_topo, text="BUSCAR", command=self.buscar_cliente_venda, bg=self.accent, fg="white",
                  font=(self.font_family, 10, "bold"), relief="flat", cursor="hand2", bd=0,
                  padx=16, pady=7).pack(side="left", padx=(4, 14))

        self.tipo_var = tk.StringVar(value="RETIRADA")
        f_radios = tk.Frame(f_topo, bg=self.bg_dark)
        f_radios.pack(side="left", padx=8)
        for tipo in ["RETIRADA", "ENTREGA", "MESA"]:
            rb = tk.Radiobutton(f_radios, text=tipo, variable=self.tipo_var, value=tipo,
                                command=self.atualizar_taxa_carrinho, bg=self.bg_soft, fg="white",
                                selectcolor=self.accent, activebackground=self.accent,
                                activeforeground="white", indicatoron=0,
                                font=(self.font_family, 10, "bold"), width=10, bd=0, cursor="hand2")
            rb.pack(side="left", padx=2, ipady=5)

        f_corpo = tk.Frame(self.aba_venda, bg=self.bg_dark)
        f_corpo.pack(fill="both", expand=True, padx=10)
        
        f_grade_container = tk.Frame(f_corpo, bg=self.bg_dark)
        f_grade_container.pack(side="left", fill="both", expand=True)

        self.canvas_grade = tk.Canvas(f_grade_container, bg=self.bg_dark, highlightthickness=0)
        self.scroll_grade = ttk.Scrollbar(f_grade_container, orient="vertical", command=self.canvas_grade.yview)
        self.canvas_grade.configure(yscrollcommand=self.scroll_grade.set)
        self.scroll_grade.pack(side="right", fill="y")
        self.canvas_grade.pack(side="left", fill="both", expand=True)

        self.canvas_grade.bind('<Enter>', lambda e: self.canvas_grade.bind_all("<MouseWheel>", lambda e2: self.canvas_grade.yview_scroll(int(-1*(e2.delta/120)), "units")))
        self.canvas_grade.bind('<Leave>', lambda e: self.canvas_grade.unbind_all("<MouseWheel>"))

        f_dir = tk.Frame(f_corpo, bg=self.bg_card, width=380,
                         highlightthickness=1, highlightbackground=self.bg_soft)
        f_dir.pack(side="right", fill="y", padx=(10, 0))
        f_dir.pack_propagate(False)

        tk.Frame(f_dir, bg=self.accent, height=3).pack(fill="x", side="top")

        f_bottom = tk.Frame(f_dir, bg=self.bg_card)
        f_bottom.pack(side="bottom", fill="x")

        tk.Button(f_bottom, text="FINALIZAR PEDIDO", bg=self.green, fg="white",
                  font=(self.font_family, 12, "bold"), height=2,
                  command=self.abrir_janela_pagamento, relief="flat", cursor="hand2",
                  bd=0).pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        tk.Button(f_bottom, text="EXCLUIR SELECIONADOS", bg=self.accent, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.remover_item,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="bottom", fill="x", padx=10, pady=(2, 4))

        f_total_box = tk.Frame(f_bottom, bg=self.bg_soft, padx=12, pady=8)
        f_total_box.pack(side="bottom", fill="x", padx=10, pady=(2, 6))
        tk.Label(f_total_box, text="TOTAL DO PEDIDO", bg=self.bg_soft, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w")
        self.lbl_total_grande = tk.Label(f_total_box, text="R$ 0,00",
                                         font=(self.font_family, 28, "bold"),
                                         fg=self.yellow, bg=self.bg_soft)
        self.lbl_total_grande.pack(anchor="w")

        self.f_info_cli = tk.Frame(f_dir, bg=self.bg_soft, pady=8, padx=10)
        self.f_info_cli.pack(side="top", fill="x", padx=8, pady=(8, 4))
        self.lbl_info_nome = tk.Label(self.f_info_cli, text="CONSUMIDOR FINAL",
                                      bg=self.bg_soft, fg=self.yellow,
                                      font=(self.font_family, 11, "bold"), anchor="w")
        self.lbl_info_nome.pack(fill="x")
        self.lbl_info_tel = tk.Label(self.f_info_cli, text="—",
                                     bg=self.bg_soft, fg="white",
                                     font=(self.font_family, 9), anchor="w")
        self.lbl_info_tel.pack(fill="x", pady=(1, 0))
        self.lbl_info_end = tk.Label(self.f_info_cli, text="",
                                     bg=self.bg_soft, fg=self.text_muted,
                                     font=(self.font_family, 9), anchor="w",
                                     wraplength=340, justify="left")
        self.lbl_info_end.pack(fill="x", pady=(1, 2))

        f_alt_end = tk.Frame(self.f_info_cli, bg=self.bg_soft)
        f_alt_end.pack(fill="x", pady=(4, 0))
        tk.Label(f_alt_end, text="END. ALT", bg=self.bg_soft, fg=self.text_muted,
                 font=(self.font_family, 8, "bold"), width=9, anchor="w").pack(side="left")
        self.ent_end_alt = tk.Entry(f_alt_end, font=(self.font_family, 10),
                                    bg=self.bg_dark, fg="white", insertbackground="white",
                                    relief="flat", highlightthickness=1,
                                    highlightbackground=self.bg_dark, highlightcolor=self.accent)
        self.ent_end_alt.pack(side="left", fill="x", expand=True, ipady=3)
        self.ent_end_alt.bind("<KeyRelease>", lambda e: self._atualizar_estado_taxa_alt())

        tk.Label(self.f_info_cli, text="TAXA DO END. ALT (R$)", bg=self.bg_soft, fg=self.text_muted,
                 font=(self.font_family, 8, "bold"), anchor="w").pack(fill="x", pady=(6, 2))
        self.var_taxa_alt = tk.StringVar(value="")
        self.rb_taxa_alt = []
        f_alt_taxas = tk.Frame(self.f_info_cli, bg=self.bg_soft)
        f_alt_taxas.pack(fill="x")
        for v in self.valores_taxas:
            rb = tk.Radiobutton(f_alt_taxas, text=v, variable=self.var_taxa_alt, value=v,
                                command=self.atualizar_taxa_carrinho,
                                font=(self.font_family, 8, "bold"), bg=self.bg_dark, fg="white",
                                selectcolor=self.accent, activebackground=self.accent,
                                activeforeground="white", indicatoron=0, cursor="hand2",
                                bd=0, state="disabled", disabledforeground=self.text_muted)
            rb.pack(side="left", fill="x", expand=True, padx=1, ipady=2)
            self.rb_taxa_alt.append(rb)

        f_tree_venda = tk.Frame(f_dir, bg=self.bg_card)
        f_tree_venda.pack(side="top", fill="both", expand=True, padx=8, pady=(6, 0))

        self.tree = ttk.Treeview(f_tree_venda, columns=("n", "q", "v"), show="headings",
                                 selectmode="extended", style="Cart.Treeview")
        self.tree.heading("n", text="ITEM"); self.tree.heading("q", text="QTD"); self.tree.heading("v", text="TOTAL")
        self.tree.column("n", width=200); self.tree.column("q", width=40, anchor="center"); self.tree.column("v", width=80, anchor="center")
        sb_cart = ttk.Scrollbar(f_tree_venda, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb_cart.set)
        sb_cart.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<Double-1>", self.editar_item_carrinho)
        self.tree.bind("<Delete>", self.remover_item)
        self.tree.bind("<Button-1>", lambda e: self.clique_vazio_tree(e, self.tree))
        
        self.canvas_grade.bind("<Configure>", self.on_canvas_resize)
        self.atualizar_labels_info()

    def on_canvas_resize(self, event):
        if self._resize_timer: self.root.after_cancel(self._resize_timer)
        self._resize_timer = self.root.after(200, lambda: self.atualizar_grade_produtos(event.width, event.height))

    def atualizar_labels_info(self):
        self.lbl_info_nome.config(text=self.cliente_atual.get('nome', 'CONSUMIDOR FINAL'))
        tel = self.cliente_atual.get('tel', '') or "—"
        self.lbl_info_tel.config(text=f"TEL: {tel}")
        end_str = self.cliente_atual.get('end', '')
        bairro = self.cliente_atual.get('bairro', '')
        if bairro: end_str += f" - {bairro}"
        comp = []
        for k in ['bloco', 'apto']:
            val = self.cliente_atual.get(k, '')
            if val: comp.append(f"{k.upper()}: {val}")
        if comp: end_str += " | " + " ".join(comp)
        self.lbl_info_end.config(text=end_str if end_str.strip() else "Sem endereço cadastrado")

    def _atualizar_estado_taxa_alt(self):
        tem_alt = bool(self.ent_end_alt.get().strip())
        novo_estado = "normal" if tem_alt else "disabled"
        for rb in getattr(self, "rb_taxa_alt", []):
            rb.config(state=novo_estado)
        if not tem_alt and self.var_taxa_alt.get():
            self.var_taxa_alt.set("")
            self.atualizar_taxa_carrinho()

    def atualizar_taxa_carrinho(self):
        if hasattr(self, 'lbl_busca_venda'):
            if self.tipo_var.get() == "ENTREGA": self.lbl_busca_venda.config(text="TELEFONE:")
            else: self.lbl_busca_venda.config(text="NOME:")

        self.carrinho = [i for i in self.carrinho if "TAXA DE ENTREGA" not in i["nome"]]
        if self.tipo_var.get() == "ENTREGA":
            t = 0.0
            tem_alt = hasattr(self, "ent_end_alt") and bool(self.ent_end_alt.get().strip())
            taxa_alt = getattr(self, "var_taxa_alt", None)
            if tem_alt and taxa_alt and taxa_alt.get():
                try: t = float(taxa_alt.get().replace(",", "."))
                except: t = 0.0
            else:
                try: t = float(self.cliente_atual.get("taxa", 0.0))
                except: t = 0.0
            if t > 0: self.carrinho.append({"nome": "TAXA DE ENTREGA", "preco": t, "qtd": 1, "obs": ""})
        self.sincronizar_treeview()
        self.atualizar_total()

    def cor_categoria(self, categoria):
        mapa = {
            "LANCHES": self.accent,
            "REFRIGERANTES": self.accent2,
            "PORÇÕES": self.yellow,
            "OUTROS": self.green,
        }
        return mapa.get(categoria, self.accent)

    def atualizar_grade_produtos(self, cw=None, ch=None):
        self.canvas_grade.delete("all")
        self.db.cursor.execute("SELECT * FROM produtos ORDER BY categoria, nome")
        produtos = self.db.cursor.fetchall()

        if not produtos: return

        if cw is None: cw = self.canvas_grade.winfo_width()
        if ch is None: ch = self.canvas_grade.winfo_height()
        if cw < 100: cw = 750

        cols = 7
        pad_x, pad_y = 6, 6
        card_w = max(80, (cw - (pad_x * (cols + 1))) // cols)
        card_h = 64
        linhas = (len(produtos) + cols - 1) // cols
        max_y = linhas * (card_h + pad_y) + pad_y + 30

        if self.original_bg_img and cw > 10 and ch > 10:
            bg_h = max(ch, max_y)
            try:
                resized = self.original_bg_img.resize((cw, bg_h), Image.LANCZOS)
                self.bg_photo = ImageTk.PhotoImage(resized)
                self.canvas_grade.create_image(0, 0, image=self.bg_photo, anchor="nw", tags="bg")
            except: pass

        for i, p in enumerate(produtos):
            categoria = p[3] if len(p) > 3 else "LANCHES"
            cor_cat = self.cor_categoria(categoria)

            f = tk.Frame(self.canvas_grade, bg=self.bg_card, bd=0,
                         highlightthickness=1, highlightbackground=self.bg_soft, cursor="hand2")

            barra_top = tk.Frame(f, bg=cor_cat, height=3)
            barra_top.pack(fill="x", side="top")

            corpo_card = tk.Frame(f, bg=self.bg_card)
            corpo_card.pack(fill="both", expand=True)

            lbl_nome = tk.Label(corpo_card, text=p[1], bg=self.bg_card, fg="white",
                                font=(self.font_family, 8, "bold"),
                                wraplength=card_w-10, justify="center", cursor="hand2")
            lbl_nome.pack(expand=True, fill="both", pady=(4, 0), padx=4)
            lbl_preco = tk.Label(corpo_card, text=f"R$ {p[2]:.2f}", bg=self.bg_card,
                                 fg=self.yellow, font=(self.font_family, 10, "bold"),
                                 cursor="hand2")
            lbl_preco.pack(side="bottom", pady=(0, 4))

            def hover_in(e, frm=f, body=corpo_card, n=lbl_nome, pr=lbl_preco, c=cor_cat):
                frm.config(highlightbackground=c)
                body.config(bg=self.bg_soft)
                n.config(bg=self.bg_soft); pr.config(bg=self.bg_soft, fg="white")
            def hover_out(e, frm=f, body=corpo_card, n=lbl_nome, pr=lbl_preco):
                frm.config(highlightbackground=self.bg_soft)
                body.config(bg=self.bg_card)
                n.config(bg=self.bg_card); pr.config(bg=self.bg_card, fg=self.yellow)

            f.bind("<Enter>", hover_in); f.bind("<Leave>", hover_out)
            for w in (f, corpo_card, lbl_nome, lbl_preco):
                w.bind("<Button-1>", lambda e, x=p: self.modal_item(x))

            x = (i % cols) * (card_w + pad_x) + pad_x
            y = (i // cols) * (card_h + pad_y) + pad_y
            self.canvas_grade.create_window(x, y, window=f, width=card_w, height=card_h, anchor="nw", tags="card")

        self.canvas_grade.configure(scrollregion=(0, 0, cw, max_y))

    def modal_item(self, p, edit_idx=None):
        self.tabela_adicionais = {nome: preco for _id, nome, preco in self.db.get_adicionais()}

        win = tk.Toplevel(self.root)
        win.title(f"OPÇÕES: {p[1]}")
        win.attributes("-toolwindow", True)
        win.grab_set()
        win.configure(bg=self.bg_dark)

        categoria = p[3] if len(p) > 3 else "LANCHES"
        altura = 300 if categoria == "REFRIGERANTES" else 560
        self.centralizar_janela_toplevel(win, 440, altura)

        cor_cat = self.cor_categoria(categoria)
        tk.Frame(win, bg=cor_cat, height=4).pack(fill="x")

        cabe = tk.Frame(win, bg=self.bg_dark)
        cabe.pack(fill="x", pady=(12, 6), padx=18)
        tk.Label(cabe, text=p[1], font=(self.font_family, 16, "bold"),
                 fg=self.yellow, bg=self.bg_dark).pack(anchor="w")
        tk.Label(cabe, text=f"R$ {p[2]:.2f}  ·  {categoria}",
                 font=(self.font_family, 10), fg=self.text_muted, bg=self.bg_dark).pack(anchor="w")
        nome_base_produto = p[1]

        f_rodape = tk.Frame(win, bg=self.bg_dark)
        f_rodape.pack(side="bottom", fill="x", padx=18, pady=(8, 14))
        btn_confirmar = tk.Button(f_rodape, text="ADICIONAR AO PEDIDO",
                                  bg=self.green, fg="white",
                                  font=(self.font_family, 12, "bold"),
                                  relief="flat", cursor="hand2", bd=0, height=2)
        btn_confirmar.pack(fill="x")

        f_scroll_outer = tk.Frame(win, bg=self.bg_dark)
        f_scroll_outer.pack(side="top", fill="both", expand=True, padx=14, pady=(4, 0))

        canvas_modal = tk.Canvas(f_scroll_outer, bg=self.bg_dark, highlightthickness=0)
        sb_modal = ttk.Scrollbar(f_scroll_outer, orient="vertical", command=canvas_modal.yview)
        canvas_modal.configure(yscrollcommand=sb_modal.set)
        sb_modal.pack(side="right", fill="y")
        canvas_modal.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas_modal, bg=self.bg_dark)
        body_id = canvas_modal.create_window((0, 0), window=body, anchor="nw")

        def _on_body_cfg(e):
            canvas_modal.configure(scrollregion=canvas_modal.bbox("all"))
        def _on_canvas_cfg(e):
            canvas_modal.itemconfig(body_id, width=e.width)
        body.bind("<Configure>", _on_body_cfg)
        canvas_modal.bind("<Configure>", _on_canvas_cfg)
        canvas_modal.bind("<Enter>", lambda e: canvas_modal.bind_all("<MouseWheel>",
                          lambda e2: canvas_modal.yview_scroll(int(-1*(e2.delta/120)), "units")))
        canvas_modal.bind("<Leave>", lambda e: canvas_modal.unbind_all("<MouseWheel>"))

        f_qtd = tk.Frame(body, bg=self.bg_card, padx=14, pady=10)
        f_qtd.pack(pady=(4, 8), padx=4, fill="x")
        tk.Label(f_qtd, text="QUANTIDADE", fg=self.text_muted, bg=self.bg_card,
                 font=(self.font_family, 9, "bold")).pack(side="left", padx=(0, 14))
        q_var = tk.IntVar(value=1)
        if edit_idx is not None: q_var.set(self.carrinho[edit_idx].get("qtd", 1))

        tk.Button(f_qtd, text="−", font=(self.font_family, 14, "bold"), width=3,
                  command=lambda: q_var.set(max(1, q_var.get()-1)), relief="flat",
                  cursor="hand2", bg=self.bg_soft, fg="white", bd=0).pack(side="left", padx=4)
        tk.Label(f_qtd, textvariable=q_var, fg=self.yellow, bg=self.bg_card,
                 font=(self.font_family, 18, "bold"), width=3).pack(side="left", padx=8)
        tk.Button(f_qtd, text="+", font=(self.font_family, 14, "bold"), width=3,
                  command=lambda: q_var.set(q_var.get()+1), relief="flat",
                  cursor="hand2", bg=self.bg_soft, fg="white", bd=0).pack(side="left", padx=4)

        valores_iniciais = {}
        if edit_idx is not None:
            nome_completo = self.carrinho[edit_idx]["nome"]
            if " (C/ " in nome_completo:
                nome_base_produto = nome_completo.split(" (C/ ")[0]
                parte_adds = nome_completo.split("(C/ ")[1].replace(")", "")
                for pedaco in [i.strip() for i in parte_adds.split(",")]:
                    if "x " in pedaco:
                        q_add, n_add = pedaco.split("x ", 1)
                        valores_iniciais[n_add.upper()] = int(q_add)
                    else: valores_iniciais[pedaco.upper()] = 1

        vars_adicionais = {}
        if categoria != "REFRIGERANTES" and self.tabela_adicionais:
            tk.Label(body, text="ADICIONAIS", fg=self.text_muted, bg=self.bg_dark,
                     font=(self.font_family, 9, "bold")).pack(anchor="w", padx=4, pady=(6, 2))
            f_meio = tk.Frame(body, bg=self.bg_card)
            f_meio.pack(fill="x", padx=4, pady=(0, 6))
            for nome, preco in self.tabela_adicionais.items():
                f = tk.Frame(f_meio, bg=self.bg_card)
                f.pack(fill="x", pady=1, padx=8, ipady=2)
                tk.Label(f, text=nome, fg="white", bg=self.bg_card,
                         font=(self.font_family, 10)).pack(side="left", padx=(4, 0))
                tk.Label(f, text=f"+R$ {preco:.2f}", fg=self.yellow, bg=self.bg_card,
                         font=(self.font_family, 9)).pack(side="left", padx=8)
                var = tk.IntVar(value=valores_iniciais.get(nome.upper(), 0))
                vars_adicionais[nome] = var
                tk.Button(f, text="+", width=3, command=lambda v=var: v.set(v.get()+1),
                          relief="flat", cursor="hand2", bg=self.bg_soft, fg="white", bd=0).pack(side="right", padx=1)
                tk.Label(f, textvariable=var, width=3, fg=self.yellow, bg=self.bg_card,
                         font=(self.font_family, 10, "bold")).pack(side="right", padx=4)
                tk.Button(f, text="−", width=3, command=lambda v=var: v.set(max(0, v.get()-1)),
                          relief="flat", cursor="hand2", bg=self.bg_soft, fg="white", bd=0).pack(side="right", padx=1)

        tk.Label(body, text="OBSERVAÇÕES", fg=self.text_muted, bg=self.bg_dark,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", padx=4, pady=(6, 2))
        obs_t = tk.Text(body, height=2, relief="flat", highlightthickness=1,
                        highlightbackground=self.bg_soft, highlightcolor=self.accent,
                        bg=self.bg_soft, fg="white", insertbackground="white",
                        font=(self.font_family, 11))
        obs_t.pack(padx=4, pady=(0, 8), fill="x", ipady=4)
        if edit_idx is not None: obs_t.insert("1.0", self.carrinho[edit_idx].get("obs", ""))
            
        def conf():
            lista_nomes = []
            preco_adds = 0
            for nome, var in vars_adicionais.items():
                if var.get() > 0: 
                    lista_nomes.append(f"{f'{var.get()}x ' if var.get()>1 else ''}{nome}")
                    preco_adds += (self.tabela_adicionais[nome] * var.get())
            res = self.db.cursor.execute("SELECT preco FROM produtos WHERE nome=?", (nome_base_produto,)).fetchone()
            preco_base = res[0] if res else p[2]
            item_data = {"nome": nome_base_produto + (f" (C/ {', '.join(lista_nomes)})" if lista_nomes else ""), "preco": preco_base + preco_adds, "qtd": q_var.get(), "obs": obs_t.get("1.0", tk.END).strip().upper()}
            if edit_idx is not None: self.carrinho[edit_idx] = item_data
            else: self.carrinho.append(item_data)
            self.sincronizar_treeview(); self.atualizar_total(); win.destroy()

        btn_confirmar.config(command=conf)
        win.bind("<Return>", lambda e: conf())

    def montar_aba_produtos(self):
        f_cad = tk.Frame(self.aba_produtos, bg=self.bg_dark)
        f_cad.pack(side="left", fill="both", expand=True, padx=24, pady=24)

        tk.Label(f_cad, text="CADASTRAR / EDITAR PRODUTO", font=(self.font_family, 15, "bold"),
                 bg=self.bg_dark, fg=self.yellow).pack(anchor="w", pady=(0, 4))
        tk.Label(f_cad, text="Adicione, atualize e remova itens do cardápio.",
                 bg=self.bg_dark, fg=self.text_muted, font=(self.font_family, 9)).pack(anchor="w", pady=(0, 18))

        tk.Label(f_cad, text="NOME DO PRODUTO", bg=self.bg_dark, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(5, 2))
        self.ent_nome_prod = self._input_premium(f_cad, font=(self.font_family, 12))
        self.ent_nome_prod.pack(fill="x", ipady=4)

        tk.Label(f_cad, text="PREÇO (R$)", bg=self.bg_dark, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(14, 2))
        self.ent_preco_prod = self._input_premium(f_cad, font=(self.font_family, 12))
        self.ent_preco_prod.pack(fill="x", ipady=4)

        tk.Label(f_cad, text="CATEGORIA", bg=self.bg_dark, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(14, 2))
        f_cat_radios = tk.Frame(f_cad, bg=self.bg_dark)
        f_cat_radios.pack(fill="x")
        self.var_cat_prod = tk.StringVar(value="LANCHES")
        for c in ["LANCHES", "REFRIGERANTES", "PORÇÕES", "OUTROS"]:
            tk.Radiobutton(f_cat_radios, text=c, variable=self.var_cat_prod, value=c,
                           font=(self.font_family, 9, "bold"), bg=self.bg_soft, fg="white",
                           selectcolor=self.accent, activebackground=self.accent,
                           activeforeground="white", indicatoron=0, cursor="hand2",
                           bd=0).pack(side="left", fill="x", expand=True, padx=2, ipady=4)

        f_btns = tk.Frame(f_cad, bg=self.bg_dark)
        f_btns.pack(pady=24, fill="x")
        self.btn_salvar_prod = tk.Button(f_btns, text="SALVAR NOVO", bg=self.green, fg="white",
                                         font=(self.font_family, 11, "bold"), command=self.salvar_produto,
                                         relief="flat", cursor="hand2", bd=0, height=2)
        self.btn_salvar_prod.pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(f_btns, text="LIMPAR", bg=self.bg_soft, fg="white",
                  font=(self.font_family, 11, "bold"), command=self.limpar_form_produto,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", fill="x", expand=True, padx=4)

        f_lista = tk.Frame(self.aba_produtos, bg=self.bg_card, width=520)
        f_lista.pack(side="left", fill="both", expand=True, padx=10, pady=24)

        cabe = tk.Frame(f_lista, bg=self.bg_card)
        cabe.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(cabe, text="CARDÁPIO DO SISTEMA", bg=self.bg_card, fg="white",
                 font=(self.font_family, 12, "bold")).pack(anchor="w")
        tk.Label(cabe, text="Toque duplo para editar.", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9)).pack(anchor="w")

        f_btns_lista = tk.Frame(f_lista, bg=self.bg_card)
        f_btns_lista.pack(side="bottom", pady=(0, 12), fill="x", padx=12)
        tk.Button(f_btns_lista, text="EDITAR", bg=self.accent2, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.carregar_produto_edicao,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", padx=4, expand=True, fill="x")
        tk.Button(f_btns_lista, text="EXCLUIR", bg=self.accent, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.excluir_produto,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="right", padx=4, expand=True, fill="x")

        f_tree_prod = tk.Frame(f_lista, bg=self.bg_card)
        f_tree_prod.pack(side="bottom", fill="both", expand=True, padx=12, pady=8)
        self.tree_prod = ttk.Treeview(f_tree_prod, columns=("id", "n", "p", "c"), show="headings", selectmode="extended")
        for col, hd in zip(("id", "n", "p", "c"), ("ID", "NOME", "PREÇO", "CATEGORIA")):
            self.tree_prod.heading(col, text=hd)
        self.tree_prod.column("id", width=50, anchor="center")
        self.tree_prod.column("n", width=240)
        self.tree_prod.column("p", width=90, anchor="center")
        self.tree_prod.column("c", width=130, anchor="center")
        sb_prod = ttk.Scrollbar(f_tree_prod, orient="vertical", command=self.tree_prod.yview)
        self.tree_prod.configure(yscrollcommand=sb_prod.set)
        sb_prod.pack(side="right", fill="y")
        self.tree_prod.pack(side="left", fill="both", expand=True)
        self.tree_prod.bind("<Button-1>", lambda e: self.clique_vazio_tree(e, self.tree_prod))
        self.tree_prod.bind("<Double-1>", lambda e: self.carregar_produto_edicao())

        f_add = tk.Frame(self.aba_produtos, bg=self.bg_card, width=360)
        f_add.pack(side="right", fill="both", padx=(10, 24), pady=24)
        f_add.pack_propagate(False)

        barra = tk.Frame(f_add, bg=self.yellow, height=3)
        barra.pack(fill="x")

        cabe2 = tk.Frame(f_add, bg=self.bg_card)
        cabe2.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(cabe2, text="ADICIONAIS DOS LANCHES", bg=self.bg_card, fg=self.yellow,
                 font=(self.font_family, 12, "bold")).pack(anchor="w")
        tk.Label(cabe2, text="Itens extras que aparecem ao montar um pedido.",
                 bg=self.bg_card, fg=self.text_muted, font=(self.font_family, 9)).pack(anchor="w")

        f_form_add = tk.Frame(f_add, bg=self.bg_card)
        f_form_add.pack(fill="x", padx=18, pady=10)

        tk.Label(f_form_add, text="NOME", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(4, 2))
        self.ent_nome_add = self._input_premium(f_form_add, font=(self.font_family, 11))
        self.ent_nome_add.pack(fill="x", ipady=3)

        tk.Label(f_form_add, text="PREÇO (R$)", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(10, 2))
        self.ent_preco_add = self._input_premium(f_form_add, font=(self.font_family, 11))
        self.ent_preco_add.pack(fill="x", ipady=3)

        f_btn_add = tk.Frame(f_form_add, bg=self.bg_card)
        f_btn_add.pack(fill="x", pady=(12, 4))
        self.btn_salvar_add = tk.Button(f_btn_add, text="SALVAR ADICIONAL", bg=self.green, fg="white",
                                        font=(self.font_family, 10, "bold"), command=self.salvar_adicional,
                                        relief="flat", cursor="hand2", bd=0, height=2)
        self.btn_salvar_add.pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(f_btn_add, text="LIMPAR", bg=self.bg_soft, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.limpar_form_adicional,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", padx=2)

        f_btns_add = tk.Frame(f_add, bg=self.bg_card)
        f_btns_add.pack(side="bottom", fill="x", padx=18, pady=(0, 14))
        tk.Button(f_btns_add, text="EDITAR", bg=self.accent2, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.carregar_adicional_edicao,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", padx=2, fill="x", expand=True)
        tk.Button(f_btns_add, text="EXCLUIR", bg=self.accent, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.excluir_adicional,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="right", padx=2, fill="x", expand=True)

        f_tree_add = tk.Frame(f_add, bg=self.bg_card)
        f_tree_add.pack(side="bottom", fill="both", expand=True, padx=18, pady=(8, 8))

        self.tree_add = ttk.Treeview(f_tree_add, columns=("id", "nome", "preco"), show="headings", selectmode="browse")
        self.tree_add.heading("id", text="ID")
        self.tree_add.heading("nome", text="NOME")
        self.tree_add.heading("preco", text="PREÇO")
        self.tree_add.column("id", width=40, anchor="center")
        self.tree_add.column("nome", width=180)
        self.tree_add.column("preco", width=80, anchor="center")
        sb_add = ttk.Scrollbar(f_tree_add, orient="vertical", command=self.tree_add.yview)
        self.tree_add.configure(yscrollcommand=sb_add.set)
        sb_add.pack(side="right", fill="y")
        self.tree_add.pack(side="left", fill="both", expand=True)
        self.tree_add.bind("<Double-1>", lambda e: self.carregar_adicional_edicao())

        self.id_adicional_em_edicao = None
        self.atualizar_lista_produtos_aba()
        self.atualizar_lista_adicionais()

    def limpar_form_produto(self):
        self.ent_nome_prod.delete(0, tk.END); self.ent_preco_prod.delete(0, tk.END); self.var_cat_prod.set("LANCHES")
        self.id_produto_em_edicao = None; self.btn_salvar_prod.config(text="SALVAR NOVO", bg=self.green)

    def carregar_produto_edicao(self):
        sel = self.tree_prod.selection()
        if len(sel) == 1:
            pid = self.tree_prod.item(sel[0])['values'][0]
            r = self.db.cursor.execute("SELECT id, nome, preco, categoria FROM produtos WHERE id=?", (pid,)).fetchone()
            if r:
                self.limpar_form_produto(); self.id_produto_em_edicao = r[0]
                self.ent_nome_prod.insert(0, r[1]); self.ent_preco_prod.insert(0, f"{r[2]:.2f}".replace('.', ','))
                self.var_cat_prod.set(r[3]); self.btn_salvar_prod.config(text="ATUALIZAR ITEM", bg=self.accent2)

    def salvar_produto(self):
        nome = self.ent_nome_prod.get().upper().strip(); preco_str = self.ent_preco_prod.get().replace(',', '.'); categoria = self.var_cat_prod.get()
        if not nome or not preco_str: return messagebox.showwarning("Aviso", "Preencha Nome e Preço!")
        try:
            preco = float(preco_str)
            if self.id_produto_em_edicao: self.db.cursor.execute("UPDATE produtos SET nome=?, preco=?, categoria=? WHERE id=?", (nome, preco, categoria, self.id_produto_em_edicao)); msg = "Produto atualizado!"
            else: self.db.cursor.execute("INSERT INTO produtos (nome, preco, categoria) VALUES (?, ?, ?)", (nome, preco, categoria)); msg = "Produto adicionado!"
            self.db.conn.commit(); messagebox.showinfo("Sucesso", msg); self.limpar_form_produto(); self.atualizar_lista_produtos_aba(); self.atualizar_grade_produtos()
        except ValueError: messagebox.showerror("Erro", "Valor de preço inválido.")

    def atualizar_lista_produtos_aba(self):
        for i in self.tree_prod.get_children(): self.tree_prod.delete(i)
        for r in self.db.cursor.execute("SELECT id, nome, preco, categoria FROM produtos ORDER BY categoria, nome").fetchall():
            self.tree_prod.insert("", "end", values=(r[0], r[1], f"R$ {r[2]:.2f}", r[3]))

    def excluir_produto(self):
        sel = self.tree_prod.selection()
        if sel and messagebox.askyesno("AVISO", "Excluir os itens selecionados do cardápio?"):
            for item in sel: self.db.cursor.execute("DELETE FROM produtos WHERE id=?", (self.tree_prod.item(item)['values'][0],))
            self.db.conn.commit(); self.limpar_form_produto(); self.atualizar_lista_produtos_aba(); self.atualizar_grade_produtos()

    def atualizar_lista_adicionais(self):
        if not hasattr(self, "tree_add"): return
        for i in self.tree_add.get_children(): self.tree_add.delete(i)
        for r in self.db.get_adicionais():
            self.tree_add.insert("", "end", values=(r[0], r[1], f"R$ {r[2]:.2f}"))
        self.tabela_adicionais = {nome: preco for _id, nome, preco in self.db.get_adicionais()}

    def limpar_form_adicional(self):
        self.ent_nome_add.delete(0, tk.END)
        self.ent_preco_add.delete(0, tk.END)
        self.id_adicional_em_edicao = None
        self.btn_salvar_add.config(text="SALVAR ADICIONAL", bg=self.green)

    def carregar_adicional_edicao(self):
        sel = self.tree_add.selection()
        if len(sel) == 1:
            ad_id = self.tree_add.item(sel[0])['values'][0]
            r = self.db.cursor.execute("SELECT id, nome, preco FROM adicionais WHERE id=?", (ad_id,)).fetchone()
            if r:
                self.limpar_form_adicional()
                self.id_adicional_em_edicao = r[0]
                self.ent_nome_add.insert(0, r[1])
                self.ent_preco_add.insert(0, f"{r[2]:.2f}".replace('.', ','))
                self.btn_salvar_add.config(text="ATUALIZAR ADICIONAL", bg=self.accent2)

    def salvar_adicional(self):
        nome = self.ent_nome_add.get().strip().upper()
        preco_str = self.ent_preco_add.get().replace(',', '.').strip()
        if not nome or not preco_str:
            return messagebox.showwarning("Aviso", "Preencha o nome e o preço do adicional.")
        try:
            preco = float(preco_str)
        except ValueError:
            return messagebox.showerror("Erro", "Preço inválido.")
        try:
            self.db.upsert_adicional(nome, preco, self.id_adicional_em_edicao)
            msg = "Adicional atualizado!" if self.id_adicional_em_edicao else "Adicional cadastrado!"
            self.limpar_form_adicional()
            self.atualizar_lista_adicionais()
            messagebox.showinfo("Sucesso", msg)
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Já existe um adicional com este nome.")

    def excluir_adicional(self):
        sel = self.tree_add.selection()
        if sel and messagebox.askyesno("AVISO", "Excluir o adicional selecionado?"):
            ad_id = self.tree_add.item(sel[0])['values'][0]
            self.db.excluir_adicional(ad_id)
            self.limpar_form_adicional()
            self.atualizar_lista_adicionais()

    def montar_aba_clientes(self):
        f_lista = tk.Frame(self.aba_cad, bg=self.bg_card, width=620,
                           highlightthickness=1, highlightbackground=self.bg_soft)
        f_lista.pack(side="right", fill="both", expand=True, padx=(12, 24), pady=24)

        f_cad_outer = tk.Frame(self.aba_cad, bg=self.bg_dark)
        f_cad_outer.pack(side="left", fill="both", expand=True)

        self.canvas_cad = tk.Canvas(f_cad_outer, bg=self.bg_dark, highlightthickness=0)
        self.scroll_cad = ttk.Scrollbar(f_cad_outer, orient="vertical", command=self.canvas_cad.yview)
        self.canvas_cad.configure(yscrollcommand=self.scroll_cad.set)
        self.scroll_cad.pack(side="right", fill="y")
        self.canvas_cad.pack(side="left", fill="both", expand=True)

        self.canvas_cad.bind('<Enter>', lambda e: self.canvas_cad.bind_all("<MouseWheel>", lambda e2: self.canvas_cad.yview_scroll(int(-1*(e2.delta/120)), "units")))
        self.canvas_cad.bind('<Leave>', lambda e: self.canvas_cad.unbind_all("<MouseWheel>"))

        f_cad = tk.Frame(self.canvas_cad, bg=self.bg_dark)
        self.window_id_cad = self.canvas_cad.create_window((0, 0), window=f_cad, anchor="nw")

        def _on_cad_frame_cfg(event):
            self.canvas_cad.configure(scrollregion=self.canvas_cad.bbox("all"))
        def _on_cad_canvas_cfg(event):
            if self.canvas_cad.winfo_exists():
                self.canvas_cad.itemconfig(self.window_id_cad, width=event.width)
        f_cad.bind("<Configure>", _on_cad_frame_cfg)
        self.canvas_cad.bind("<Configure>", _on_cad_canvas_cfg)

        f_cad_inner = tk.Frame(f_cad, bg=self.bg_dark)
        f_cad_inner.pack(fill="both", expand=True, padx=24, pady=24)
        f_cad = f_cad_inner

        tk.Label(f_cad, text="CADASTRO DE CLIENTE", font=(self.font_family, 15, "bold"),
                 bg=self.bg_dark, fg=self.yellow).pack(anchor="w", pady=(0, 4))
        tk.Label(f_cad, text="Digite o telefone para buscar automaticamente.",
                 bg=self.bg_dark, fg=self.text_muted, font=(self.font_family, 9)).pack(anchor="w", pady=(0, 18))

        self.ents_cli = {}
        campos = [("TELEFONE", "tel"), ("NOME", "nome"), ("ENDEREÇO", "end"),
                  ("BAIRRO", "bair"), ("TAXA DE ENTREGA (R$)", "tax"),
                  ("BLOCO / TORRE", "blo"), ("APARTAMENTO / Nº", "apt")]
        for texto, key in campos:
            tk.Label(f_cad, text=texto, bg=self.bg_dark, fg=self.text_muted,
                     font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(8, 2))
            if key == "tax":
                f_taxas = tk.Frame(f_cad, bg=self.bg_dark)
                f_taxas.pack(fill="x")
                self.var_taxa_cli = tk.StringVar(value=self.valores_taxas[0])
                for v in self.valores_taxas:
                    tk.Radiobutton(f_taxas, text=f"{v}", variable=self.var_taxa_cli, value=v,
                                   font=(self.font_family, 9, "bold"), bg=self.bg_soft, fg="white",
                                   selectcolor=self.accent, activebackground=self.accent,
                                   activeforeground="white", indicatoron=0, cursor="hand2",
                                   bd=0).pack(side="left", fill="x", expand=True, padx=1, ipady=4)
                self.ents_cli[key] = self.var_taxa_cli
            else:
                ent = self._input_premium(f_cad, font=(self.font_family, 12))
                ent.pack(fill="x", ipady=4)
                self.ents_cli[key] = ent
        self.ents_cli["tel"].bind("<KeyRelease>", self.busca_automatica_cliente)

        f_btns = tk.Frame(f_cad, bg=self.bg_dark)
        f_btns.pack(pady=22, fill="x")
        tk.Button(f_btns, text="SALVAR", bg=self.green, fg="white", font=(self.font_family, 11, "bold"),
                  command=self.salvar_cliente, relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(f_btns, text="LIMPAR", bg=self.bg_soft, fg="white", font=(self.font_family, 11, "bold"),
                  command=self.limpar_campos_cliente, relief="flat", cursor="hand2", bd=0, height=2).pack(side="left", fill="x", expand=True, padx=4)

        tk.Frame(f_lista, bg=self.accent2, height=3).pack(fill="x")

        cabe = tk.Frame(f_lista, bg=self.bg_card)
        cabe.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(cabe, text="CLIENTES CADASTRADOS", bg=self.bg_card, fg="white",
                 font=(self.font_family, 12, "bold")).pack(anchor="w")
        tk.Label(cabe, text="Toque uma linha para carregar no formulário ao lado.",
                 bg=self.bg_card, fg=self.text_muted, font=(self.font_family, 9)).pack(anchor="w")

        tk.Button(f_lista, text="EXCLUIR SELECIONADOS", bg=self.accent, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.excluir_cliente,
                  relief="flat", cursor="hand2", bd=0, height=2).pack(side="bottom", pady=12, fill="x", padx=14)

        f_tree_cli = tk.Frame(f_lista, bg=self.bg_card)
        f_tree_cli.pack(side="bottom", fill="both", expand=True, padx=14, pady=8)
        self.tree_cli = ttk.Treeview(f_tree_cli, columns=("n", "t", "e"), show="headings", selectmode="extended")
        self.tree_cli.heading("n", text="NOME"); self.tree_cli.heading("t", text="TELEFONE"); self.tree_cli.heading("e", text="ENDEREÇO")
        self.tree_cli.column("n", width=170); self.tree_cli.column("t", width=110); self.tree_cli.column("e", width=270)
        sb_cli = ttk.Scrollbar(f_tree_cli, orient="vertical", command=self.tree_cli.yview)
        self.tree_cli.configure(yscrollcommand=sb_cli.set)
        sb_cli.pack(side="right", fill="y")
        self.tree_cli.pack(side="left", fill="both", expand=True)
        self.tree_cli.bind("<<TreeviewSelect>>", self.carregar_cliente_selecionado)
        self.tree_cli.bind("<Button-1>", lambda e: self.clique_vazio_tree(e, self.tree_cli))
        self.atualizar_lista_clientes()

    def busca_automatica_cliente(self, event=None):
        tel = self.ents_cli["tel"].get().strip()
        if len(tel) >= 8:
            res = self.db.cursor.execute("SELECT nome, endereco, bairro, taxa, bloco, apto FROM clientes WHERE telefone=?", (tel,)).fetchone()
            if res:
                self.ents_cli["nome"].delete(0, tk.END); self.ents_cli["nome"].insert(0, str(res[0]))
                self.ents_cli["end"].delete(0, tk.END); self.ents_cli["end"].insert(0, str(res[1]))
                self.ents_cli["bair"].delete(0, tk.END); self.ents_cli["bair"].insert(0, str(res[2]))
                tax_val = f"{float(res[3]):.2f}" if res[3] else self.valores_taxas[0]
                if tax_val in self.valores_taxas: self.ents_cli["tax"].set(tax_val)
                self.ents_cli["blo"].delete(0, tk.END); self.ents_cli["blo"].insert(0, str(res[4] or ""))
                self.ents_cli["apt"].delete(0, tk.END); self.ents_cli["apt"].insert(0, str(res[5] or ""))

    def carregar_cliente_selecionado(self, event):
        sel = self.tree_cli.selection()
        if len(sel) == 1:
            tel_item = self.tree_cli.item(sel[0])['values'][1]
            res = self.db.cursor.execute("SELECT nome, telefone, endereco, bairro, bloco, apto, taxa FROM clientes WHERE telefone=?", (str(tel_item),)).fetchone()
            if res:
                self.limpar_campos_cliente()
                self.ents_cli["tel"].insert(0, str(res[1])); self.ents_cli["nome"].insert(0, str(res[0]))
                self.ents_cli["end"].insert(0, str(res[2])); self.ents_cli["bair"].insert(0, str(res[3])); self.ents_cli["blo"].insert(0, str(res[4] or "")); self.ents_cli["apt"].insert(0, str(res[5] or ""))
                tax_val = f"{float(res[6]):.2f}" if res[6] else self.valores_taxas[0]
                if tax_val in self.valores_taxas: self.ents_cli["tax"].set(tax_val)

    def salvar_cliente(self):
        t, n = self.ents_cli["tel"].get(), self.ents_cli["nome"].get().upper(); e, b = self.ents_cli["end"].get().upper(), self.ents_cli["bair"].get().upper()
        bl, ap = self.ents_cli["blo"].get().upper(), self.ents_cli["apt"].get().upper()
        try: tx = float(self.ents_cli["tax"].get().replace(',', '.')) if self.ents_cli["tax"].get() else 0.0
        except ValueError: tx = 0.0
        if not t or not n: messagebox.showwarning("Aviso", "Telefone e Nome são obrigatórios."); return
        try:
            self.db.cursor.execute("INSERT INTO clientes (telefone, nome, endereco, bairro, taxa, bloco, apto) VALUES (?,?,?,?,?,?,?) ON CONFLICT(telefone) DO UPDATE SET nome=excluded.nome, endereco=excluded.endereco, bairro=excluded.bairro, taxa=excluded.taxa, bloco=excluded.bloco, apto=excluded.apto", (t, n, e, b, tx, bl, ap))
            self.db.conn.commit(); self.atualizar_lista_clientes(); self.limpar_campos_cliente()
        except Exception as err: messagebox.showerror("Erro", str(err))

    def atualizar_lista_clientes(self):
        for i in self.tree_cli.get_children(): self.tree_cli.delete(i)
        for r in self.db.cursor.execute("SELECT nome, telefone, endereco FROM clientes ORDER BY nome").fetchall(): self.tree_cli.insert("", "end", values=r)

    def excluir_cliente(self):
        sel = self.tree_cli.selection()
        if sel and messagebox.askyesno("AVISO", f"Excluir {len(sel)} cliente(s)?"):
            for item in sel: self.db.cursor.execute("DELETE FROM clientes WHERE telefone=?", (str(self.tree_cli.item(item)['values'][1]),))
            self.db.conn.commit(); self.atualizar_lista_clientes(); self.limpar_campos_cliente()

    def limpar_campos_cliente(self):
        for key, e in self.ents_cli.items(): 
            if key == "tax": e.set(self.valores_taxas[0])
            else: e.delete(0, tk.END)

    def abrir_janela_pagamento(self):
        if not self.carrinho: return
        win = tk.Toplevel(self.root)
        win.attributes("-toolwindow", True)
        win.grab_set()
        self.centralizar_janela_toplevel(win, 480, 560)
        win.configure(bg=self.bg_dark)

        tk.Frame(win, bg=self.green, height=4).pack(fill="x")

        tk.Label(win, text="PAGAMENTO", font=(self.font_family, 14, "bold"),
                 bg=self.bg_dark, fg=self.text_muted).pack(pady=(24, 2))
        total = self.total_atual
        tk.Label(win, text=f"R$ {total:.2f}", font=(self.font_family, 38, "bold"),
                 bg=self.bg_dark, fg=self.green).pack(pady=4)

        tk.Label(win, text="FORMA DE PAGAMENTO", font=(self.font_family, 9, "bold"),
                 fg=self.text_muted, bg=self.bg_dark).pack(pady=(18, 6))
        v_pg = tk.StringVar(value="DINHEIRO")
        f_pg_radios = tk.Frame(win, bg=self.bg_dark)
        f_pg_radios.pack()
        for op in ["DINHEIRO", "CARTÃO", "PIX"]:
            tk.Radiobutton(f_pg_radios, text=op, variable=v_pg, value=op,
                           font=(self.font_family, 11, "bold"), bg=self.bg_soft, fg="white",
                           selectcolor=self.accent, activebackground=self.accent,
                           activeforeground="white", indicatoron=0, width=10,
                           cursor="hand2", bd=0).pack(side="left", padx=4, ipady=8)

        tk.Label(win, text="VALOR RECEBIDO", font=(self.font_family, 9, "bold"),
                 fg=self.text_muted, bg=self.bg_dark).pack(pady=(20, 4))
        e_rec = tk.Entry(win, font=(self.font_family, 22, "bold"), justify="center", width=14,
                         bg=self.bg_soft, fg="white", relief="flat",
                         highlightthickness=1, highlightbackground=self.bg_soft,
                         highlightcolor=self.accent, insertbackground="white")
        e_rec.pack(pady=4, ipady=8)

        total_arredondado = round(total, 2)
        e_rec.insert(0, str(int(total_arredondado)) if total_arredondado == int(total_arredondado) else f"{total_arredondado:.2f}")
        e_rec.select_range(0, tk.END)
        e_rec.focus_set()

        f_troco = tk.Frame(win, bg=self.bg_card, padx=20, pady=10)
        f_troco.pack(pady=18, padx=40, fill="x")
        tk.Label(f_troco, text="TROCO", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w")
        lbl_troco = tk.Label(f_troco, text="R$ 0,00", font=(self.font_family, 20, "bold"),
                             bg=self.bg_card, fg=self.yellow)
        lbl_troco.pack(anchor="w")

        def atualizar_troco(event):
            try: lbl_troco.config(text=f"R$ {max(0, float(e_rec.get().replace(',','.') or 0) - total):.2f}")
            except: pass
        e_rec.bind("<KeyRelease>", atualizar_troco)

        def disparar_finalizacao(event=None): self.finalizar_venda(v_pg.get(), e_rec.get(), win)

        tk.Button(win, text="FINALIZAR & IMPRIMIR (ENTER)", bg=self.green, fg="white",
                  font=(self.font_family, 12, "bold"), height=2, command=disparar_finalizacao,
                  relief="flat", cursor="hand2", bd=0).pack(fill="x", padx=40, pady=10)
        win.bind('<Return>', disparar_finalizacao)

    def finalizar_venda(self, pg, rec_str, win):
        try:
            total = self.total_atual
            data = datetime.now().strftime("%d/%m/%Y %H:%M")
            tipo_pedido = self.tipo_var.get()
            end_alt = self.ent_end_alt.get().strip().upper()
            try: rec = float(rec_str.replace(',','.'))
            except: rec = total
            nome_f = self.cliente_atual['nome']
            
            if self.id_pedido_em_edicao:
                self.db.cursor.execute("UPDATE vendas SET data_hora=?, cliente=?, tipo_pedido=?, total=?, pagamento=?, itens_json=?, endereco_entrega=? WHERE id=?", (data, nome_f, tipo_pedido, total, pg, json.dumps(self.carrinho), end_alt, self.id_pedido_em_edicao))
                num_v = self.id_pedido_em_edicao
                self.id_pedido_em_edicao = None
            else:
                self.db.cursor.execute("INSERT INTO vendas (data_hora, cliente, tipo_pedido, total, pagamento, itens_json, endereco_entrega) VALUES (?,?,?,?,?,?,?)", (data, nome_f, tipo_pedido, total, pg, json.dumps(self.carrinho), end_alt))
                num_v = self.db.cursor.lastrowid
            self.db.conn.commit()
            
            texto = self.gerar_texto_comanda(num_v, data, tipo_pedido, total, pg, rec, max(0, rec-total), self.carrinho, nome_f, end_alt)
            self.enviar_impressora(texto, num_v)

            self.carrinho = []; self.cliente_atual = self.cliente_padrao.copy(); self.atualizar_labels_info()
            self.ent_busca_venda.delete(0, tk.END); self.ent_end_alt.delete(0, tk.END)
            self._atualizar_estado_taxa_alt()
            self.tipo_var.set("RETIRADA"); self.atualizar_taxa_carrinho()
            win.destroy()
            self.carregar_pedidos_dia()
            self.lbl_num_pedido.config(text=f"{self.obter_proximo_id()}")
        except Exception as e: messagebox.showerror("Erro", str(e))

    def gerar_texto_comanda(self, num, data, tipo, total, pagto, rec, troco, itens, nome_exibicao, end_alt):
        l = 48 if self.prefs.get("largura", "80mm") == "80mm" else 32
        l_big = l // 2

        corpo = f"<BIG_INV>{self.nome_loja[:l_big]:^{l_big}}</BIG_INV>\n"
        slogan = (getattr(self, "loja_slogan", "") or "").strip()
        if slogan:
            corpo += f"{slogan[:l]:^{l}}\n"
        endereco = (getattr(self, "loja_endereco", "") or "").strip()
        if endereco:
            corpo += f"{endereco[:l]:^{l}}\n"
        info_dupla = []
        telefone = (getattr(self, "loja_telefone", "") or "").strip()
        if telefone: info_dupla.append(f"TEL: {telefone}")
        cnpj = (getattr(self, "loja_cnpj", "") or "").strip()
        if cnpj: info_dupla.append(f"CNPJ: {cnpj}")
        if info_dupla:
            linha_info = "  |  ".join(info_dupla)
            corpo += f"{linha_info[:l]:^{l}}\n"
        instagram = (getattr(self, "loja_instagram", "") or "").strip()
        if instagram:
            corpo += f"{instagram[:l]:^{l}}\n"
        corpo += f"{'='*l}\n"

        corpo += f"PEDIDO: {num}  DATA: {data}\nCLIENTE: {nome_exibicao}\n"
        if self.cliente_atual.get('tel'): corpo += f"TEL: {self.cliente_atual['tel']}\n"
        
        corpo += f"<BIG_INV>{f'TIPO: {tipo}':^{l_big}}</BIG_INV>\n"
        
        taxa_entrega = 0.0
        for i in itens:
            if i['nome'] == "TAXA DE ENTREGA":
                taxa_entrega += i['preco'] * i['qtd']

        if tipo == "ENTREGA":
            try:
                dt_pedido = datetime.strptime(data, "%d/%m/%Y %H:%M")
                minutos_entrega = int(self.prefs.get("tempos", {}).get(f"{taxa_entrega:.2f}", "60"))
                minimo = (dt_pedido + timedelta(minutes=minutos_entrega)).strftime('%H:%M')
                maximo = (dt_pedido + timedelta(minutes=minutos_entrega + 40)).strftime('%H:%M')
                corpo += f"PREVISAO DE ENTREGA: {minimo} AS {maximo}\n"
            except: pass

            if end_alt: corpo += f"END: {end_alt}\n"
            else:
                corpo += f"END: {self.cliente_atual.get('end', '')}\n"
                if self.cliente_atual.get('bairro'): corpo += f"BAIRRO: {self.cliente_atual['bairro']}\n"
                comp = []
                if self.cliente_atual.get('bloco'): comp.append(f"BLO: {self.cliente_atual['bloco']}")
                if self.cliente_atual.get('apto'): comp.append(f"APT: {self.cliente_atual['apto']}")
                if comp: corpo += " | ".join(comp) + "\n"
        elif tipo == "RETIRADA":
            try:
                dt_pedido = datetime.strptime(data, "%d/%m/%Y %H:%M")
                minutos_retirada = int(self.prefs.get("tempo_retirada", "20"))
                pronto = (dt_pedido + timedelta(minutes=minutos_retirada)).strftime('%H:%M')
                corpo += f"PRONTO PARA RETIRADA AS: {pronto}\n"
            except: pass
        
        corpo += f"{'-'*l}\n"
        for i in itens:
            if i['nome'] == "TAXA DE ENTREGA": continue
            linha_nome = f"{i['qtd']}x {i['nome']}"
            linha_preco = f"R$ {i['preco']*i['qtd']:.2f}"
            corpo += f"{linha_nome} - {linha_preco}\n"
            if i.get('obs'): corpo += f" OBS: {i['obs']}\n"
            
        corpo += f"{'-'*l}\n"
        
        if taxa_entrega > 0:
            corpo += f"TAXA DE ENTREGA: R$ {taxa_entrega:.2f}\n"
            
        corpo += f"TOTAL: R$ {total:.2f}\nPAGAMENTO: {pagto}\n"
        
        if pagto == "DINHEIRO": 
            corpo += f"RECEBIDO: R$ {rec:.2f}\n"
            if troco > 0:
                corpo += f"<BIG_INV>{f'TROCO: R$ {troco:.2f}':^{l_big}}</BIG_INV>\n"
            else:
                corpo += f"TROCO: R$ {troco:.2f}\n"
        
        corpo += f"\n<INV>{'OBRIGADO PELA PREFERENCIA!':^{l}}</INV>\n"
        corpo += f"<INV>{'DESENVOLVIDO BY ADOLFO NOVAES':^{l}}</INV>\n\n\n\n"
        return corpo

    def enviar_impressora(self, texto, num):
        if not IMPRESSAO_DISPONIVEL: 
            print(f"--- SIMULAÇÃO #{num} ---\n{texto}")
            return
        try:
            hPrinter = win32print.OpenPrinter(self.prefs["impressora"])
            try:
                hJob = win32print.StartDocPrinter(hPrinter, 1, (f"Pedido {num}", None, "RAW"))
                win32print.StartPagePrinter(hPrinter)
                t = self.prefs["tamanho_fonte"]
                if t == "GRANDE": cmd = b'\x1b\x21\x30'
                elif t == "MÉDIA": cmd = b'\x1b\x21\x10'
                elif t == "PERSONALIZADO": cmd = b'\x1b\x21' + bytes([int(self.prefs.get("fonte_num", "0"))])
                else: cmd = b'\x1b\x21\x00'
                
                texto_limpo = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
                encoded_texto = texto_limpo.encode('cp850')
                
                cmd_big_inv = b'\x1b\x21\x30\x1b\x45\x01\x1d\x42\x01'
                cmd_inv = cmd + b'\x1b\x45\x01\x1d\x42\x01' 
                cmd_normal = cmd + b'\x1b\x45\x01\x1d\x42\x00' 
                
                encoded_texto = encoded_texto.replace(b'<BIG_INV>', cmd_big_inv).replace(b'</BIG_INV>', cmd_normal)
                encoded_texto = encoded_texto.replace(b'<INV>', cmd_inv).replace(b'</INV>', cmd_normal)
                
                win32print.WritePrinter(hPrinter, cmd + b'\x1b\x45\x01' + encoded_texto + b'\x1b\x45\x00\x1b\x21\x00' + b"\n\n\x1d\x56\x01")
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
            finally: win32print.ClosePrinter(hPrinter)
        except Exception as e: messagebox.showwarning("Impressora", str(e))

    def reimprimir_pedido(self):
        s = self.tree_ped.selection()
        if not s: return
        for item in s:
            pid = self.tree_ped.item(item)['values'][0]
            r = self.db.cursor.execute("SELECT * FROM vendas WHERE id=?", (pid,)).fetchone()
            c = self.db.cursor.execute("SELECT nome, telefone, endereco, bairro, bloco, apto, taxa FROM clientes WHERE nome=?", (r[3],)).fetchone()
            if c:
                self.cliente_atual = {"nome": c[0], "tel": c[1], "end": c[2], "bairro": c[3], "bloco": c[4], "apto": c[5], "taxa": c[6]}
            else:
                self.cliente_atual = {"nome": r[3], "tel": "", "end": "", "bairro": "", "bloco": "", "apto": "", "taxa": 0}
            end_alt_gravado = r[8] if len(r) > 8 and r[8] else ""
            self.enviar_impressora(self.gerar_texto_comanda(r[0], r[1], r[4], r[2], r[5], r[2], 0, json.loads(r[7]), r[3], end_alt_gravado), r[0])

    def buscar_cliente_venda(self, by_enter=False):
        txt = self.ent_busca_venda.get().upper().strip()
        if not txt: return
        r = self.db.cursor.execute("SELECT nome, telefone, endereco, bairro, bloco, apto, taxa FROM clientes WHERE telefone=?", (txt,)).fetchone()
        if r: 
            self.cliente_atual = {"nome": r[0], "tel": r[1], "end": r[2], "bairro": r[3], "bloco": r[4], "apto": r[5], "taxa": r[6]}
            
            if r[2] and r[2].strip(): 
                self.tipo_var.set("ENTREGA")
            else:
                self.tipo_var.set("RETIRADA")
            
            self.ent_end_alt.delete(0, tk.END); self.ent_busca_venda.delete(0, tk.END); self.ent_busca_venda.insert(0, r[1])
        else:
            self.cliente_atual = {"nome": txt, "tel": txt if any(c.isdigit() for c in txt) else "", "end": "", "bairro": "", "taxa": 0}
            self.ent_end_alt.delete(0, tk.END)
        self._atualizar_estado_taxa_alt()
        self.atualizar_taxa_carrinho()
        self.atualizar_labels_info()

    def atalho_enter(self, event): 
        if self.root.focus_get() == self.ent_busca_venda: self.buscar_cliente_venda(by_enter=True)

    def sincronizar_treeview(self): 
        taxas = [i for i in self.carrinho if i["nome"] == "TAXA DE ENTREGA"]
        outros = [i for i in self.carrinho if i["nome"] != "TAXA DE ENTREGA"]
        self.carrinho = outros + taxas 
        for i in self.tree.get_children(): self.tree.delete(i)
        for i in self.carrinho: self.tree.insert("", "end", values=(i["nome"], i["qtd"], f"R$ {i['preco']*i['qtd']:.2f}"))
            
    def atualizar_total(self): 
        self.total_atual = sum(i['preco']*i['qtd'] for i in self.carrinho)
        self.lbl_total_grande.config(text=f"R$ {self.total_atual:.2f}")
        
    def centralizar_janela_toplevel(self, win, w, h): 
        tela_w = self.root.winfo_screenwidth()
        tela_h = self.root.winfo_screenheight()
        x, y = (tela_w // 2) - (w // 2), (tela_h // 2) - (h // 2) - 40 
        win.geometry(f"{w}x{h}+{x}+{max(0, y)}") 
        
    def obter_proximo_id(self): 
        res = self.db.cursor.execute("SELECT MAX(id) FROM vendas").fetchone()
        return (res[0] + 1) if res and res[0] else 1
        
    def editar_item_carrinho(self, e):
        s = self.tree.selection()
        if len(s) == 1 and "TAXA DE ENTREGA" not in self.carrinho[self.tree.index(s[0])]["nome"]:
            idx = self.tree.index(s[0])
            r = self.db.cursor.execute("SELECT id, nome, preco, categoria FROM produtos WHERE nome=?", (self.carrinho[idx]["nome"].split(" (C/")[0],)).fetchone()
            if r: self.modal_item(r, edit_idx=idx)
            else: self.modal_item([0, self.carrinho[idx]["nome"].split(" (C/")[0], self.carrinho[idx]["preco"], "LANCHES"], edit_idx=idx)

    def remover_item(self, event=None):
        s = self.tree.selection()
        if s and messagebox.askyesno("Confirmar", "Excluir itens selecionados?"):
            for idx in sorted([self.tree.index(i) for i in s], reverse=True): del self.carrinho[idx]
            self.sincronizar_treeview(); self.atualizar_total()

    def montar_aba_pedidos(self):
        f_top = tk.Frame(self.aba_pedidos, bg=self.bg_dark, pady=18)
        f_top.pack(fill="x", padx=24)
        tk.Label(f_top, text="PEDIDOS DO DIA", font=(self.font_family, 16, "bold"),
                 bg=self.bg_dark, fg=self.yellow).pack(side="left")
        tk.Label(f_top, text="Filtre por data, reimprima, edite ou visualize uma comanda.",
                 bg=self.bg_dark, fg=self.text_muted, font=(self.font_family, 9)).pack(side="left", padx=12)

        f_filtros = tk.Frame(self.aba_pedidos, bg=self.bg_card, padx=18, pady=12)
        f_filtros.pack(fill="x", padx=24, pady=(0, 12))

        tk.Label(f_filtros, text="DATA", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(side="left", padx=(0, 8))
        self.ent_data_ped = DateEntry(f_filtros, width=12, background=self.accent,
                                      foreground='white', borderwidth=0,
                                      date_pattern='dd/mm/yyyy', font=(self.font_family, 11))
        self.ent_data_ped.pack(side="left")
        tk.Button(f_filtros, text="FILTRAR", command=self.carregar_pedidos_dia,
                  bg=self.accent, fg="white", relief="flat", cursor="hand2", bd=0,
                  font=(self.font_family, 10, "bold"), padx=18, pady=6).pack(side="left", padx=12)

        btns = tk.Frame(self.aba_pedidos, bg=self.bg_dark)
        btns.pack(side="bottom", fill="x", pady=18, padx=24)
        tk.Button(btns, text="REIMPRIMIR", bg=self.accent2, fg="white",
                  font=(self.font_family, 11, "bold"), command=self.reimprimir_pedido,
                  relief="flat", cursor="hand2", bd=0, height=2, padx=18).pack(side="left", padx=4)
        tk.Button(btns, text="EDITAR", bg=self.yellow, fg=self.bg_dark,
                  font=(self.font_family, 11, "bold"), command=self.carregar_pedido_para_edicao,
                  relief="flat", cursor="hand2", bd=0, height=2, padx=18).pack(side="left", padx=4)
        tk.Button(btns, text="VISUALIZAR", bg="#7E57C2", fg="white",
                  font=(self.font_family, 11, "bold"), command=self.visualizar_pedido,
                  relief="flat", cursor="hand2", bd=0, height=2, padx=18).pack(side="left", padx=4)
        tk.Button(btns, text="EXCLUIR", bg=self.accent, fg="white",
                  font=(self.font_family, 11, "bold"), command=self.excluir_pedido,
                  relief="flat", cursor="hand2", bd=0, height=2, padx=18).pack(side="right", padx=4)

        f_tree_ped = tk.Frame(self.aba_pedidos, bg=self.bg_card,
                              highlightthickness=1, highlightbackground=self.bg_soft)
        f_tree_ped.pack(fill="both", expand=True, padx=24)

        self.tree_ped = ttk.Treeview(f_tree_ped, columns=("id", "hora", "cli", "tipo", "total", "pgto"), show="headings", selectmode="extended")
        self.tree_ped.heading("id", text="PEDIDO")
        self.tree_ped.heading("hora", text="HORA")
        self.tree_ped.heading("cli", text="CLIENTE")
        self.tree_ped.heading("tipo", text="TIPO")
        self.tree_ped.heading("total", text="TOTAL")
        self.tree_ped.heading("pgto", text="PAGAMENTO")
        self.tree_ped.column("id", width=100, anchor="center")
        self.tree_ped.column("hora", width=130, anchor="center")
        self.tree_ped.column("cli", width=260)
        self.tree_ped.column("tipo", width=100, anchor="center")
        self.tree_ped.column("total", width=110, anchor="center")
        self.tree_ped.column("pgto", width=130, anchor="center")

        sb_ped = ttk.Scrollbar(f_tree_ped, orient="vertical", command=self.tree_ped.yview)
        self.tree_ped.configure(yscrollcommand=sb_ped.set)
        sb_ped.pack(side="right", fill="y")
        self.tree_ped.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        self.carregar_pedidos_dia()

    def carregar_pedidos_dia(self):
        for i in self.tree_ped.get_children(): self.tree_ped.delete(i)
        for r in self.db.cursor.execute("SELECT id, data_hora, cliente, tipo_pedido, total, pagamento FROM vendas WHERE data_hora LIKE ? ORDER BY id DESC", (f"{self.ent_data_ped.get()}%",)).fetchall(): 
            self.tree_ped.insert("", "end", values=(r[0], r[1], r[2], r[3], f"R$ {r[4]:.2f}", r[5]))

    def excluir_pedido(self):
        s = self.tree_ped.selection()
        if s and messagebox.askyesno("AVISO", f"Excluir {len(s)} pedido(s)?"):
            for item in s: self.db.cursor.execute("DELETE FROM vendas WHERE id=?", (self.tree_ped.item(item)['values'][0],))
            self.db.conn.commit()
            self.carregar_pedidos_dia()

    def carregar_pedido_para_edicao(self):
        s = self.tree_ped.selection()
        if len(s) == 1:
            pid = self.tree_ped.item(s[0])['values'][0]
            r = self.db.cursor.execute("SELECT * FROM vendas WHERE id=?", (pid,)).fetchone()
            self.tabs.select(0)
            self.carrinho, self.id_pedido_em_edicao = json.loads(r[7]), pid
            self.lbl_num_pedido.config(text=f"{pid}")
            self.tipo_var.set(r[4])
            c = self.db.cursor.execute("SELECT nome, telefone, endereco, bairro, bloco, apto, taxa FROM clientes WHERE nome=?", (r[3],)).fetchone()
            if c: 
                self.cliente_atual = {"nome": c[0], "tel": c[1], "end": c[2], "bairro": c[3], "bloco": c[4], "apto": c[5], "taxa": c[6]}
                self.ent_busca_venda.delete(0, tk.END); self.ent_busca_venda.insert(0, c[1]) 
            else: 
                self.cliente_atual = {"nome": r[3], "tel": "", "end": "", "bairro": "", "taxa": 0}
                self.ent_busca_venda.delete(0, tk.END); self.ent_busca_venda.insert(0, r[3])
            
            self.ent_end_alt.delete(0, tk.END)
            if len(r) > 8 and r[8]: self.ent_end_alt.insert(0, r[8])
            self._atualizar_estado_taxa_alt()
            self.atualizar_labels_info(); self.sincronizar_treeview(); self.atualizar_total()

    def visualizar_pedido(self):
        s = self.tree_ped.selection()
        if len(s) == 1: self.abrir_modal_visualizar_pedido(self.tree_ped.item(s[0])['values'][0])

    def visualizar_pedido_finan(self):
        s = self.tree_fin.selection()
        if len(s) == 1: self.abrir_modal_visualizar_pedido(self.tree_fin.item(s[0])['values'][0])
        
    def abrir_modal_visualizar_pedido(self, pid):
        r = self.db.cursor.execute("SELECT * FROM vendas WHERE id=?", (pid,)).fetchone()
        if not r: return
        win = tk.Toplevel(self.root)
        win.attributes("-toolwindow", True)
        win.grab_set()
        win.configure(bg=self.bg_dark)
        self.centralizar_janela_toplevel(win, 450, 600)
        
        tk.Frame(win, bg=self.accent, height=4).pack(fill="x")
        tk.Label(win, text=f"PEDIDO Nº {pid}", font=(self.font_family, 18, "bold"),
                 fg=self.yellow, bg=self.bg_dark).pack(pady=15)
        f_info = tk.Frame(win, bg=self.bg_card, padx=18, pady=14,
                          highlightthickness=1, highlightbackground=self.bg_soft)
        f_info.pack(fill="x", padx=24)
        for label, valor in (("DATA", r[1]), ("CLIENTE", r[3]), ("TIPO", r[4]), ("PAGAMENTO", r[5])):
            f_l = tk.Frame(f_info, bg=self.bg_card)
            f_l.pack(fill="x", pady=2)
            tk.Label(f_l, text=label, bg=self.bg_card, fg=self.text_muted,
                     font=(self.font_family, 9, "bold"), width=12, anchor="w").pack(side="left")
            tk.Label(f_l, text=str(valor), bg=self.bg_card, fg="white",
                     font=(self.font_family, 10), anchor="w").pack(side="left")

        tk.Label(win, text="ITENS DO PEDIDO", font=(self.font_family, 11, "bold"),
                 fg=self.text_muted, bg=self.bg_dark).pack(pady=(15, 5), anchor="w", padx=24)
        txt_itens = tk.Text(win, bg=self.bg_soft, fg="white", font=("Consolas", 11),
                            bd=0, height=12, relief="flat", highlightthickness=1,
                            highlightbackground=self.bg_soft, padx=12, pady=10)
        txt_itens.pack(fill="both", expand=True, padx=24, pady=5)
        texto = ""
        for i in json.loads(r[7]):
            texto += f"{i['qtd']}x {i['nome']} - R$ {i['preco']*i['qtd']:.2f}\n"
            if i.get('obs'): texto += f"   OBS: {i['obs']}\n"
            texto += "-"*40 + "\n"
        txt_itens.insert("1.0", texto); txt_itens.config(state="disabled")

        f_total = tk.Frame(win, bg=self.bg_card, padx=18, pady=12)
        f_total.pack(fill="x", padx=24, pady=14)
        tk.Label(f_total, text="TOTAL DO PEDIDO", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(anchor="w")
        tk.Label(f_total, text=f"R$ {r[2]:.2f}", font=(self.font_family, 24, "bold"),
                 fg=self.green, bg=self.bg_card).pack(anchor="w")

    def criar_card_financeiro(self, master, col, titulo, cor_destaque, cor_fundo):
        frame_card = tk.Frame(master, bg=self.bg_card, bd=0,
                              highlightthickness=1, highlightbackground=self.bg_soft)
        frame_card.grid(row=0, column=col, padx=8, sticky="nsew")
        tk.Frame(frame_card, bg=cor_destaque, height=4).pack(fill="x")
        inner = tk.Frame(frame_card, bg=self.bg_card, padx=20, pady=18)
        inner.pack(fill="both", expand=True)
        tk.Label(inner, text=titulo, bg=self.bg_card, fg=cor_destaque,
                 font=(self.font_family, 10, "bold")).pack(anchor="w")
        lbl_val = tk.Label(inner, text="R$ 0,00", bg=self.bg_card, fg="white",
                           font=(self.font_family, 22, "bold"))
        lbl_val.pack(anchor="w", pady=(8, 0))
        return lbl_val

    def montar_financeiro(self):
        f_top = tk.Frame(self.aba_finan, bg=self.bg_dark)
        f_top.pack(fill="x", padx=24, pady=(18, 6))
        tk.Label(f_top, text="DASHBOARD FINANCEIRO", font=(self.font_family, 16, "bold"),
                 fg=self.yellow, bg=self.bg_dark).pack(side="left")
        tk.Label(f_top, text="Filtros, indicadores e impressão de relatório de caixa.",
                 bg=self.bg_dark, fg=self.text_muted, font=(self.font_family, 9)).pack(side="left", padx=12)

        f_filtros = tk.Frame(self.aba_finan, bg=self.bg_card,
                             highlightthickness=1, highlightbackground=self.bg_soft)
        f_filtros.pack(fill="x", padx=24, pady=8)
        tk.Frame(f_filtros, bg=self.accent, height=3).pack(fill="x")

        f_filtros_inner = tk.Frame(f_filtros, bg=self.bg_card, padx=18, pady=12)
        f_filtros_inner.pack(fill="x")

        tk.Label(f_filtros_inner, text="PERÍODO", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(side="left", padx=(0, 10))
        self.per_var = tk.StringVar(value="HOJE")
        cb = ttk.Combobox(f_filtros_inner, textvariable=self.per_var,
                          values=["HOJE", "ESTA SEMANA", "ESTE MÊS", "PERSONALIZADO"],
                          state="readonly", width=18, font=(self.font_family, 11))
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self.toggle_dat_fin())

        self.f_dat_fin = tk.Frame(f_filtros_inner, bg=self.bg_card)
        self.ent_f1 = DateEntry(self.f_dat_fin, width=12, background=self.accent,
                                foreground='white', borderwidth=0,
                                date_pattern='dd/mm/yyyy', font=(self.font_family, 11))
        self.ent_f1.pack(side="left", padx=5)
        tk.Label(self.f_dat_fin, text="ATÉ", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).pack(side="left")
        self.ent_f2 = DateEntry(self.f_dat_fin, width=12, background=self.accent,
                                foreground='white', borderwidth=0,
                                date_pattern='dd/mm/yyyy', font=(self.font_family, 11))
        self.ent_f2.pack(side="left", padx=5)

        tk.Button(f_filtros_inner, text="GERAR DADOS", bg=self.green, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.gerar_relatorio_moderno,
                  relief="flat", cursor="hand2", bd=0, padx=18, pady=8).pack(side="left", padx=18)
        tk.Button(f_filtros_inner, text="IMPRIMIR RELATÓRIO", bg=self.accent2, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.imprimir_relatorio_financeiro,
                  relief="flat", cursor="hand2", bd=0, padx=14, pady=8).pack(side="right")
        tk.Button(f_filtros_inner, text="VISUALIZAR PEDIDO", bg="#7E57C2", fg="white",
                  font=(self.font_family, 10, "bold"), command=self.visualizar_pedido_finan,
                  relief="flat", cursor="hand2", bd=0, padx=14, pady=8).pack(side="right", padx=8)
        
        self.f_cards = tk.Frame(self.aba_finan, bg=self.bg_dark)
        self.f_cards.pack(fill="x", padx=10, pady=10)
        for i in range(4): self.f_cards.columnconfigure(i, weight=1)
        self.lbl_val_geral = self.criar_card_financeiro(self.f_cards, 0, "TOTAL GERAL", self.yellow, "#333300")
        self.lbl_val_din   = self.criar_card_financeiro(self.f_cards, 1, "💵 DINHEIRO", self.green, "#002b11")
        self.lbl_val_car   = self.criar_card_financeiro(self.f_cards, 2, "💳 CARTÃO", "#3498DB", "#002244")
        self.lbl_val_pix   = self.criar_card_financeiro(self.f_cards, 3, "📱 PIX", "#9B59B6", "#2a0033")
        
        f_tabela = tk.Frame(self.aba_finan, bg=self.bg_dark)
        f_tabela.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        tk.Label(f_tabela, text="HISTÓRICO DE TRANSAÇÕES", bg=self.bg_dark, fg="gray", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        
        f_tree_fin = tk.Frame(f_tabela, bg=self.bg_dark)
        f_tree_fin.pack(fill="both", expand=True)
        
        self.tree_fin = ttk.Treeview(f_tree_fin, columns=("id", "data", "cli", "tipo", "pgto", "valor"), show="headings", height=10)
        self.tree_fin.heading("id", text="Nº PEDIDO")
        self.tree_fin.heading("data", text="DATA/HORA")
        self.tree_fin.heading("cli", text="CLIENTE")
        self.tree_fin.heading("tipo", text="TIPO")
        self.tree_fin.heading("pgto", text="PAGAMENTO")
        self.tree_fin.heading("valor", text="VALOR")
        self.tree_fin.column("id", width=80, anchor="center")
        self.tree_fin.column("data", width=140, anchor="center")
        self.tree_fin.column("cli", width=220)
        self.tree_fin.column("tipo", width=100, anchor="center")
        self.tree_fin.column("pgto", width=120, anchor="center")
        self.tree_fin.column("valor", width=100, anchor="center")
        
        sb_fin = ttk.Scrollbar(f_tree_fin, orient="vertical", command=self.tree_fin.yview)
        self.tree_fin.configure(yscrollcommand=sb_fin.set)
        sb_fin.pack(side="right", fill="y")
        self.tree_fin.pack(side="left", fill="both", expand=True)
        
        self.gerar_relatorio_moderno()

    def salvar_dados_loja(self):
        n = self.ents_loja["nome"].get().strip().upper()
        s = self.ent_conf_senha.get().strip()
        end = self.ents_loja["endereco"].get().strip()
        cnpj = self.ents_loja["cnpj"].get().strip()
        tel = self.ents_loja["telefone"].get().strip()
        insta = self.ents_loja["instagram"].get().strip()
        slogan = self.ents_loja["slogan"].get().strip()
        if not n or not s:
            return messagebox.showwarning("Aviso", "Nome da loja e senha são obrigatórios!")

        self.db.set_config(n, s, end, cnpj, tel, insta, slogan)
        self.nome_loja = n
        self.senha_finan = s
        self.loja_endereco = end
        self.loja_cnpj = cnpj
        self.loja_telefone = tel
        self.loja_instagram = insta
        self.loja_slogan = slogan
        self.root.title(f"AURA SYSTEM - {self.nome_loja}")
        messagebox.showinfo("Sucesso", "Dados da loja atualizados com sucesso!\nA partir do próximo pedido eles aparecerão na comanda impressa.")

    def toggle_dat_fin(self):
        if self.per_var.get() == "PERSONALIZADO": self.f_dat_fin.pack(side="left")
        else: self.f_dat_fin.pack_forget()

    def gerar_relatorio_moderno(self):
        hoje = datetime.now().date()
        if self.per_var.get() == "HOJE": d1 = d2 = hoje
        elif self.per_var.get() == "ESTA SEMANA": d1, d2 = hoje - timedelta(days=hoje.weekday()), hoje
        elif self.per_var.get() == "ESTE MÊS": d1, d2 = hoje.replace(day=1), hoje
        else:
            try: d1, d2 = datetime.strptime(self.ent_f1.get(), "%d/%m/%Y").date(), datetime.strptime(self.ent_f2.get(), "%d/%m/%Y").date()
            except: return messagebox.showwarning("Erro", "Formato de data inválido. Use DD/MM/AAAA")
            
        tots = {"DINHEIRO": 0.0, "CARTÃO": 0.0, "PIX": 0.0}; t_geral = 0.0
        for i in self.tree_fin.get_children(): self.tree_fin.delete(i)
        
        for r in self.db.cursor.execute("SELECT id, data_hora, total, pagamento, cliente, tipo_pedido FROM vendas").fetchall():
            try:
                dt = datetime.strptime(r[1].split()[0], "%d/%m/%Y").date()
            except (ValueError, AttributeError, IndexError):
                continue
            if d1 <= dt <= d2:
                if r[3] in tots:
                    tots[r[3]] += r[2]
                t_geral += r[2]
                self.tree_fin.insert("", "end", values=(r[0], r[1], r[4], r[5], r[3], f"R$ {r[2]:.2f}"))
            
        self.dados_relatorio_atual = {"d1": d1, "d2": d2, "t_geral": t_geral, "dinheiro": tots["DINHEIRO"], "cartao": tots["CARTÃO"], "pix": tots["PIX"]}
        self.lbl_val_geral.config(text=f"R$ {t_geral:.2f}"); self.lbl_val_din.config(text=f"R$ {tots['DINHEIRO']:.2f}")
        self.lbl_val_car.config(text=f"R$ {tots['CARTÃO']:.2f}"); self.lbl_val_pix.config(text=f"R$ {tots['PIX']:.2f}")

    def imprimir_relatorio_financeiro(self):
        if not hasattr(self, 'dados_relatorio_atual'):
            return messagebox.showwarning("Aviso", "Gere um relatório primeiro.")
        d1, d2 = self.dados_relatorio_atual["d1"].strftime('%d/%m/%Y'), self.dados_relatorio_atual["d2"].strftime('%d/%m/%Y')
        l = 48 if self.prefs.get("largura", "80mm") == "80mm" else 32
        l_big = l // 2

        corpo = f"<BIG_INV>{'AURA SYSTEM':^{l_big}}</BIG_INV>\n"
        corpo += f"{self.nome_loja:^{l}}\n"
        endereco = (getattr(self, "loja_endereco", "") or "").strip()
        if endereco:
            corpo += f"{endereco[:l]:^{l}}\n"
        cnpj = (getattr(self, "loja_cnpj", "") or "").strip()
        telefone = (getattr(self, "loja_telefone", "") or "").strip()
        info = []
        if cnpj: info.append(f"CNPJ: {cnpj}")
        if telefone: info.append(f"TEL: {telefone}")
        if info:
            linha = "  |  ".join(info)
            corpo += f"{linha[:l]:^{l}}\n"
        corpo += f"{'='*l}\n"

        corpo += f"<INV>{'RELATORIO FINANCEIRO':^{l}}</INV>\n"
        corpo += f"{'-'*l}\n"
        corpo += f"DATA: {d1}\n" if d1 == d2 else f"PERIODO: {d1} A {d2}\n"
        corpo += f"EMITIDO EM: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        corpo += f"{'-'*l}\n"

        def alinhar_linha(t, v): return f"{t}{' ' * max(1, l - len(t) - len(f'R$ {v:.2f}'))}R$ {v:.2f}\n"

        corpo += alinhar_linha("DINHEIRO:", self.dados_relatorio_atual["dinheiro"])
        corpo += alinhar_linha("CARTAO:", self.dados_relatorio_atual["cartao"])
        corpo += alinhar_linha("PIX:", self.dados_relatorio_atual["pix"])
        corpo += f"{'-'*l}\n"
        total_str = f"TOTAL GERAL: R$ {self.dados_relatorio_atual['t_geral']:.2f}"
        corpo += f"<INV>{total_str:^{l}}</INV>\n"
        corpo += f"{'='*l}\n"
        corpo += f"<INV>{'AURA SYSTEM - VIA DE CAIXA':^{l}}</INV>\n"
        corpo += f"{'DESENVOLVIDO POR ADOLFO NOVAES':^{l}}\n\n\n\n"

        self.enviar_impressora(corpo, "RELATORIO")
        messagebox.showinfo("OK", "Relatório impresso!")

    def _seccao_premium(self, parent, titulo, subtitulo=""):
        wrap = tk.Frame(parent, bg=self.bg_card, highlightthickness=1, highlightbackground=self.bg_soft)
        wrap.pack(fill="x", padx=80, pady=10)

        barra = tk.Frame(wrap, bg=self.accent, height=3)
        barra.pack(fill="x")

        cabe = tk.Frame(wrap, bg=self.bg_card)
        cabe.pack(fill="x", padx=24, pady=(16, 4))
        tk.Label(cabe, text=titulo, bg=self.bg_card, fg=self.yellow,
                 font=(self.font_family, 12, "bold")).pack(anchor="w")
        if subtitulo:
            tk.Label(cabe, text=subtitulo, bg=self.bg_card, fg=self.text_muted,
                     font=(self.font_family, 9)).pack(anchor="w", pady=(2, 0))

        body = tk.Frame(wrap, bg=self.bg_card, padx=24, pady=16)
        body.pack(fill="x")
        return body

    def _input_premium(self, parent, **kwargs):
        kwargs.setdefault("bg", self.bg_soft)
        kwargs.setdefault("fg", "white")
        kwargs.setdefault("font", (self.font_family, 11))
        kwargs.setdefault("relief", "flat")
        kwargs.setdefault("highlightthickness", 1)
        kwargs.setdefault("highlightbackground", self.bg_soft)
        kwargs.setdefault("highlightcolor", self.accent)
        kwargs.setdefault("insertbackground", "white")
        return tk.Entry(parent, **kwargs)

    def montar_configuracoes(self):
        for w in self.aba_config.winfo_children(): w.destroy()

        self.canvas_config = tk.Canvas(self.aba_config, bg=self.bg_dark, highlightthickness=0)
        self.scroll_config = ttk.Scrollbar(self.aba_config, orient="vertical", command=self.canvas_config.yview)
        self.canvas_config.configure(yscrollcommand=self.scroll_config.set)

        self.scroll_config.pack(side="right", fill="y")
        self.canvas_config.pack(side="left", fill="both", expand=True)

        self.canvas_config.bind('<Enter>', lambda e: self.canvas_config.bind_all("<MouseWheel>", lambda e2: self.canvas_config.yview_scroll(int(-1*(e2.delta/120)), "units")))
        self.canvas_config.bind('<Leave>', lambda e: self.canvas_config.unbind_all("<MouseWheel>"))

        f = tk.Frame(self.canvas_config, bg=self.bg_dark, pady=20)
        self.window_id = self.canvas_config.create_window((0, 0), window=f, anchor="nw")

        def on_configure_frame(event):
            self.canvas_config.configure(scrollregion=self.canvas_config.bbox("all"))

        def on_configure_canvas(event):
            if self.canvas_config.winfo_exists():
                self.canvas_config.itemconfig(self.window_id, width=event.width)

        f.bind("<Configure>", on_configure_frame)
        self.canvas_config.bind("<Configure>", on_configure_canvas)

        tk.Label(f, text="CONFIGURAÇÕES DO SISTEMA", bg=self.bg_dark, fg="white",
                 font=(self.font_family, 18, "bold")).pack(pady=(10, 4))
        tk.Label(f, text="Personalize sua loja, impressão e tema visual.",
                 bg=self.bg_dark, fg=self.text_muted, font=(self.font_family, 10)).pack(pady=(0, 14))

        body = self._seccao_premium(f, "IDENTIDADE DA LOJA",
                                    "Estes dados sairão impressos no topo de toda comanda.")

        campos_loja = [
            ("Nome da Loja", "nome", self.nome_loja),
            ("Slogan / Frase de impacto", "slogan", getattr(self, "loja_slogan", "")),
            ("Endereço", "endereco", getattr(self, "loja_endereco", "")),
            ("CNPJ", "cnpj", getattr(self, "loja_cnpj", "")),
            ("Telefone", "telefone", getattr(self, "loja_telefone", "")),
            ("Instagram / Site", "instagram", getattr(self, "loja_instagram", "")),
        ]
        self.ents_loja = {}
        for i, (label, key, valor) in enumerate(campos_loja):
            tk.Label(body, text=label.upper(), bg=self.bg_card, fg=self.text_muted,
                     font=(self.font_family, 9, "bold")).grid(row=i, column=0, sticky="w", pady=(8, 2), padx=(0, 12))
            ent = self._input_premium(body, width=42)
            ent.grid(row=i, column=1, sticky="we", pady=(8, 2))
            ent.insert(0, valor or "")
            self.ents_loja[key] = ent
        body.columnconfigure(1, weight=1)

        i_senha = len(campos_loja)
        tk.Label(body, text="SENHA DO FINANCEIRO", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=i_senha, column=0, sticky="w", pady=(14, 2), padx=(0, 12))
        self.ent_conf_senha = self._input_premium(body, width=42, show="*")
        self.ent_conf_senha.grid(row=i_senha, column=1, sticky="we", pady=(14, 2))
        self.ent_conf_senha.insert(0, self.senha_finan)

        self.var_show = tk.BooleanVar()
        def toggle_senha(): self.ent_conf_senha.config(show="" if self.var_show.get() else "*")
        tk.Checkbutton(body, text="Mostrar senha", variable=self.var_show, command=toggle_senha,
                       bg=self.bg_card, fg=self.text_muted, selectcolor=self.bg_soft,
                       activebackground=self.bg_card, activeforeground="white",
                       font=(self.font_family, 9)).grid(row=i_senha+1, column=1, sticky="w", pady=(2, 8))

        tk.Button(body, text="SALVAR DADOS DA LOJA", bg=self.green, fg="white",
                  font=(self.font_family, 10, "bold"), command=self.salvar_dados_loja,
                  relief="flat", cursor="hand2", bd=0, padx=22, pady=10).grid(
                      row=i_senha+2, column=1, pady=(14, 0), sticky="w")

        body = self._seccao_premium(f, "TEMPOS DE PREPARO E ENTREGA",
                                    "Em minutos. Usado para imprimir a previsão na comanda.")

        tk.Label(body, text="ENTREGA — POR VALOR DA TAXA",
                 bg=self.bg_card, fg=self.yellow,
                 font=(self.font_family, 9, "bold")).grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 6))
        self.ents_tempos = {}
        for i, t in enumerate(self.valores_taxas):
            r, c = (i // 4) + 1, (i % 4) * 2
            tk.Label(body, text=f"R$ {t}", bg=self.bg_card, fg=self.text_muted,
                     font=(self.font_family, 9, "bold")).grid(row=r, column=c, sticky="e", pady=8, padx=(8, 6))
            ent = self._input_premium(body, width=8, justify="center")
            ent.grid(row=r, column=c+1, sticky="w", padx=(0, 18), pady=8)
            ent.insert(0, self.prefs.get("tempos", {}).get(t, "60"))
            self.ents_tempos[t] = ent

        tk.Frame(body, bg=self.bg_soft, height=1).grid(row=99, column=0, columnspan=8, sticky="we", pady=(12, 8))

        tk.Label(body, text="RETIRADA NO BALCÃO",
                 bg=self.bg_card, fg=self.yellow,
                 font=(self.font_family, 9, "bold")).grid(row=100, column=0, columnspan=8, sticky="w", pady=(0, 6))
        tk.Label(body, text="Tempo de preparo (min)", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=101, column=0, sticky="e", pady=8, padx=(8, 6))
        self.ent_tempo_retirada = self._input_premium(body, width=8, justify="center")
        self.ent_tempo_retirada.grid(row=101, column=1, sticky="w", padx=(0, 18), pady=8)
        self.ent_tempo_retirada.insert(0, self.prefs.get("tempo_retirada", "20"))

        body = self._seccao_premium(f, "IMPRESSÃO & TEMA VISUAL",
                                    "Impressora padrão, largura do papel, fonte e paleta.")

        tk.Label(body, text="IMPRESSORA", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.cb_imp = ttk.Combobox(body, width=38, state="readonly", font=(self.font_family, 11))
        self.cb_imp.grid(row=1, column=0, sticky="w", pady=(0, 12))
        tk.Button(body, text="ATUALIZAR LISTA", bg=self.bg_soft, fg="white",
                  font=(self.font_family, 9, "bold"), command=self.atualizar_lista_impressoras,
                  relief="flat", cursor="hand2", bd=0, padx=14, pady=6).grid(row=1, column=1, padx=12, sticky="w")

        tk.Label(body, text="LARGURA DO PAPEL", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.cb_larg = ttk.Combobox(body, values=["80mm", "60mm"], state="readonly", width=20, font=(self.font_family, 11))
        self.cb_larg.set(self.prefs.get("largura", "80mm"))
        self.cb_larg.grid(row=3, column=0, sticky="w", pady=(0, 12))

        tk.Label(body, text="TAMANHO DA FONTE", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.cb_fon = ttk.Combobox(body, values=["NORMAL", "MÉDIA", "GRANDE", "PERSONALIZADO"], state="readonly", width=20, font=(self.font_family, 11))
        self.cb_fon.set(self.prefs.get("tamanho_fonte", "NORMAL"))
        self.cb_fon.grid(row=5, column=0, sticky="w", pady=(0, 12))

        self.ent_fon = self._input_premium(body, width=10, justify="center")
        self.ent_fon.grid(row=5, column=1, sticky="w", padx=12)
        self.ent_fon.insert(0, self.prefs.get("fonte_num", "0"))

        tk.Label(body, text="TEMA DE CORES", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=6, column=0, sticky="w", pady=(0, 2))
        temas = ["PRETO E VERMELHO", "VERMELHO E AMARELO", "AZUL E VERDE", "LARANJA E VERMELHO"]
        self.cb_tema = ttk.Combobox(body, values=temas, state="readonly", width=30, font=(self.font_family, 11))
        self.cb_tema.set(self.prefs.get("tema", "PRETO E VERMELHO"))
        self.cb_tema.grid(row=7, column=0, sticky="w", pady=(0, 12))

        tk.Label(body, text="IMAGEM DE FUNDO (TELA DE VENDAS)", bg=self.bg_card, fg=self.text_muted,
                 font=(self.font_family, 9, "bold")).grid(row=8, column=0, sticky="w", pady=(0, 2))
        self.ent_bg = self._input_premium(body, width=42)
        self.ent_bg.grid(row=9, column=0, sticky="w", pady=(0, 4))
        self.ent_bg.insert(0, self.prefs.get("bg_image", ""))
        tk.Button(body, text="PROCURAR", bg=self.bg_soft, fg="white", font=(self.font_family, 9, "bold"),
                  command=self.escolher_imagem, relief="flat", cursor="hand2", bd=0, padx=14, pady=6).grid(row=9, column=1, padx=12, sticky="w")

        tk.Button(f, text="SALVAR CONFIGURAÇÕES DE IMPRESSÃO E TEMA", bg=self.green, fg="white",
                  font=(self.font_family, 12, "bold"), height=2, width=50,
                  command=self.salvar_config_btn, relief="flat", cursor="hand2", bd=0).pack(pady=24)

        tk.Label(f, text="Developed by Adolfo Rodrigues Novaes",
                 font=(self.font_family, 10, "italic"), fg=self.text_muted, bg=self.bg_dark).pack(side="bottom", pady=10)

        self.atualizar_lista_impressoras()

    def escolher_imagem(self):
        path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
        if path:
            self.ent_bg.delete(0, tk.END)
            self.ent_bg.insert(0, path)

    def atualizar_lista_impressoras(self):
        if IMPRESSAO_DISPONIVEL: 
            try:
                self.cb_imp['values'] = [p[2] for p in win32print.EnumPrinters(2)]
                if self.prefs.get("impressora") in self.cb_imp['values']:
                    self.cb_imp.set(self.prefs["impressora"])
                elif self.cb_imp['values']:
                    self.cb_imp.set(self.cb_imp['values'][0])
            except Exception:
                self.cb_imp['values'] = ["SIMULADOR"]
                self.cb_imp.set("SIMULADOR")
        else: 
            self.cb_imp['values'] = ["SIMULADOR (Falta pywin32)"]
            self.cb_imp.set("SIMULADOR (Falta pywin32)")

    def salvar_config_btn(self):
        self.prefs.update({
            "impressora": self.cb_imp.get(),
            "largura": self.cb_larg.get(),
            "tamanho_fonte": self.cb_fon.get(),
            "fonte_num": self.ent_fon.get(),
            "tema": self.cb_tema.get(),
            "bg_image": self.ent_bg.get(),
            "tempos": {t: self.ents_tempos[t].get() for t in self.valores_taxas},
            "tempo_retirada": self.ent_tempo_retirada.get().strip() or "20"
        })
        self.salvar_configuracoes()
        messagebox.showinfo("OK", "Configurações salvas com sucesso!\n\nO sistema será fechado para aplicar as alterações com segurança. Abra novamente.")
        self.root.quit(); self.root.destroy(); sys.exit(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppPDV(root)
    root.mainloop()