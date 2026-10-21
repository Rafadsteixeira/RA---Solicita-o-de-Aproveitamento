"""R.A - Equivalencias para cadastro de processos de equivalência em Excel."""

from __future__ import annotations

import html
import json
import re
import tkinter as tk
import webbrowser
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from storage import (
    EQUIVALENCE_COLUMNS,
    ExcelStorage,
    ExcelStorageError,
    PROCESS_COLUMNS,
    StudentStorage,
)
from pdf_import import PDFStudentImportError, extract_students_data


APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "portal_config.json"
DEFAULT_WORKBOOK = APP_DIR / "Base_Equivalencias_Senac.xlsx"

BLUE = "#004A8D"
BLUE_DARK = "#003366"
BLUE_LIGHT = "#EBF3FA"
ORANGE = "#F47920"
ORANGE_DARK = "#D96512"
PAGE_BG = "#F4F7F9"
TEXT = "#1E293B"
MUTED = "#718096"
WHITE = "#FFFFFF"
GREEN = "#14804A"
RED = "#B42318"
LINE = "#D8E1EA"

FIELD_LABELS = {
    "id_processo": "ID do processo",
    "data_requerimento": "Data do requerimento",
    "id_aluno": "ID do aluno / matrícula",
    "nome_aluno": "Nome do aluno",
    "cpf": "CPF",
    "rg": "RG",
    "telefone": "Telefone",
    "email": "E-mail",
    "situacao_interessado": "Situação do interessado",
    "curso_origem": "Curso de origem",
    "ano_cursado": "Ano cursado",
    "tipo_aluno": "Tipo de aluno",
    "curso_destino": "Curso de destino",
    "pc_cursado": "PC cursado",
    "pc_destino": "PC de destino",
    "codigo_oferta": "Código da oferta",
    "periodo_letivo": "Período letivo",
    "ch_total_oferta": "CH total da oferta",
    "ch_a_cursar": "CH a cursar",
    "valor_total": "Valor total",
    "qtd_parcelas": "Quantidade de parcelas",
    "valor_parcela": "Valor da parcela",
    "status_processo": "Status do processo",
    "responsavel_financeiro": "Responsável financeiro",
    "cpf_responsavel": "CPF do responsável",
    "prazo_resposta_dias": "Prazo para resposta",
    "observacoes": "Observações",
}

LEFT_FIELDS = [
    "id_processo", "data_requerimento", "id_aluno", "nome_aluno", "cpf", "rg",
    "telefone", "email", "situacao_interessado", "curso_origem", "ano_cursado",
    "tipo_aluno", "curso_destino", "pc_cursado",
]
RIGHT_FIELDS = [
    "pc_destino", "codigo_oferta", "periodo_letivo", "ch_total_oferta", "ch_a_cursar",
    "valor_total", "qtd_parcelas", "valor_parcela", "status_processo",
    "responsavel_financeiro", "cpf_responsavel", "prazo_resposta_dias",
]

CHOICES = {
    "situacao_interessado": [
        "Ex-aluno Senac", "Aluno de curso em andamento", "Outra instituição de ensino", "Outros"
    ],
    "tipo_aluno": ["Aluno", "Ex-aluno", "Bolsista Integral", "Bolsista Parcial", "Particular", "Outro"],
    "status_processo": [
        "Recebido", "Em análise pedagógica", "Aguardando documentos", "Deferido", "Indeferido", "Concluído"
    ],
}

EQ_FIELDS = [field for _heading, field in EQUIVALENCE_COLUMNS if field != "id_processo"]
UC_OPTIONS = [f"UC-{number}" for number in range(1, 26)]
EQ_LABELS = {
    "modulo_uc_cursado": "Nº do módulo / UC cursado",
    "nome_uc_cursado": "Nome do módulo ou UC cursado",
    "mencao": "Menção",
    "numero_uc_destino": "Nº da UC a cursar",
    "nome_uc_destino": "Nome da UC a cursar",
    "numero_aula": "Nº da aula",
    "id_curso": "ID do curso",
}


def center_window(window, width: int, height: int, parent=None, min_width: int = 480, min_height: int = 360):
    """Define um tamanho que caiba no monitor e centraliza a janela."""
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    width = min(width, max(320, screen_width - 40))
    height = min(height, max(320, screen_height - 80))

    if parent is not None:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    else:
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
    x = max(0, min(x, screen_width - width))
    y = max(0, min(y, screen_height - height))

    window.geometry(f"{width}x{height}+{x}+{y}")
    window.minsize(min(min_width, width), min(min_height, height))


class ScrollableFrame(ttk.Frame):
    def __init__(self, master, background=PAGE_BG):
        super().__init__(master)
        self.canvas = tk.Canvas(self, bg=background, highlightthickness=0, bd=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=background)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.inner.bind("<Configure>", self._configure_scrollregion)
        self.canvas.bind("<Configure>", self._configure_width)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    def _configure_scrollregion(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _configure_width(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _on_mousewheel(self, event):
        if self.winfo_ismapped():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class LoginScreen:
    """Tela de autenticação e criação inicial do administrador."""

    def __init__(self, root: tk.Tk, storage: ExcelStorage, on_authenticated):
        self.root = root
        self.storage = storage
        self.on_authenticated = on_authenticated
        self.first_user = not storage.list_users()

        root.title("R.A - Equivalencias | Acesso")
        login_height = 610 if self.first_user else 500
        center_window(root, 520, login_height, min_width=480, min_height=460)
        root.configure(bg=PAGE_BG)
        self.frame = tk.Frame(root, bg=PAGE_BG, padx=28, pady=28)
        self.frame.pack(fill="both", expand=True)

        card = tk.Frame(self.frame, bg=WHITE, highlightbackground=LINE, highlightthickness=1,
                        padx=34, pady=30)
        card.pack(expand=True, fill="x")
        tk.Label(card, text="S", bg="#14588D", fg=ORANGE, font=("Segoe UI", 22, "bold"),
                 width=2, padx=6, pady=4).pack(anchor="w")
        tk.Label(card, text="R.A - Equivalencias", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 17, "bold")).pack(anchor="w", pady=(14, 2))

        if self.first_user:
            tk.Label(card, text="Crie o primeiro acesso de administrador.", bg=WHITE, fg=MUTED,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 20))
            self.username_var = self._field(card, "USUÁRIO")
            self.password_var = self._field(card, "SENHA", secret=True)
            self.confirm_var = self._field(card, "CONFIRMAR SENHA", secret=True)
            tk.Label(card, text="Função: Administrador", bg=WHITE, fg=TEXT,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12, 18))
            button_text = "Criar administrador e entrar"
        else:
            tk.Label(card, text="Entre com seu usuário e senha.", bg=WHITE, fg=MUTED,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 20))
            self.username_var = self._field(card, "USUÁRIO")
            self.password_var = self._field(card, "SENHA", secret=True)
            self.confirm_var = None
            button_text = "Entrar"

        self.error_label = tk.Label(card, text="", bg=WHITE, fg=RED,
                                    font=("Segoe UI", 9), wraplength=390, justify="left")
        self.error_label.pack(anchor="w", pady=(0, 8))
        tk.Button(card, text=button_text, command=self.submit, bg=ORANGE, fg=WHITE,
                  activebackground=ORANGE_DARK, activeforeground=WHITE, relief="flat", bd=0,
                  padx=16, pady=11, font=("Segoe UI", 10, "bold"), cursor="hand2").pack(fill="x")
        tk.Label(card, text="As senhas são armazenadas como hash protegido na planilha.",
                 bg=WHITE, fg=MUTED, font=("Segoe UI", 8), wraplength=390).pack(anchor="w", pady=(16, 0))

        root.bind("<Return>", lambda _event: self.submit())
        self.username_entry.focus_set()

    def _field(self, parent, label, secret=False):
        tk.Label(parent, text=label, bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(8, 5))
        variable = tk.StringVar()
        entry = ttk.Entry(parent, textvariable=variable, show="•" if secret else "")
        entry.pack(fill="x", ipady=3)
        if label == "USUÁRIO":
            self.username_entry = entry
        elif label == "SENHA":
            self.password_entry = entry
        elif label == "CONFIRMAR SENHA":
            self.confirm_entry = entry
        return variable

    def submit(self):
        username = self.username_var.get().strip()
        password = self.password_var.get()
        if self.first_user:
            if password != self.confirm_var.get():
                self.error_label.configure(text="As senhas não coincidem.")
                return
            try:
                self.storage.create_user(username, password, "Administrador", must_change_password=False)
            except ExcelStorageError as exc:
                self.error_label.configure(text=str(exc))
                return
            user = {
                "username": username,
                "function": "Administrador",
                "must_change_password": False,
            }
        else:
            try:
                user = self.storage.authenticate_user(username, password)
            except ExcelStorageError as exc:
                self.error_label.configure(text=str(exc))
                return
            if user is None:
                self.error_label.configure(text="Usuário ou senha inválidos.")
                self.password_var.set("")
                self.password_entry.focus_set()
                return

        if user.get("must_change_password"):
            self.root.unbind("<Return>")
            PasswordChangeDialog(
                self.root,
                self.storage,
                user["username"],
                required=True,
                on_success=lambda: self._open_portal(user),
            )
            return
        self._open_portal(user)

    def _open_portal(self, user):
        user["must_change_password"] = False
        self.root.unbind("<Return>")
        self.frame.destroy()
        self.on_authenticated(user)


class PasswordChangeDialog:
    def __init__(self, parent, storage: ExcelStorage, username: str, required=False, on_success=None):
        self.storage = storage
        self.username = username
        self.required = required
        self.on_success = on_success
        self.window = tk.Toplevel(parent)
        self.window.title("Trocar senha")
        self.window.resizable(True, True)
        center_window(self.window, 480, 430, parent=parent, min_width=450, min_height=400)
        self.window.configure(bg=PAGE_BG)
        self.window.transient(parent)
        self.window.protocol("WM_DELETE_WINDOW", self._close)

        card = tk.Frame(self.window, bg=WHITE, highlightbackground=LINE, highlightthickness=1,
                        padx=28, pady=24)
        card.pack(fill="both", expand=True, padx=22, pady=22)
        title = "Crie sua nova senha" if required else "Alterar senha"
        tk.Label(card, text=title, bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        explanation = (
            "Para continuar, troque a senha provisória por uma senha pessoal."
            if required else "Informe sua senha atual e escolha uma nova senha."
        )
        tk.Label(card, text=explanation, bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 9), wraplength=390, justify="left").pack(anchor="w", pady=(5, 14))
        self.current_var = self._field(card, "SENHA ATUAL", secret=True)
        self.new_var = self._field(card, "NOVA SENHA (mínimo 8 caracteres)", secret=True)
        self.confirm_var = self._field(card, "CONFIRMAR NOVA SENHA", secret=True)
        self.error_label = tk.Label(card, text="", bg=WHITE, fg=RED,
                                    font=("Segoe UI", 9), wraplength=390, justify="left")
        self.error_label.pack(anchor="w", pady=(7, 5))
        tk.Button(card, text="Salvar nova senha", command=self.submit, bg=ORANGE, fg=WHITE,
                  activebackground=ORANGE_DARK, activeforeground=WHITE, relief="flat", bd=0,
                  padx=16, pady=10, font=("Segoe UI", 10, "bold"), cursor="hand2").pack(fill="x")
        if not required:
            tk.Button(card, text="Cancelar", command=self._close, bg=WHITE, fg=MUTED,
                      activebackground=BLUE_LIGHT, relief="flat", bd=0, padx=12, pady=7,
                      font=("Segoe UI", 9), cursor="hand2").pack(anchor="e", pady=(7, 0))
        self.window.bind("<Return>", lambda _event: self.submit())
        self.window.grab_set()
        self.current_entry.focus_set()

    def _field(self, parent, label, secret=False):
        tk.Label(parent, text=label, bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(7, 4))
        variable = tk.StringVar()
        entry = ttk.Entry(parent, textvariable=variable, show="•" if secret else "")
        entry.pack(fill="x", ipady=3)
        if label == "SENHA ATUAL":
            self.current_entry = entry
        return variable

    def submit(self):
        if self.new_var.get() != self.confirm_var.get():
            self.error_label.configure(text="As novas senhas não coincidem.")
            return
        if self.new_var.get() == self.current_var.get():
            self.error_label.configure(text="Escolha uma senha diferente da senha atual.")
            return
        try:
            self.storage.change_user_password(
                self.username, self.current_var.get(), self.new_var.get()
            )
        except ExcelStorageError as exc:
            self.error_label.configure(text=str(exc))
            return

        self.window.unbind("<Return>")
        self.window.grab_release()
        self.window.destroy()
        if self.on_success:
            self.on_success()

    def _close(self):
        if self.required:
            return
        self.window.grab_release()
        self.window.destroy()


class PortalApp:
    def __init__(self, root: tk.Tk, current_user: dict[str, str]):
        self.root = root
        self.current_user = current_user
        self.root.title("R.A - Equivalencias | Equivalência & Alunos")
        center_window(self.root, 1460, 930, min_width=1120, min_height=740)
        self.root.configure(bg=PAGE_BG)

        config = self._read_config()
        self.workbook_path = Path(config.get("workbook_path") or DEFAULT_WORKBOOK)
        configured_student_path = config.get("student_workbook_path")
        self.student_workbook_path = Path(
            configured_student_path or self.workbook_path.with_name("Base_Alunos.xlsx")
        )
        if self.student_workbook_path.resolve() == self.workbook_path.resolve():
            self.student_workbook_path = self.workbook_path.with_name("Base_Alunos.xlsx")
        self.auto_search = tk.BooleanVar(value=bool(config.get("auto_search", True)))
        self.validate_duplicate_offer = tk.BooleanVar(
            value=bool(config.get("validate_duplicate_offer", True))
        )
        try:
            self.school_year = int(config.get("school_year", date.today().year))
        except (TypeError, ValueError):
            self.school_year = date.today().year
        self.process_code_prefix = self._normalize_process_prefix(
            config.get("process_code_prefix", "PROC")
        ) or "PROC"
        self.settings_school_year_var = tk.StringVar(value=str(self.school_year))
        self.settings_process_prefix_var = tk.StringVar(value=self.process_code_prefix)
        self.storage = ExcelStorage(self.workbook_path)
        self.student_storage = StudentStorage(self.student_workbook_path)
        self.data_ready = False
        self.current_process_id: str | None = None
        self.process_labels: dict[str, str] = {}
        self.student_process_labels: dict[str, str] = {}
        self.student_vars: dict[str, tk.StringVar] = {}
        self.student_tree: ttk.Treeview | None = None
        self.student_records: list[dict[str, str]] = []
        self.current_student_id: str | None = None
        self.student_id_entry: ttk.Entry | None = None
        self.student_process_selector: ttk.Combobox | None = None
        self.form_vars: dict[str, tk.StringVar] = {}
        self.form_widgets: dict[str, tk.Widget] = {}
        self.nav_buttons: dict[str, tk.Button] = {}
        self.pages: dict[str, tk.Frame] = {}
        self.equiv_rows: list[dict[str, str]] = []
        self.dashboard_status_counts: dict[str, int] = {}
        self.dashboard_total_value: tk.Label | None = None
        self.dashboard_active_value: tk.Label | None = None
        self.dashboard_wait_value: tk.Label | None = None
        self.dashboard_wait_detail: tk.Label | None = None
        self.status_chart_canvas: tk.Canvas | None = None
        self._search_after_id = None

        self._configure_styles()
        self._build_shell()
        self._build_pages()
        self.show_page("home")
        try:
            legacy_students = self.storage.list_students()
            self.storage.initialize()
            self.student_storage.initialize()
            if legacy_students and not self.student_storage.list_students():
                self.student_storage.save_students(legacy_students)
            if not configured_student_path or self.student_workbook_path.resolve() != Path(configured_student_path).resolve():
                config["student_workbook_path"] = str(self.student_workbook_path.resolve())
                self._write_config(config)
            self.data_ready = True
            self._refresh_process_selector()
            self._refresh_student_list()
            self._refresh_student_process_selector()
            self._new_student_form()
            self._refresh_dashboard()
            self._update_path_labels()
            self._load_new_record()
        except (ExcelStorageError, OSError) as exc:
            self.root.after(
                150,
                lambda error=str(exc): messagebox.showerror("Planilha", error, parent=self.root),
            )

    @staticmethod
    def _read_config() -> dict:
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return {}

    @staticmethod
    def _normalize_process_prefix(value: str) -> str:
        return re.sub(r"[^A-Za-z0-9]+", "-", str(value or "")).strip("-").upper()

    @staticmethod
    def _write_config(config: dict):
        CONFIG_FILE.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    def _configure_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=PAGE_BG)
        style.configure("Card.TFrame", background=WHITE)
        style.configure("TLabel", background=PAGE_BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=WHITE, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Muted.Card.TLabel", background=WHITE, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Heading.Card.TLabel", background=WHITE, foreground=BLUE, font=("Segoe UI", 11, "bold"))
        style.configure("TEntry", padding=(8, 7), fieldbackground="#F8FAFC", foreground=TEXT)
        style.configure("TCombobox", padding=(7, 6), fieldbackground="#F8FAFC")
        style.configure("Treeview", rowheight=31, font=("Segoe UI", 9), background=WHITE, fieldbackground=WHITE)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background=BLUE, foreground=WHITE)
        style.map("Treeview.Heading", background=[("active", BLUE_DARK)])
        style.configure("TCheckbutton", background=WHITE, font=("Segoe UI", 9))
        style.configure("Vertical.TScrollbar", arrowsize=12)

    def _build_shell(self):
        self.shell = tk.Frame(self.root, bg=PAGE_BG)
        self.shell.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(self.shell, bg=BLUE_DARK, width=276)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=BLUE_DARK, padx=22, pady=22)
        brand.pack(fill="x")
        logo = tk.Label(brand, text="S", bg="#14588D", fg=ORANGE, font=("Segoe UI", 20, "bold"),
                        width=2, height=1, relief="flat", padx=5, pady=4)
        logo.pack(side="left", padx=(0, 12))
        brand_text = tk.Frame(brand, bg=BLUE_DARK)
        brand_text.pack(side="left", fill="x")
        tk.Label(brand_text, text="R.A - Equivalencias", bg=BLUE_DARK, fg=WHITE,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(brand_text, text="Equivalência & Alunos", bg=BLUE_DARK, fg="#BDD5EA",
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Frame(self.sidebar, bg="#28608B", height=1).pack(fill="x")

        nav = tk.Frame(self.sidebar, bg=BLUE_DARK, padx=12, pady=18)
        nav.pack(fill="x")
        nav_items = [
            ("cadastro", "♟  Cadastro & Busca Aluno"),
            ("alunos", "▣  Cadastro de Alunos"),
            ("requerimento", "▤  Requerimento Oficial"),
            ("equivalencia", "▦  Equivalência de UCs"),
            ("configuracoes", "⚙  Configurações (Planilha)"),
        ]
        nav_items[0], nav_items[1] = nav_items[1], nav_items[0]
        nav_items[1] = ("cadastro", "Solicitar Aproveitamento")
        nav_items.insert(0, ("home", "Dashboard"))
        for key, caption in nav_items:
            button = tk.Button(
                nav, text=caption, anchor="w", justify="left", cursor="hand2",
                command=lambda page=key: self.show_page(page), relief="flat", bd=0,
                padx=15, pady=13, bg=BLUE_DARK, fg="#D7E8F7", activebackground="#1B5B90",
                activeforeground=WHITE, font=("Segoe UI", 10, "bold"), wraplength=220,
            )
            button.pack(fill="x", pady=3)
            self.nav_buttons[key] = button

        status_box = tk.Frame(self.sidebar, bg="#073D6C", highlightbackground="#28608B", highlightthickness=1,
                              padx=14, pady=12)
        status_box.pack(side="bottom", fill="x", padx=15, pady=(0, 10))
        tk.Label(status_box, text="●  Origem da planilha", bg="#073D6C", fg="#A9D8C0",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        self.sidebar_path = tk.Label(status_box, text="", bg="#073D6C", fg="#C4D5E5",
                                     font=("Consolas", 8), wraplength=220, justify="left")
        self.sidebar_path.pack(anchor="w", pady=(5, 0))
        tk.Label(self.sidebar, text="R.A - Equivalencias  •  Python", bg=BLUE_DARK, fg="#AFC5D8",
                 font=("Segoe UI", 8), pady=13).pack(side="bottom", fill="x")

        self.main = tk.Frame(self.shell, bg=PAGE_BG)
        self.main.pack(side="left", fill="both", expand=True)
        self.topbar = tk.Frame(self.main, bg=WHITE, height=68, padx=28,
                               highlightbackground=LINE, highlightthickness=1)
        self.topbar.pack(fill="x")
        self.topbar.pack_propagate(False)
        left = tk.Frame(self.topbar, bg=WHITE)
        left.pack(side="left", fill="y")
        tk.Label(left, text="PROCESSO ATIVO", bg=WHITE, fg="#8A9AAF",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 14))
        self.process_selector = ttk.Combobox(left, state="readonly", width=43, font=("Segoe UI", 10, "bold"))
        self.process_selector.pack(side="left", pady=16)
        self.process_selector.bind("<<ComboboxSelected>>", self._on_process_selected)

        right = tk.Frame(self.topbar, bg=WHITE)
        right.pack(side="right", fill="y")
        self.connection_badge = tk.Label(right, text="●  Base pronta", bg="#E7F7EF", fg=GREEN,
                                          font=("Segoe UI", 9, "bold"), padx=12, pady=7)
        self.connection_badge.pack(side="left", pady=13)
        tk.Frame(right, bg=LINE, width=1, height=30).pack(side="left", padx=12, pady=18)
        identity = tk.Frame(right, bg=WHITE)
        identity.pack(side="left", padx=(0, 10))
        tk.Label(identity, text=self.current_user.get("username", ""), bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(identity, text=self.current_user.get("function", ""), bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8)).pack(anchor="w")
        self._button(right, "Alterar senha", self.change_my_password, "secondary").pack(side="left", padx=4)
        if self.current_user.get("function", "").strip().casefold() == "administrador":
            self._button(right, "Usuários", self.open_user_manager, "secondary").pack(side="left", padx=4)
        self._button(right, "Sair", self.logout, "secondary").pack(side="left", padx=(4, 0))

        self.page_area = tk.Frame(self.main, bg=PAGE_BG)
        self.page_area.pack(fill="both", expand=True)

    def _build_pages(self):
        self._build_student_page()
        self._build_cadastro_page()
        self._build_home_page()
        self._build_requerimento_page()
        self._build_equivalencia_page()
        self._build_settings_page()

    def _new_scroll_page(self, key: str):
        page = tk.Frame(self.page_area, bg=PAGE_BG)
        scroll = ScrollableFrame(page)
        scroll.pack(fill="both", expand=True)
        inner = scroll.inner
        inner.configure(padx=28, pady=24)
        self.pages[key] = page
        return inner

    def show_page(self, key: str):
        for page in self.pages.values():
            page.pack_forget()
        if key in self.pages:
            self.pages[key].pack(fill="both", expand=True)
        for nav_key, button in self.nav_buttons.items():
            selected = nav_key == key
            button.configure(
                bg="#155891" if selected else BLUE_DARK,
                fg=WHITE if selected else "#D7E8F7",
                highlightthickness=0,
            )
        if key == "requerimento":
            self._update_request_preview()
        elif key == "equivalencia":
            self._load_equivalences()
        elif key == "configuracoes":
            self._sync_settings_fields()
        elif key == "alunos":
            self._refresh_student_list()
        elif key == "home":
            self._refresh_dashboard()

    def _build_home_page(self):
        content = self._new_scroll_page("home")
        _, intro = self._card(content, padx=24, pady=20)
        self._section_title(
            intro,
            "Dashboard",
            "Acompanhe os status e veja há quantos dias está aberta a solicitação com maior espera.",
        )

        metrics = tk.Frame(content, bg=PAGE_BG)
        metrics.pack(fill="x", pady=(0, 18))
        for column in range(3):
            metrics.grid_columnconfigure(column, weight=1, uniform="dashboard_metric")
        self.dashboard_total_value = self._dashboard_metric(
            metrics, 0, "TOTAL DE SOLICITAÇÕES", "0", "Processos cadastrados"
        )
        self.dashboard_active_value = self._dashboard_metric(
            metrics, 1, "EM ANDAMENTO", "0", "Aguardando conclusão"
        )

        wait_card = tk.Frame(metrics, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        wait_card.grid(row=0, column=2, sticky="nsew", padx=(6, 0))
        wait_body = tk.Frame(wait_card, bg=WHITE, padx=18, pady=16)
        wait_body.pack(fill="both", expand=True)
        tk.Label(wait_body, text="MAIOR ESPERA EM ABERTO", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        self.dashboard_wait_value = tk.Label(
            wait_body, text="—", bg=WHITE, fg=ORANGE_DARK, font=("Segoe UI", 22, "bold")
        )
        self.dashboard_wait_value.pack(anchor="w", pady=(5, 0))
        self.dashboard_wait_detail = tk.Label(
            wait_body, text="Sem solicitações em aberto", bg=WHITE, fg=TEXT,
            font=("Segoe UI", 8), justify="left", anchor="w", wraplength=280,
        )
        self.dashboard_wait_detail.pack(fill="x", anchor="w", pady=(2, 0))

        _, chart_body = self._card(content, padx=22, pady=18)
        tk.Label(chart_body, text="Solicitações por status", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))
        self.status_chart_canvas = tk.Canvas(
            chart_body, height=210, bg=WHITE, highlightthickness=0, bd=0
        )
        self.status_chart_canvas.pack(fill="x", expand=True)
        self.status_chart_canvas.bind(
            "<Configure>", lambda event: self._draw_status_chart(event.width)
        )

    @staticmethod
    def _dashboard_metric(parent, column, title, value, subtitle):
        card = tk.Frame(parent, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        card.grid(
            row=0, column=column, sticky="nsew",
            padx=(0 if column == 0 else 6, 6),
        )
        body = tk.Frame(card, bg=WHITE, padx=18, pady=16)
        body.pack(fill="both", expand=True)
        tk.Label(body, text=title, bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        value_label = tk.Label(
            body, text=value, bg=WHITE, fg=BLUE, font=("Segoe UI", 24, "bold")
        )
        value_label.pack(anchor="w", pady=(5, 0))
        tk.Label(body, text=subtitle, bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
        return value_label

    @staticmethod
    def _parse_request_date(value):
        text = str(value or "").strip()
        if not text:
            return None
        for pattern in (
            "%d/%m/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f",
        ):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text).date()
        except ValueError:
            return None

    def _refresh_dashboard(self):
        if self.dashboard_total_value is None:
            return
        try:
            records = self.storage.list_processes() if self.data_ready else []
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            records = []

        closed_statuses = {
            "deferido", "indeferido", "concluido", "concluida", "cancelado",
            "concluído", "concluída",
            "cancelada", "encerrado", "encerrada", "arquivado", "arquivada",
        }
        counts = {}
        open_records = []
        for record in records:
            status = record.get("status_processo", "").strip() or "Sem status"
            counts[status] = counts.get(status, 0) + 1
            if status.casefold() not in closed_statuses:
                requested = self._parse_request_date(record.get("data_requerimento", ""))
                if requested:
                    days_waiting = max(0, (date.today() - requested).days)
                    open_records.append((days_waiting, record))

        self.dashboard_status_counts = counts
        self.dashboard_total_value.configure(text=str(len(records)))
        self.dashboard_active_value.configure(
            text=str(sum(
                count for status, count in counts.items()
                if status.casefold() not in closed_statuses
            ))
        )
        if open_records:
            days_waiting, record = max(open_records, key=lambda item: item[0])
            unit = "dia" if days_waiting == 1 else "dias"
            self.dashboard_wait_value.configure(text=f"{days_waiting} {unit}")
            detail = " · ".join(
                part for part in (
                    record.get("id_processo", "").strip(),
                    record.get("nome_aluno", "").strip(),
                    record.get("status_processo", "").strip(),
                ) if part
            )
            self.dashboard_wait_detail.configure(text=detail or "Processo sem identificação")
        else:
            self.dashboard_wait_value.configure(text="—")
            self.dashboard_wait_detail.configure(
                text="Sem solicitações em aberto com data válida"
            )
        self._draw_status_chart()

    def _draw_status_chart(self, width=None):
        canvas = self.status_chart_canvas
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete("all")
        width = width or canvas.winfo_width()
        if width < 100:
            width = 700
        preferred = CHOICES.get("status_processo", [])
        ordered = [status for status in preferred if status in self.dashboard_status_counts]
        ordered.extend(sorted(
            (status for status in self.dashboard_status_counts if status not in ordered),
            key=str.casefold,
        ))
        if not ordered:
            canvas.configure(height=100)
            canvas.create_text(
                12, 36, text="Ainda não há solicitações cadastradas.", anchor="w",
                fill=MUTED, font=("Segoe UI", 10),
            )
            return

        row_height = 38
        canvas.configure(height=max(100, len(ordered) * row_height + 12))
        max_count = max(self.dashboard_status_counts.values(), default=1)
        label_width = min(190, max(125, int(width * 0.25)))
        left = 12
        bar_left = left + label_width
        bar_right = max(bar_left + 80, width - 52)
        track_width = bar_right - bar_left
        for index, status in enumerate(ordered):
            count = self.dashboard_status_counts[status]
            y = 8 + index * row_height
            canvas.create_text(
                left, y + 10, text=status, anchor="w", fill=TEXT,
                font=("Segoe UI", 9),
            )
            canvas.create_rectangle(
                bar_left, y + 2, bar_right, y + 20, fill="#EAF0F6", outline=""
            )
            bar_width = max(3, int(track_width * count / max_count))
            color = ORANGE if status.casefold() in {
                "aguardando documentos", "sem status"
            } else BLUE
            canvas.create_rectangle(
                bar_left, y + 2, bar_left + bar_width, y + 20,
                fill=color, outline="",
            )
            canvas.create_text(
                bar_right + 10, y + 11, text=str(count), anchor="w",
                fill=TEXT, font=("Segoe UI", 9, "bold"),
            )

    def _card(self, parent, padx=22, pady=18):
        card = tk.Frame(parent, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x", pady=(0, 18))
        content = tk.Frame(card, bg=WHITE, padx=padx, pady=pady)
        content.pack(fill="x")
        return card, content

    def _section_title(self, parent, title: str, subtitle: str | None = None):
        tk.Label(parent, text=title, bg=WHITE, fg=BLUE, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(parent, text=subtitle, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

    def _build_student_page(self):
        content = self._new_scroll_page("alunos")
        _, search = self._card(content, padx=22, pady=16)
        self._section_title(
            search,
            "Cadastro de Alunos",
            "Cadastre alunos manualmente ou importe um relatório PDF como pré-cadastro.",
        )
        search_row = tk.Frame(search, bg=WHITE)
        search_row.pack(fill="x", pady=(14, 0))
        self.student_search_var = tk.StringVar()
        entry = ttk.Entry(search_row, textvariable=self.student_search_var)
        entry.pack(side="left", fill="x", expand=True, ipady=3)
        entry.bind("<Return>", lambda _event: self._refresh_student_list())
        self._button(search_row, "Buscar", self._refresh_student_list, "secondary").pack(side="left", padx=8)
        self._button(search_row, "Importar alunos (PDF)", self.import_students_pdf, "orange").pack(side="left")

        _, body = self._card(content, padx=18, pady=16)
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        list_panel = tk.Frame(body, bg=WHITE)
        list_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        form_panel = tk.Frame(body, bg=WHITE)
        form_panel.grid(row=0, column=1, sticky="nsew", padx=(16, 0))
        tk.Label(list_panel, text="Alunos pré-cadastrados", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 9))

        tree_frame = tk.Frame(list_panel, bg=WHITE)
        tree_frame.pack(fill="both", expand=True)
        columns = ("id", "nome", "rg", "cpf", "nascimento", "matricula")
        self.student_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)
        headings = [
            ("id", "ID", 105), ("nome", "Nome", 200), ("rg", "RG", 100),
            ("cpf", "CPF", 125), ("nascimento", "Data nasc.", 105),
            ("matricula", "Data matrícula", 115),
        ]
        for column, title, width in headings:
            self.student_tree.heading(column, text=title)
            self.student_tree.column(column, width=width, minwidth=70, anchor="w")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.student_tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.student_tree.xview)
        self.student_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.student_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.student_tree.bind("<<TreeviewSelect>>", self._on_student_selected)

        tk.Label(form_panel, text="Dados do aluno", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 9))
        student_fields = [
            ("id_aluno", "ID"), ("nome_aluno", "Nome"), ("rg", "RG"),
            ("cpf", "CPF"), ("data_nascimento", "Data de nascimento"),
            ("data_matricula", "Data de matrícula"),
        ]
        for field, label in student_fields:
            row = tk.Frame(form_panel, bg=WHITE)
            row.pack(fill="x", pady=5)
            tk.Label(row, text=label, bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold"),
                     width=20, anchor="w").pack(side="left", padx=(0, 7))
            variable = tk.StringVar()
            self.student_vars[field] = variable
            input_field = ttk.Entry(row, textvariable=variable)
            input_field.pack(side="left", fill="x", expand=True, ipady=2)
            if field == "id_aluno":
                self.student_id_entry = input_field

        actions = tk.Frame(form_panel, bg=WHITE)
        actions.pack(fill="x", pady=(18, 0))
        self._button(actions, "Novo aluno", self._new_student_form, "secondary").pack(side="left")
        self._button(actions, "Salvar pré-cadastro", self.save_student_form, "primary").pack(side="left", padx=7)
        self._button(actions, "Excluir", self.delete_student_form, "danger").pack(side="left")

    def _refresh_student_list(self, _event=None):
        if not self.data_ready or self.student_tree is None:
            return
        try:
            records = self.student_storage.find_students(self.student_search_var.get())
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        self.student_records = records
        for item in self.student_tree.get_children():
            self.student_tree.delete(item)
        for index, record in enumerate(records):
            self.student_tree.insert(
                "", "end", iid=str(index),
                values=(record.get("id_aluno", ""), record.get("nome_aluno", ""),
                        record.get("rg", ""), record.get("cpf", ""),
                        record.get("data_nascimento", ""), record.get("data_matricula", "")),
            )

    def _on_student_selected(self, _event=None):
        if self.student_tree is None:
            return
        selection = self.student_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        if 0 <= index < len(self.student_records):
            self._load_student_form(self.student_records[index])

    def _load_student_form(self, record: dict[str, str]):
        for field, variable in self.student_vars.items():
            variable.set(record.get(field, ""))
        self.current_student_id = record.get("id_aluno") or None
        if self.student_id_entry is not None:
            self.student_id_entry.configure(state="readonly")

    def _new_student_form(self):
        self.current_student_id = None
        if self.student_id_entry is not None:
            self.student_id_entry.configure(state="normal")
        for field, variable in self.student_vars.items():
            variable.set("")
        if self.data_ready:
            try:
                self.student_vars["id_aluno"].set(self.student_storage.next_student_id())
            except ExcelStorageError as exc:
                self._show_storage_error(exc)
        if self.student_tree is not None:
            self.student_tree.selection_remove(*self.student_tree.selection())

    def save_student_form(self):
        if not self.data_ready:
            messagebox.showerror("Alunos", "Conecte uma planilha Excel antes de salvar.", parent=self.root)
            return
        record = {field: variable.get().strip() for field, variable in self.student_vars.items()}
        if not record.get("nome_aluno"):
            messagebox.showwarning("Alunos", "Informe o nome do aluno.", parent=self.root)
            return
        try:
            record["id_aluno"] = self.student_storage.save_student(record)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        self.current_student_id = record["id_aluno"]
        self._refresh_student_list()
        self._refresh_student_process_selector()
        saved = self.student_storage.get_student(record["id_aluno"])
        if saved:
            self._load_student_form(saved)
        messagebox.showinfo("Alunos", "Pré-cadastro do aluno salvo.", parent=self.root)

    def delete_student_form(self):
        if not self.current_student_id:
            messagebox.showinfo("Alunos", "Selecione um aluno para excluir.", parent=self.root)
            return
        if not messagebox.askyesno("Excluir aluno", "Excluir este pré-cadastro?", parent=self.root):
            return
        try:
            self.student_storage.delete_student(self.current_student_id)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        self._new_student_form()
        self._refresh_student_list()
        self._refresh_student_process_selector()

    def import_students_pdf(self):
        if not self.data_ready:
            messagebox.showerror("Importar alunos", "Conecte uma planilha Excel antes de importar.", parent=self.root)
            return
        pdf_path = filedialog.askopenfilename(
            parent=self.root,
            title="Selecione o relatório PDF dos alunos",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if not pdf_path:
            return
        try:
            records = extract_students_data(pdf_path)
            students = [record for record in records if record.get("nome_aluno", "").strip()]
            if not students:
                raise PDFStudentImportError("Não encontrei nenhum nome de aluno para criar pré-cadastros.")
            student_ids = self.student_storage.save_students(students)
        except (PDFStudentImportError, ExcelStorageError) as exc:
            messagebox.showerror(
                "Importar alunos",
                f"{exc}\n\nNenhum pré-cadastro foi alterado nesta importação.",
                parent=self.root,
            )
            return
        self._refresh_student_list()
        self._refresh_student_process_selector()
        if student_ids:
            imported = self.student_storage.get_student(student_ids[-1])
            if imported:
                self._load_student_form(imported)
        messagebox.showinfo(
            "Importação concluída",
            f"Foram criados ou atualizados {len(student_ids)} pré-cadastro(s) na aba Alunos. "
            "Eles já podem ser selecionados na tela de aproveitamento.",
            parent=self.root,
        )

    def _build_cadastro_page(self):
        content = self._new_scroll_page("cadastro")
        _, heading = self._card(content, padx=22, pady=18)
        self._section_title(
            heading,
            "Solicitar Aproveitamento",
            "Localize um processo existente ou preencha os dados para registrar uma nova solicitação.",
        )
        _, search = self._card(content, padx=22, pady=18)
        tk.Label(search, text="⌕  PUXAR ALUNO POR ID DO PROCESSO OU MATRÍCULA", bg=WHITE,
                 fg=BLUE, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 9))
        search_row = tk.Frame(search, bg=WHITE)
        search_row.pack(fill="x")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self.search_entry.bind("<Return>", lambda _event: self.search_processes())
        ttk.Button(search_row, text="⌕  Puxar dados", command=self.search_processes).pack(side="left", padx=(9, 0), ipady=2)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        self.search_status = tk.Label(search, text="", bg=WHITE, fg=BLUE, font=("Segoe UI", 8), anchor="w")
        self.search_status.pack(fill="x", pady=(7, 0))
        self.db_path_label = tk.Label(search, text="Banco de dados ativo: —", bg=WHITE, fg=MUTED,
                                       font=("Consolas", 8), anchor="e")
        self.db_path_label.pack(fill="x", pady=(8, 0))

        form_card = tk.Frame(content, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        form_card.pack(fill="x", pady=(0, 18))
        form_head = tk.Frame(form_card, bg=BLUE, padx=22, pady=15)
        form_head.pack(fill="x")
        tk.Label(form_head, text="▣  Ficha do aluno e dados do processo", bg=BLUE, fg=WHITE,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        self.process_badge = tk.Label(form_head, text="Novo processo", bg="#2D6597", fg=WHITE,
                                      font=("Consolas", 9, "bold"), padx=10, pady=5)
        self.process_badge.pack(side="right")

        form_body = tk.Frame(form_card, bg=WHITE, padx=22, pady=20)
        form_body.pack(fill="x")
        student_pick = tk.Frame(form_body, bg=WHITE)
        student_pick.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        tk.Label(student_pick, text="Selecionar aluno pré-cadastrado", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 5))
        self.student_process_selector = ttk.Combobox(student_pick, state="readonly")
        self.student_process_selector.pack(fill="x")
        self.student_process_selector.bind("<<ComboboxSelected>>", self._on_process_student_selected)
        left = tk.Frame(form_body, bg=WHITE)
        right = tk.Frame(form_body, bg=WHITE)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 18))
        right.grid(row=1, column=1, sticky="nsew", padx=(18, 0))
        form_body.grid_columnconfigure(0, weight=1)
        form_body.grid_columnconfigure(1, weight=1)
        self._build_field_group(left, "IDENTIFICAÇÃO DO INTERESSADO", LEFT_FIELDS)
        self._build_field_group(right, "MATRÍCULA, OFERTA E FINANCEIRO", RIGHT_FIELDS)

        separator = tk.Frame(form_body, bg="#E8EDF2", height=1)
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(18, 14))
        tk.Label(form_body, text="Observações", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=3, column=0, columnspan=2, sticky="w")
        self.observations = tk.Text(form_body, height=3, wrap="word", relief="flat", bd=0,
                                    bg="#F8FAFC", fg=TEXT, font=("Segoe UI", 9),
                                    highlightthickness=1, highlightbackground=LINE, padx=10, pady=8)
        self.observations.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        buttons = tk.Frame(content, bg=PAGE_BG)
        buttons.pack(fill="x", pady=(0, 24))
        self._button(buttons, "＋  Novo cadastro", self._load_new_record, "secondary").pack(side="left")
        self._button(buttons, "Salvar cadastro", self.save_process, "primary").pack(side="left", padx=8)
        self._button(buttons, "Excluir", self.delete_process, "danger").pack(side="left")
        self._button(buttons, "Ver equivalência de UCs  →", lambda: self.show_page("equivalencia"), "secondary").pack(side="right")
        self._button(buttons, "Gerar requerimento", lambda: self.show_page("requerimento"), "orange").pack(side="right", padx=8)

    def _build_field_group(self, parent, heading: str, fields: list[str]):
        tk.Label(parent, text=heading, bg=WHITE, fg="#8494A8", font=("Segoe UI", 9, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 10))
        tk.Frame(parent, bg="#E8EDF2", height=1).pack(fill="x", pady=(0, 8))
        for field in fields:
            row = tk.Frame(parent, bg=WHITE)
            row.pack(fill="x", pady=4)
            label = tk.Label(row, text=FIELD_LABELS[field] + ":", bg=WHITE, fg="#425267",
                             font=("Segoe UI", 8), width=23, anchor="w")
            label.pack(side="left", padx=(0, 7))
            variable = tk.StringVar()
            self.form_vars[field] = variable
            if field in CHOICES:
                widget = ttk.Combobox(row, textvariable=variable, values=CHOICES[field], font=("Segoe UI", 8))
                if field == "status_processo":
                    widget.configure(state="readonly")
            else:
                widget = ttk.Entry(row, textvariable=variable, font=("Segoe UI", 8))
            widget.pack(side="left", fill="x", expand=True, ipady=2)
            self.form_widgets[field] = widget
            if field == "id_aluno":
                variable.trace_add(
                    "write",
                    lambda *_args, value=variable: self._sync_student_process_choice(value.get()),
                )

    def _build_requerimento_page(self):
        content = self._new_scroll_page("requerimento")
        top_card, top = self._card(content, padx=22, pady=16)
        left = tk.Frame(top, bg=WHITE)
        left.pack(side="left", fill="x", expand=True)
        self._section_title(left, "Visualização do requerimento", "Documento preenchido com os dados do processo selecionado")
        self._button(top, "Abrir versão para impressão", self.export_request, "orange").pack(side="right", padx=(10, 0))

        doc = tk.Frame(content, bg=WHITE, highlightbackground=LINE, highlightthickness=1, padx=46, pady=34)
        doc.pack(fill="x", padx=42, pady=(0, 26))
        tk.Label(doc, text="REQUERIMENTO", bg=WHITE, fg=BLUE, font=("Segoe UI", 16, "bold")).pack()
        tk.Label(doc, text="APROVEITAMENTO DE COMPETÊNCIAS PARA CONTINUIDADE DE ESTUDOS",
                 bg=WHITE, fg="#596B80", font=("Segoe UI", 9, "bold")).pack(pady=(4, 18))
        tk.Frame(doc, bg=LINE, height=2).pack(fill="x", pady=(0, 16))

        self.request_values: dict[str, tk.StringVar] = {}
        identity = tk.Frame(doc, bg=WHITE)
        identity.pack(fill="x")
        request_fields = [
            ("nome_aluno", "NOME"), ("cpf", "CPF"), ("rg", "RG"),
            ("telefone", "TELEFONE"), ("email", "E-MAIL"),
            ("situacao_interessado", "SITUAÇÃO DO INTERESSADO"), ("curso_origem", "TÍTULO DO CURSO"),
            ("ano_cursado", "ANO EM QUE CURSOU"), ("data_requerimento", "DATA DO REQUERIMENTO"),
            ("prazo_resposta_dias", "PRAZO PARA RESPOSTA"),
        ]
        for idx, (field, caption) in enumerate(request_fields):
            cell = tk.Frame(identity, bg=WHITE, padx=4, pady=8)
            cell.grid(row=idx // 2, column=idx % 2, sticky="ew")
            tk.Label(cell, text=caption, bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            var = tk.StringVar(value="—")
            self.request_values[field] = var
            tk.Label(cell, textvariable=var, bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold"),
                     wraplength=390, justify="left").pack(anchor="w", pady=(3, 0))
        identity.grid_columnconfigure(0, weight=1)
        identity.grid_columnconfigure(1, weight=1)

        tk.Frame(doc, bg=LINE, height=1).pack(fill="x", pady=(10, 12))
        tk.Label(doc, text="OBSERVAÇÕES", bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        self.request_observations = tk.Label(doc, text="—", bg="#F7F9FB", fg=TEXT,
                                             font=("Consolas", 9), justify="left", anchor="w",
                                             padx=12, pady=12, wraplength=840)
        self.request_observations.pack(fill="x", pady=(5, 18))
        signatures = tk.Frame(doc, bg=WHITE)
        signatures.pack(fill="x", pady=(24, 30))
        for title in ("Visto da Secretaria Escolar", "Visto do Interessado"):
            block = tk.Frame(signatures, bg=WHITE, padx=18)
            block.pack(side="left", fill="x", expand=True)
            tk.Frame(block, bg="#5B6572", height=1).pack(fill="x", pady=(18, 5))
            tk.Label(block, text=title, bg=WHITE, fg=TEXT, font=("Segoe UI", 8)).pack()
        tk.Frame(doc, bg=LINE, height=2).pack(fill="x", pady=(4, 15))
        tk.Label(doc, text="CANHOTO / PROTOCOLO", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack()
        self.request_protocol = tk.Label(doc, text="", bg=WHITE, fg=TEXT,
                                         font=("Segoe UI", 9), justify="center", pady=12)
        self.request_protocol.pack(fill="x")

    def _build_equivalencia_page(self):
        content = self._new_scroll_page("equivalencia")
        _, head = self._card(content, padx=22, pady=18)
        title_area = tk.Frame(head, bg=WHITE)
        title_area.pack(side="left", fill="x", expand=True)
        tk.Label(title_area, text="FORMULÁRIO OFICIAL DE EQUIVALÊNCIA", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        self._section_title(title_area, "Formulário Equivalência – Alteração de PC")
        self.eq_process_label = tk.Label(head, text="Sem processo selecionado", bg="#FFF3E9", fg=ORANGE_DARK,
                                         font=("Consolas", 9, "bold"), padx=12, pady=8)
        self.eq_process_label.pack(side="right")

        _, table_area = self._card(content, padx=0, pady=0)
        columns = ["modulo_uc_cursado", "nome_uc_cursado", "mencao", "numero_uc_destino",
                   "nome_uc_destino", "numero_aula", "id_curso"]
        self.eq_tree = ttk.Treeview(table_area, columns=columns, show="headings", height=12)
        widths = [125, 260, 95, 115, 280, 125, 100]
        for field, width in zip(columns, widths):
            self.eq_tree.heading(field, text=EQ_LABELS[field])
            self.eq_tree.column(field, width=width, minwidth=80, stretch=True)
        self.eq_tree.pack(fill="x", padx=1, pady=1)
        self.eq_tree.bind("<Double-1>", lambda _event: self.edit_equivalence())
        self.eq_total = tk.Label(table_area, text="0 unidades curriculares", bg="#F7F9FB", fg=MUTED,
                                 font=("Segoe UI", 9), padx=16, pady=10, anchor="w")
        self.eq_total.pack(fill="x", pady=(1, 0))

        actions = tk.Frame(content, bg=PAGE_BG)
        actions.pack(fill="x", pady=(0, 24))
        self._button(actions, "＋  Adicionar UC", self.add_equivalence, "primary").pack(side="left")
        self._button(actions, "Editar selecionada", self.edit_equivalence, "secondary").pack(side="left", padx=8)
        self._button(actions, "Remover", self.delete_equivalence, "danger").pack(side="left")
        self._button(actions, "Exportar mapeamento", self.export_equivalences, "orange").pack(side="right")

    def _build_settings_page(self):
        content = self._new_scroll_page("configuracoes")
        card = tk.Frame(content, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x", padx=80, pady=(0, 20))
        body = tk.Frame(card, bg=WHITE, padx=32, pady=28)
        body.pack(fill="x")
        tk.Label(body, text="▦", bg=BLUE_LIGHT, fg=BLUE, font=("Segoe UI", 22, "bold"),
                 padx=12, pady=6).pack(anchor="w")
        tk.Label(body, text="Configuração da planilha Excel", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 17, "bold")).pack(anchor="w", pady=(14, 4))
        tk.Label(body, text="Selecione a pasta de trabalho usada para armazenar processos e equivalências.",
                 bg=WHITE, fg=MUTED, font=("Segoe UI", 9), wraplength=650, justify="left").pack(anchor="w")

        code_settings = tk.Frame(body, bg=WHITE)
        code_settings.pack(fill="x", pady=(20, 4))
        code_settings.grid_columnconfigure(0, weight=1)
        code_settings.grid_columnconfigure(1, weight=2)
        tk.Label(code_settings, text="ANO LETIVO", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 14))
        tk.Label(code_settings, text="INÍCIO DO CÓDIGO DO PROCESSO (PREFIXO)", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w")
        ttk.Entry(code_settings, textvariable=self.settings_school_year_var, width=10).grid(
            row=1, column=0, sticky="ew", padx=(0, 14), pady=(6, 0), ipady=3
        )
        ttk.Entry(code_settings, textvariable=self.settings_process_prefix_var).grid(
            row=1, column=1, sticky="ew", pady=(6, 0), ipady=3
        )
        tk.Label(body, text="Novos IDs serão gerados no formato PREFIXO-ANO-001, por exemplo PROC-2026-001.",
                 bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 12))

        tk.Label(body, text="CAMINHO DA PLANILHA (.XLSX)", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(26, 7))
        path_row = tk.Frame(body, bg=WHITE)
        path_row.pack(fill="x")
        self.settings_path_var = tk.StringVar()
        self.settings_path_entry = ttk.Entry(path_row, textvariable=self.settings_path_var, font=("Consolas", 9))
        self.settings_path_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self._button(path_row, "Procurar…", self.browse_workbook, "secondary").pack(side="left", padx=(8, 0))
        tk.Label(body, text="A planilha selecionada fica salva para a próxima vez que abrir o sistema.",
                 bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 16))

        tk.Label(body, text="BASE SEPARADA DE ALUNOS (.XLSX)", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 7))
        student_path_row = tk.Frame(body, bg=WHITE)
        student_path_row.pack(fill="x")
        self.settings_student_path_var = tk.StringVar()
        self.settings_student_path_entry = ttk.Entry(
            student_path_row, textvariable=self.settings_student_path_var, font=("Consolas", 9)
        )
        self.settings_student_path_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self._button(
            student_path_row, "Procurar...", self.browse_student_workbook, "secondary"
        ).pack(side="left", padx=(8, 0))
        tk.Label(body, text="Os alunos ficam neste arquivo separado; o v\u00ednculo com o processo usa o ID do processo.",
                 bg=WHITE, fg=MUTED, font=("Segoe UI", 8), wraplength=650,
                 justify="left").pack(anchor="w", pady=(6, 16))

        options = tk.Frame(body, bg="#F7F9FB", highlightbackground=LINE, highlightthickness=1, padx=14, pady=12)
        options.pack(fill="x")
        ttk.Checkbutton(options, text="Pesquisar automaticamente enquanto digito", variable=self.auto_search).pack(anchor="w", pady=3)
        ttk.Checkbutton(options, text="Avisar se aluno e oferta já constarem em outro processo",
                        variable=self.validate_duplicate_offer).pack(anchor="w", pady=3)

        btn_row = tk.Frame(body, bg=WHITE)
        btn_row.pack(fill="x", pady=(20, 0))
        self._button(btn_row, "Salvar e conectar", self.save_settings, "orange").pack(side="left", fill="x", expand=True)
        self._button(btn_row, "Padrão", self.reset_settings, "secondary").pack(side="left", padx=(9, 0))

    def open_user_manager(self):
        if self.current_user.get("function", "").strip().casefold() != "administrador":
            messagebox.showwarning("Usuários", "Somente um administrador pode cadastrar usuários.", parent=self.root)
            return

        window = tk.Toplevel(self.root)
        window.title("Usuários do sistema")
        center_window(window, 700, 610, parent=self.root, min_width=620, min_height=540)
        window.configure(bg=PAGE_BG)
        window.transient(self.root)

        content = tk.Frame(window, bg=PAGE_BG, padx=24, pady=22)
        content.pack(fill="both", expand=True)
        tk.Label(content, text="Cadastro de usuários", bg=PAGE_BG, fg=BLUE,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(content, text="As senhas são gravadas na aba Usuarios como hashes protegidos.",
                 bg=PAGE_BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(3, 14))

        table_frame = tk.Frame(content, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        table_frame.pack(fill="both", expand=True, pady=(0, 16))
        tree = ttk.Treeview(table_frame, columns=("username", "function"), show="headings", height=9)
        tree.heading("username", text="Usuário")
        tree.heading("function", text="Função")
        tree.column("username", width=220, anchor="w")
        tree.column("function", width=350, anchor="w")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", padx=(0, 8), pady=8)

        def refresh_users():
            tree.delete(*tree.get_children())
            for user in self.storage.list_users():
                tree.insert("", "end", values=(user["username"], user["function"]))

        refresh_users()
        form = tk.Frame(content, bg=WHITE, highlightbackground=LINE, highlightthickness=1,
                        padx=16, pady=14)
        form.pack(fill="x")
        tk.Label(form, text="Novo usuário", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        username_var = tk.StringVar()
        function_var = tk.StringVar()
        password_var = tk.StringVar()
        confirm_var = tk.StringVar()
        ttk.Label(form, text="Usuário", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12))
        ttk.Label(form, text="Função", style="Card.TLabel").grid(row=1, column=1, sticky="w")
        ttk.Entry(form, textvariable=username_var).grid(row=2, column=0, sticky="ew", padx=(0, 12), pady=(4, 8))
        ttk.Combobox(form, textvariable=function_var, values=(
            "Administrador", "Secretaria Escolar", "Coordenação Pedagógica", "Consulta", "Outro"
        )).grid(row=2, column=1, sticky="ew", pady=(4, 8))
        ttk.Label(form, text="Senha (mínimo 8 caracteres)", style="Card.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 12))
        ttk.Label(form, text="Confirmar senha", style="Card.TLabel").grid(row=3, column=1, sticky="w")
        ttk.Entry(form, textvariable=password_var, show="•").grid(row=4, column=0, sticky="ew", padx=(0, 12), pady=(4, 8))
        ttk.Entry(form, textvariable=confirm_var, show="•").grid(row=4, column=1, sticky="ew", pady=(4, 8))
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        error_label = tk.Label(form, text="", bg=WHITE, fg=RED, font=("Segoe UI", 9), anchor="w")
        error_label.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        def create_account():
            if password_var.get() != confirm_var.get():
                error_label.configure(text="As senhas não coincidem.")
                return
            try:
                self.storage.create_user(username_var.get(), password_var.get(), function_var.get())
            except ExcelStorageError as exc:
                error_label.configure(text=str(exc))
                return
            refresh_users()
            username_var.set("")
            function_var.set("")
            password_var.set("")
            confirm_var.set("")
            error_label.configure(text="Usuário cadastrado.", fg=GREEN)

        self._button(form, "Cadastrar usuário", create_account, "primary").grid(
            row=6, column=0, columnspan=2, sticky="ew", pady=(4, 0)
        )
        window.grab_set()

    def change_my_password(self):
        PasswordChangeDialog(
            self.root,
            self.storage,
            self.current_user["username"],
            on_success=lambda: messagebox.showinfo(
                "Senha", "Sua senha foi alterada.", parent=self.root
            ),
        )

    def logout(self):
        for child in self.root.winfo_children():
            child.destroy()
        LoginScreen(self.root, self.storage, self._launch_authenticated)

    def _launch_authenticated(self, user):
        self.root.portal_app = PortalApp(self.root, user)

    def _button(self, parent, text, command, kind="primary"):
        palette = {
            "primary": (BLUE, WHITE, BLUE_DARK),
            "orange": (ORANGE, WHITE, ORANGE_DARK),
            "secondary": ("#EAF0F5", "#344256", "#DEE7EF"),
            "danger": ("#FCE9E7", RED, "#F7D5D1"),
        }
        bg, fg, active = palette[kind]
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=active,
                         activeforeground=fg, relief="flat", bd=0, padx=15, pady=9,
                         font=("Segoe UI", 9, "bold"), cursor="hand2")

    def _update_path_labels(self):
        self.sidebar_path.configure(
            text=f"Processos: {self.workbook_path.name}\nAlunos: {self.student_workbook_path.name}"
        )
        self.db_path_label.configure(
            text=f"Base de processos: {self.workbook_path}\nBase de alunos: {self.student_workbook_path}"
        )

    def _refresh_student_process_selector(self):
        if self.student_process_selector is None or not self.data_ready:
            return
        try:
            records = self.student_storage.list_students()
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            records = []
        self.student_process_labels.clear()
        labels = []
        for record in records:
            student_id = record.get("id_aluno", "").strip()
            if not student_id:
                continue
            name = record.get("nome_aluno", "").strip()
            label = f"{student_id} - {name}" if name else student_id
            labels.append(label)
            self.student_process_labels[label] = student_id
        self.student_process_selector.configure(values=["Selecione um pré-cadastro", *labels])
        self._sync_student_process_choice(self.form_vars["id_aluno"].get())

    def _sync_student_process_choice(self, student_id: str):
        if self.student_process_selector is None:
            return
        label = next(
            (label for label, value in self.student_process_labels.items()
             if value.casefold() == str(student_id or "").strip().casefold()),
            "Selecione um pré-cadastro",
        )
        self.student_process_selector.set(label)

    def _on_process_student_selected(self, _event=None):
        if self.student_process_selector is None:
            return
        student_id = self.student_process_labels.get(self.student_process_selector.get())
        if not student_id:
            return
        try:
            student = self.student_storage.get_student(student_id)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        if not student:
            return
        for field in ("id_aluno", "nome_aluno", "rg", "cpf", "telefone", "email", "curso_origem", "ano_cursado"):
            if field in self.form_vars:
                self.form_vars[field].set(student.get(field, ""))
        self._update_request_preview()
        self.search_status.configure(text="Pré-cadastro preenchido. Salve o processo para registrar o vínculo.", fg=GREEN)

    def _refresh_process_selector(self):
        try:
            records = self.storage.list_processes()
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            records = []
        labels = []
        self.process_labels.clear()
        for record in records:
            process_id = record.get("id_processo", "").strip()
            if not process_id:
                continue
            name = record.get("nome_aluno", "").strip()
            label = f"{process_id}  —  {name}" if name else process_id
            labels.append(label)
            self.process_labels[label] = process_id
        new_label = "Novo processo (não salvo)"
        self.process_selector.configure(values=[new_label, *labels])
        wanted = self.current_process_id
        selected_label = next((label for label, pid in self.process_labels.items() if pid == wanted), None)
        if selected_label:
            self.process_selector.set(selected_label)
        else:
            self.process_selector.set(new_label)

    def _on_process_selected(self, _event=None):
        label = self.process_selector.get()
        if label == "Novo processo (não salvo)":
            self._load_new_record()
            return
        process_id = self.process_labels.get(label)
        if not process_id:
            return
        record = self.storage.get_process(process_id)
        if record:
            self._set_form(record)
            self.current_process_id = process_id
            self._refresh_process_selector()
            self._update_request_preview()
            if "equivalencia" in self.pages:
                self._load_equivalences()

    def _on_search_key(self, _event=None):
        if not self.auto_search.get():
            return
        if self._search_after_id:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(450, lambda: self.search_processes(auto=True))

    def search_processes(self, auto=False):
        query = self.search_var.get().strip()
        if not query:
            self.search_status.configure(text="")
            return
        try:
            matches = self.storage.find_processes(query)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        if not matches:
            self.search_status.configure(text="Nenhum processo encontrado.", fg=MUTED)
            if not auto:
                messagebox.showinfo("Busca", "Nenhum processo encontrado para essa busca.", parent=self.root)
        elif len(matches) == 1:
            self.search_status.configure(text="1 processo encontrado.", fg=GREEN)
            self.current_process_id = matches[0].get("id_processo", "")
            self._set_form(matches[0])
            self._refresh_process_selector()
            self._update_request_preview()
            self._load_equivalences()
        elif auto:
            self.search_status.configure(text=f"{len(matches)} processos encontrados. Clique em “Puxar dados” para escolher.", fg=BLUE)
        else:
            self.search_status.configure(text=f"{len(matches)} processos encontrados.", fg=BLUE)
            self._choose_search_result(matches)

    def _choose_search_result(self, records):
        dialog = tk.Toplevel(self.root)
        dialog.title("Resultados da busca")
        center_window(dialog, 760, 390, parent=self.root, min_width=620, min_height=320)
        dialog.transient(self.root)
        dialog.grab_set()
        tk.Label(dialog, text="Selecione o processo para abrir", font=("Segoe UI", 12, "bold"),
                 fg=BLUE, padx=16, pady=14).pack(anchor="w")
        tree = ttk.Treeview(dialog, columns=("id", "nome", "matricula", "curso"), show="headings", height=10)
        for col, title, width in [("id", "ID do processo", 145), ("nome", "Nome do aluno", 280),
                                  ("matricula", "ID / matrícula", 135), ("curso", "Curso de origem", 180)]:
            tree.heading(col, text=title)
            tree.column(col, width=width)
        for idx, record in enumerate(records):
            tree.insert("", "end", iid=str(idx), values=(record.get("id_processo", ""),
                        record.get("nome_aluno", ""), record.get("id_aluno", ""), record.get("curso_origem", "")))
        tree.pack(fill="both", expand=True, padx=16)

        def choose(_event=None):
            selection = tree.selection()
            if not selection:
                return
            record = records[int(selection[0])]
            self.current_process_id = record.get("id_processo", "")
            self._set_form(record)
            self._refresh_process_selector()
            self._update_request_preview()
            self._load_equivalences()
            dialog.destroy()

        tree.bind("<Double-1>", choose)
        self._button(dialog, "Abrir processo", choose, "primary").pack(anchor="e", padx=16, pady=12)

    def _set_form(self, record: dict[str, str]):
        for field, variable in self.form_vars.items():
            variable.set(record.get(field, ""))
        self.observations.delete("1.0", "end")
        self.observations.insert("1.0", record.get("observacoes", ""))
        self.current_process_id = record.get("id_processo") or None
        self.process_badge.configure(text=self.current_process_id or "Novo processo")
        widget = self.form_widgets.get("id_processo")
        if widget:
            widget.configure(state="readonly" if self.current_process_id else "normal")
        self._sync_student_process_choice(record.get("id_aluno", ""))

    def _load_new_record(self):
        record = {field: "" for _heading, field in PROCESS_COLUMNS}
        year = self.school_year
        prefix = self.process_code_prefix
        if self.data_ready:
            try:
                record["id_processo"] = self.storage.next_process_id(year, prefix)
            except ExcelStorageError:
                record["id_processo"] = f"{prefix}-{year}-001"
        else:
            record["id_processo"] = f"{prefix}-{year}-001"
        record["data_requerimento"] = date.today().strftime("%d/%m/%Y")
        record["status_processo"] = "Recebido"
        record["prazo_resposta_dias"] = "20 dias úteis"
        self._set_form(record)
        self.current_process_id = None
        self.process_badge.configure(text="Novo processo")
        self.form_widgets["id_processo"].configure(state="normal")
        self.search_var.set("")
        self.search_status.configure(text="")
        self.process_selector.set("Novo processo (não salvo)")
        self._update_request_preview()
        self._load_equivalences()

    def _collect_form(self):
        record = {field: variable.get().strip() for field, variable in self.form_vars.items()}
        record["observacoes"] = self.observations.get("1.0", "end-1c").strip()
        return record

    def save_process(self):
        record = self._collect_form()
        process_id = record.get("id_processo", "").strip()
        if not process_id:
            messagebox.showwarning("Cadastro", "Informe o ID do processo.", parent=self.root)
            return
        if not record.get("nome_aluno"):
            messagebox.showwarning("Cadastro", "Informe o nome do aluno.", parent=self.root)
            return
        if not self.data_ready:
            messagebox.showerror("Planilha", "Conecte uma planilha Excel em Configurações antes de salvar.", parent=self.root)
            return
        try:
            record["id_aluno"] = self.student_storage.resolve_student_id(record)
            previous_process = (
                self.storage.get_process(self.current_process_id)
                if self.current_process_id else None
            )
            records = self.storage.list_processes()
            duplicate_id = next((r for r in records if r.get("id_processo", "").casefold() == process_id.casefold()
                                 and r.get("id_processo", "").casefold() != (self.current_process_id or "").casefold()), None)
            if duplicate_id:
                messagebox.showwarning("Cadastro", "Esse ID de processo já existe. Gere outro ID para o novo cadastro.", parent=self.root)
                return
            if self.validate_duplicate_offer.get() and record.get("id_aluno") and record.get("codigo_oferta"):
                duplicate = next((r for r in records
                                  if r.get("id_processo", "").casefold() != (self.current_process_id or "").casefold()
                                  and r.get("id_aluno", "").strip().casefold() == record["id_aluno"].casefold()
                                  and r.get("codigo_oferta", "").strip().casefold() == record["codigo_oferta"].casefold()), None)
                if duplicate:
                    answer = messagebox.askyesno(
                        "Possível duplicidade",
                        f"Este aluno já tem a oferta em {duplicate.get('id_processo', '')}.\nDeseja salvar mesmo assim?",
                        parent=self.root,
                    )
                    if not answer:
                        return
            self.storage.save_process(record, original_id=self.current_process_id)
            self.student_storage.save_process_student(record, process_id)
            if previous_process:
                old_student_id = previous_process.get("id_aluno", "").strip()
                old_process_id = previous_process.get("id_processo", "").strip()
                if old_student_id and (
                    old_student_id.casefold() != record.get("id_aluno", "").strip().casefold()
                    or old_process_id.casefold() != process_id.casefold()
                ):
                    self.student_storage.remove_process_link(old_student_id, old_process_id)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        self.current_process_id = process_id
        self._set_form(record)
        self._refresh_process_selector()
        self._refresh_student_process_selector()
        self._update_request_preview()
        self._load_equivalences()
        messagebox.showinfo(
            "Cadastro",
            "Processo salvo na base de equivalências e vínculo atualizado na base separada de alunos.",
            parent=self.root,
        )

    def delete_process(self):
        if not self.current_process_id:
            messagebox.showinfo("Excluir", "Selecione um processo salvo primeiro.", parent=self.root)
            return
        if not messagebox.askyesno("Excluir processo", f"Excluir {self.current_process_id} e suas equivalências?", parent=self.root):
            return
        try:
            self.storage.delete_process(self.current_process_id)
            self.student_storage.unlink_process(self.current_process_id)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            return
        self.current_process_id = None
        self._refresh_process_selector()
        self._load_new_record()
        messagebox.showinfo("Excluir", "Processo excluído da planilha.", parent=self.root)

    def _update_request_preview(self):
        if not hasattr(self, "request_values"):
            return
        if self.form_vars:
            record = self._collect_form()
        elif self.current_process_id and self.data_ready:
            record = self.storage.get_process(self.current_process_id) or {}
        else:
            record = {}
        for field, variable in self.request_values.items():
            variable.set(record.get(field, "").strip() or "—")
        self.request_observations.configure(text=record.get("observacoes", "").strip() or "—")
        name = record.get("nome_aluno", "").strip() or "—"
        course = record.get("curso_origem", "").strip() or "—"
        process_id = record.get("id_processo", "").strip() or "—"
        self.request_protocol.configure(text=f"{name}\nCurso: {course}   •   Processo: {process_id}")

    def export_request(self):
        record = self._collect_form()
        if not record.get("id_processo") or not record.get("nome_aluno"):
            messagebox.showwarning("Requerimento", "Selecione ou salve um processo com o nome do aluno.", parent=self.root)
            return
        suggested = re.sub(r"[^A-Za-z0-9_-]+", "_", record.get("id_processo", "requerimento"))
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Salvar requerimento para impressão",
            initialfile=f"Requerimento_{suggested}.html",
            defaultextension=".html", filetypes=[("Página HTML", "*.html")],
        )
        if not path:
            return
        try:
            Path(path).write_text(self._request_html(record), encoding="utf-8")
            webbrowser.open(Path(path).resolve().as_uri())
        except OSError as exc:
            messagebox.showerror("Requerimento", f"Não foi possível criar o arquivo:\n{exc}", parent=self.root)

    def _request_html(self, record: dict[str, str]) -> str:
        def val(field):
            return html.escape(record.get(field, "").strip() or "—")

        situation = record.get("situacao_interessado", "").strip().casefold()
        choices = ["Ex-aluno Senac", "Aluno de curso em andamento", "Outra instituição de ensino", "Outros"]
        checkbox_html = "".join(
            f"<li>{'[ X ]' if situation == choice.casefold() else '[   ]'} {html.escape(choice)}</li>"
            for choice in choices
        )
        observations = html.escape(record.get("observacoes", "").strip() or "—").replace("\n", "<br>")
        return f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Requerimento {val('id_processo')}</title>
<style>
body{{font:14px Arial,sans-serif;color:#1e293b;margin:0;background:#eef2f6}} .page{{max-width:850px;margin:30px auto;background:white;padding:48px 58px;box-shadow:0 2px 12px #b7c1ca}}
h1{{color:#004a8d;text-align:center;font-size:21px;margin:0}} .subtitle{{text-align:center;font-weight:bold;color:#66758a;font-size:11px;margin:7px 0 20px}}
hr{{border:0;border-top:2px solid #dce3e9;margin:18px 0}} .grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px 28px}} .item strong{{display:block;color:#748397;font-size:10px;margin-bottom:5px}} .item span{{font-weight:bold;white-space:pre-wrap}}
.section{{margin-top:20px}} .section h2{{font-size:12px;color:#334155}} .obs{{background:#f7f9fb;border:1px solid #dce3e9;padding:14px;line-height:1.5}}
.signatures{{display:grid;grid-template-columns:1fr 1fr;gap:42px;text-align:center;margin:55px 0 40px;font-size:12px}} .signatures div{{border-top:1px solid #334155;padding-top:7px}} .protocol{{border-top:2px dashed #ccd5dd;padding-top:20px;text-align:center}}
.actions{{max-width:850px;margin:0 auto 26px;text-align:right}} button{{border:0;background:#f47920;color:white;padding:11px 18px;border-radius:7px;font-weight:bold;cursor:pointer}}
@media print{{body{{background:white}} .page{{margin:0 auto;box-shadow:none;padding:20mm 15mm}} .actions{{display:none}}}}
</style></head><body><div class="actions"><button onclick="window.print()">Imprimir / Salvar como PDF</button></div>
<main class="page"><h1>REQUERIMENTO</h1><div class="subtitle">APROVEITAMENTO DE COMPETÊNCIAS PARA CONTINUIDADE DE ESTUDOS</div><hr>
<div class="grid">
<div class="item"><strong>NOME</strong><span>{val('nome_aluno')}</span></div><div class="item"><strong>CPF</strong><span>{val('cpf')}</span></div>
<div class="item"><strong>RG</strong><span>{val('rg')}</span></div><div class="item"><strong>TELEFONE</strong><span>{val('telefone')}</span></div>
<div class="item"><strong>E-MAIL</strong><span>{val('email')}</span></div><div class="item"><strong>CURSO</strong><span>{val('curso_origem')}</span></div>
<div class="item"><strong>ANO EM QUE CURSOU</strong><span>{val('ano_cursado')}</span></div><div class="item"><strong>DATA DO REQUERIMENTO</strong><span>{val('data_requerimento')}</span></div>
</div><div class="section"><h2>SITUAÇÃO DO INTERESSADO</h2><ul>{checkbox_html}</ul></div>
<div class="section"><h2>OBSERVAÇÕES</h2><div class="obs">{observations}</div></div>
<div class="grid section"><div class="item"><strong>PRAZO PARA RESPOSTA</strong><span>{val('prazo_resposta_dias')}</span></div><div class="item"><strong>ID DO PROCESSO</strong><span>{val('id_processo')}</span></div></div>
<div class="signatures"><div>Visto da Secretaria Escolar</div><div>Visto do Interessado</div></div>
<div class="protocol"><strong>CANHOTO / PROTOCOLO</strong><p>Requerimento de aproveitamento para fins de continuidade de estudos</p>
<p>{val('nome_aluno')} · {val('curso_origem')} · {val('data_requerimento')}</p><p>Processo: {val('id_processo')}</p></div>
</main></body></html>"""

    def _load_equivalences(self):
        if not hasattr(self, "eq_tree"):
            return
        for item in self.eq_tree.get_children():
            self.eq_tree.delete(item)
        process_id = self.current_process_id or ""
        self.eq_process_label.configure(text=process_id or "Sem processo selecionado")
        if not process_id or not self.data_ready:
            self.equiv_rows = []
            self.eq_total.configure(text="Selecione um processo salvo para consultar as equivalências.")
            return
        try:
            self.equiv_rows = self.storage.list_equivalences(process_id)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            self.equiv_rows = []
        for idx, row in enumerate(self.equiv_rows):
            self.eq_tree.insert("", "end", iid=str(idx), values=tuple(row.get(field, "") for field in EQ_FIELDS))
        count = len(self.equiv_rows)
        self.eq_total.configure(text=f"Total de {count} unidade(s) curricular(es) vinculada(s) a {process_id}.")

    def _equivalence_dialog(self, initial: dict[str, str] | None = None):
        if not self.current_process_id:
            messagebox.showinfo("Equivalência", "Salve ou selecione um processo antes de cadastrar UCs.", parent=self.root)
            return None
        dialog = tk.Toplevel(self.root)
        dialog.title("Equivalência de UC")
        center_window(dialog, 660, 560, parent=self.root, min_width=520, min_height=460)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=WHITE)
        tk.Label(dialog, text="Dados da equivalência", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 15, "bold"), padx=24, pady=20).pack(anchor="w")
        form = tk.Frame(dialog, bg=WHITE, padx=24)
        form.pack(fill="both", expand=True)
        variables = {}
        for field in EQ_FIELDS:
            tk.Label(form, text=EQ_LABELS[field], bg=WHITE, fg=TEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 3))
            var = tk.StringVar(value=(initial or {}).get(field, ""))
            variables[field] = var
            if field in {"modulo_uc_cursado", "numero_uc_destino"}:
                widget = ttk.Combobox(
                    form,
                    textvariable=var,
                    values=UC_OPTIONS,
                    state="readonly",
                    font=("Segoe UI", 9),
                )
            else:
                widget = ttk.Entry(form, textvariable=var, font=("Segoe UI", 9))
            widget.pack(fill="x", ipady=2)
        result = {"value": None}

        def save():
            row = {field: var.get().strip() for field, var in variables.items()}
            if not row.get("modulo_uc_cursado") and not row.get("nome_uc_cursado"):
                messagebox.showwarning("Equivalência", "Informe ao menos o módulo ou nome da UC cursada.", parent=dialog)
                return
            result["value"] = row
            dialog.destroy()

        self._button(dialog, "Salvar UC", save, "primary").pack(anchor="e", padx=24, pady=18)
        self.root.wait_window(dialog)
        return result["value"]

    def add_equivalence(self):
        row = self._equivalence_dialog()
        if row is None:
            return
        self.equiv_rows.append(row)
        self._persist_equivalences()

    def edit_equivalence(self):
        selected = self.eq_tree.selection()
        if not selected:
            messagebox.showinfo("Equivalência", "Selecione uma linha para editar.", parent=self.root)
            return
        idx = int(selected[0])
        row = self._equivalence_dialog(self.equiv_rows[idx])
        if row is None:
            return
        self.equiv_rows[idx] = row
        self._persist_equivalences()

    def delete_equivalence(self):
        selected = self.eq_tree.selection()
        if not selected:
            messagebox.showinfo("Equivalência", "Selecione uma linha para remover.", parent=self.root)
            return
        idx = int(selected[0])
        if not messagebox.askyesno("Remover UC", "Remover a equivalência selecionada?", parent=self.root):
            return
        self.equiv_rows.pop(idx)
        self._persist_equivalences()

    def _persist_equivalences(self):
        try:
            self.storage.save_equivalences(self.current_process_id or "", self.equiv_rows)
        except ExcelStorageError as exc:
            self._show_storage_error(exc)
            self._load_equivalences()
            return
        self._load_equivalences()

    def export_equivalences(self):
        if not self.current_process_id:
            messagebox.showinfo("Equivalências", "Selecione um processo para exportar.", parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Exportar mapeamento de equivalências",
            initialfile=f"Mapeamento_{self.current_process_id}.xlsx", defaultextension=".xlsx",
            filetypes=[("Pasta de trabalho Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            openpyxl, Alignment, Font, PatternFill = ExcelStorage._load_openpyxl()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Mapeamento"
            ws.append([heading for heading, _field in EQUIVALENCE_COLUMNS])
            for row in self.equiv_rows:
                ws.append([
                    self.current_process_id if field == "id_processo" else row.get(field, "")
                    for _heading, field in EQUIVALENCE_COLUMNS
                ])
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="004A8D")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for column in ws.columns:
                ws.column_dimensions[column[0].column_letter].width = min(
                    max(max(len(str(cell.value or "")) for cell in column) + 3, 16), 46
                )
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            wb.save(path)
        except (OSError, ExcelStorageError) as exc:
            messagebox.showerror("Exportação", str(exc), parent=self.root)
            return
        messagebox.showinfo("Exportação", "Mapeamento exportado para uma nova planilha Excel.", parent=self.root)

    def _sync_settings_fields(self):
        if hasattr(self, "settings_path_var"):
            self.settings_path_var.set(str(self.workbook_path))
        if hasattr(self, "settings_student_path_var"):
            self.settings_student_path_var.set(str(self.student_workbook_path))
        if hasattr(self, "settings_school_year_var"):
            self.settings_school_year_var.set(str(self.school_year))
        if hasattr(self, "settings_process_prefix_var"):
            self.settings_process_prefix_var.set(self.process_code_prefix)

    def browse_workbook(self):
        path = filedialog.askopenfilename(
            parent=self.root, title="Selecionar planilha Excel",
            filetypes=[("Pasta de trabalho Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.settings_path_var.set(path)

    def browse_student_workbook(self):
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Selecionar ou criar a base separada de alunos",
            initialdir=str(self.student_workbook_path.parent),
            initialfile=self.student_workbook_path.name,
            defaultextension=".xlsx",
            filetypes=[("Pasta de trabalho Excel", "*.xlsx")],
        )
        if path:
            self.settings_student_path_var.set(path)

    def save_settings(self):
        raw_year = self.settings_school_year_var.get().strip()
        if not re.fullmatch(r"\d{4}", raw_year):
            messagebox.showwarning(
                "Configurações", "Informe o ano letivo com quatro dígitos, por exemplo 2026.",
                parent=self.root,
            )
            return
        school_year = int(raw_year)
        if school_year < 2000 or school_year > 9999:
            messagebox.showwarning(
                "Configurações", "Informe um ano letivo válido.", parent=self.root
            )
            return
        raw_prefix = self.settings_process_prefix_var.get().strip()
        process_prefix = self._normalize_process_prefix(raw_prefix)
        if not raw_prefix or not process_prefix:
            messagebox.showwarning(
                "Configurações", "Informe o início do código do processo, por exemplo PROC.",
                parent=self.root,
            )
            return

        raw_path = self.settings_path_var.get().strip().strip('"')
        if not raw_path:
            messagebox.showwarning("Configurações", "Informe ou selecione o caminho da planilha.", parent=self.root)
            return
        path = Path(raw_path).expanduser()
        if path.suffix.casefold() != ".xlsx":
            path = path.with_suffix(".xlsx")
        raw_student_path = self.settings_student_path_var.get().strip().strip('"')
        student_path = Path(raw_student_path).expanduser() if raw_student_path else path.with_name("Base_Alunos.xlsx")
        if student_path.suffix.casefold() != ".xlsx":
            student_path = student_path.with_suffix(".xlsx")
        if student_path.resolve() == path.resolve():
            messagebox.showerror(
                "Configurações", "A base de alunos precisa ser um arquivo separado da planilha de processos.",
                parent=self.root,
            )
            return
        storage = ExcelStorage(path)
        student_storage = StudentStorage(student_path)
        try:
            existing_students = self.student_storage.list_students()
            legacy_students = storage.list_students()
            storage.initialize()
            storage.copy_users_from(self.storage)
            student_storage.initialize()
            if existing_students:
                student_storage.save_students(existing_students)
            if legacy_students:
                student_storage.save_students(legacy_students)
            config = {
                "workbook_path": str(path.resolve()),
                "student_workbook_path": str(student_path.resolve()),
                "school_year": school_year,
                "process_code_prefix": process_prefix,
                "auto_search": self.auto_search.get(),
                "validate_duplicate_offer": self.validate_duplicate_offer.get(),
            }
            self._write_config(config)
        except (ExcelStorageError, OSError) as exc:
            messagebox.showerror("Configurações", str(exc), parent=self.root)
            return
        self.workbook_path = path.resolve()
        self.student_workbook_path = student_path.resolve()
        self.school_year = school_year
        self.process_code_prefix = process_prefix
        self.settings_school_year_var.set(str(school_year))
        self.settings_process_prefix_var.set(process_prefix)
        self.storage = storage
        self.student_storage = student_storage
        self.data_ready = True
        self.current_process_id = None
        self._update_path_labels()
        self._refresh_process_selector()
        self._refresh_student_list()
        self._refresh_student_process_selector()
        self._load_new_record()
        messagebox.showinfo(
            "Configurações", "Bases conectadas. Processos e usuários ficam na planilha principal; alunos, no arquivo separado.",
            parent=self.root,
        )

    def reset_settings(self):
        self.settings_path_var.set(str(DEFAULT_WORKBOOK))
        self.settings_student_path_var.set(str(DEFAULT_WORKBOOK.with_name("Base_Alunos.xlsx")))
        self.settings_school_year_var.set(str(date.today().year))
        self.settings_process_prefix_var.set("PROC")
        self.auto_search.set(True)
        self.validate_duplicate_offer.set(True)

    def _show_storage_error(self, error):
        self.connection_badge.configure(text="●  Verificar planilha", bg="#FFF4D6", fg="#9A5B00")
        messagebox.showerror("Planilha Excel", str(error), parent=self.root)


def main():
    root = tk.Tk()
    config = PortalApp._read_config()
    workbook_path = Path(config.get("workbook_path") or DEFAULT_WORKBOOK)
    storage = ExcelStorage(workbook_path)
    try:
        storage.initialize()
    except ExcelStorageError as exc:
        root.withdraw()
        messagebox.showerror("Planilha", str(exc), parent=root)
        root.destroy()
        return

    def open_portal(user):
        root.portal_app = PortalApp(root, user)

    try:
        LoginScreen(root, storage, open_portal)
    except ExcelStorageError as exc:
        root.withdraw()
        messagebox.showerror("Usuários", str(exc), parent=root)
        root.destroy()
        return

    def show_window():
        try:
            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            root.focus_force()
            root.after(800, lambda: root.attributes("-topmost", False))
        except tk.TclError:
            pass

    root.after(120, show_window)
    root.mainloop()


if __name__ == "__main__":
    main()
