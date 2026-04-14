# ui.py
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox, ttk
import threading
import queue
from agent import ExcelAgent, get_available_models
from excel_controller import ExcelController
from history import CommandHistory
from code_validator import validate_code
from dry_run import analyze_code
from templates import TEMPLATES, TEMPLATE_NAMES
from session_manager import save_session, load_session, clear_session
from batch_scheduler import BatchJob, BatchScheduler
from prompts import COMMON_COMMANDS


class ExcelAIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ExcelAI - AI-Powered Excel Controller")
        self.root.geometry("1200x750")
        self.root.configure(bg="#1e1e2e")
        self.root.minsize(900, 600)

        self.agent = ExcelAgent()
        self.excel = ExcelController()
        self.history = CommandHistory()
        self.batch_scheduler = BatchScheduler()

        self.auto_run_var = tk.BooleanVar(value=True)
        self.dry_run_var = tk.BooleanVar(value=False)
        self.analysis_mode_var = tk.BooleanVar(value=False)
        self.current_cmd = ""

        self.q = queue.Queue()
        self.root.after(100, self._process_queue)

        self._autocomplete_list = None
        self._autocomplete_index = -1

        self._build_ui()
        self._load_session()
        self._refresh_models()
        self._refresh_sheets()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self._build_top_bar()
        self._build_main_area()
        self._build_bottom_bar()

    # ================================================================
    # TOP BAR
    # ================================================================
    def _build_top_bar(self):
        top = tk.Frame(self.root, bg="#313244", pady=4)
        top.pack(fill="x")

        tk.Label(
            top,
            text="ExcelAI",
            font=("Segoe UI", 13, "bold"),
            bg="#313244",
            fg="#cdd6f4",
        ).pack(side="left", padx=8)

        btn_frame = tk.Frame(top, bg="#313244")
        btn_frame.pack(side="left", padx=4)

        for text, cmd, color in [
            ("Open", self._open_file, "#89b4fa"),
            ("Connect", self._connect_running, "#a6e3a1"),
            ("New", self._new_workbook, "#f38ba8"),
        ]:
            tk.Button(
                btn_frame,
                text=text,
                command=cmd,
                bg=color,
                fg="#1e1e2e",
                relief="flat",
                padx=6,
                font=("Segoe UI", 9, "bold"),
            ).pack(side="left", padx=2)

        self.status_label = tk.Label(
            top, text="Not connected", font=("Segoe UI", 9), bg="#313244", fg="#f38ba8"
        )
        self.status_label.pack(side="left", padx=8)

        right_frame = tk.Frame(top, bg="#313244")
        right_frame.pack(side="right", padx=8)

        tk.Button(
            right_frame,
            text="Export Macro",
            command=self._export_macro,
            bg="#fab387",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=2)

        tk.Button(
            right_frame,
            text="Undo",
            command=self._undo,
            bg="#cba6f7",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=2)

        tk.Button(
            right_frame,
            text="Redo",
            command=self._redo,
            bg="#cba6f7",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=2)

    # ================================================================
    # MAIN AREA (Paned: Left=Chat, Right=Code+Controls)
    # ================================================================
    def _build_main_area(self):
        self.paned = tk.PanedWindow(
            self.root, orient=tk.HORIZONTAL, bg="#1e1e2e", sashwidth=6
        )
        self.paned.pack(fill="both", expand=True, padx=6, pady=4)

        left_frame = tk.Frame(self.paned, bg="#1e1e2e")
        self.paned.add(left_frame, minsize=350, stretch="always")

        right_frame = tk.Frame(self.paned, bg="#1e1e2e")
        self.paned.add(right_frame, minsize=300)

        self._build_chat(left_frame)
        self._build_right_panel(right_frame)

        self.paned.paneconfigure(left_frame, width=550)

    # ================================================================
    # CHAT AREA (Left Pane)
    # ================================================================
    def _build_chat(self, parent):
        self.chat = scrolledtext.ScrolledText(
            parent,
            font=("Consolas", 10),
            bg="#181825",
            fg="#cdd6f4",
            insertbackground="white",
            wrap=tk.WORD,
            state="disabled",
            pady=8,
            padx=8,
        )
        self.chat.pack(fill="both", expand=True)

        self.chat.tag_config(
            "user", foreground="#89b4fa", font=("Consolas", 10, "bold")
        )
        self.chat.tag_config("ai", foreground="#a6e3a1")
        self.chat.tag_config("code", foreground="#f9e2af", background="#11111b")
        self.chat.tag_config("error", foreground="#f38ba8")
        self.chat.tag_config("success", foreground="#94e2d5")
        self.chat.tag_config("system", foreground="#6c7086")
        self.chat.tag_config("analysis", foreground="#89dceb")
        self.chat.tag_config("dryrun", foreground="#fab387")

        input_frame = tk.Frame(parent, bg="#1e1e2e", pady=4)
        input_frame.pack(fill="x")

        self.autocomplete_listbox = None

        self.entry = tk.Entry(
            input_frame,
            font=("Segoe UI", 11),
            bg="#313244",
            fg="#cdd6f4",
            insertbackground="white",
            relief="flat",
            bd=6,
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=4)
        self.entry.bind("<Return>", lambda e: self._send())
        self.entry.bind("<KeyRelease>", self._on_key_release)
        self.entry.bind("<Up>", self._autocomplete_up)
        self.entry.bind("<Down>", self._autocomplete_down)
        self.entry.bind("<Tab>", self._autocomplete_select)
        self.entry.bind("<Escape>", self._hide_autocomplete)

        btn_frame = tk.Frame(input_frame, bg="#1e1e2e")
        btn_frame.pack(side="left", padx=(4, 0))

        tk.Button(
            btn_frame,
            text="Send",
            command=self._send,
            bg="#cba6f7",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=8,
        ).pack(side="left", padx=2)

        tk.Button(
            btn_frame,
            text="Clear",
            command=self._clear_chat,
            bg="#45475a",
            fg="#cdd6f4",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
        ).pack(side="left", padx=2)

    # ================================================================
    # RIGHT PANEL (Code Editor + Controls)
    # ================================================================
    def _build_right_panel(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("TNotebook", background="#1e1e2e")
        style.configure(
            "TNotebook.Tab",
            background="#313244",
            foreground="#1e1e2e",
            padding=[10, 4],
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#89b4fa")],
            foreground=[("selected", "#1e1e2e")],
        )

        code_frame = tk.Frame(notebook, bg="#1e1e2e")
        notebook.add(code_frame, text="Code Editor")

        controls_frame = tk.Frame(notebook, bg="#1e1e2e")
        notebook.add(controls_frame, text="Controls")

        templates_frame = tk.Frame(notebook, bg="#1e1e2e")
        notebook.add(templates_frame, text="Templates")

        batch_frame = tk.Frame(notebook, bg="#1e1e2e")
        notebook.add(batch_frame, text="Batch/Schedule")

        self._build_code_editor(code_frame)
        self._build_controls(controls_frame)
        self._build_templates(templates_frame)
        self._build_batch_panel(batch_frame)

    def _build_code_editor(self, parent):
        editor_top = tk.Frame(parent, bg="#1e1e2e")
        editor_top.pack(fill="x", pady=(0, 4))

        tk.Label(
            editor_top,
            text="Code Review / Manual Edit",
            font=("Segoe UI", 10, "bold"),
            bg="#1e1e2e",
            fg="#bac2de",
        ).pack(side="left")

        self.editor = scrolledtext.ScrolledText(
            parent,
            font=("Consolas", 11),
            bg="#11111b",
            fg="#f9e2af",
            insertbackground="white",
            wrap=tk.NONE,
            pady=8,
            padx=8,
        )
        self.editor.pack(fill="both", expand=True)

        btn_row = tk.Frame(parent, bg="#1e1e2e")
        btn_row.pack(fill="x", pady=(4, 0))

        tk.Button(
            btn_row,
            text="Execute Code",
            command=self._run_from_editor,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Validate",
            command=self._validate_editor_code,
            bg="#89b4fa",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Dry Run",
            command=self._dry_run_editor_code,
            bg="#fab387",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(side="left", fill="x", expand=True, padx=2)

    def _build_controls(self, parent):
        canvas = tk.Canvas(parent, bg="#1e1e2e", highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg="#1e1e2e")

        scrollable.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        row = 0

        section_font = ("Segoe UI", 11, "bold")
        label_font = ("Segoe UI", 9)

        tk.Label(
            scrollable,
            text="Model Selection",
            font=section_font,
            bg="#1e1e2e",
            fg="#cdd6f4",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        row += 1

        tk.Label(
            scrollable,
            text="Ollama Model:",
            font=label_font,
            bg="#1e1e2e",
            fg="#bac2de",
        ).grid(row=row, column=0, sticky="w", padx=8)
        self.model_var = tk.StringVar(value="gemma4:e4b")
        self.model_combo = ttk.Combobox(
            scrollable, textvariable=self.model_var, state="readonly", width=25
        )
        self.model_combo.grid(row=row, column=1, padx=8, pady=2, sticky="ew")
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_change)
        row += 1

        tk.Button(
            scrollable,
            text="Refresh Models",
            command=self._refresh_models,
            bg="#45475a",
            fg="#cdd6f4",
            font=label_font,
            relief="flat",
        ).grid(row=row, column=0, columnspan=2, padx=8, pady=2, sticky="ew")
        row += 1

        tk.Frame(scrollable, bg="#45475a", height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=8
        )
        row += 1

        tk.Label(
            scrollable,
            text="Sheet Management",
            font=section_font,
            bg="#1e1e2e",
            fg="#cdd6f4",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 4))
        row += 1

        tk.Label(
            scrollable,
            text="Active Sheet:",
            font=label_font,
            bg="#1e1e2e",
            fg="#bac2de",
        ).grid(row=row, column=0, sticky="w", padx=8)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(
            scrollable, textvariable=self.sheet_var, state="readonly", width=25
        )
        self.sheet_combo.grid(row=row, column=1, padx=8, pady=2, sticky="ew")
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_change)
        row += 1

        tk.Button(
            scrollable,
            text="Refresh Sheets",
            command=self._refresh_sheets,
            bg="#45475a",
            fg="#cdd6f4",
            font=label_font,
            relief="flat",
        ).grid(row=row, column=0, columnspan=2, padx=8, pady=2, sticky="ew")
        row += 1

        tk.Frame(scrollable, bg="#45475a", height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=8
        )
        row += 1

        tk.Label(
            scrollable,
            text="Cross-File Operations",
            font=section_font,
            bg="#1e1e2e",
            fg="#cdd6f4",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 4))
        row += 1

        tk.Button(
            scrollable,
            text="Open Another Workbook",
            command=self._open_other_wb,
            bg="#89b4fa",
            fg="#1e1e2e",
            font=label_font,
            relief="flat",
        ).grid(row=row, column=0, columnspan=2, padx=8, pady=2, sticky="ew")
        row += 1

        tk.Label(
            scrollable,
            text="Open Workbooks:",
            font=label_font,
            bg="#1e1e2e",
            fg="#bac2de",
        ).grid(row=row, column=0, sticky="w", padx=8)
        self.wb_listbox = tk.Listbox(
            scrollable,
            bg="#181825",
            fg="#cdd6f4",
            font=("Consolas", 9),
            height=4,
            selectbackground="#45475a",
        )
        self.wb_listbox.grid(row=row, column=1, padx=8, pady=2, sticky="ew")
        row += 1

        tk.Button(
            scrollable,
            text="Switch to Selected Workbook",
            command=self._switch_workbook,
            bg="#45475a",
            fg="#cdd6f4",
            font=label_font,
            relief="flat",
        ).grid(row=row, column=0, columnspan=2, padx=8, pady=2, sticky="ew")
        row += 1

        tk.Frame(scrollable, bg="#45475a", height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=8
        )
        row += 1

        tk.Label(
            scrollable, text="Session", font=section_font, bg="#1e1e2e", fg="#cdd6f4"
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 4))
        row += 1

        sess_btn_frame = tk.Frame(scrollable, bg="#1e1e2e")
        sess_btn_frame.grid(
            row=row, column=0, columnspan=2, padx=8, pady=2, sticky="ew"
        )

        tk.Button(
            sess_btn_frame,
            text="Save Session",
            command=self._save_session,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=label_font,
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            sess_btn_frame,
            text="Clear Session",
            command=self._clear_session,
            bg="#f38ba8",
            fg="#1e1e2e",
            font=label_font,
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=2)
        row += 1

        scrollable.columnconfigure(1, weight=1)

    def _build_templates(self, parent):
        tk.Label(
            parent,
            text="Template Library",
            font=("Segoe UI", 12, "bold"),
            bg="#1e1e2e",
            fg="#cdd6f4",
        ).pack(pady=(8, 4))

        list_frame = tk.Frame(parent, bg="#1e1e2e")
        list_frame.pack(fill="both", expand=True, padx=8, pady=4)

        self.template_listbox = tk.Listbox(
            list_frame,
            bg="#181825",
            fg="#cdd6f4",
            font=("Segoe UI", 10),
            selectbackground="#45475a",
            selectforeground="#cdd6f4",
            activestyle="none",
        )
        self.template_listbox.pack(fill="both", expand=True, side="left")

        for name in TEMPLATE_NAMES:
            self.template_listbox.insert(tk.END, name)

        template_scroll = tk.Scrollbar(list_frame, command=self.template_listbox.yview)
        template_scroll.pack(side="right", fill="y")
        self.template_listbox.config(yscrollcommand=template_scroll.set)

        self.template_desc = tk.Label(
            parent,
            text="Select a template to see its description.",
            font=("Segoe UI", 9),
            bg="#1e1e2e",
            fg="#6c7086",
            wraplength=400,
            justify="left",
        )
        self.template_desc.pack(fill="x", padx=8, pady=4)

        self.template_listbox.bind("<<ListboxSelect>>", self._on_template_select)

        btn_row = tk.Frame(parent, bg="#1e1e2e")
        btn_row.pack(fill="x", padx=8, pady=4)

        tk.Button(
            btn_row,
            text="Load to Editor",
            command=self._load_template_to_editor,
            bg="#89b4fa",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Execute Directly",
            command=self._execute_template,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(side="left", fill="x", expand=True, padx=2)

    def _build_batch_panel(self, parent):
        tk.Label(
            parent,
            text="Batch Replay & Scheduler",
            font=("Segoe UI", 12, "bold"),
            bg="#1e1e2e",
            fg="#cdd6f4",
        ).pack(pady=(8, 4))

        tk.Label(
            parent,
            text="Create a batch job from successful commands:",
            font=("Segoe UI", 9),
            bg="#1e1e2e",
            fg="#6c7086",
        ).pack(padx=8)

        name_frame = tk.Frame(parent, bg="#1e1e2e")
        name_frame.pack(fill="x", padx=8, pady=4)

        tk.Label(
            name_frame,
            text="Job Name:",
            font=("Segoe UI", 9),
            bg="#1e1e2e",
            fg="#bac2de",
        ).pack(side="left", padx=4)
        self.batch_name_entry = tk.Entry(
            name_frame,
            font=("Segoe UI", 10),
            bg="#313244",
            fg="#cdd6f4",
            insertbackground="white",
            relief="flat",
            bd=4,
        )
        self.batch_name_entry.pack(side="left", fill="x", expand=True, padx=4)

        sched_frame = tk.Frame(parent, bg="#1e1e2e")
        sched_frame.pack(fill="x", padx=8, pady=4)

        tk.Label(
            sched_frame,
            text="Schedule (HH:MM):",
            font=("Segoe UI", 9),
            bg="#1e1e2e",
            fg="#bac2de",
        ).pack(side="left", padx=4)
        self.schedule_entry = tk.Entry(
            sched_frame,
            font=("Segoe UI", 10),
            bg="#313244",
            fg="#cdd6f4",
            insertbackground="white",
            relief="flat",
            bd=4,
            width=8,
        )
        self.schedule_entry.pack(side="left", padx=4)

        self.repeat_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            sched_frame,
            text="Repeat Daily",
            variable=self.repeat_var,
            bg="#1e1e2e",
            fg="#cdd6f4",
            selectcolor="#45475a",
            activebackground="#1e1e2e",
            activeforeground="#cdd6f4",
            font=("Segoe UI", 9),
        ).pack(side="left", padx=8)

        btn_row = tk.Frame(parent, bg="#1e1e2e")
        btn_row.pack(fill="x", padx=8, pady=4)

        tk.Button(
            btn_row,
            text="Create from History",
            command=self._create_batch_from_history,
            bg="#cba6f7",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=3,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Save as Batch Job",
            command=self._save_batch_job,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=3,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Run Selected",
            command=self._run_batch_job,
            bg="#89b4fa",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=3,
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            btn_row,
            text="Delete Selected",
            command=self._delete_batch_job,
            bg="#f38ba8",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=3,
        ).pack(side="left", fill="x", expand=True, padx=2)

        self.batch_listbox = tk.Listbox(
            parent,
            bg="#181825",
            fg="#cdd6f4",
            font=("Consolas", 9),
            selectbackground="#45475a",
            height=8,
        )
        self.batch_listbox.pack(fill="both", expand=True, padx=8, pady=4)

        sched_btn_frame = tk.Frame(parent, bg="#1e1e2e")
        sched_btn_frame.pack(fill="x", padx=8, pady=4)

        tk.Button(
            sched_btn_frame,
            text="Start Scheduler",
            command=self._start_scheduler,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=2)

        tk.Button(
            sched_btn_frame,
            text="Stop Scheduler",
            command=self._stop_scheduler,
            bg="#f38ba8",
            fg="#1e1e2e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=2)

        self.scheduler_status = tk.Label(
            sched_btn_frame,
            text="Scheduler: Stopped",
            font=("Segoe UI", 9),
            bg="#1e1e2e",
            fg="#6c7086",
        )
        self.scheduler_status.pack(side="left", padx=8)

        self._refresh_batch_list()

    # ================================================================
    # BOTTOM BAR
    # ================================================================
    def _build_bottom_bar(self):
        bottom = tk.Frame(self.root, bg="#313244", pady=3)
        bottom.pack(fill="x")

        tk.Checkbutton(
            bottom,
            text="Auto-Run",
            variable=self.auto_run_var,
            bg="#313244",
            fg="#cdd6f4",
            selectcolor="#45475a",
            activebackground="#313244",
            activeforeground="#cdd6f4",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=8)

        tk.Checkbutton(
            bottom,
            text="Dry-Run",
            variable=self.dry_run_var,
            bg="#313244",
            fg="#fab387",
            selectcolor="#45475a",
            activebackground="#313244",
            activeforeground="#fab387",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=8)

        tk.Checkbutton(
            bottom,
            text="Analysis Mode",
            variable=self.analysis_mode_var,
            bg="#313244",
            fg="#89dceb",
            selectcolor="#45475a",
            activebackground="#313244",
            activeforeground="#89dceb",
            font=("Segoe UI", 9, "bold"),
            command=self._toggle_analysis_mode,
        ).pack(side="left", padx=8)

        self.mode_label = tk.Label(
            bottom,
            text="Mode: Automation",
            font=("Segoe UI", 9),
            bg="#313244",
            fg="#a6e3a1",
        )
        self.mode_label.pack(side="right", padx=8)

    # ================================================================
    # CONNECTION
    # ================================================================
    def _open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            ok, msg = self.excel.connect_or_open(filepath=path)
            self._update_status(ok, msg)
            self._refresh_sheets()

    def _connect_running(self):
        ok, msg = self.excel.connect_or_open()
        self._update_status(ok, msg)
        self._refresh_sheets()

    def _new_workbook(self):
        ok, msg = self.excel.connect_or_open()
        self._update_status(ok, msg)
        self._refresh_sheets()

    def _update_status(self, ok, msg):
        if ok:
            self.status_label.config(text=f"Connected: {msg}", fg="#a6e3a1")
            self._log("success", f"[Connected] {msg}\n")
        else:
            self.status_label.config(text="Connection failed", fg="#f38ba8")
            self._log("error", f"[Error] {msg}\n")

    # ================================================================
    # MAIN SEND FLOW (Threaded)
    # ================================================================
    def _send(self):
        self._hide_autocomplete()
        cmd = self.entry.get().strip()
        if not cmd:
            return
        self.entry.delete(0, tk.END)

        if not self.excel.is_connected():
            self._log("error", "Connect to Excel first using the buttons above.\n")
            return

        self._log("user", f"\nYou: {cmd}\n")

        if self.analysis_mode_var.get():
            self._log("analysis", "Analysis mode: Reading sheet data...\n")
        else:
            self._log("system", "Thinking... (Reading sheet context)\n")

        self.entry.config(state="disabled")
        context = (
            self.excel.get_full_context()
            if self.analysis_mode_var.get()
            else self.excel.get_sheet_context()
        )

        threading.Thread(
            target=self._ai_worker, args=(cmd, context), daemon=True
        ).start()

    def _ai_worker(self, cmd, context):
        try:
            self.agent.set_analysis_mode(self.analysis_mode_var.get())
            code = self.agent.ask(cmd, context)
            self.q.put(("ai_done", cmd, code))
        except Exception as e:
            self.q.put(("ai_error", str(e), ""))

    def _process_queue(self):
        try:
            while True:
                msg_type, arg1, arg2 = self.q.get_nowait()
                if msg_type == "ai_done":
                    self.entry.config(state="normal")
                    self._handle_ai_response(cmd=arg1, code=arg2)
                elif msg_type == "ai_error":
                    self.entry.config(state="normal")
                    self._log("error", f"Error connecting to AI: {arg1}\n")
        except queue.Empty:
            pass
        self.root.after(100, self._process_queue)

    def _handle_ai_response(self, cmd, code):
        is_analysis = self.analysis_mode_var.get()

        if is_analysis:
            self._log("ai", "AI Generated Analysis Code:\n")
            self._log("code", f"{code}\n")

            if self.auto_run_var.get():
                self._log("system", "Running analysis...\n")
                success, result = self.excel.execute_analysis(code)
                if success:
                    self._log("analysis", f"\n--- Analysis Results ---\n{result}\n")
                else:
                    self._log("error", f"{result}\n")
                self.history.add(cmd, code, success)
            else:
                self.editor.delete("1.0", tk.END)
                self.editor.insert("1.0", code)
                self.current_cmd = cmd
            return

        if self.dry_run_var.get():
            self._log("dryrun", "DRY RUN MODE - Previewing changes:\n")
            self._log("code", f"{code}\n")
            validation = validate_code(code)
            if not validation.is_safe:
                self._log(
                    "error",
                    "Code blocked for safety:\n" + "\n".join(validation.issues) + "\n",
                )
                return
            result = analyze_code(code)
            self._log("dryrun", result.summary() + "\n")
            self.editor.delete("1.0", tk.END)
            self.editor.insert("1.0", code)
            self.current_cmd = cmd
            self._log(
                "system", "Code sent to editor. Uncheck Dry-Run and execute to apply.\n"
            )
            return

        if self.auto_run_var.get():
            self._log("ai", "AI Generated Code:\n")
            self._log("code", f"{code}\n")
            self._log("system", "Auto-running code...\n")
            self._execute_code(cmd, code)
        else:
            self._log("system", "Code generated. Sent to Code Editor for review.\n")
            self.editor.delete("1.0", tk.END)
            self.editor.insert("1.0", code)
            self.current_cmd = cmd

    def _execute_code(self, cmd, code):
        success, result = self.excel.execute(code)
        self.history.add(cmd, code, success)
        if success:
            self._log("success", f"Done.\n")
            self._refresh_sheets()
        else:
            self._log("error", f"{result}\n")

    def _run_from_editor(self):
        code = self.editor.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Empty", "Code editor is empty!")
            return
        cmd = getattr(self, "current_cmd", "Manual Editor Execution")
        self._log("system", "Running from Editor:\n")
        self._log("code", f"{code}\n")

        if self.analysis_mode_var.get():
            success, result = self.excel.execute_analysis(code)
            if success:
                self._log("analysis", f"\n--- Analysis Results ---\n{result}\n")
            else:
                self._log("error", f"{result}\n")
            self.history.add(cmd, code, success)
        else:
            self._execute_code(cmd, code)

    # ================================================================
    # CODE VALIDATION & DRY-RUN FROM EDITOR
    # ================================================================
    def _validate_editor_code(self):
        code = self.editor.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Empty", "Code editor is empty!")
            return
        result = validate_code(code)
        if result.is_safe:
            self._log("success", "Code validation: PASSED - Safe to execute.\n")
        else:
            self._log(
                "error", "Code validation: BLOCKED\n" + "\n".join(result.issues) + "\n"
            )

    def _dry_run_editor_code(self):
        code = self.editor.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Empty", "Code editor is empty!")
            return
        validation = validate_code(code)
        if not validation.is_safe:
            self._log(
                "error",
                "Code blocked for safety:\n" + "\n".join(validation.issues) + "\n",
            )
            return
        result = analyze_code(code)
        self._log("dryrun", result.summary() + "\n")

    # ================================================================
    # UNDO / REDO
    # ================================================================
    def _undo(self):
        ok, msg = self.excel.undo()
        if ok:
            self._log("success", f"Undo applied.\n")
            self._refresh_sheets()
        else:
            self._log("error", f"Undo failed: {msg}\n")

    def _redo(self):
        ok, msg = self.excel.redo()
        if ok:
            self._log("success", f"Redo applied.\n")
            self._refresh_sheets()
        else:
            self._log("error", f"Redo failed: {msg}\n")

    # ================================================================
    # MODEL SELECTION
    # ================================================================
    def _refresh_models(self):
        models = get_available_models()
        self.model_combo["values"] = models
        if self.model_var.get() not in models and models:
            self.model_var.set(models[0])

    def _on_model_change(self, event=None):
        self.agent.set_model(self.model_var.get())
        self._log("system", f"Model switched to: {self.model_var.get()}\n")

    # ================================================================
    # SHEET SWITCHER
    # ================================================================
    def _refresh_sheets(self):
        if self.excel.is_connected():
            sheets = self.excel.get_sheet_names()
            self.sheet_combo["values"] = sheets
            current = self.excel.get_current_sheet_name()
            if current:
                self.sheet_var.set(current)
            elif sheets:
                self.sheet_var.set(sheets[0])

            wbs = self.excel.list_open_workbooks()
            self.wb_listbox.delete(0, tk.END)
            for wb_name in wbs:
                self.wb_listbox.insert(tk.END, wb_name)
                if self.excel.wb and wb_name == self.excel.wb.name:
                    self.wb_listbox.selection_set(tk.END)

    def _on_sheet_change(self, event=None):
        name = self.sheet_var.get()
        if name:
            ok = self.excel.switch_sheet(name)
            if ok:
                self._log("success", f"Switched to sheet: {name}\n")
            else:
                self._log("error", f"Could not switch to sheet: {name}\n")

    # ================================================================
    # CROSS-FILE OPERATIONS
    # ================================================================
    def _open_other_wb(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            ok, msg = self.excel.open_other_workbook(path)
            if ok:
                self._log("success", f"{msg}\n")
                self._refresh_sheets()
            else:
                self._log("error", f"Failed: {msg}\n")

    def _switch_workbook(self):
        sel = self.wb_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a workbook from the list first.")
            return
        name = self.wb_listbox.get(sel[0])
        wb = self.excel.get_open_workbook(name)
        if wb:
            self.excel.wb = wb
            self.excel.ws = wb.sheets.active
            self._log("success", f"Switched to workbook: {name}\n")
            self._refresh_sheets()
        else:
            self._log("error", f"Could not find workbook: {name}\n")

    # ================================================================
    # TEMPLATES
    # ================================================================
    def _on_template_select(self, event=None):
        sel = self.template_listbox.curselection()
        if not sel:
            return
        name = TEMPLATE_NAMES[sel[0]]
        desc = TEMPLATES[name]["description"]
        self.template_desc.config(text=desc)

    def _load_template_to_editor(self):
        sel = self.template_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a template first.")
            return
        name = TEMPLATE_NAMES[sel[0]]
        code = TEMPLATES[name]["code"]
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", code)
        self.current_cmd = f"Template: {name}"
        self._log("system", f"Template '{name}' loaded to editor.\n")

    def _execute_template(self):
        sel = self.template_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a template first.")
            return
        name = TEMPLATE_NAMES[sel[0]]
        code = TEMPLATES[name]["code"]
        self.current_cmd = f"Template: {name}"
        self._log("system", f"Executing template: {name}\n")
        self._execute_code(f"Template: {name}", code)

    # ================================================================
    # BATCH / SCHEDULER
    # ================================================================
    def _create_batch_from_history(self):
        log = self.history.get_all()
        successful = [item for item in log if item.get("success")]
        if not successful:
            messagebox.showinfo("Info", "No successful commands in history.")
            return
        name = (
            self.batch_name_entry.get().strip()
            or f"Batch_{len(self.batch_scheduler.jobs)}"
        )
        schedule_time = self.schedule_entry.get().strip()
        repeat = self.repeat_var.get()
        filepath = self.excel.wb.fullname if self.excel.is_connected() else ""

        job = BatchJob(
            name=name,
            commands=successful,
            filepath=filepath,
            schedule_time=schedule_time,
            repeat=repeat,
        )
        self.batch_scheduler.add_job(job)
        self._refresh_batch_list()
        self._log(
            "success", f"Batch job '{name}' created with {len(successful)} commands.\n"
        )

    def _save_batch_job(self):
        name = self.batch_name_entry.get().strip()
        if not name:
            messagebox.showwarning("Warning", "Enter a job name.")
            return
        schedule_time = self.schedule_entry.get().strip()
        repeat = self.repeat_var.get()
        filepath = self.excel.wb.fullname if self.excel.is_connected() else ""

        log = self.history.get_all()
        successful = [item for item in log if item.get("success")]
        if not successful:
            messagebox.showinfo("Info", "No successful commands to save.")
            return

        job = BatchJob(
            name=name,
            commands=successful,
            filepath=filepath,
            schedule_time=schedule_time,
            repeat=repeat,
        )
        self.batch_scheduler.add_job(job)
        self._refresh_batch_list()
        self._log("success", f"Batch job '{name}' saved.\n")

    def _run_batch_job(self):
        sel = self.batch_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a batch job first.")
            return
        job = self.batch_scheduler.jobs[sel[0]]
        self._log("system", f"Running batch job: {job.name}\n")
        ok, msg = self.batch_scheduler.execute_job(job, self.excel)
        if ok:
            self._log("success", f"Batch job completed:\n{msg}\n")
            self._refresh_sheets()
        else:
            self._log("error", f"Batch job failed: {msg}\n")

    def _delete_batch_job(self):
        sel = self.batch_listbox.curselection()
        if not sel:
            return
        self.batch_scheduler.remove_job(sel[0])
        self._refresh_batch_list()

    def _refresh_batch_list(self):
        self.batch_listbox.delete(0, tk.END)
        for job in self.batch_scheduler.get_jobs():
            sched = f" [{job.schedule_time}]" if job.schedule_time else ""
            cnt = f" ({len(job.commands)} cmds)"
            self.batch_listbox.insert(tk.END, f"{job.name}{cnt}{sched}")

    def _start_scheduler(self):
        self.batch_scheduler.start_scheduler(
            self.excel,
            callback=self._scheduler_callback,
        )
        self.scheduler_status.config(text="Scheduler: Running", fg="#a6e3a1")
        self._log("success", "Scheduler started.\n")

    def _stop_scheduler(self):
        self.batch_scheduler.stop_scheduler()
        self.scheduler_status.config(text="Scheduler: Stopped", fg="#6c7086")
        self._log("system", "Scheduler stopped.\n")

    def _scheduler_callback(self, job, ok, msg):
        self.q.put(("scheduler_run", job.name, "OK" if ok else f"FAIL: {msg}"))

    # ================================================================
    # SESSION SAVE/LOAD
    # ================================================================
    def _save_session(self):
        ok = save_session(
            conversation_history=self.agent.get_history(),
            model=self.model_var.get(),
            auto_run=self.auto_run_var.get(),
            dry_run=self.dry_run_var.get(),
            analysis_mode=self.analysis_mode_var.get(),
            last_file=self.excel.wb.fullname if self.excel.is_connected() else "",
        )
        if ok:
            self._log("success", "Session saved.\n")
        else:
            self._log("error", "Failed to save session.\n")

    def _clear_session(self):
        clear_session()
        self._log("system", "Session cleared.\n")

    def _load_session(self):
        session = load_session()
        if not session:
            self._log(
                "system",
                "Welcome to ExcelAI!\n"
                "1. Connect to Excel\n"
                "2. Type a command or use templates\n"
                "3. Toggle Dry-Run to preview before executing\n\n",
            )
            return

        if session.get("model"):
            self.model_var.set(session["model"])
            self.agent.set_model(session["model"])

        if session.get("auto_run") is not None:
            self.auto_run_var.set(session["auto_run"])

        if session.get("dry_run") is not None:
            self.dry_run_var.set(session["dry_run"])

        if session.get("analysis_mode") is not None:
            self.analysis_mode_var.set(session["analysis_mode"])

        if session.get("conversation_history"):
            self.agent.set_history(session["conversation_history"])

        self._log("system", "Previous session restored.\n")

    # ================================================================
    # ANALYSIS MODE TOGGLE
    # ================================================================
    def _toggle_analysis_mode(self):
        if self.analysis_mode_var.get():
            self.mode_label.config(text="Mode: Analysis", fg="#89dceb")
            self._log(
                "analysis",
                "Analysis mode ON - AI will analyze data instead of modifying it.\n",
            )
        else:
            self.mode_label.config(text="Mode: Automation", fg="#a6e3a1")
            self._log(
                "system",
                "Automation mode ON - AI will generate code to modify Excel.\n",
            )

    # ================================================================
    # AUTOCOMPLETE
    # ================================================================
    def _on_key_release(self, event):
        if event.keysym in ("Up", "Down", "Tab", "Escape", "Return"):
            return

        typed = self.entry.get()
        if len(typed) < 2:
            self._hide_autocomplete()
            return

        matches = []
        all_commands = list(set(COMMON_COMMANDS + self.history.get_commands()))

        typed_lower = typed.lower()
        for cmd in all_commands:
            if cmd.lower().startswith(typed_lower):
                matches.append(cmd)

        if not matches:
            self._hide_autocomplete()
            return

        self._show_autocomplete(matches)

    def _show_autocomplete(self, matches):
        self._hide_autocomplete()

        self._autocomplete_list = matches
        self._autocomplete_index = -1

        list_frame = tk.Frame(self.root, bg="#313244", bd=1, relief="solid")
        self._autocomplete_frame = list_frame

        self._autocomplete_listbox = tk.Listbox(
            list_frame,
            bg="#313244",
            fg="#cdd6f4",
            font=("Segoe UI", 9),
            selectbackground="#45475a",
            selectforeground="#cdd6f4",
            activestyle="none",
            height=min(len(matches), 6),
        )
        self._autocomplete_listbox.pack(fill="both")

        for m in matches[:6]:
            self._autocomplete_listbox.insert(tk.END, m)

        self._autocomplete_listbox.bind("<Button-1>", self._autocomplete_click)

        x = self.entry.winfo_rootx() - self.root.winfo_rootx()
        y = (
            self.entry.winfo_rooty()
            - self.root.winfo_rooty()
            + self.entry.winfo_height()
            + 8
        )
        list_frame.place(x=x, y=y, width=350)

    def _hide_autocomplete(self, event=None):
        if hasattr(self, "_autocomplete_frame") and self._autocomplete_frame:
            self._autocomplete_frame.destroy()
            self._autocomplete_frame = None
        self._autocomplete_list = None
        self._autocomplete_index = -1

    def _autocomplete_up(self, event=None):
        if not self._autocomplete_list:
            return
        if self._autocomplete_listbox:
            cur = self._autocomplete_listbox.curselection()
            if cur:
                idx = cur[0] - 1
                if idx >= 0:
                    self._autocomplete_listbox.selection_clear(0, tk.END)
                    self._autocomplete_listbox.selection_set(idx)
            return "break"

    def _autocomplete_down(self, event=None):
        if not self._autocomplete_list:
            return
        if self._autocomplete_listbox:
            cur = self._autocomplete_listbox.curselection()
            size = self._autocomplete_listbox.size()
            if not cur:
                self._autocomplete_listbox.selection_set(0)
            elif cur[0] < size - 1:
                self._autocomplete_listbox.selection_clear(0, tk.END)
                self._autocomplete_listbox.selection_set(cur[0] + 1)
            return "break"

    def _autocomplete_select(self, event=None):
        if not self._autocomplete_list or not self._autocomplete_listbox:
            return
        sel = self._autocomplete_listbox.curselection()
        if sel:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self._autocomplete_listbox.get(sel[0]))
        self._hide_autocomplete()
        return "break"

    def _autocomplete_click(self, event=None):
        sel = self._autocomplete_listbox.curselection()
        if sel:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self._autocomplete_listbox.get(sel[0]))
        self._hide_autocomplete()

    # ================================================================
    # EXPORT MACRO
    # ================================================================
    def _export_macro(self):
        log = self.history.get_all()
        successful_items = [item for item in log if item["success"]]

        if not successful_items:
            messagebox.showinfo(
                "Export Macro", "No successful commands in history to export."
            )
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".py",
            filetypes=[("Python Files", "*.py")],
            title="Save Macro As",
        )
        if not path:
            return

        script = [
            "import xlwings as xw",
            "",
            "def run_macro():",
            "    wb = xw.books.active",
            "    ws = wb.sheets.active",
            "    ",
        ]

        for item in successful_items:
            script.append(f"    # COMMAND: {item['command']}")
            for line in item["code"].split("\n"):
                script.append(f"    {line}")
            script.append("")

        script.extend(
            [
                "if __name__ == '__main__':",
                "    run_macro()",
                "    print('Macro completed successfully.')",
            ]
        )

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(script))
            messagebox.showinfo("Success", f"Macro saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save macro:\n{str(e)}")

    # ================================================================
    # UI HELPERS
    # ================================================================
    def _log(self, tag, text):
        self.chat.config(state="normal")
        self.chat.insert(tk.END, text, tag)
        self.chat.see(tk.END)
        self.chat.config(state="disabled")

    def _clear_chat(self):
        self.chat.config(state="normal")
        self.chat.delete("1.0", tk.END)
        self.chat.config(state="disabled")
        self.agent.reset_memory()
        self.editor.delete("1.0", tk.END)
        self._log("system", "Chat cleared. AI memory reset.\n")

    def _on_close(self):
        self._save_session()
        self.excel.cleanup()
        self.batch_scheduler.stop_scheduler()
        self.root.destroy()
