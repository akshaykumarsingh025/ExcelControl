# ui.py
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
import threading
import queue
from agent import ExcelAgent
from excel_controller import ExcelController
from history import CommandHistory


class ExcelAIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ExcelAI — Powered by gemma4:e4b")
        self.root.geometry("1050x650")
        self.root.configure(bg="#1e1e2e")

        self.agent = ExcelAgent()
        self.excel = ExcelController()
        self.history = CommandHistory()

        self.auto_run_var = tk.BooleanVar(value=True)
        self.q = queue.Queue()
        self.root.after(100, self._process_queue)

        self._build_ui()

    def _build_ui(self):
        # ── Top Bar ──────────────────────────────────────────────
        top = tk.Frame(self.root, bg="#313244", pady=6)
        top.pack(fill="x")

        # Left side controls
        tk.Label(
            top,
            text="📊 ExcelAI",
            font=("Segoe UI", 14, "bold"),
            bg="#313244",
            fg="#cdd6f4",
        ).pack(side="left", padx=12)

        tk.Button(
            top,
            text="📂 Open",
            command=self._open_file,
            bg="#89b4fa",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=4)

        tk.Button(
            top,
            text="🔗 Connect",
            command=self._connect_running,
            bg="#a6e3a1",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=4)

        tk.Button(
            top,
            text="🆕 New",
            command=self._new_workbook,
            bg="#f38ba8",
            fg="#1e1e2e",
            relief="flat",
            padx=6,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=4)

        self.status_label = tk.Label(
            top,
            text="⚠ Not connected",
            font=("Segoe UI", 9),
            bg="#313244",
            fg="#f38ba8",
        )
        self.status_label.pack(side="left", padx=12)

        # Right side controls
        tk.Button(
            top,
            text="💾 Export Macro.py",
            command=self._export_macro,
            bg="#fab387",
            fg="#1e1e2e",
            relief="flat",
            padx=8,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=12)

        tk.Checkbutton(
            top,
            text="⚡ Auto-Run Code",
            variable=self.auto_run_var,
            bg="#313244",
            fg="#cdd6f4",
            selectcolor="#45475a",
            activebackground="#313244",
            activeforeground="#cdd6f4",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=8)

        # ── Main Layout: PanedWindow ──────────────────────────────
        self.paned = tk.PanedWindow(
            self.root, orient=tk.HORIZONTAL, bg="#1e1e2e", sashwidth=6
        )
        self.paned.pack(fill="both", expand=True, padx=10, pady=6)

        # --- Left Pane: Chat ---
        left_frame = tk.Frame(self.paned, bg="#1e1e2e")
        self.paned.add(left_frame, minsize=400, stretch="always")

        self.chat = scrolledtext.ScrolledText(
            left_frame,
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

        # Chat Input Bar
        chat_bottom = tk.Frame(left_frame, bg="#1e1e2e", pady=6)
        chat_bottom.pack(fill="x")

        self.entry = tk.Entry(
            chat_bottom,
            font=("Segoe UI", 11),
            bg="#313244",
            fg="#cdd6f4",
            insertbackground="white",
            relief="flat",
            bd=6,
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=4)
        self.entry.bind("<Return>", lambda e: self._send())

        tk.Button(
            chat_bottom,
            text="▶ Send",
            command=self._send,
            bg="#cba6f7",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=10,
        ).pack(side="left", padx=(6, 0))

        tk.Button(
            chat_bottom,
            text="🧹 Clear",
            command=self._clear_chat,
            bg="#45475a",
            fg="#cdd6f4",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
        ).pack(side="left", padx=4)

        # --- Right Pane: Code Editor ---
        right_frame = tk.Frame(self.paned, bg="#1e1e2e")
        self.paned.add(right_frame, minsize=300)

        editor_top = tk.Frame(right_frame, bg="#1e1e2e")
        editor_top.pack(fill="x", pady=(0, 6))

        tk.Label(
            editor_top,
            text="📝 Code Review / Manual Edit",
            font=("Segoe UI", 10, "bold"),
            bg="#1e1e2e",
            fg="#bac2de",
        ).pack(side="left")

        self.editor = scrolledtext.ScrolledText(
            right_frame,
            font=("Consolas", 11),
            bg="#11111b",
            fg="#f9e2af",
            insertbackground="white",
            wrap=tk.NONE,
            pady=8,
            padx=8,
        )
        self.editor.pack(fill="both", expand=True)

        tk.Button(
            right_frame,
            text="▶ Execute Code in Editor",
            command=self._run_from_editor,
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=4,
        ).pack(fill="x", pady=(6, 0))

        # Initial Welcome
        self._log(
            "system",
            "Welcome to ExcelAI!\n1. Connect to Excel\n2. Try asking for a change\n\nIf Auto-Run is OFF, code appears on the right for review.\n",
        )
        self.paned.paneconfigure(left_frame, width=600)

    # ── Connection Methods ────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            ok, msg = self.excel.connect_or_open(filepath=path)
            self._update_status(ok, msg)

    def _connect_running(self):
        ok, msg = self.excel.connect_or_open()
        self._update_status(ok, msg)

    def _new_workbook(self):
        ok, msg = self.excel.connect_or_open()
        self._update_status(ok, msg)

    def _update_status(self, ok, msg):
        if ok:
            self.status_label.config(text=f"✅ {msg}", fg="#a6e3a1")
            self._log("success", f"[Connected] {msg}\n")
        else:
            self.status_label.config(text="❌ Connection failed", fg="#f38ba8")
            self._log("error", f"[Error] {msg}\n")

    # ── Main Flow (Threaded) ──────────────────────────────────────

    def _send(self):
        cmd = self.entry.get().strip()
        if not cmd:
            return
        self.entry.delete(0, tk.END)

        if not self.excel.is_connected():
            self._log("error", "⚠ Connect to Excel first using the buttons above.\n")
            return

        self._log("user", f"\n👤 You: {cmd}\n")
        self._log("system", "🤖 Thinking... (Reading sheet context)\n")
        self.entry.config(state="disabled")  # lock input while thinking

        # Grab sheet context before AI thinks
        context = self.excel.get_sheet_context()

        # Run AI generation in background thread to keep UI smooth
        threading.Thread(
            target=self._ai_worker, args=(cmd, context), daemon=True
        ).start()

    def _ai_worker(self, cmd, context):
        try:
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
                    self._log("error", f"❌ Error connecting to AI: {arg1}\n")
        except queue.Empty:
            pass
        self.root.after(100, self._process_queue)

    def _handle_ai_response(self, cmd, code):
        if self.auto_run_var.get():
            self._log("ai", "🤖 AI Generated Code:\n")
            self._log("code", f"{code}\n")
            self._log("system", "⚡ Auto-running code...\n")
            self._execute_code(cmd, code)
        else:
            self._log(
                "system",
                "⏸ Code generated. Sent to Code Editor on the right for review.\n",
            )
            self.editor.delete("1.0", tk.END)
            self.editor.insert("1.0", code)
            self.current_cmd = cmd  # Remember what command this was for

    def _run_from_editor(self):
        code = self.editor.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Empty", "Code editor is empty!")
            return
        cmd = getattr(self, "current_cmd", "Manual Editor Execution")

        self._log("system", "⚡ Running from Editor:\n")
        self._log("code", f"{code}\n")
        self._execute_code(cmd, code)

    def _execute_code(self, cmd, code):
        success, result = self.excel.execute(code)
        self.history.add(cmd, code, success)

        if success:
            self._log("success", f"✅ {result}\n")
        else:
            self._log("error", f"❌ {result}\n")

    # ── Export Feature ────────────────────────────────────────────

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
            "    # Connect to the active workbook",
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
            messagebox.showinfo("Success", f"Macro saved successfully to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save macro:\n{str(e)}")

    # ── UI Helpers ────────────────────────────────────────────────

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
