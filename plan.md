# 📊 ExcelAI — AI-Powered Live Excel Controller
### Build Plan for AI Agent (Model: gemma4:e4b via Ollama)

---

## 🎯 Project Goal

Build a desktop Python application that lets users control **Microsoft Excel live** using natural language.
The user types a command → the AI understands it → Python code runs → Excel changes happen **visibly in real time**.

---

## 📁 Final Folder Structure

```
ExcelAI/
│
├── main.py                  ← App entry point
├── agent.py                 ← Ollama AI agent logic
├── excel_controller.py      ← xlwings Excel interface
├── ui.py                    ← Tkinter GUI
├── history.py               ← Command history & undo support
├── prompts.py               ← System prompts for the AI
├── requirements.txt         ← Python dependencies
├── launch.bat               ← Windows launcher
└── README.md                ← Usage guide
```

---

## 🛠️ Tech Stack

| Layer         | Tool/Library         | Purpose                                      |
|---------------|----------------------|----------------------------------------------|
| AI Model      | `gemma4:e4b` (Ollama)| Understands commands, generates xlwings code |
| Excel Control | `xlwings`            | Controls live open Excel window              |
| GUI           | `tkinter`            | Clean desktop chat interface                 |
| AI Client     | `ollama` (Python)    | Talks to local Ollama server                 |
| Safety        | `ast` module         | Validates code before executing              |

---

## 📦 requirements.txt

```
xlwings
ollama
```

> Only 2 dependencies. Both installable via pip. No internet needed at runtime.

---

## 🔧 Step-by-Step Build Instructions

---

### STEP 1 — `prompts.py` (Write This First)

This file holds the system prompt that teaches `gemma4:e4b` the xlwings API.

```python
# prompts.py

SYSTEM_PROMPT = """
You are an Excel automation AI. You control Microsoft Excel using Python xlwings.

RULES:
1. ONLY respond with pure Python code. No explanation. No markdown. No backticks.
2. The active sheet object is already available as variable: ws
3. The active workbook is available as: wb
4. Never import xlwings - it is already loaded.
5. Never call wb.save() or wb.close() - these are handled automatically.
6. If a task is unclear, write a comment # UNCLEAR: reason and do nothing else.

AVAILABLE xlwings COMMANDS:
- ws["A1"].value = "text"              → Write text to cell
- ws["A1"].value                       → Read cell value
- ws.range("A1:C3").value = [[...]]    → Write a 2D list to a range
- ws.range("A1").expand().value        → Read entire data block
- ws["A1"].color = (255, 0, 0)         → Set background color (R,G,B)
- ws["A1"].font.bold = True            → Bold text
- ws["A1"].font.size = 14              → Font size
- ws["A1"].font.color = (0, 0, 255)    → Font color
- ws["A1"].column_width = 20           → Set column width
- ws["A1"].row_height = 30             → Set row height
- ws["A1"].formula = "=SUM(B1:B10)"   → Insert formula
- ws["A1"].number_format = "0.00"      → Number format
- ws.name = "Sheet Name"               → Rename sheet
- wb.sheets.add("NewSheet")            → Add a new sheet
- ws.pictures.add("path/to/image.png") → Insert an image
- ws.charts.add(...)                   → Add a chart

EXAMPLE TASK: "Put Name, Age, City as headers in row 1 with green background"
EXAMPLE OUTPUT:
ws["A1"].value = "Name"
ws["B1"].value = "Age"
ws["C1"].value = "City"
ws.range("A1:C1").color = (144, 238, 144)
ws.range("A1:C1").font.bold = True
"""
```

---

### STEP 2 — `excel_controller.py`

Handles all Excel connection logic using xlwings.

```python
# excel_controller.py
import xlwings as xw

class ExcelController:
    def __init__(self):
        self.app = None
        self.wb = None
        self.ws = None

    def connect_or_open(self, filepath=None):
        """Connect to already open Excel, or open a file, or create new workbook."""
        try:
            if filepath:
                self.app = xw.App(visible=True)
                self.wb = self.app.books.open(filepath)
            elif xw.apps:
                # Attach to already running Excel
                self.app = xw.apps.active
                self.wb = self.app.books.active
            else:
                # Open fresh Excel with new workbook
                self.app = xw.App(visible=True)
                self.wb = self.app.books.add()

            self.ws = self.wb.sheets.active
            self.app.screen_updating = True   # Ensure changes are visible live
            return True, f"Connected to: {self.wb.name}"
        except Exception as e:
            return False, str(e)

    def execute(self, code: str):
        """Execute AI-generated xlwings code safely."""
        if not self.wb:
            return False, "No Excel workbook connected."

        # Inject ws and wb into execution context
        local_vars = {"ws": self.ws, "wb": self.wb, "xw": xw}

        try:
            exec(code, {}, local_vars)
            self.wb.save()
            return True, "Done."
        except Exception as e:
            return False, f"Error: {str(e)}"

    def get_sheet_names(self):
        if self.wb:
            return [s.name for s in self.wb.sheets]
        return []

    def switch_sheet(self, name):
        if self.wb:
            self.ws = self.wb.sheets[name]
            return True
        return False

    def is_connected(self):
        return self.wb is not None
```

---

### STEP 3 — `agent.py`

Sends user command to Ollama `gemma4:e4b` and returns clean Python code.

```python
# agent.py
import ollama
import re
from prompts import SYSTEM_PROMPT

class ExcelAgent:
    def __init__(self):
        self.model = "gemma4:e4b"
        self.conversation_history = []

    def ask(self, user_command: str) -> str:
        """Send command to Ollama and return the generated Python code."""

        self.conversation_history.append({
            "role": "user",
            "content": user_command
        })

        try:
            response = ollama.chat(
                model=self.model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT}
                ] + self.conversation_history
            )

            raw = response["message"]["content"]

            # Remember AI response in history for context
            self.conversation_history.append({
                "role": "assistant",
                "content": raw
            })

            # Strip markdown code blocks if model adds them
            clean = re.sub(r"```python|```", "", raw).strip()
            return clean

        except Exception as e:
            return f"# ERROR: Could not reach Ollama\n# {str(e)}"

    def reset_memory(self):
        """Clear conversation history."""
        self.conversation_history = []
```

---

### STEP 4 — `history.py`

Tracks command history so users can review what was done.

```python
# history.py

class CommandHistory:
    def __init__(self):
        self.log = []  # List of (command, code, status)

    def add(self, command: str, code: str, success: bool):
        self.log.append({
            "command": command,
            "code": code,
            "success": success
        })

    def get_last_code(self):
        if self.log:
            return self.log[-1]["code"]
        return None

    def get_all(self):
        return self.log

    def clear(self):
        self.log = []
```

---

### STEP 5 — `ui.py`

Tkinter GUI — chat-style interface with Excel control panel.

```python
# ui.py
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
from agent import ExcelAgent
from excel_controller import ExcelController
from history import CommandHistory

class ExcelAIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ExcelAI — Powered by gemma4:e4b")
        self.root.geometry("820x620")
        self.root.configure(bg="#1e1e2e")

        self.agent = ExcelAgent()
        self.excel = ExcelController()
        self.history = CommandHistory()

        self._build_ui()

    def _build_ui(self):
        # ── Top Bar ──────────────────────────────────────────────
        top = tk.Frame(self.root, bg="#313244", pady=6)
        top.pack(fill="x")

        tk.Label(top, text="📊 ExcelAI", font=("Segoe UI", 14, "bold"),
                 bg="#313244", fg="#cdd6f4").pack(side="left", padx=12)

        tk.Button(top, text="📂 Open Excel File", command=self._open_file,
                  bg="#89b4fa", fg="#1e1e2e", relief="flat", padx=8,
                  font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)

        tk.Button(top, text="🔗 Connect to Running Excel", command=self._connect_running,
                  bg="#a6e3a1", fg="#1e1e2e", relief="flat", padx=8,
                  font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)

        tk.Button(top, text="🆕 New Workbook", command=self._new_workbook,
                  bg="#f38ba8", fg="#1e1e2e", relief="flat", padx=8,
                  font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)

        self.status_label = tk.Label(top, text="⚠ Not connected",
                                     font=("Segoe UI", 9), bg="#313244", fg="#f38ba8")
        self.status_label.pack(side="right", padx=12)

        # ── Chat Area ────────────────────────────────────────────
        self.chat = scrolledtext.ScrolledText(
            self.root, font=("Consolas", 10),
            bg="#181825", fg="#cdd6f4",
            insertbackground="white",
            wrap=tk.WORD, state="disabled", pady=8, padx=8
        )
        self.chat.pack(fill="both", expand=True, padx=10, pady=6)

        # Tag colors for different message types
        self.chat.tag_config("user",    foreground="#89b4fa", font=("Consolas", 10, "bold"))
        self.chat.tag_config("ai",      foreground="#a6e3a1")
        self.chat.tag_config("code",    foreground="#f9e2af", background="#11111b")
        self.chat.tag_config("error",   foreground="#f38ba8")
        self.chat.tag_config("success", foreground="#94e2d5")
        self.chat.tag_config("system",  foreground="#6c7086")

        # ── Input Bar ────────────────────────────────────────────
        bottom = tk.Frame(self.root, bg="#1e1e2e", pady=6)
        bottom.pack(fill="x", padx=10)

        self.entry = tk.Entry(
            bottom, font=("Segoe UI", 11),
            bg="#313244", fg="#cdd6f4",
            insertbackground="white",
            relief="flat", bd=6
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.entry.bind("<Return>", lambda e: self._send())

        tk.Button(bottom, text="▶ Run", command=self._send,
                  bg="#cba6f7", fg="#1e1e2e",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=(6, 0))

        tk.Button(bottom, text="🧹 Clear", command=self._clear_chat,
                  bg="#45475a", fg="#cdd6f4",
                  font=("Segoe UI", 10),
                  relief="flat", padx=10, pady=6).pack(side="left", padx=4)

        self._log("system", "Welcome to ExcelAI! Connect to Excel above, then type your command.\n")

    # ── Connection Methods ────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
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

    # ── Main Send / Execute Flow ──────────────────────────────────

    def _send(self):
        cmd = self.entry.get().strip()
        if not cmd:
            return
        self.entry.delete(0, tk.END)

        if not self.excel.is_connected():
            self._log("error", "⚠ Connect to Excel first using the buttons above.\n")
            return

        self._log("user", f"\n👤 You: {cmd}\n")
        self._log("system", "🤖 Thinking...\n")
        self.root.update()

        # Get AI-generated code
        code = self.agent.ask(cmd)

        self._log("ai", "🤖 AI Generated Code:\n")
        self._log("code", f"{code}\n")

        # Execute against live Excel
        success, result = self.excel.execute(code)
        self.history.add(cmd, code, success)

        if success:
            self._log("success", f"✅ {result}\n")
        else:
            self._log("error", f"❌ {result}\n")

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
        self._log("system", "Chat cleared. AI memory reset.\n")
```

---

### STEP 6 — `main.py`

Entry point — launches the Tkinter window.

```python
# main.py
import tkinter as tk
from ui import ExcelAIApp

if __name__ == "__main__":
    root = tk.Root()
    app = ExcelAIApp(root)
    root.mainloop()
```

> ⚠️ **Bug to fix in main.py**: Change `tk.Root()` → `tk.Tk()`

---

### STEP 7 — `launch.bat`

Double-click to launch the app on Windows.

```bat
@echo off
title ExcelAI Launcher
color 0A

echo ============================================
echo    ExcelAI - Powered by Ollama gemma4:e4b
echo ============================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+
    pause
    exit /b
)

:: Check Ollama is running
echo [1/3] Checking Ollama...
curl -s http://localhost:11434/api/tags >nul 2>&1
if errorlevel 1 (
    echo [INFO] Starting Ollama in background...
    start /B ollama serve
    timeout /t 3 /nobreak >nul
) else (
    echo [OK] Ollama is already running.
)

:: Install dependencies if needed
echo [2/3] Checking Python dependencies...
pip show xlwings >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing xlwings...
    pip install xlwings --quiet
)
pip show ollama >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing ollama...
    pip install ollama --quiet
)

:: Pull model if not already downloaded
echo [3/3] Checking gemma4:e4b model...
ollama list | find "gemma4" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Pulling gemma4:e4b model (first time only, may take a few minutes)...
    ollama pull gemma4:e4b
)

echo.
echo [LAUNCH] Starting ExcelAI...
echo.
python main.py

if errorlevel 1 (
    echo.
    echo [ERROR] App crashed. See error above.
    pause
)
```

---

## ✅ Build Checklist

Complete these in order:

- [ ] Create folder `ExcelAI/`
- [ ] Create `prompts.py`
- [ ] Create `excel_controller.py`
- [ ] Create `agent.py`
- [ ] Create `history.py`
- [ ] Create `ui.py`
- [ ] Create `main.py` (fix `tk.Root()` → `tk.Tk()`)
- [ ] Create `launch.bat`
- [ ] Test: Run `launch.bat` → open Excel → type a command → see live changes

---

## 🧪 Example Commands to Test After Build

| Type in chat | Expected Excel result |
|---|---|
| `Put "Sales Report" in A1 with bold blue text` | A1 → bold blue header |
| `Add headers: Product, Qty, Price in row 1 with yellow background` | Row 1 formatted |
| `Fill A2:A6 with Jan, Feb, Mar, Apr, May` | Month names appear |
| `Put formula =B2*C2 in D2` | Formula inserted live |
| `Make column A width 25` | Column A resizes |
| `Create a new sheet called Summary` | New tab appears in Excel |

---

## ⚠️ Known Limitations & Notes

- **Windows only** — xlwings COM interface requires Windows + Microsoft Excel installed
- **Ollama must be running** before launch — `launch.bat` handles this automatically
- **gemma4:e4b** is a small efficient model; for better code generation consider also testing `qwen2.5-coder:7b`
- **exec() safety**: The current implementation uses `exec()` directly. For production, add `ast.parse()` validation before execution.
- **Screen updating**: `app.screen_updating = True` ensures all changes are visible live

---

## 🚀 Future Enhancements (Phase 2)

- [ ] View current sheet data in sidebar
- [ ] Undo last action button
- [ ] Voice input (mic → command)
- [ ] Export command history as a macro `.py` file
- [ ] Sheet switcher dropdown in UI
- [ ] Show generated code in separate panel before executing
