# agent.py
import ollama
import re
import subprocess
from prompts import SYSTEM_PROMPT, ANALYSIS_SYSTEM_PROMPT


def get_available_models() -> list[str]:
    try:
        result = subprocess.run(
            ["ollama", "list"], capture_output=True, text=True, timeout=10
        )
        lines = result.stdout.strip().split("\n")
        models = []
        for line in lines[1:]:
            parts = line.split()
            if parts:
                name = parts[0]
                if name and not name.startswith("NAME"):
                    models.append(name)
        return models if models else ["gemma4:e4b"]
    except Exception:
        return ["gemma4:e4b"]


class ExcelAgent:
    def __init__(self, model: str = "gemma4:e4b"):
        self.model = model
        self.conversation_history = []
        self.analysis_mode = False

    def set_model(self, model: str):
        self.model = model

    def set_analysis_mode(self, enabled: bool):
        self.analysis_mode = enabled

    def ask(self, user_command: str, context: str = "") -> str:
        prompt = user_command
        if context:
            prompt = f"--- CURRENT SHEET CONTEXT (First 5 rows) ---\n{context}\n\n--- TASK ---\n{user_command}"

        self.conversation_history.append({"role": "user", "content": prompt})

        system_prompt = ANALYSIS_SYSTEM_PROMPT if self.analysis_mode else SYSTEM_PROMPT

        try:
            response = ollama.chat(
                model=self.model,
                messages=[{"role": "system", "content": system_prompt}]
                + self.conversation_history,
            )

            raw = response["message"]["content"]
            self.conversation_history.append({"role": "assistant", "content": raw})
            clean = re.sub(r"```python|```", "", raw).strip()
            return clean

        except Exception as e:
            return f"# ERROR: Could not reach Ollama\n# {str(e)}"

    def execute_analysis(self, code: str, ws, wb) -> str:
        import io
        import contextlib

        local_vars = {"ws": ws, "wb": wb, "xw": __import__("xlwings")}
        local_vars["statistics"] = __import__("statistics")
        local_vars["math"] = __import__("math")
        local_vars["collections"] = __import__("collections")
        local_vars["datetime"] = __import__("datetime")

        output = io.StringIO()
        try:
            with contextlib.redirect_stdout(output):
                exec(code, {"__builtins__": {}}, local_vars)
            return output.getvalue()
        except Exception as e:
            return f"Analysis Error: {str(e)}"

    def reset_memory(self):
        self.conversation_history = []

    def get_history(self) -> list:
        return self.conversation_history

    def set_history(self, history: list):
        self.conversation_history = history
