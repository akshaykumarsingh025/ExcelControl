# agent.py
import ollama
import re
from prompts import SYSTEM_PROMPT


class ExcelAgent:
    def __init__(self):
        self.model = "gemma4:e4b"
        self.conversation_history = []

    def ask(self, user_command: str, context: str = "") -> str:
        """Send command to Ollama and return the generated Python code."""

        # Embed context invisibly if provided
        prompt = user_command
        if context:
            prompt = f"--- CURRENT SHEET CONTEXT (First 5 rows) ---\n{context}\n\n--- TASK ---\n{user_command}"

        self.conversation_history.append({"role": "user", "content": prompt})

        try:
            response = ollama.chat(
                model=self.model,
                messages=[{"role": "system", "content": SYSTEM_PROMPT}]
                + self.conversation_history,
            )

            raw = response["message"]["content"]

            # Remember AI response in history for context
            self.conversation_history.append({"role": "assistant", "content": raw})

            # Strip markdown code blocks if model adds them
            clean = re.sub(r"```python|```", "", raw).strip()
            return clean

        except Exception as e:
            return f"# ERROR: Could not reach Ollama\n# {str(e)}"

    def reset_memory(self):
        """Clear conversation history."""
        self.conversation_history = []
