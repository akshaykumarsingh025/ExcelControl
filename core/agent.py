import re

import ollama

from prompts import (
    SYSTEM_PROMPT,
    ANALYSIS_SYSTEM_PROMPT,
    VISION_SYSTEM_PROMPT,
)


OLLAMA_HOST = "http://localhost:11434"


def get_available_models() -> list[str]:
    try:
        client = ollama.Client(host=OLLAMA_HOST)
        models_resp = client.list()
        models = []
        for model_info in models_resp.get("models", []):
            name = model_info.get("name", "")
            if name:
                models.append(name)
        return models if models else ["gemma4:e4b"]
    except Exception:
        return ["gemma4:e4b"]


class ExcelAgent:
    def __init__(self, model: str = "gemma4:e4b"):
        self.model = model
        self.conversation_history = []
        self.analysis_mode = False
        self._client = ollama.Client(host=OLLAMA_HOST)

    def set_model(self, model: str):
        self.model = model
        self._client = ollama.Client(host=OLLAMA_HOST)

    def set_analysis_mode(self, enabled: bool):
        self.analysis_mode = enabled

    def ask(self, user_command: str, context: str = "", images: list = None) -> str:
        prompt = user_command
        if context:
            prompt = f"--- CURRENT SHEET CONTEXT (First 5 rows) ---\n{context}\n\n--- TASK ---\n{user_command}"

        system_prompt = ANALYSIS_SYSTEM_PROMPT if self.analysis_mode else SYSTEM_PROMPT

        user_msg = {"role": "user", "content": prompt}
        if images:
            user_msg["images"] = images
            system_prompt = VISION_SYSTEM_PROMPT
        self.conversation_history.append(user_msg)

        try:
            response = self._client.chat(
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

    def ask_with_context(self, user_command: str, context: str = "",
                         mode: str = "automation") -> str:
        if mode == "analysis":
            self.analysis_mode = True
        else:
            self.analysis_mode = False

        return self.ask(user_command, context=context)

    def chat_with_data(self, question: str, sheet_data: list[list]) -> str:
        import json

        preview_rows = sheet_data[:25]
        data_str = json.dumps(preview_rows, indent=2, ensure_ascii=False)

        prompt = (
            f"--- SHEET DATA (first {len(preview_rows)} rows as JSON) ---\n"
            f"{data_str}\n\n"
            f"--- QUESTION ---\n{question}\n\n"
            f"Answer the question based on the data above. "
            f"If you need to compute something, write Python code to analyze it. "
            f"If a simple text answer suffices, respond plainly."
        )

        system_prompt = ANALYSIS_SYSTEM_PROMPT

        self.conversation_history.append({"role": "user", "content": prompt})

        try:
            response = self._client.chat(
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

    def generate_formula(self, description: str, context: str = "") -> str:
        prompt_parts = [
            "Generate an Excel formula based on this description.",
            f"Description: {description}",
        ]
        if context:
            prompt_parts.append(f"Context about the sheet:\n{context}")

        prompt_parts.append(
            "\nRespond with ONLY the formula string (no explanation, no backticks). "
            "For example: =SUM(A1:A10) or =VLOOKUP(B2,Sheet2!A:C,3,FALSE)"
        )

        prompt = "\n\n".join(prompt_parts)

        system_prompt = (
            "You are an Excel formula expert. "
            "Respond with ONLY a valid Excel formula string. "
            "No explanation, no markdown, no backticks. "
            "Just the formula starting with ="
        )

        self.conversation_history.append({"role": "user", "content": prompt})

        try:
            response = self._client.chat(
                model=self.model,
                messages=[{"role": "system", "content": system_prompt}]
                + self.conversation_history[-3:],
            )
            raw = response["message"]["content"]
            self.conversation_history.append({"role": "assistant", "content": raw})
            formula = raw.strip()
            formula = re.sub(r"^```.*?\n?", "", formula)
            formula = re.sub(r"\n?```$", "", formula)
            formula = formula.strip()
            if not formula.startswith("="):
                formula = "=" + formula
            return formula
        except Exception as e:
            return f"# ERROR: {str(e)}"

    def call_vision_api(self, prompt: str, image_bytes: bytes,
                        system_prompt: str = None,
                        json_mode: bool = False) -> str:
        import base64

        b64 = base64.b64encode(image_bytes).decode("utf-8")
        if system_prompt is None:
            system_prompt = VISION_SYSTEM_PROMPT

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt, "images": [b64]},
        ]

        kwargs = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            "options": {"temperature": 0, "num_predict": 8192},
        }

        if json_mode:
            kwargs["format"] = "json"

        try:
            response = self._client.chat(**kwargs)
            raw = response["message"]["content"]
            return re.sub(r"```python|```json|```", "", raw).strip()
        except Exception as e:
            return f"# ERROR: {str(e)}"

    def reset_memory(self):
        self.conversation_history = []

    def get_history(self) -> list:
        return self.conversation_history

    def set_history(self, history: list):
        self.conversation_history = history
