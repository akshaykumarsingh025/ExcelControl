# code_validator.py
import ast


BLOCKED_IMPORTS = {
    "os",
    "subprocess",
    "shutil",
    "sys",
    "pathlib",
    "socket",
    "http",
    "urllib",
    "requests",
    "pickle",
    "shelve",
    "ctypes",
    "multiprocessing",
    "importlib",
    "code",
    "codeop",
    "compileall",
    "webbrowser",
    "ftplib",
    "smtplib",
    "telnetlib",
}

BLOCKED_BUILTINS = {
    "eval",
    "exec",
    "compile",
    "__import__",
    "open",
    "globals",
    "locals",
    "vars",
    "input",
    "breakpoint",
    "exit",
    "quit",
}

ALLOWED_OPEN = False


class ValidationResult:
    def __init__(self, is_safe: bool, issues: list[str], code: str):
        self.is_safe = is_safe
        self.issues = issues
        self.code = code

    def __bool__(self):
        return self.is_safe


def validate_code(code: str) -> ValidationResult:
    issues = []

    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        return ValidationResult(False, [f"Syntax Error: {e}"], code)

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                root_module = alias.name.split(".")[0]
                if root_module in BLOCKED_IMPORTS:
                    issues.append(f"Blocked import: {alias.name}")

        elif isinstance(node, ast.ImportFrom):
            if node.module:
                root_module = node.module.split(".")[0]
                if root_module in BLOCKED_IMPORTS:
                    issues.append(f"Blocked import from: {node.module}")

        elif isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Name) and func.id in BLOCKED_BUILTINS:
                if func.id == "open" and ALLOWED_OPEN:
                    continue
                issues.append(f"Blocked builtin call: {func.id}")

            if isinstance(func, ast.Attribute):
                if isinstance(func.value, ast.Name):
                    if func.value.id in BLOCKED_IMPORTS:
                        issues.append(
                            f"Blocked attribute call: {func.value.id}.{func.attr}"
                        )

        elif isinstance(node, ast.Attribute):
            if isinstance(node.value, ast.Name):
                if node.value.id == "__builtins__":
                    issues.append("Blocked access to __builtins__")

    is_safe = len(issues) == 0
    return ValidationResult(is_safe, issues, code)
