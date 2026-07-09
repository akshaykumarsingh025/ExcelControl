import ast
import importlib

_MODULES = {
    "json": __import__("json"),
    "math": __import__("math"),
    "datetime": __import__("datetime"),
    "re": __import__("re"),
    "collections": __import__("collections"),
    "statistics": __import__("statistics"),
    "time": __import__("time"),
    "random": __import__("random"),
}

SAFE_IMPORTS = frozenset(_MODULES.keys())

BLOCKED_BUILTINS = frozenset({
    "eval", "exec", "compile", "open", "globals", "locals", "vars",
    "input", "breakpoint", "exit", "quit",
    "getattr", "setattr", "delattr", "hasattr",
    "__import__",
})

SAFE_BUILTINS = {
    "print": print,
    "range": range,
    "len": len,
    "int": int,
    "float": float,
    "str": str,
    "bool": bool,
    "list": list,
    "dict": dict,
    "tuple": tuple,
    "set": set,
    "enumerate": enumerate,
    "zip": zip,
    "map": map,
    "filter": filter,
    "sorted": sorted,
    "reversed": reversed,
    "abs": abs,
    "min": min,
    "max": max,
    "sum": sum,
    "round": round,
    "isinstance": isinstance,
    "type": type,
    "True": True,
    "False": False,
    "None": None,
}


class SandboxEscapeError(Exception):
    pass


def _safe_import(name, *args, **kwargs):
    if name in SAFE_IMPORTS:
        return importlib.import_module(name)
    raise ImportError(f"Module '{name}' is not allowed in this sandbox")


def _make_type_wrapper():
    real_type = type

    def _restricted_type(*args, **kwargs):
        return real_type(*args, **kwargs)

    _restricted_type.__subclasses__ = lambda: (_ for _ in ())
    _restricted_type.mro = lambda self=None: []
    _restricted_type.__bases__ = ()
    _restricted_type.__mro__ = ()
    return _restricted_type


def build_sandbox_globals():
    globals_dict = {
        "__builtins__": {
            "__import__": _safe_import,
            **{k: v for k, v in SAFE_BUILTINS.items() if k != "type"},
            "type": _make_type_wrapper(),
        },
    }
    globals_dict.update(_MODULES)
    return globals_dict


def build_sandbox_locals(ws, wb):
    import xlwings as xw
    return {"ws": ws, "wb": wb, "xw": xw}


def compile_restricted(code: str):
    tree = ast.parse(code)

    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute):
            attr_name = node.attr
            if attr_name.startswith("__") and attr_name.endswith("__"):
                dunder_attrs = {
                    "__class__", "__mro__", "__bases__", "__subclasses__",
                    "__globals__", "__code__", "__func__", "__self__",
                    "__dict__", "__init__", "__new__",
                }
                if attr_name in dunder_attrs:
                    raise SandboxEscapeError(
                        f"Access to attribute '{attr_name}' is blocked"
                    )

        if isinstance(node, ast.Call):
            if isinstance(node.func, ast.Attribute):
                if node.func.attr == "__subclasses__":
                    raise SandboxEscapeError(
                        "Call to __subclasses__() is blocked"
                    )
                if node.func.attr in ("__init_subclass__",):
                    raise SandboxEscapeError(
                        f"Call to {node.func.attr} is blocked"
                    )

        if isinstance(node, ast.Name):
            if node.id in BLOCKED_BUILTINS:
                raise SandboxEscapeError(
                    f"Use of builtin '{node.id}' is blocked"
                )

    compiled = compile(tree, "<sandbox>", "exec")
    return compiled
