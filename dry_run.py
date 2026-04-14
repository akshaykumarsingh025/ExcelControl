# dry_run.py
import ast


class DryRunResult:
    def __init__(
        self,
        cells_read: list,
        cells_written: list,
        formulas: list,
        formatting: list,
        sheet_ops: list,
        warnings: list,
    ):
        self.cells_read = cells_read
        self.cells_written = cells_written
        self.formulas = formulas
        self.formatting = formatting
        self.sheet_ops = sheet_ops
        self.warnings = warnings

    def summary(self) -> str:
        lines = ["=== DRY RUN PREVIEW ===\n"]
        if self.cells_written:
            lines.append(f"Cells to WRITE: {len(self.cells_written)}")
            for c in self.cells_written[:10]:
                lines.append(f"  {c}")
            if len(self.cells_written) > 10:
                lines.append(f"  ... and {len(self.cells_written) - 10} more")
        if self.formulas:
            lines.append(f"\nFormulas to INSERT: {len(self.formulas)}")
            for f in self.formulas:
                lines.append(f"  {f}")
        if self.formatting:
            lines.append(f"\nFormatting CHANGES: {len(self.formatting)}")
            for f in self.formatting[:10]:
                lines.append(f"  {f}")
            if len(self.formatting) > 10:
                lines.append(f"  ... and {len(self.formatting) - 10} more")
        if self.sheet_ops:
            lines.append(f"\nSheet OPERATIONS: {len(self.sheet_ops)}")
            for s in self.sheet_ops:
                lines.append(f"  {s}")
        if self.warnings:
            lines.append(f"\nWARNINGS:")
            for w in self.warnings:
                lines.append(f"  {w}")
        if not any(
            [self.cells_written, self.formulas, self.formatting, self.sheet_ops]
        ):
            lines.append("No detectable changes (code may use dynamic references).")
        lines.append("\n=== END PREVIEW ===")
        return "\n".join(lines)


def analyze_code(code: str) -> DryRunResult:
    cells_read = []
    cells_written = []
    formulas = []
    formatting = []
    sheet_ops = []
    warnings = []

    try:
        tree = ast.parse(code)
    except SyntaxError:
        warnings.append("Code has syntax errors, cannot analyze.")
        return DryRunResult(
            cells_read, cells_written, formulas, formatting, sheet_ops, warnings
        )

    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                attrs = _get_attribute_chain(target)
                if not attrs:
                    continue
                obj_path = ".".join(attrs)

                if ".value" in obj_path or ".formula" in obj_path:
                    cell_ref = _extract_cell_ref(obj_path)
                    if ".formula" in obj_path:
                        val = _try_get_value(node.value)
                        formulas.append(f"{cell_ref} = {val}")
                    else:
                        val = _try_get_value(node.value)
                        cells_written.append(f"{cell_ref} = {val}")

                elif ".color" in obj_path or ".bold" in obj_path or ".size" in obj_path:
                    cell_ref = _extract_cell_ref(obj_path)
                    val = _try_get_value(node.value)
                    formatting.append(f"{cell_ref} {obj_path.split('.')[-1]} = {val}")

                elif ".column_width" in obj_path or ".row_height" in obj_path:
                    cell_ref = _extract_cell_ref(obj_path)
                    val = _try_get_value(node.value)
                    formatting.append(f"{cell_ref} {obj_path.split('.')[-1]} = {val}")

                elif ".number_format" in obj_path:
                    cell_ref = _extract_cell_ref(obj_path)
                    val = _try_get_value(node.value)
                    formatting.append(f"{cell_ref} number_format = {val}")

                elif ".name" in obj_path and "ws." in obj_path:
                    val = _try_get_value(node.value)
                    sheet_ops.append(f"Rename sheet to {val}")

        elif isinstance(node, ast.Call):
            attrs = _get_attribute_chain(node.func)
            if not attrs:
                if isinstance(node.func, ast.Name):
                    if node.func.id == "wb" and hasattr(node, "args"):
                        pass
                continue
            obj_path = ".".join(attrs)

            if obj_path.endswith(".sheets.add") or obj_path.endswith(".sheets"):
                sheet_ops.append("Add new sheet")
            elif ".charts.add" in obj_path:
                sheet_ops.append("Add chart")
            elif ".pictures.add" in obj_path:
                sheet_ops.append("Add picture")
            elif ".delete" in obj_path:
                sheet_ops.append("Delete operation")
            elif ".autofit" in obj_path:
                formatting.append("Autofit columns/rows")
            elif ".clear" in obj_path:
                sheet_ops.append("Clear operation")

    return DryRunResult(
        cells_read, cells_written, formulas, formatting, sheet_ops, warnings
    )


def _get_attribute_chain(node) -> list[str]:
    parts = []
    current = node
    while isinstance(current, ast.Attribute):
        parts.append(current.attr)
        current = current.value
    if isinstance(current, ast.Name):
        parts.append(current.id)
    parts.reverse()
    return parts


def _extract_cell_ref(obj_path: str) -> str:
    for i, part in enumerate(obj_path.split(".")):
        if part in ("ws", "wb") and i + 1 < len(obj_path.split(".")):
            remaining = obj_path.split(".")[i + 1 :]
            return ".".join(remaining)
    return obj_path


def _try_get_value(node) -> str:
    if isinstance(node, ast.Constant):
        return repr(node.value)
    elif isinstance(node, ast.List):
        items = [_try_get_value(e) for e in node.elts]
        return "[" + ", ".join(items) + "]"
    elif isinstance(node, ast.Call):
        if isinstance(node.func, ast.Name):
            return f"{node.func.id}(...)"
        return "..."
    elif isinstance(node, ast.Name):
        return node.id
    elif isinstance(node, ast.BinOp):
        left = _try_get_value(node.left)
        right = _try_get_value(node.right)
        return f"{left} op {right}"
    return "..."
