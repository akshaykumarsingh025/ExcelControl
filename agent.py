# agent.py
import ast
import base64
import io
import json
import os
import re
import subprocess
import tempfile
import ollama
from prompts import (
    SYSTEM_PROMPT,
    ANALYSIS_SYSTEM_PROMPT,
    VISION_SYSTEM_PROMPT,
    STRUCTURE_ANALYSIS_PROMPT,
    JSON_EXTRACTION_SYSTEM_PROMPT,
    make_json_extraction_prompt,
    make_right_half_extraction_prompt,
)
from validators import validate_and_correct_table

OLLAMA_HOST = "http://localhost:11434"


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
            client = ollama.Client(host=OLLAMA_HOST)
            response = client.chat(
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

    # ── Image preprocessing ───────────────────────────────────────────────

    def _normalize_and_enhance(self, img):
        from PIL import Image as PILImage
        import PIL.ImageEnhance as IE
        import PIL.ImageFilter as IF
        import PIL.ImageChops as IC
        import PIL.ImageOps as IO

        if img.mode == "RGBA":
            img = img.convert("RGB")

        img = img.convert("L")
        img = img.filter(IF.MedianFilter(size=3))

        bg = img.filter(IF.GaussianBlur(radius=25))
        img = IC.subtract(img, bg, scale=1.0, offset=128)

        img = IO.autocontrast(img, cutoff=2)
        img = IE.Contrast(img).enhance(1.8)
        img = img.filter(IF.SHARPEN)

        try:
            import numpy as np
            arr = np.array(img).astype(np.float32)
            hist, _ = np.histogram(arr.flatten(), bins=256, range=(0, 256))
            total = arr.size
            sum_total = np.sum(np.arange(256) * hist)
            sum_bg = 0.0
            weight_bg = 0
            max_var = 0.0
            threshold = 128
            for t in range(256):
                weight_bg += hist[t]
                if weight_bg == 0:
                    continue
                weight_fg = total - weight_bg
                if weight_fg == 0:
                    break
                sum_bg += t * hist[t]
                mean_bg = sum_bg / weight_bg
                mean_fg = (sum_total - sum_bg) / weight_fg
                var_between = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2
                if var_between > max_var:
                    max_var = var_between
                    threshold = t
            k = 0.03
            arr = 255.0 / (1.0 + np.exp(-k * (arr - threshold)))
            arr = np.clip(arr, 0, 255).astype(np.uint8)
            img = PILImage.fromarray(arr)
        except ImportError:
            pass

        img = img.convert("RGB")
        return img

    def _deskew_image(self, img):
        try:
            import numpy as np
            from PIL import Image as PILImage

            gray = img.convert("L") if img.mode != "L" else img
            arr = np.array(gray)
            inv = 255 - arr

            best_angle = 0
            best_score = 0
            for angle_10x in range(-50, 51, 5):
                angle = angle_10x / 10.0
                rotated = img.rotate(angle, expand=False, fillcolor=255)
                rot_arr = 255 - np.array(rotated.convert("L"))
                projection = np.sum(rot_arr, axis=1)
                score = np.var(projection)
                if score > best_score:
                    best_score = score
                    best_angle = angle

            if abs(best_angle) > 0.1:
                img = img.rotate(best_angle, expand=True, fillcolor=255)
            return img
        except ImportError:
            return img

    def _preprocess_image(self, image_path: str) -> bytes:
        try:
            from PIL import Image as PILImage
        except ImportError:
            with open(image_path, "rb") as f:
                return f.read()

        img = PILImage.open(image_path)

        if img.mode == "RGBA":
            img = img.convert("RGB")

        img = self._deskew_image(img)

        max_dim = 4096
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)

        if max(img.size) < 3072:
            scale = 3072 / max(img.size)
            img = img.resize(
                (int(img.width * scale), int(img.height * scale)),
                PILImage.LANCZOS,
            )

        img = self._normalize_and_enhance(img)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    def _preprocess_strip(self, img):
        from PIL import Image as PILImage

        if img.mode == "RGBA":
            img = img.convert("RGB")

        max_dim = 4096
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)

        if max(img.size) < 3072:
            scale = 3072 / max(img.size)
            img = img.resize(
                (int(img.width * scale), int(img.height * scale)),
                PILImage.LANCZOS,
            )

        img = self._normalize_and_enhance(img)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    # ── API calls ─────────────────────────────────────────────────────────

    def _call_ollama_api(self, prompt: str, image_bytes: bytes,
                         system_prompt: str = None,
                         json_mode: bool = False) -> str:
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        if system_prompt is None:
            system_prompt = VISION_SYSTEM_PROMPT

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt, "images": [b64]},
            ],
            "stream": False,
            "options": {"temperature": 0, "num_predict": 8192},
        }

        if json_mode:
            payload["format"] = "json"

        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        )
        try:
            json.dump(payload, tmp)
            tmp.close()

            result = subprocess.run(
                [
                    "curl", "-s", "-X", "POST",
                    f"{OLLAMA_HOST}/api/chat",
                    "-d", f"@{tmp.name}",
                ],
                capture_output=True,
                text=True,
                timeout=600,
            )

            if result.returncode != 0:
                return f"# ERROR: curl failed: {result.stderr}"

            resp = json.loads(result.stdout)
            raw = resp["message"]["content"]
            return re.sub(r"```python|```json|```", "", raw).strip()

        except Exception as e:
            return f"# ERROR: {str(e)}"
        finally:
            os.unlink(tmp.name)

    # ── JSON parsing ──────────────────────────────────────────────────────

    def _parse_json_rows(self, text: str) -> list | None:
        text = text.strip()
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                try:
                    parsed = json.loads(match.group())
                except json.JSONDecodeError:
                    return None
            else:
                match = re.search(r"\[.*\]", text, re.DOTALL)
                if match:
                    try:
                        parsed = json.loads(match.group())
                        if isinstance(parsed, list):
                            return parsed
                    except json.JSONDecodeError:
                        return None
                return None

        if isinstance(parsed, dict):
            if "rows" in parsed:
                return parsed["rows"]
            if "data" in parsed:
                return parsed["data"]
            for v in parsed.values():
                if isinstance(v, list):
                    return v
        elif isinstance(parsed, list):
            return parsed

        return None

    def _extract_data_list(self, text: str):
        idx = text.find("data = [")
        if idx < 0:
            return None
        depth = 1
        i = idx + len("data = [")
        while i < len(text):
            ch = text[i]
            if ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0:
                    end = i + 1
                    break
            i += 1
        else:
            return None
        try:
            return ast.literal_eval(text[idx : idx + len("data = ") + (end - idx)])
        except Exception:
            return None

    # ── Core extraction: Pass 1 ───────────────────────────────────────────

    def _analyze_structure(self, image_bytes: bytes) -> dict | None:
        prompt = "Analyze this table image. Count the data rows and identify all column headers."
        raw = self._call_ollama_api(
            prompt, image_bytes,
            system_prompt=STRUCTURE_ANALYSIS_PROMPT,
            json_mode=True,
        )
        try:
            result = json.loads(raw)
            if "num_rows" in result and "headers" in result:
                return result
        except (json.JSONDecodeError, TypeError):
            pass
        return None

    # ── Core extraction: Pass 2 (single image or strip) ───────────────────

    def _extract_json_data(self, image_bytes: bytes, headers: list[str] | None,
                            expected_rows: int | None = None,
                            strip_label: str = "") -> list | None:
        prompt = make_json_extraction_prompt(
            headers=headers,
            expected_rows=expected_rows,
            strip_label=strip_label,
        )
        raw = self._call_ollama_api(
            prompt, image_bytes,
            system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
            json_mode=True,
        )
        return self._parse_json_rows(raw)

    # ── Row helpers ───────────────────────────────────────────────────────

    @staticmethod
    def _is_null(val) -> bool:
        if val is None:
            return True
        s = str(val).strip().lower()
        return s in ("null", "none", "")

    @staticmethod
    def _try_int(val) -> int | None:
        try:
            return int(str(val).strip())
        except (ValueError, TypeError):
            return None

    def _rows_are_same_record(self, row_a, row_b):
        if len(row_a) > 0 and len(row_b) > 0:
            a_sno = self._try_int(row_a[0])
            b_sno = self._try_int(row_b[0])
            if a_sno is not None and b_sno is not None:
                if a_sno == b_sno:
                    return True

        min_len = min(len(row_a), len(row_b))
        if min_len >= 2:
            a_name = str(row_a[1]).strip().lower() if row_a[1] is not None and not self._is_null(row_a[1]) else ""
            b_name = str(row_b[1]).strip().lower() if row_b[1] is not None and not self._is_null(row_b[1]) else ""
            if a_name and b_name and len(a_name) >= 3 and a_name == b_name:
                return True

        return False

    def _merge_rows(self, row_a, row_b):
        max_len = max(len(row_a), len(row_b))
        merged = []
        for i in range(max_len):
            a = row_a[i] if i < len(row_a) else None
            b = row_b[i] if i < len(row_b) else None

            a_null = self._is_null(a)
            b_null = self._is_null(b)

            if a_null and b_null:
                merged.append(None)
            elif a_null:
                merged.append(b)
            elif b_null:
                merged.append(a)
            elif str(a).strip() == str(b).strip():
                merged.append(a)
            elif len(str(b).strip()) > len(str(a).strip()):
                merged.append(b)
            else:
                merged.append(a)
        return merged

    # ── Strip-based extraction for tall images ────────────────────────────

    def _extract_with_strips(self, img, headers, expected_rows, extract_fn=None):
        """Extract data from a PIL Image using horizontal strips.

        Args:
            img: PIL Image (already preprocessed or original)
            headers: Column headers list
            expected_rows: Expected number of data rows
            extract_fn: Function(image_bytes, headers, expected_rows, strip_label) -> list|None
                        Defaults to self._extract_json_data
        """
        if extract_fn is None:
            extract_fn = self._extract_json_data

        w, h = img.size
        rows_per_strip = 10
        num_strips = max(2, (expected_rows + rows_per_strip - 1) // rows_per_strip)
        strip_height = h // num_strips
        overlap = int(strip_height * 0.3)

        strips = []
        for i in range(num_strips):
            y_start = max(0, i * strip_height - (overlap if i > 0 else 0))
            y_end = min(h, (i + 1) * strip_height + (overlap if i < num_strips - 1 else 0))
            strip_img = img.crop((0, y_start, w, y_end))
            labels = ["top", "upper-middle", "middle", "lower-middle", "bottom"]
            label = labels[min(i, len(labels) - 1)] if num_strips <= 5 else f"strip {i+1} of {num_strips}"
            strips.append((label, strip_img))

        all_rows = []
        for label, strip_img in strips:
            strip_bytes = self._preprocess_strip(strip_img)
            rows_in_strip = extract_fn(strip_bytes, headers, rows_per_strip, label)
            if rows_in_strip:
                for row in rows_in_strip:
                    row = [None if self._is_null(c) else c for c in row]
                    merged = False
                    for i, existing in enumerate(all_rows):
                        if self._rows_are_same_record(existing, row):
                            all_rows[i] = self._merge_rows(existing, row)
                            merged = True
                            break
                    if not merged:
                        all_rows.append(row)

        return all_rows

    # ── Two-page split extraction ──────────────────────────────────────────

    def _split_headers(self, headers: list[str]) -> tuple[list[str], list[str], int]:
        """Determine left-page and right-page headers from the full header list.

        For two-page spreads, the right page typically starts at the first column
        that looks like a numeric/code column after text columns.

        Returns (left_headers, right_headers, split_index).
        split_index is the index in the full headers where the right page begins.
        """
        if not headers:
            return headers, [], len(headers)

        split_idx = len(headers)

        text_keywords = {"name", "designation", "title", "type", "category", "holder", "serial", "s.no", "s no"}
        code_keywords = {"account", "ifsc", "bank", "branch", "code", "number", "no", "amount", "id", "micr", "upi"}

        for i, h in enumerate(headers):
            h_lower = h.lower().strip()
            for kw in code_keywords:
                if kw in h_lower and i > 0:
                    prev_lower = headers[i - 1].lower().strip()
                    prev_is_text = any(kw2 in prev_lower for kw2 in text_keywords)
                    curr_is_code = any(kw2 in h_lower for kw2 in code_keywords)
                    if prev_is_text and curr_is_code:
                        split_idx = i
                        break
            if split_idx < len(headers):
                break

        if split_idx >= len(headers):
            mid = len(headers) // 2
            if mid > 1:
                split_idx = mid

        left_headers = headers[:split_idx]
        right_headers = ["S.No"] + headers[split_idx:]
        return left_headers, right_headers, split_idx

    def _extract_left_page(self, image_path: str, headers: list[str],
                           expected_rows: int) -> list:
        """Extract left-page columns by cropping the left half of the image."""
        from PIL import Image as PILImage

        img = PILImage.open(image_path)
        w, h = img.size

        margin = int(w * 0.04)
        right_boundary = (w // 2) + margin
        left_img = img.crop((0, 0, right_boundary, h))

        left_headers, _, _ = self._split_headers(headers)

        if expected_rows > 10:
            return self._extract_with_strips(left_img, left_headers, expected_rows)
        else:
            left_bytes = self._preprocess_strip(left_img)
            rows = self._extract_json_data(left_bytes, left_headers, expected_rows, "left page")
            return rows if rows else []

    def _extract_right_page(self, image_path: str, headers: list[str],
                            expected_rows: int) -> list:
        """Extract right-page columns by cropping the right half of the image."""
        from PIL import Image as PILImage

        img = PILImage.open(image_path)
        w, h = img.size

        margin = int(w * 0.04)
        left_boundary = (w // 2) - margin
        right_img = img.crop((left_boundary, 0, w, h))

        _, right_headers, _ = self._split_headers(headers)

        def extract_right_strip(image_bytes, hdrs, exp_rows, strip_label):
            prompt = make_right_half_extraction_prompt(hdrs, exp_rows, strip_label)
            raw = self._call_ollama_api(
                prompt, image_bytes,
                system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
                json_mode=True,
            )
            return self._parse_json_rows(raw)

        if expected_rows > 10:
            return self._extract_with_strips(right_img, right_headers, expected_rows, extract_right_strip)
        else:
            right_bytes = self._preprocess_strip(right_img)
            prompt = make_right_half_extraction_prompt(right_headers, expected_rows, "")
            raw = self._call_ollama_api(
                prompt, right_bytes,
                system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
                json_mode=True,
            )
            rows = self._parse_json_rows(raw)
            return rows if rows else []

    def _merge_left_right(self, left_rows: list, right_rows: list,
                          headers: list[str]) -> list:
        """Merge left-page and right-page extraction results row by row.

        Strategy:
        1. Match by S.No (col 0 in both left and right rows)
        2. Fall back to row order for unmatched rows
        3. Right-page rows have format: [S.No, right_col1, right_col2, ...]
           Left-page rows have format: [S.No, left_col1, left_col2, ..., None, None, ...]
        """
        _, _, split_idx = self._split_headers(headers)
        right_data_start = split_idx

        right_by_sno = {}
        right_unmatched = []
        for row in right_rows:
            row = [None if self._is_null(c) else c for c in row]
            sno = self._try_int(row[0]) if len(row) > 0 else None
            if sno is not None:
                right_by_sno[sno] = row
            else:
                right_unmatched.append(row)

        merged = []
        used_sno = set()

        for left_row in left_rows:
            left_row = [None if self._is_null(c) else c for c in left_row]
            full_row = list(left_row)
            while len(full_row) < len(headers):
                full_row.append(None)

            matched_right = None
            sno = self._try_int(left_row[0]) if len(left_row) > 0 else None

            if sno is not None and sno in right_by_sno:
                matched_right = right_by_sno[sno]
                used_sno.add(sno)

            if matched_right is not None:
                right_vals = matched_right[1:]
                for ri, val in enumerate(right_vals):
                    target_col = right_data_start + ri
                    if target_col < len(full_row):
                        if self._is_null(full_row[target_col]) and not self._is_null(val):
                            full_row[target_col] = val

            merged.append(full_row)

        remaining_right = [row for sno, row in right_by_sno.items() if sno not in used_sno]
        remaining_right.extend(right_unmatched)

        if remaining_right:
            main_missing = []
            for i, row in enumerate(merged):
                has_right = any(
                    not self._is_null(row[j]) for j in range(right_data_start, len(row))
                )
                if not has_right:
                    main_missing.append(i)

            ri = 0
            for main_idx in main_missing:
                if ri >= len(remaining_right):
                    break
                right_row = remaining_right[ri]
                right_vals = right_row[1:]
                for rvi, val in enumerate(right_vals):
                    target_col = right_data_start + rvi
                    if target_col < len(merged[main_idx]):
                        if self._is_null(merged[main_idx][target_col]) and not self._is_null(val):
                            merged[main_idx][target_col] = val
                ri += 1

        return merged

    # ── Main entry point ──────────────────────────────────────────────────

    def ask_with_image_json(self, image_path: str, user_command: str = "") -> str:
        try:
            raw_bytes = self._preprocess_image(image_path)

            if user_command:
                return self.ask_with_image(image_path, user_command)

            # Pass 1: Structure analysis
            structure = self._analyze_structure(raw_bytes)
            if structure:
                headers = structure.get("headers", [])
                expected_rows = structure.get("num_rows", None)
                is_two_page = structure.get("is_two_page_spread", False)
            else:
                headers = None
                expected_rows = None
                is_two_page = False

            # Detect two-page spread from image dimensions
            if not is_two_page:
                try:
                    from PIL import Image as PILImage
                    img = PILImage.open(image_path)
                    w, h = img.size
                    if w > h * 1.2:
                        is_two_page = True
                except ImportError:
                    pass

            if not headers:
                return self.ask_with_image(image_path, "")

            padded_expected = int((expected_rows or 25) * 1.3)

            # ── Pass 2a: Full-image extraction (reads left-side data best) ──
            full_img_rows = None
            try:
                from PIL import Image as PILImage
                full_img = PILImage.open(io.BytesIO(raw_bytes))
                if padded_expected > 15:
                    full_img_rows = self._extract_with_strips(full_img, headers, padded_expected)
                else:
                    full_img_rows = self._extract_json_data(raw_bytes, headers, padded_expected)
                    if full_img_rows:
                        full_img_rows = [[None if self._is_null(c) else c for c in row] for row in full_img_rows]
            except Exception:
                pass

            # ── Pass 2b: Split extraction for two-page spreads ──────────────
            split_rows = None
            if is_two_page:
                try:
                    left_rows = self._extract_left_page(image_path, headers, padded_expected)
                    right_rows = self._extract_right_page(image_path, headers, padded_expected)
                    if left_rows and right_rows:
                        split_rows = self._merge_left_right(left_rows, right_rows, headers)
                    elif left_rows:
                        split_rows = left_rows
                    elif right_rows:
                        split_rows = right_rows
                except Exception:
                    pass

            # ── Merge both extraction passes for maximum accuracy ───────────
            if full_img_rows and split_rows:
                all_rows = self._merge_multi_pass(full_img_rows, split_rows, headers)
            elif full_img_rows:
                all_rows = full_img_rows
            elif split_rows:
                all_rows = split_rows
            else:
                return self.ask_with_image(image_path, "")

            # Clean up null strings
            all_rows = [
                [None if self._is_null(c) else c for c in row]
                for row in all_rows
            ]

            # Ensure all rows have exactly len(headers) columns (trim or pad)
            num_cols = len(headers)
            all_rows = [
                row[:num_cols] + [None] * max(0, num_cols - len(row))
                for row in all_rows
            ]

            # Build full data with headers
            full_data = [headers] + all_rows

            # Post-process: run validators
            try:
                full_data, warnings = validate_and_correct_table(full_data)
                if warnings:
                    warning_comment = "# Validator corrections:\n" + "".join(
                        f"# - {w}\n" for w in warnings[:20]
                    )
                else:
                    warning_comment = "# All validations passed\n"
            except Exception:
                warning_comment = "# Validators could not run\n"

            import pprint
            data_str = pprint.pformat(full_data, indent=4)
            return (
                f"# Total rows: {len(full_data) - 1}\n"
                f"{warning_comment}\n"
                f"data = {data_str}\n\n"
                f'ws.range("A1").value = data\n'
                f'ws.range("A1").expand("right").font.bold = True\n'
                f"ws.autofit()\n"
            )

        except Exception as e:
            return f"# ERROR: Image processing failed\n# {str(e)}"

    def _merge_multi_pass(self, pass_a_rows: list, pass_b_rows: list,
                           headers: list[str]) -> list:
        """Merge results from two extraction passes for maximum accuracy.

        pass_a = full-image extraction (better left-side data)
        pass_b = split extraction (better right-side data)

        For each row in pass_a, find matching row in pass_b and merge,
        preferring non-null values from either pass.
        """
        idx_b_by_sno = {}
        idx_b_by_name = {}
        for idx, row in enumerate(pass_b_rows):
            if len(row) > 0 and row[0] is not None:
                sno = self._try_int(row[0])
                if sno is not None:
                    idx_b_by_sno[sno] = idx
            if len(row) > 1 and row[1] is not None:
                name_key = str(row[1]).strip().lower()
                if name_key and len(name_key) >= 3:
                    idx_b_by_name.setdefault(name_key, idx)

        used_b = set()
        merged = []

        for row_a in pass_a_rows:
            row_a = [None if self._is_null(c) else c for c in row_a]
            match_b_idx = None

            if len(row_a) > 0 and row_a[0] is not None:
                sno = self._try_int(row_a[0])
                if sno is not None and sno in idx_b_by_sno:
                    match_b_idx = idx_b_by_sno[sno]

            if match_b_idx is None and len(row_a) > 1 and row_a[1] is not None:
                name_key = str(row_a[1]).strip().lower()
                if name_key and len(name_key) >= 3 and name_key in idx_b_by_name:
                    candidate = idx_b_by_name[name_key]
                    if candidate not in used_b:
                        match_b_idx = candidate

            if match_b_idx is not None and match_b_idx not in used_b:
                row_b = pass_b_rows[match_b_idx]
                row_b = [None if self._is_null(c) else c for c in row_b]
                merged_row = self._merge_rows(row_a, row_b)
                used_b.add(match_b_idx)
            else:
                merged_row = list(row_a)

            while len(merged_row) < len(headers):
                merged_row.append(None)
            merged.append(merged_row)

        for idx, row_b in enumerate(pass_b_rows):
            if idx in used_b:
                continue
            row_b = [None if self._is_null(c) else c for c in row_b]
            while len(row_b) < len(headers):
                row_b.append(None)
            merged.append(row_b)

        return merged

    # ── Legacy methods ─────────────────────────────────────────────────────

    def ask_with_image(self, image_path: str, user_command: str = "") -> str:
        try:
            raw_bytes = self._preprocess_image(image_path)

            try:
                from PIL import Image as PILImage
                img = PILImage.open(image_path)
                w, h = img.size
                is_tall = h > w * 1.2
            except ImportError:
                is_tall = False

            if is_tall:
                return self._ask_with_strips(image_path, raw_bytes, user_command)

            if not user_command:
                user_command = self._extraction_prompt("")

            return self._call_ollama_api(user_command, raw_bytes)

        except Exception as e:
            return f"# ERROR: Image processing failed\n# {str(e)}"

    def _extraction_prompt(self, extra_note: str) -> str:
        return (
            "Extract EVERY row from this HANDWRITTEN table and write it to the sheet starting at A1.\n\n"
            "STEP 1 — LIST ALL DATA:\n"
            "Write: data = [\n"
            '    ["header1", "header2", ...],\n'
            '    ["row1col0", "row1col1", ...],\n'
            "    ... every single row ...\n"
            "]\n"
            "Numbers as int/float, text as strings, empty cells as None.\n\n"
            "Identify ALL actual columns from the table — do not merge or split columns. "
            "Read each column header exactly as written and preserve the column order.\n\n"
            "STEP 2 — XLWINGS CODE:\n"
            'ws.range("A1").value = data\n'
            'ws.range("A1").expand("right").font.bold = True\n'
            "ws.autofit()\n\n"
            f"{extra_note}"
            "CRITICAL: Extract absolutely every row — missing even one row is a failure."
        )

    def _ask_with_strips(
        self, image_path: str, raw_bytes: bytes, user_command: str
    ) -> str:
        from PIL import Image as PILImage

        img = PILImage.open(image_path)
        w, h = img.size

        overlap = int(h * 0.05)
        third = h // 3
        strips = [
            ("top", img.crop((0, 0, w, third * 2 + overlap))),
            ("middle", img.crop((0, third - overlap, w, third * 2 + overlap))),
            ("bottom", img.crop((0, third * 2 - overlap, w, h))),
        ]

        results = []
        prev_headers = None
        for i, (label, strip_img) in enumerate(strips):
            strip_bytes = self._preprocess_strip(strip_img)
            if user_command:
                prompt = user_command
            elif i == 0:
                prompt = self._extraction_prompt(
                    f"This is the top portion of the table. Extract ALL rows visible here.\n"
                )
            else:
                extra = f"This is the {label} portion of the table. Extract ALL rows visible here.\n"
                if prev_headers:
                    hdr = json.dumps(prev_headers)
                    extra += (
                        f"CRITICAL: The table has exactly {len(prev_headers)} columns: {hdr}.\n"
                        f"Use these EXACT column names in the same order. Do NOT rename, reorder, or add columns.\n"
                    )
                prompt = self._extraction_prompt(extra)
            result = self._call_ollama_api(prompt, strip_bytes)
            results.append((label, result))

            if i == 0:
                parsed = self._extract_data_list(result)
                if parsed and len(parsed) > 0:
                    prev_headers = parsed[0]

        combined = self._merge_strip_results(results)
        return combined

    def _merge_strip_results(self, results):
        all_data = []
        header = None
        header_cols = 0

        for label, text in results:
            parsed = self._extract_data_list(text)
            if not parsed or len(parsed) < 2:
                continue

            strip_header = parsed[0]
            strip_rows = parsed[1:]
            strip_cols = len(strip_header)

            if header is None:
                header = strip_header
                header_cols = strip_cols
                all_data.extend(strip_rows)
            else:
                for row in strip_rows:
                    padded = row + [None] * max(0, header_cols - len(row))
                    trimmed = padded[:header_cols]
                    merged = False
                    for i, existing in enumerate(all_data):
                        if self._rows_are_same_record(existing, trimmed):
                            all_data[i] = self._merge_rows(existing, trimmed)
                            merged = True
                            break
                    if not merged:
                        all_data.append(trimmed)

        if header is None:
            return results[0][1] if results else "# ERROR: No data extracted"

        combined = [header] + all_data
        data_str = json.dumps(combined, indent=4, ensure_ascii=False)

        return (
            f"data = {data_str}\n\n"
            f'ws.range("A1").value = data\n'
            f'ws.range("A1").expand("right").font.bold = True\n'
            f"ws.autofit()\n"
        )

    def execute_analysis(self, code: str, ws, wb) -> str:
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
