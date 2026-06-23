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

        # Background subtraction to handle uneven lighting (register crease)
        bg = img.filter(IF.GaussianBlur(radius=25))
        img = IC.subtract(img, bg, scale=1.0, offset=128)

        img = IO.autocontrast(img, cutoff=2)
        img = IE.Contrast(img).enhance(1.8)

        # Use SHARPEN instead of EDGE_ENHANCE_MORE — preserves thin handwritten strokes
        img = img.filter(IF.SHARPEN)

        # Soft contrast enhancement using Otsu's method (NOT hard binarization)
        # Hard B&W destroys handwriting detail — keep grayscale for the vision model
        try:
            import numpy as np
            arr = np.array(img).astype(np.float32)
            # Otsu threshold for reference point
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
            # Soft sigmoid-based contrast around the Otsu threshold
            # Pulls dark text darker and light background lighter, but preserves gradients
            k = 0.03  # sigmoid steepness (lower = softer)
            arr = 255.0 / (1.0 + np.exp(-k * (arr - threshold)))
            arr = np.clip(arr, 0, 255).astype(np.uint8)
            img = PILImage.fromarray(arr)
        except ImportError:
            pass  # Fall back to non-enhanced image

        img = img.convert("RGB")
        return img

    def _deskew_image(self, img):
        """Deskew a grayscale image using horizontal projection profile."""
        try:
            import numpy as np
            from PIL import Image as PILImage

            gray = img.convert("L") if img.mode != "L" else img
            arr = np.array(gray)
            # Invert so text is white
            inv = 255 - arr

            best_angle = 0
            best_score = 0
            # Search ±5 degrees in 0.5 degree steps
            for angle_10x in range(-50, 51, 5):
                angle = angle_10x / 10.0
                rotated = img.rotate(angle, expand=False, fillcolor=255)
                rot_arr = 255 - np.array(rotated.convert("L"))
                # Horizontal projection: sum each row
                projection = np.sum(rot_arr, axis=1)
                # Score = variance of projection (higher = more aligned)
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

        # Deskew before any resizing
        img = self._deskew_image(img)

        max_dim = 4096
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)

        # Upscale small images to at least 3072px for readable handwriting
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

    def _call_ollama_api(self, prompt: str, image_bytes: bytes,
                         system_prompt: str = None,
                         json_mode: bool = False) -> str:
        """Call Ollama chat API with an image.

        Args:
            prompt: The user message.
            image_bytes: PNG image as bytes.
            system_prompt: Override system prompt (defaults to VISION_SYSTEM_PROMPT).
            json_mode: If True, force JSON output via Ollama format parameter.
        """
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

    def _parse_json_rows(self, text: str) -> list | None:
        """Parse JSON response from Ollama into a list of rows.

        Handles both {"rows": [...]} format and bare [[...]] format.
        """
        text = text.strip()
        # Remove markdown code fences if present
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            # Try to extract JSON from surrounding text
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                try:
                    parsed = json.loads(match.group())
                except json.JSONDecodeError:
                    return None
            else:
                # Try bare array
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
            # Look for "rows" key
            if "rows" in parsed:
                return parsed["rows"]
            # Look for "data" key
            if "data" in parsed:
                return parsed["data"]
            # Look for first list value
            for v in parsed.values():
                if isinstance(v, list):
                    return v
        elif isinstance(parsed, list):
            return parsed

        return None

    def _extract_data_list(self, text: str):
        """Legacy: parse Python-style data = [...] from text."""
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

    # ── Two-Pass Extraction Architecture ──────────────────────────────────

    def _analyze_structure(self, image_bytes: bytes) -> dict | None:
        """Pass 1: Detect table structure (rows, columns, headers)."""
        prompt = "Analyze this handwritten table. Count the data rows and identify all column headers."
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

    def _extract_json_data(self, image_bytes: bytes, headers: list[str] | None,
                            expected_rows: int | None = None,
                            strip_label: str = "") -> list | None:
        """Pass 2: Extract table data as JSON rows."""
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

    # Standard column set for Indian bank registers (used as fallback)
    _BANK_REGISTER_HEADERS = [
        "S.No", "Name", "Father Name", "Designation",
        "Account Holder", "Account No", "IFSC Code", "Bank Name", "Branch"
    ]

    def ask_with_image_json(self, image_path: str, user_command: str = "") -> str:
        """Extract table data from an image using two-pass JSON architecture.

        Pass 1: Analyze structure (row count, headers)
        Pass 2: Extract data as JSON (with strips for tall images)
        Post-process: Validate and correct IFSC, bank names, etc.
        Output: xlwings Python code ready to execute.
        """
        try:
            raw_bytes = self._preprocess_image(image_path)

            # If user provided a specific command, use old path
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

            # Supplement weak structure analysis for bank registers
            # If the model detected too few columns but found bank-related headers,
            # use the standard bank register column set
            if headers and len(headers) < 8:
                has_bank_indicators = any(
                    keyword in h.lower()
                    for h in headers
                    for keyword in ["bank", "ifsc", "account", "name", "designation"]
                )
                if has_bank_indicators or is_two_page:
                    headers = self._BANK_REGISTER_HEADERS

            # Determine if image needs strip-based extraction
            # Use strips when many rows expected (regardless of aspect ratio)
            use_strips = False
            if expected_rows and expected_rows > 20:
                use_strips = True
            else:
                try:
                    from PIL import Image as PILImage
                    img = PILImage.open(image_path)
                    w, h = img.size
                    if h > w * 1.2 and (expected_rows is None or expected_rows > 10):
                        use_strips = True
                except ImportError:
                    pass

            # For two-page spreads with many expected rows, always use strips
            if is_two_page and (expected_rows is None or expected_rows > 10):
                use_strips = True
                if expected_rows is None:
                    expected_rows = 28  # reasonable default for register pages

            if use_strips and expected_rows:
                # Add a 30% buffer to expected rows so strip prompts don't stop early
                padded_expected = int(expected_rows * 1.3)
                all_rows = self._extract_with_json_strips(
                    image_path, raw_bytes, headers, padded_expected
                )
            else:
                # Single-pass extraction
                all_rows = self._extract_json_data(
                    raw_bytes, headers, expected_rows
                )

            if not all_rows:
                # Fallback to legacy method
                return self.ask_with_image(image_path, "")

            # Remove hallucinated rows (model sometimes invents data)
            all_rows = self._remove_hallucinated_rows(all_rows, headers)

            # Build full data with headers
            if headers:
                full_data = [headers] + all_rows
            else:
                # First row might be header
                full_data = all_rows

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

            # Generate xlwings code (using pprint to ensure valid Python syntax like None instead of null)
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

    def _remove_hallucinated_rows(self, rows: list, headers: list | None) -> list:
        """Detect and remove rows that appear to be hallucinated by the model.

        Common hallucination patterns:
        - Sequential account numbers (100000000001, 100000000002, ...)
        - Repeated father names across many rows (all "LateRam")
        - Names following a pattern (Smt. Shanti, Smt. Meena, Smt. Anita, ...)
        - Identical fake account numbers (100000000000)
        """
        if len(rows) < 5:
            return rows

        clean_rows = []
        sequential_count = 0
        identical_acct_count = 0

        for i, row in enumerate(rows):
            is_suspicious = False

            # Check for sequential or identical account-number patterns
            if i > 0 and len(row) > 4 and len(clean_rows) > 0:
                prev_row = clean_rows[-1] if clean_rows else None
                if prev_row and len(prev_row) > 4:
                    try:
                        curr_acct = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                        prev_acct = str(prev_row[5]).strip() if len(prev_row) > 5 and prev_row[5] else ""
                        
                        if curr_acct and prev_acct:
                            if curr_acct == prev_acct and len(curr_acct) >= 8:
                                identical_acct_count += 1
                                if identical_acct_count >= 2:
                                    is_suspicious = True
                            elif curr_acct.isdigit() and prev_acct.isdigit() and int(curr_acct) == int(prev_acct) + 1:
                                sequential_count += 1
                                if sequential_count >= 2:
                                    is_suspicious = True
                            else:
                                sequential_count = 0
                                identical_acct_count = 0
                        else:
                            sequential_count = 0
                            identical_acct_count = 0
                    except (ValueError, TypeError, IndexError):
                        sequential_count = 0
                        identical_acct_count = 0

            # Check for repeated father-name patterns (hallucination signal)
            if i >= 5 and len(row) > 2:
                recent_fathers = [
                    str(r[2]).strip() if len(r) > 2 and r[2] else ""
                    for r in rows[max(0, i-4):i]
                ]
                curr_father = str(row[2]).strip() if row[2] else ""
                if curr_father and recent_fathers.count(curr_father) >= 3:
                    is_suspicious = True

            if not is_suspicious:
                clean_rows.append(row)

        return clean_rows

    def _extract_with_json_strips(self, image_path: str, raw_bytes: bytes,
                                   headers: list[str] | None,
                                   expected_rows: int) -> list:
        """Extract data from a tall image using row-based strips with JSON output."""
        from PIL import Image as PILImage

        img = PILImage.open(image_path)
        w, h = img.size

        # Calculate strip sizes based on expected row count
        # Aim for ~7 rows per strip with overlap to ensure we hit the bottom rows
        rows_per_strip = 7
        num_strips = max(2, (expected_rows + rows_per_strip - 1) // rows_per_strip)

        # Height per strip with overlap
        strip_height = h // num_strips
        overlap = int(strip_height * 0.3)  # 30% overlap for safety

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
            rows_in_strip = self._extract_json_data(
                strip_bytes, headers,
                expected_rows=rows_per_strip,
                strip_label=label,
            )
            if rows_in_strip:
                for row in rows_in_strip:
                    # Deduplicate against existing rows
                    is_dup = False
                    for existing in all_rows:
                        if self._rows_match(existing, row):
                            is_dup = True
                            break
                    if not is_dup:
                        all_rows.append(row)

        return all_rows

    def ask_with_image(self, image_path: str, user_command: str = "") -> str:
        """Legacy extraction method using Python code output."""
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
                    is_dup = False
                    for existing in all_data:
                        if self._rows_match(existing, trimmed):
                            is_dup = True
                            break
                    if not is_dup:
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

    def _rows_match(self, row_a, row_b, threshold=0.7):
        min_len = min(len(row_a), len(row_b))
        if min_len == 0:
            return False
        matches = 0
        for i in range(min_len):
            a = str(row_a[i]).strip().lower() if row_a[i] is not None else ""
            b = str(row_b[i]).strip().lower() if row_b[i] is not None else ""
            if a == b:
                matches += 1
            elif a and b and (a in b or b in a):
                matches += 0.8
        return matches / min_len >= threshold

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
