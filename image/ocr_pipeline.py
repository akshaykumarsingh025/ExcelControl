import ast
import io
import json
import re

from PIL import Image as PILImage

from image.preprocessor import ImagePreprocessor
from prompts import (
    STRUCTURE_ANALYSIS_PROMPT,
    JSON_EXTRACTION_SYSTEM_PROMPT,
    VISION_SYSTEM_PROMPT,
    make_json_extraction_prompt,
    make_right_half_extraction_prompt,
)
from validators import validate_and_correct_table


class OcrPipeline:

    def __init__(self, agent, preprocessor: ImagePreprocessor = None):
        self.agent = agent
        self.preprocessor = preprocessor or ImagePreprocessor()

    @staticmethod
    def is_null(val) -> bool:
        if val is None:
            return True
        s = str(val).strip().lower()
        return s in ("null", "none", "")

    @staticmethod
    def try_int(val) -> int | None:
        try:
            return int(str(val).strip())
        except (ValueError, TypeError):
            return None

    def rows_are_same_record(self, row_a, row_b):
        if len(row_a) > 0 and len(row_b) > 0:
            a_sno = self.try_int(row_a[0])
            b_sno = self.try_int(row_b[0])
            if a_sno is not None and b_sno is not None:
                if a_sno == b_sno:
                    return True

        min_len = min(len(row_a), len(row_b))
        if min_len >= 2:
            a_name = str(row_a[1]).strip().lower() if row_a[1] is not None and not self.is_null(row_a[1]) else ""
            b_name = str(row_b[1]).strip().lower() if row_b[1] is not None and not self.is_null(row_b[1]) else ""
            if a_name and b_name and len(a_name) >= 3 and a_name == b_name:
                return True

        return False

    def merge_rows(self, row_a, row_b):
        max_len = max(len(row_a), len(row_b))
        merged = []
        for i in range(max_len):
            a = row_a[i] if i < len(row_a) else None
            b = row_b[i] if i < len(row_b) else None

            a_null = self.is_null(a)
            b_null = self.is_null(b)

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

    def analyze_structure(self, image_bytes: bytes) -> dict | None:
        prompt = "Analyze this table image. Count the data rows and identify all column headers."
        raw = self.agent.call_vision_api(
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

    def extract_json_data(self, image_bytes: bytes, headers: list[str] | None,
                          expected_rows: int | None = None,
                          strip_label: str = "") -> list | None:
        prompt = make_json_extraction_prompt(
            headers=headers,
            expected_rows=expected_rows,
            strip_label=strip_label,
        )
        raw = self.agent.call_vision_api(
            prompt, image_bytes,
            system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
            json_mode=True,
        )
        return self.parse_json_rows(raw)

    def parse_json_rows(self, text: str) -> list | None:
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

    def extract_data_list(self, text: str):
        idx = text.find("data = [")
        if idx < 0:
            return None
        depth = 1
        i = idx + len("data = [")
        end = None
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
        if end is None:
            return None
        try:
            return ast.literal_eval(text[idx: idx + len("data = ") + (end - idx)])
        except Exception:
            return None

    def extract_with_strips(self, img, headers, expected_rows, extract_fn=None):
        if extract_fn is None:
            extract_fn = self.extract_json_data

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
            strip_bytes = self.preprocessor.preprocess_strip(strip_img)
            rows_in_strip = extract_fn(strip_bytes, headers, rows_per_strip, label)
            if rows_in_strip:
                for row in rows_in_strip:
                    row = [None if self.is_null(c) else c for c in row]
                    merged = False
                    for i, existing in enumerate(all_rows):
                        if self.rows_are_same_record(existing, row):
                            all_rows[i] = self.merge_rows(existing, row)
                            merged = True
                            break
                    if not merged:
                        all_rows.append(row)

        return all_rows

    def split_headers(self, headers: list[str]) -> tuple[list[str], list[str], int]:
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

    def extract_left_page(self, image_path: str, headers: list[str],
                          expected_rows: int) -> list:
        img = PILImage.open(image_path)
        w, h = img.size

        margin = int(w * 0.04)
        right_boundary = (w // 2) + margin
        left_img = img.crop((0, 0, right_boundary, h))

        left_headers, _, _ = self.split_headers(headers)

        if expected_rows > 10:
            return self.extract_with_strips(left_img, left_headers, expected_rows)
        else:
            left_bytes = self.preprocessor.preprocess_strip(left_img)
            rows = self.extract_json_data(left_bytes, left_headers, expected_rows, "left page")
            return rows if rows else []

    def extract_right_page(self, image_path: str, headers: list[str],
                           expected_rows: int) -> list:
        img = PILImage.open(image_path)
        w, h = img.size

        margin = int(w * 0.04)
        left_boundary = (w // 2) - margin
        right_img = img.crop((left_boundary, 0, w, h))

        _, right_headers, _ = self.split_headers(headers)

        def extract_right_strip(image_bytes, hdrs, exp_rows, strip_label):
            prompt = make_right_half_extraction_prompt(hdrs, exp_rows, strip_label)
            raw = self.agent.call_vision_api(
                prompt, image_bytes,
                system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
                json_mode=True,
            )
            return self.parse_json_rows(raw)

        if expected_rows > 10:
            return self.extract_with_strips(right_img, right_headers, expected_rows, extract_right_strip)
        else:
            right_bytes = self.preprocessor.preprocess_strip(right_img)
            prompt = make_right_half_extraction_prompt(right_headers, expected_rows, "")
            raw = self.agent.call_vision_api(
                prompt, right_bytes,
                system_prompt=JSON_EXTRACTION_SYSTEM_PROMPT,
                json_mode=True,
            )
            rows = self.parse_json_rows(raw)
            return rows if rows else []

    def merge_left_right(self, left_rows: list, right_rows: list,
                         headers: list[str]) -> list:
        _, _, split_idx = self.split_headers(headers)
        right_data_start = split_idx

        right_by_sno = {}
        right_unmatched = []
        for row in right_rows:
            row = [None if self.is_null(c) else c for c in row]
            sno = self.try_int(row[0]) if len(row) > 0 else None
            if sno is not None:
                right_by_sno[sno] = row
            else:
                right_unmatched.append(row)

        merged = []
        used_sno = set()

        for left_row in left_rows:
            left_row = [None if self.is_null(c) else c for c in left_row]
            full_row = list(left_row)
            while len(full_row) < len(headers):
                full_row.append(None)

            matched_right = None
            sno = self.try_int(left_row[0]) if len(left_row) > 0 else None

            if sno is not None and sno in right_by_sno:
                matched_right = right_by_sno[sno]
                used_sno.add(sno)

            if matched_right is not None:
                right_vals = matched_right[1:]
                for ri, val in enumerate(right_vals):
                    target_col = right_data_start + ri
                    if target_col < len(full_row):
                        if self.is_null(full_row[target_col]) and not self.is_null(val):
                            full_row[target_col] = val

            merged.append(full_row)

        remaining_right = [row for sno, row in right_by_sno.items() if sno not in used_sno]
        remaining_right.extend(right_unmatched)

        if remaining_right:
            main_missing = []
            for i, row in enumerate(merged):
                has_right = any(
                    not self.is_null(row[j]) for j in range(right_data_start, len(row))
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
                        if self.is_null(merged[main_idx][target_col]) and not self.is_null(val):
                            merged[main_idx][target_col] = val
                ri += 1

        return merged

    def merge_multi_pass(self, pass_a_rows: list, pass_b_rows: list,
                         headers: list[str]) -> list:
        idx_b_by_sno = {}
        idx_b_by_name = {}
        for idx, row in enumerate(pass_b_rows):
            if len(row) > 0 and row[0] is not None:
                sno = self.try_int(row[0])
                if sno is not None:
                    idx_b_by_sno[sno] = idx
            if len(row) > 1 and row[1] is not None:
                name_key = str(row[1]).strip().lower()
                if name_key and len(name_key) >= 3:
                    idx_b_by_name.setdefault(name_key, idx)

        used_b = set()
        merged = []

        for row_a in pass_a_rows:
            row_a = [None if self.is_null(c) else c for c in row_a]
            match_b_idx = None

            if len(row_a) > 0 and row_a[0] is not None:
                sno = self.try_int(row_a[0])
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
                row_b = [None if self.is_null(c) else c for c in row_b]
                merged_row = self.merge_rows(row_a, row_b)
                used_b.add(match_b_idx)
            else:
                merged_row = list(row_a)

            while len(merged_row) < len(headers):
                merged_row.append(None)
            merged.append(merged_row)

        for idx, row_b in enumerate(pass_b_rows):
            if idx in used_b:
                continue
            row_b = [None if self.is_null(c) else c for c in row_b]
            while len(row_b) < len(headers):
                row_b.append(None)
            merged.append(row_b)

        return merged

    def ask_with_image_json(self, image_path: str, user_command: str = "") -> str:
        try:
            raw_bytes = self.preprocessor.preprocess_image(image_path)

            if user_command:
                return self._ask_with_image_legacy(image_path, user_command)

            structure = self.analyze_structure(raw_bytes)
            if structure:
                headers = structure.get("headers", [])
                expected_rows = structure.get("num_rows", None)
                is_two_page = structure.get("is_two_page_spread", False)
            else:
                headers = None
                expected_rows = None
                is_two_page = False

            if not is_two_page:
                try:
                    img = PILImage.open(image_path)
                    w, h = img.size
                    if w > h * 1.2:
                        is_two_page = True
                except Exception:
                    pass

            if not headers:
                return self._ask_with_image_legacy(image_path, "")

            padded_expected = int((expected_rows or 25) * 1.3)

            full_img_rows = None
            try:
                full_img = PILImage.open(io.BytesIO(raw_bytes))
                if padded_expected > 15:
                    full_img_rows = self.extract_with_strips(full_img, headers, padded_expected)
                else:
                    full_img_rows = self.extract_json_data(raw_bytes, headers, padded_expected)
                    if full_img_rows:
                        full_img_rows = [[None if self.is_null(c) else c for c in row] for row in full_img_rows]
            except Exception:
                pass

            split_rows = None
            if is_two_page:
                try:
                    left_rows = self.extract_left_page(image_path, headers, padded_expected)
                    right_rows = self.extract_right_page(image_path, headers, padded_expected)
                    if left_rows and right_rows:
                        split_rows = self.merge_left_right(left_rows, right_rows, headers)
                    elif left_rows:
                        split_rows = left_rows
                    elif right_rows:
                        split_rows = right_rows
                except Exception:
                    pass

            if full_img_rows and split_rows:
                all_rows = self.merge_multi_pass(full_img_rows, split_rows, headers)
            elif full_img_rows:
                all_rows = full_img_rows
            elif split_rows:
                all_rows = split_rows
            else:
                return self._ask_with_image_legacy(image_path, "")

            all_rows = [
                [None if self.is_null(c) else c for c in row]
                for row in all_rows
            ]

            num_cols = len(headers)
            all_rows = [
                row[:num_cols] + [None] * max(0, num_cols - len(row))
                for row in all_rows
            ]

            full_data = [headers] + all_rows

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

    def _ask_with_image_legacy(self, image_path: str, user_command: str = "") -> str:
        try:
            raw_bytes = self.preprocessor.preprocess_image(image_path)

            try:
                img = PILImage.open(image_path)
                w, h = img.size
                is_tall = h > w * 1.2
            except Exception:
                is_tall = False

            if is_tall:
                return self._ask_with_strips(image_path, raw_bytes, user_command)

            if not user_command:
                user_command = self._extraction_prompt("")

            return self.agent.call_vision_api(user_command, raw_bytes)

        except Exception as e:
            return f"# ERROR: Image processing failed\n# {str(e)}"

    def _ask_with_strips(self, image_path: str, raw_bytes: bytes,
                         user_command: str) -> str:
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
            strip_bytes = self.preprocessor.preprocess_strip(strip_img)
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
            result = self.agent.call_vision_api(prompt, strip_bytes)
            results.append((label, result))

            if i == 0:
                parsed = self.extract_data_list(result)
                if parsed and len(parsed) > 0:
                    prev_headers = parsed[0]

        combined = self._merge_strip_results(results)
        return combined

    def _merge_strip_results(self, results):
        all_data = []
        header = None
        header_cols = 0

        for label, text in results:
            parsed = self.extract_data_list(text)
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
                        if self.rows_are_same_record(existing, trimmed):
                            all_data[i] = self.merge_rows(existing, trimmed)
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
