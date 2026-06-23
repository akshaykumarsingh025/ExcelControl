# test_json_extract.py — Test the new two-pass JSON extraction pipeline
"""
Tests the improved OCR pipeline on the Fateh Singh Gang test image.
Validates:
  1. Structure analysis (Pass 1) detects correct rows/columns
  2. Data extraction (Pass 2) returns valid JSON with all rows
  3. Post-processing validators correct IFSC, bank names, account numbers
  4. Final output is valid xlwings Python code
"""
import sys
import json
import re
import time

# Force UTF-8 output to avoid Windows CP1252 encoding errors
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from agent import ExcelAgent

IMAGE_PATH = r"D:\Software\Projects\ExcelControl\testFiles\Fateh Singh Gang.jpeg"
EXPECTED_MIN_ROWS = 25  # At least 25 of 28 rows should be extracted
EXPECTED_COLS_MIN = 8   # At least 8 columns


def test_structure_analysis():
    """Test Pass 1: Structure detection."""
    print("=" * 70)
    print("TEST 1: Structure Analysis (Pass 1)")
    print("=" * 70)

    agent = ExcelAgent(model="gemma4:31b-cloud")
    raw_bytes = agent._preprocess_image(IMAGE_PATH)
    print(f"  Preprocessed image: {len(raw_bytes):,} bytes")

    print("  Analyzing structure...")
    t0 = time.time()
    structure = agent._analyze_structure(raw_bytes)
    t1 = time.time()
    print(f"  Time: {t1-t0:.1f}s")

    if structure is None:
        print("  [FAIL] FAILED: Structure analysis returned None")
        return None

    print(f"  Detected rows: {structure.get('num_rows', '?')}")
    print(f"  Detected cols: {structure.get('num_cols', '?')}")
    print(f"  Headers: {structure.get('headers', [])}")
    print(f"  Two-page spread: {structure.get('is_two_page_spread', '?')}")
    print(f"  Notes: {structure.get('notes', '')}")

    num_rows = structure.get("num_rows", 0)
    headers = structure.get("headers", [])

    if num_rows >= EXPECTED_MIN_ROWS:
        print(f"  [OK] Row count OK ({num_rows} >= {EXPECTED_MIN_ROWS})")
    else:
        print(f"  [WARN] Row count low ({num_rows} < {EXPECTED_MIN_ROWS})")

    if len(headers) >= EXPECTED_COLS_MIN:
        print(f"  [OK] Column count OK ({len(headers)} >= {EXPECTED_COLS_MIN})")
    else:
        print(f"  [WARN] Column count low ({len(headers)} < {EXPECTED_COLS_MIN})")

    return structure


def test_full_extraction():
    """Test the full two-pass pipeline with validators."""
    print("\n" + "=" * 70)
    print("TEST 2: Full JSON Extraction Pipeline")
    print("=" * 70)

    agent = ExcelAgent(model="gemma4:31b-cloud")

    print("  Running ask_with_image_json...")
    t0 = time.time()
    result = agent.ask_with_image_json(IMAGE_PATH)
    t1 = time.time()
    print(f"  Time: {t1-t0:.1f}s")

    if result.startswith("# ERROR"):
        print(f"  [FAIL] FAILED: {result}")
        return None

    # Parse the data from the result
    data_match = re.search(r'data\s*=\s*(\[.*?\])\s*\n\nws\.', result, re.DOTALL)
    if not data_match:
        print("  [FAIL] FAILED: Could not parse data = [...] from result")
        print("  First 500 chars:", result[:500])
        return result

    import ast
    try:
        data = ast.literal_eval(data_match.group(1))
    except (SyntaxError, ValueError):
        print("  [FAIL] FAILED: Data is not valid Python list")
        return result

    headers = data[0] if data else []
    rows = data[1:] if data else []

    print(f"\n  Headers ({len(headers)}): {headers}")
    print(f"  Data rows: {len(rows)}")

    # Validate row count
    if len(rows) >= EXPECTED_MIN_ROWS:
        print(f"  [OK] Row count OK ({len(rows)} >= {EXPECTED_MIN_ROWS})")
    else:
        print(f"  [WARN] Row count low ({len(rows)} < {EXPECTED_MIN_ROWS})")

    # Validate IFSC codes
    ifsc_col = None
    for i, h in enumerate(headers):
        if "ifsc" in str(h).lower() or "code" in str(h).lower():
            ifsc_col = i
            break

    if ifsc_col is not None:
        valid_ifsc = 0
        for row in rows:
            if ifsc_col < len(row) and row[ifsc_col]:
                code = str(row[ifsc_col]).strip()
                if re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", code):
                    valid_ifsc += 1
        print(f"  IFSC validation: {valid_ifsc}/{len(rows)} valid ({valid_ifsc/max(len(rows),1)*100:.0f}%)")

    # Validate bank names
    bank_col = None
    for i, h in enumerate(headers):
        if "bank" in str(h).lower() and "name" in str(h).lower():
            bank_col = i
            break
    if bank_col is None:
        for i, h in enumerate(headers):
            if "bank" in str(h).lower():
                bank_col = i
                break

    if bank_col is not None:
        known_banks = {"INDIAN BANK", "BANK OF INDIA", "UNION BANK", "UNION BANK OF INDIA",
                       "BANK OF BARODA", "PUNJAB NATIONAL BANK", "FINO PAYMENTS BANK",
                       "AIRTEL PAYMENTS BANK", "INDIA POST"}
        matched = 0
        for row in rows:
            if bank_col < len(row) and row[bank_col]:
                if str(row[bank_col]).strip().upper() in known_banks:
                    matched += 1
        print(f"  Bank name match: {matched}/{len(rows)} known ({matched/max(len(rows),1)*100:.0f}%)")

    # Print validator warnings from comments
    for line in result.split("\n"):
        if line.startswith("# -"):
            print(f"  [FIX] {line[4:]}")

    # Print first 5 rows as sample
    print(f"\n  Sample rows:")
    for i, row in enumerate(rows[:5]):
        print(f"    Row {i+1}: {row}")
    if len(rows) > 5:
        print(f"    ... ({len(rows) - 5} more rows)")

    return result


if __name__ == "__main__":
    print("Testing improved OCR pipeline on: Fateh Singh Gang.jpeg")
    print(f"Using model: gemma4:31b-cloud")
    print()

    structure = test_structure_analysis()
    result = test_full_extraction()

    if result:
        # Save the full output for inspection
        out_path = r"D:\Software\Projects\ExcelControl\testFiles\_last_extraction.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"\n  Full output saved to: {out_path}")

    print("\n" + "=" * 70)
    print("DONE")
    print("=" * 70)
