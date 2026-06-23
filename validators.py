# validators.py — Post-OCR validation and correction layer
import re
import difflib


# ── Known bank names in India ─────────────────────────────────────────────
KNOWN_BANKS = [
    "INDIAN BANK",
    "BANK OF INDIA",
    "UNION BANK OF INDIA",
    "UNION BANK",
    "BANK OF BARODA",
    "PUNJAB NATIONAL BANK",
    "FINO PAYMENTS BANK",
    "AIRTEL PAYMENTS BANK",
    "STATE BANK OF INDIA",
    "CANARA BANK",
    "CENTRAL BANK OF INDIA",
    "INDIAN OVERSEAS BANK",
    "UCO BANK",
    "BANK OF MAHARASHTRA",
    "PUNJAB & SIND BANK",
    "ALLAHABAD BANK",
    "ANDHRA BANK",
    "CORPORATION BANK",
    "DENA BANK",
    "ORIENTAL BANK OF COMMERCE",
    "SYNDICATE BANK",
    "UNITED BANK OF INDIA",
    "VIJAYA BANK",
    "IDBI BANK",
    "AXIS BANK",
    "HDFC BANK",
    "ICICI BANK",
    "KOTAK MAHINDRA BANK",
    "YES BANK",
    "INDIA POST PAYMENTS BANK",
    "INDIA POST",
    "POST OFFICE",
    "GRAMIN BANK",
    "ARYAVART BANK",
]

# Common OCR substitutions for letter↔digit confusion
_LETTER_TO_DIGIT = str.maketrans("OoIlSZBG", "00115289")
_DIGIT_TO_LETTER = str.maketrans("015289", "OOlSZB")


def validate_ifsc(code) -> tuple[bool, str | None]:
    """Validate IFSC code format: 4 letters + '0' + 6 alphanumeric chars.

    Returns (is_valid, corrected_code_or_None).
    Auto-corrects common OCR letter↔digit errors.
    """
    if code is None or not isinstance(code, str):
        return False, None

    code = code.strip().upper().replace(" ", "").replace(".", "")

    if len(code) < 11:
        return False, None

    # Take first 11 characters (sometimes trailing junk)
    code = code[:11]

    if len(code) != 11:
        return False, None

    # Part 1: first 4 must be letters
    prefix = code[:4]
    corrected_prefix = ""
    for ch in prefix:
        if ch.isalpha():
            corrected_prefix += ch
        elif ch.isdigit():
            # Try digit→letter substitution
            mapped = ch.translate(_DIGIT_TO_LETTER)
            corrected_prefix += mapped
        else:
            return False, None

    # Part 2: 5th character must be '0'
    fifth = code[4]
    if fifth == "O":
        fifth = "0"
    if fifth != "0":
        return False, None

    # Part 3: last 6 must be alphanumeric
    suffix = code[5:]
    if not suffix.isalnum():
        # Try cleaning
        cleaned = re.sub(r"[^A-Z0-9]", "", suffix)
        if len(cleaned) == 6:
            suffix = cleaned
        else:
            return False, None

    corrected = corrected_prefix + "0" + suffix
    is_valid = (
        len(corrected) == 11
        and corrected[:4].isalpha()
        and corrected[4] == "0"
        and corrected[5:].isalnum()
    )
    return is_valid, corrected if is_valid else None


def validate_account_number(num) -> tuple[bool, str]:
    """Clean and validate account numbers.

    Strips leading dots/spaces, ensures digits only.
    Returns (is_valid, cleaned_number).
    """
    if num is None:
        return False, ""

    raw = str(num).strip()

    # Strip leading dots (common OCR artifact)
    raw = raw.lstrip(".")

    # Remove spaces
    raw = raw.replace(" ", "")

    # Check if it's all digits
    if raw.isdigit() and len(raw) >= 4:
        return True, raw

    # Try extracting just digits
    digits_only = re.sub(r"[^0-9]", "", raw)
    if len(digits_only) >= 4:
        return True, digits_only

    return False, raw


def match_bank_name(raw: str) -> tuple[str, float]:
    """Fuzzy-match a raw OCR bank name to known bank names.

    Returns (matched_name, confidence_score 0.0-1.0).
    """
    if not raw or not isinstance(raw, str):
        return raw or "", 0.0

    cleaned = raw.strip().upper()
    # Remove doubled words (e.g., "INDIAN BANK BANK" → "INDIAN BANK")
    words = cleaned.split()
    deduped = []
    for w in words:
        if not deduped or w != deduped[-1]:
            deduped.append(w)
    cleaned = " ".join(deduped)

    # Exact match
    if cleaned in KNOWN_BANKS:
        return cleaned, 1.0

    # Fuzzy match
    matches = difflib.get_close_matches(cleaned, KNOWN_BANKS, n=1, cutoff=0.5)
    if matches:
        ratio = difflib.SequenceMatcher(None, cleaned, matches[0]).ratio()
        return matches[0], ratio

    # Partial containment check
    for bank in KNOWN_BANKS:
        if bank in cleaned or cleaned in bank:
            return bank, 0.8

    return cleaned, 0.3  # low confidence, return cleaned version


def validate_serial_numbers(data: list) -> tuple[list, list[str]]:
    """Ensure serial numbers in column 0 are sequential 1..N.

    Returns (corrected_data, list_of_warnings).
    """
    if not data or len(data) < 2:
        return data, []

    warnings = []
    corrected = [data[0]]  # keep header

    for i, row in enumerate(data[1:], start=1):
        new_row = list(row)
        expected_sno = i

        actual = row[0] if row else None
        try:
            actual_int = int(actual) if actual is not None else None
        except (ValueError, TypeError):
            actual_int = None

        if actual_int != expected_sno:
            warnings.append(
                f"Row {i}: S.No was {actual!r}, corrected to {expected_sno}"
            )
            new_row[0] = expected_sno

        corrected.append(new_row)

    return corrected, warnings


def validate_row_completeness(row: list, expected_cols: int) -> float:
    """Return confidence score 0.0-1.0 for how complete a row is.

    Counts non-None cells vs expected column count.
    """
    if not row:
        return 0.0

    filled = sum(1 for cell in row if cell is not None and str(cell).strip())
    return filled / max(expected_cols, 1)


def validate_and_correct_table(data: list, ifsc_col: int = -3,
                                acct_col: int = -4,
                                bank_col: int = -2) -> tuple[list, list[str]]:
    """Run all validators on extracted table data.

    Args:
        data: 2D list where data[0] is the header row.
        ifsc_col: column index for IFSC codes (negative = from end).
        acct_col: column index for account numbers.
        bank_col: column index for bank names.

    Returns:
        (corrected_data, list_of_all_warnings)
    """
    if not data or len(data) < 2:
        return data, ["No data to validate"]

    all_warnings = []
    num_cols = len(data[0])

    # Resolve negative indices
    if ifsc_col < 0:
        ifsc_col = num_cols + ifsc_col
    if acct_col < 0:
        acct_col = num_cols + acct_col
    if bank_col < 0:
        bank_col = num_cols + bank_col

    # Validate serial numbers
    data, sno_warnings = validate_serial_numbers(data)
    all_warnings.extend(sno_warnings)

    # Validate each data row
    for i, row in enumerate(data[1:], start=1):
        # Pad row to expected columns
        while len(row) < num_cols:
            row.append(None)

        # IFSC code validation
        if 0 <= ifsc_col < len(row) and row[ifsc_col]:
            is_valid, corrected = validate_ifsc(row[ifsc_col])
            if corrected and corrected != str(row[ifsc_col]).strip().upper():
                all_warnings.append(
                    f"Row {i}: IFSC '{row[ifsc_col]}' -> '{corrected}'"
                )
                row[ifsc_col] = corrected
            elif not is_valid:
                all_warnings.append(
                    f"Row {i}: IFSC '{row[ifsc_col]}' is invalid (could not auto-correct)"
                )

        # Account number validation
        if 0 <= acct_col < len(row) and row[acct_col]:
            is_valid, cleaned = validate_account_number(row[acct_col])
            if cleaned != str(row[acct_col]).strip():
                all_warnings.append(
                    f"Row {i}: Account '{row[acct_col]}' -> '{cleaned}'"
                )
                row[acct_col] = cleaned

        # Bank name validation
        if 0 <= bank_col < len(row) and row[bank_col]:
            matched, confidence = match_bank_name(str(row[bank_col]))
            if matched != str(row[bank_col]).strip().upper():
                all_warnings.append(
                    f"Row {i}: Bank '{row[bank_col]}' -> '{matched}' "
                    f"(confidence: {confidence:.0%})"
                )
                row[bank_col] = matched

        # Row completeness
        completeness = validate_row_completeness(row, num_cols)
        if completeness < 0.5:
            all_warnings.append(
                f"Row {i}: Low completeness ({completeness:.0%}) — may need review"
            )

    return data, all_warnings
