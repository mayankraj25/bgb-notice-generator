"""
Read and parse the BGB Excel input sheet.

Handles real-world quirks:
- Variable number of leading columns (Ser No, File No, etc.)
- Multiple header rows (merged/split headers like "Due date of Notice" + "1st / 2nd")
- Contractor/address in a single cell with \\n separators
- BGB field with "BGB No" prefix and "for Rs" / "of Rs" variants
- GE field with \\n separators and dash-prefixed sub-fields
- Dates as datetime objects or strings
"""

import re
from datetime import date, datetime
from typing import Optional

import openpyxl

from utils import parse_bgb_field, parse_ca_field, parse_bank_field, parse_ge_bank_field


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _str(cell) -> str:
    v = cell.value
    if v is None:
        return ""
    return str(v).strip()


def _parse_date(raw) -> Optional[date]:
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    raw = str(raw).strip()
    if not raw:
        return None
    import dateutil.parser as dp
    try:
        # Try dayfirst (common in India: 31-10-2026)
        return dp.parse(raw, dayfirst=True).date()
    except Exception:
        pass
    # Manual fallbacks
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d-%b-%Y",
                "%Y-%m-%d", "%d.%m.%y", "%d.%m.%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


# ---------------------------------------------------------------------------
# Column auto-detection
# ---------------------------------------------------------------------------

# Keywords that identify each logical column
_COL_PATTERNS = {
    'ca':          re.compile(r'ca\s*no|name\s*of\s*work',            re.I),
    'contractor':  re.compile(r'name\s*of\s*contractor',              re.I),
    'bgb':         re.compile(r'bgb\s*no',                            re.I),
    'account':     re.compile(r'account\s*details\s*of\s*contractor', re.I),
    'bank':        re.compile(r'name\s*of\s*bank',                    re.I),
    'validity':    re.compile(r'valid\s*upto',                        re.I),
    'due_notice':  re.compile(r'due\s*date\s*of\s*notice',            re.I),
    'status':      re.compile(r'present\s*status',                    re.I),
    'ge':          re.compile(r'ge\s*&|ge\s*and',                     re.I),
}


def _detect_columns(ws) -> dict:
    """
    Scan up to the first 5 rows to auto-detect which Excel column index
    holds each logical field.  Returns a dict: logical_name -> col_index (0-based).

    For "due date of notice" the next row may have "1st" / "2nd" sub-headers;
    we return due1_col and due2_col.
    """
    col_map = {}
    due_notice_col = None   # column that has the "Due date of Notice" header

    for row_idx, row in enumerate(ws.iter_rows(max_row=5)):
        for ci, cell in enumerate(row):
            v = _str(cell)
            if not v:
                continue
            for name, pat in _COL_PATTERNS.items():
                if name not in col_map and pat.search(v):
                    col_map[name] = ci
                    if name == 'due_notice':
                        due_notice_col = ci

            # Detect "1st" / "2nd" sub-headers for due dates
            if re.search(r'\b1st\b', v, re.I) and 'due1' not in col_map:
                col_map['due1'] = ci
            if re.search(r'\b2nd\b', v, re.I) and 'due2' not in col_map:
                col_map['due2'] = ci

    # If "1st"/"2nd" sub-headers weren't found, fall back to due_notice_col + 1
    if 'due1' not in col_map and due_notice_col is not None:
        col_map['due1'] = due_notice_col
    if 'due2' not in col_map and due_notice_col is not None:
        col_map['due2'] = due_notice_col + 1

    return col_map


# ---------------------------------------------------------------------------
# Is this row a header / sub-header row?
# ---------------------------------------------------------------------------

def _is_header_row(row_values: list, col_map: dict) -> bool:
    """
    A row is a header if it contains known column-header keywords,
    OR if all its logically-important cells are empty/header text.
    """
    # NOTE: do NOT include 'bgb no' here — it appears in data cell values too.
    header_keywords = re.compile(
        r'\b(ca\s*no|valid\s*upto|due\s*date\s*of\s*notice|serial|ser\s*no|'
        r'file\s*no|name\s*of\s*work|name\s*of\s*contractor|'
        r'name\s*of\s*bank|present\s*status|account\s*details\s*of)\b',
        re.I
    )
    for v in row_values:
        if v and header_keywords.search(v):
            return True
    return False


# ---------------------------------------------------------------------------
# CA field: supports both "CEJZ/GNR/18 of 2021-22, Work name"
#           and  "CEJZ/GNR/18 of 2021-22 : Work name"
# ---------------------------------------------------------------------------

def _parse_ca(raw: str) -> dict:
    if not raw:
        return {'ca_number': '', 'work_name': ''}
    # Try ": " separator first, then ","
    for sep in (' : ', ': ', ','):
        idx = raw.find(sep)
        if idx != -1:
            return {
                'ca_number': raw[:idx].strip(),
                'work_name':  raw[idx + len(sep):].strip(),
            }
    return {'ca_number': raw.strip(), 'work_name': ''}


# ---------------------------------------------------------------------------
# Contractor: cell may have \n between name and address
# ---------------------------------------------------------------------------

def _parse_contractor(raw: str) -> dict:
    if not raw:
        return {'contractor_name': '', 'contractor_address': ''}
    # Split on first \n or first ","
    if '\n' in raw:
        lines = [l.strip().rstrip(',') for l in raw.split('\n') if l.strip()]
        return {
            'contractor_name':    lines[0],
            'contractor_address': ', '.join(lines[1:]),
        }
    parts = raw.split(',', 1)
    return {
        'contractor_name':    parts[0].strip().rstrip(','),
        'contractor_address': parts[1].strip() if len(parts) > 1 else '',
    }


# ---------------------------------------------------------------------------
# BGB field: handles
#   "142GT02213330004 dt 29 Nov 2021 of Rs 825000/- against PSD"
#   "BGB No 154271121000010 dt 15 Dec 2021 for Rs 3,94,075/- against PSD"
#   "0064NDDG00015423 dt 28 Sep 2022 for Rs. 810000/-PSD"
# ---------------------------------------------------------------------------

def _parse_bgb(raw: str) -> dict:
    if not raw:
        return {'bgb_number': '', 'bgb_date_str': '',
                'bgb_amount_int': 0, 'bgb_purpose': ''}

    # Strip leading "BGB No " prefix if present
    cleaned = re.sub(r'^BGB\s*No\s*', '', raw.strip(), flags=re.I)

    # BGB Number: text before 'dt'
    m = re.match(r'^\s*(\S+)\s+dt\b', cleaned, re.IGNORECASE)
    bgb_number = m.group(1).strip() if m else cleaned.split()[0]

    # BGB Date: between 'dt' and next keyword
    m = re.search(r'\bdt\s+(.+?)\s+(?:of\b|for\b)', cleaned, re.IGNORECASE)
    bgb_date_str = m.group(1).strip() if m else ''

    # Amount: after "Rs" / "Rs." / "for Rs" and before "/-"
    m = re.search(r'\bRs\.?\s*([\d,]+)\s*/-', cleaned, re.IGNORECASE)
    if m:
        try:
            bgb_amount_int = int(m.group(1).replace(',', ''))
        except ValueError:
            bgb_amount_int = 0
    else:
        bgb_amount_int = 0

    # Purpose
    m = re.search(r'against\s+(PSD|Retention\s+Money)', cleaned, re.IGNORECASE)
    bgb_purpose = m.group(0).strip() if m else 'against Performance Security Deposit'

    return {
        'bgb_number':     bgb_number,
        'bgb_date_str':   bgb_date_str,
        'bgb_amount_int': bgb_amount_int,
        'bgb_purpose':    bgb_purpose,
    }


# ---------------------------------------------------------------------------
# GE / bank account field
# Handles both comma-separated and newline/dash formats:
#   "GE (A) Gandhinagar\n\nName of A/C- Garison Engg (proj) Div\nBank - SBI..."
#   "GE (A) Gandhinagar, Name of A/C: Garrison Engg, Bank: SBI, ..."
# ---------------------------------------------------------------------------

def _parse_ge(raw: str) -> dict:
    if not raw:
        return {}

    # Normalise newlines and dash-style labels to comma+colon style
    text = raw.replace('\n\n', '\n').replace('\n', ', ')
    # Convert "Label- value" to "Label: value"
    text = re.sub(r'\b(Name\s+of\s+A/C|Bank|Branch|A/C\s+No|IFSC)\s*[-–]\s*',
                  lambda m: m.group(1).replace(' ', ' ') + ': ', text)

    result = {}

    m = re.search(r'GE\s*\(A\)\s+([^,\n]+)', text, re.I)
    result['ge_location'] = m.group(1).strip() if m else ''

    m = re.search(r'Name\s+of\s+A/C\s*[:\-]\s*([^,\n]+)', text, re.I)
    result['account_name'] = m.group(1).strip() if m else ''

    # Bank: avoid matching "Bank Account", grab just the bank name
    m = re.search(r'\bBank\s*[:\-]\s*([^,\n]+)', text, re.I)
    result['bank_name'] = m.group(1).strip() if m else ''

    m = re.search(r'Branch\s*[:\-]\s*([^,\n]+)', text, re.I)
    result['branch_name'] = m.group(1).strip() if m else ''

    m = re.search(r'A/C\s+No\s*[:\-]\s*([\d]+)', text, re.I)
    result['account_number'] = m.group(1).strip() if m else ''

    m = re.search(r'IFSC\s*[:\-]?\s*(\S+)', text, re.I)
    result['ifsc_code'] = m.group(1).strip() if m else ''

    return result


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_excel(filepath: str) -> list:
    """
    Parse the BGB Excel file and return a list of row dicts.
    Auto-detects column positions from header rows.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Step 1: auto-detect column positions
    col_map = _detect_columns(ws)

    # Helper: safely get cell value by logical name
    def get(row_cells, name, default=""):
        idx = col_map.get(name)
        if idx is None or idx >= len(row_cells):
            return default
        return _str(row_cells[idx])

    def get_raw(row_cells, name):
        """Return raw cell value (not string-converted) for date parsing."""
        idx = col_map.get(name)
        if idx is None or idx >= len(row_cells):
            return None
        return row_cells[idx].value

    rows_data = []
    skipped_header = 0

    for row_idx, row in enumerate(ws.iter_rows(), 1):
        row_str = [_str(c) for c in row]

        # Skip completely blank rows
        if all(v == "" for v in row_str):
            continue

        # Skip header/sub-header rows
        if _is_header_row(row_str, col_map):
            skipped_header += 1
            continue

        # Skip rows where the CA column (most important) is empty
        ca_raw = get(row, 'ca')
        if not ca_raw:
            continue

        # Parse all fields
        ca_parsed          = _parse_ca(ca_raw)
        contractor_parsed  = _parse_contractor(get(row, 'contractor'))
        bgb_parsed         = _parse_bgb(get(row, 'bgb'))
        bank_parsed        = parse_bank_field(get(row, 'bank'))
        ge_parsed          = _parse_ge(get(row, 'ge'))

        validity_date = _parse_date(get_raw(row, 'validity'))
        due_date_1    = _parse_date(get_raw(row, 'due1'))
        due_date_2    = _parse_date(get_raw(row, 'due2'))

        rows_data.append({
            'row_index':          row_idx,
            # CA
            'ca_number':          ca_parsed['ca_number'],
            'work_name':          ca_parsed['work_name'],
            # Contractor
            'contractor_name':    contractor_parsed['contractor_name'],
            'contractor_address': contractor_parsed['contractor_address'],
            # BGB
            'bgb_number':         bgb_parsed['bgb_number'],
            'bgb_date_str':       bgb_parsed['bgb_date_str'],
            'bgb_amount_int':     bgb_parsed['bgb_amount_int'],
            'bgb_purpose':        bgb_parsed['bgb_purpose'],
            # Bank
            'bank_name':          bank_parsed['bank_name'],
            'bank_address':       bank_parsed['bank_address'],
            # Dates
            'validity_date':      validity_date,
            'due_date_1':         due_date_1,
            'due_date_2':         due_date_2,
            # Status
            'status':             get(row, 'status'),
            # GE / bank account
            'ge_location':        ge_parsed.get('ge_location', ''),
            'account_name':       ge_parsed.get('account_name', ''),
            'ge_bank_name':       ge_parsed.get('bank_name', ''),
            'branch_name':        ge_parsed.get('branch_name', ''),
            'account_number':     ge_parsed.get('account_number', ''),
            'ifsc_code':          ge_parsed.get('ifsc_code', ''),
        })

    return rows_data
