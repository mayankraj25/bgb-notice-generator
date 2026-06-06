"""Date helpers, filename sanitizer, and field parsers."""

import re
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def validity_in_window(validity_date: date, today: date = None) -> bool:
    """Return True if validity is within 60 days from today but not yet expired."""
    if today is None:
        today = date.today()
    delta = (validity_date - today).days
    return 0 <= delta <= 60


def extended_validity(validity_date: date) -> date:
    """Return validity_date + exactly 1 year."""
    return validity_date + relativedelta(years=1)


def submission_deadline(validity_date: date) -> date:
    """Return validity_date - 15 days."""
    return validity_date - timedelta(days=15)


def format_date(d: date) -> str:
    """Format date as '29 Nov 2021'."""
    return d.strftime("%-d %b %Y")


def format_date_short(d: date) -> str:
    """Format date as '31-10-2026' (from Excel column F style)."""
    return d.strftime("%d-%m-%Y")


# ---------------------------------------------------------------------------
# Filename sanitizer
# ---------------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """Replace characters illegal in filenames with underscores; collapse runs."""
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = re.sub(r'\s+', '_', name.strip())
    name = re.sub(r'_+', '_', name)
    return name


def build_output_filename(contractor_name: str, ca_number: str, letter_type: str) -> str:
    """
    Build output filename like:
      DhariwaltTradersPrivateLtd_CEJZ-GNR-18of2021-22_Extension_Notice.docx
    """
    short_contractor = sanitize_filename(contractor_name.split(",")[0].strip())
    short_ca = sanitize_filename(ca_number.strip())
    return f"{short_contractor}_{short_ca}_{letter_type}_Notice.docx"


# ---------------------------------------------------------------------------
# Field parsers for BGB column (Column C)
# ---------------------------------------------------------------------------

def parse_bgb_field(raw: str):
    """
    Parse a string like:
      '142GT02213330004 dt 29 Nov 2021 of Rs 825000/- against PSD'
    Returns dict with keys: bgb_number, bgb_date_str, bgb_amount_int, bgb_purpose
    Returns None on failure.
    """
    if not raw or not isinstance(raw, str):
        return None

    result = {}

    # BGB Number: text before 'dt'
    m = re.match(r'^\s*(\S+)\s+dt\b', raw, re.IGNORECASE)
    if m:
        result['bgb_number'] = m.group(1).strip()
    else:
        result['bgb_number'] = raw.split()[0] if raw.split() else ""

    # BGB Date: text between 'dt' and 'of'
    m = re.search(r'\bdt\s+(.+?)\s+(?:of\b|for\b)', raw, re.IGNORECASE)
    if m:
        result['bgb_date_str'] = m.group(1).strip()
    else:
        result['bgb_date_str'] = ""

    # Amount: number after 'Rs' and before '/-'
    m = re.search(r'\bRs\.?\s*([\d,]+)\s*/-', raw, re.IGNORECASE)
    if m:
        try:
            result['bgb_amount_int'] = int(m.group(1).replace(",", ""))
        except ValueError:
            result['bgb_amount_int'] = 0
    else:
        result['bgb_amount_int'] = 0

    # Purpose: 'against PSD' or 'against Retention Money'
    m = re.search(r'against\s+(PSD|Retention\s+Money|retention\s+money)', raw, re.IGNORECASE)
    result['bgb_purpose'] = m.group(0).strip() if m else "against Performance Security Deposit"

    return result


# ---------------------------------------------------------------------------
# Field parser for GE / Bank account column (Column J)
# ---------------------------------------------------------------------------

def parse_ge_bank_field(raw: str):
    """
    Parse a string like:
      'GE (A) Gandhinagar, Name of A/C: Garrison Engg (Proj) Div, Bank: SBI,
       Branch: Gandhinagar, A/C No: 10325061176, IFSC: SBIN0001355'
    Returns dict with keys: ge_location, account_name, bank_name, branch_name,
                             account_number, ifsc_code
    """
    if not raw or not isinstance(raw, str):
        return {}

    result = {}

    # GE Location: city after 'GE (A)'
    m = re.search(r'GE\s*\(A\)\s+([^,]+)', raw, re.IGNORECASE)
    result['ge_location'] = m.group(1).strip() if m else ""

    # Account Name
    m = re.search(r'Name\s+of\s+A/C\s*:\s*([^,]+)', raw, re.IGNORECASE)
    result['account_name'] = m.group(1).strip() if m else ""

    # Bank Name
    m = re.search(r'Bank\s*:\s*([^,]+)', raw, re.IGNORECASE)
    result['bank_name'] = m.group(1).strip() if m else ""

    # Branch
    m = re.search(r'Branch\s*:\s*([^,]+)', raw, re.IGNORECASE)
    result['branch_name'] = m.group(1).strip() if m else ""

    # Account Number
    m = re.search(r'A/C\s+No\s*:\s*([\d]+)', raw, re.IGNORECASE)
    result['account_number'] = m.group(1).strip() if m else ""

    # IFSC
    m = re.search(r'IFSC\s*:\s*(\S+)', raw, re.IGNORECASE)
    result['ifsc_code'] = m.group(1).strip() if m else ""

    return result


# ---------------------------------------------------------------------------
# CA No & Name of Work parser (Column A)
# ---------------------------------------------------------------------------

def parse_ca_field(raw: str):
    """
    Parse 'CEJZ/GNR/18 of 2021-22, Provn of Defi OTM Accn ...'
    Returns dict: ca_number, work_name
    """
    if not raw or not isinstance(raw, str):
        return {'ca_number': '', 'work_name': raw or ''}

    # CA number is everything up to the first comma
    parts = raw.split(',', 1)
    return {
        'ca_number': parts[0].strip(),
        'work_name': parts[1].strip() if len(parts) > 1 else '',
    }


# ---------------------------------------------------------------------------
# Bank name / address split (Column E)
# ---------------------------------------------------------------------------

def indian_amount_format(amount: int) -> str:
    """
    Format amount in Indian comma system: 825000 -> '8,25,000.00'
    """
    s = str(amount)
    if len(s) <= 3:
        return s + ".00"
    result = s[-3:]
    s = s[:-3]
    while s:
        chunk = s[-2:]
        result = chunk + "," + result
        s = s[:-2]
    return result + ".00"


def parse_bank_field(raw: str):
    """
    Parse 'HDFC, Aashapurna Mension, Station Road, Barmer'
    Returns dict: bank_name, bank_address
    """
    if not raw or not isinstance(raw, str):
        return {'bank_name': '', 'bank_address': ''}
    parts = raw.split(',', 1)
    return {
        'bank_name': parts[0].strip(),
        'bank_address': parts[1].strip() if len(parts) > 1 else '',
    }
