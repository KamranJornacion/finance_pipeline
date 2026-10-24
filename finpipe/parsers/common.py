"""Format-tolerant helpers shared by all bank parsers.

Every parser returns a DataFrame with columns date, description, amount where
amount is negative for money out and positive for money in.
"""
import csv
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from ..models import PARSED_COLUMNS


class ParseError(Exception):
    """Raised with a plain-English message when a file can't be understood."""


# ---------------------------------------------------------------- reading

def read_rows(path: Path) -> list[list]:
    """Return the file as a list of rows (lists of cells), skipping blank rows."""
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        try:
            df = pd.read_excel(path, header=None, dtype=object)
        except ImportError:
            raise ParseError("this is an old Excel (.xls) file; download it as CSV instead")
        rows = [[None if pd.isna(c) else c for c in r] for r in df.itertuples(index=False)]
    else:
        text = None
        for enc in ("utf-8-sig", "cp1252"):
            try:
                text = path.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ParseError("could not read the file's text encoding")
        rows = list(csv.reader(text.splitlines()))
    return [r for r in rows if any(str(c).strip() for c in r if c is not None)]


# ---------------------------------------------------------------- values

def parse_amount(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 0.0 if pd.isna(value) else float(value)
    s = str(value).strip().replace("$", "").replace(",", "").replace(" ", "")
    if not s or s in {"-", "--"}:
        return 0.0
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if s.endswith("-"):  # trailing minus, e.g. "12.50-"
        negative, s = True, s[:-1]
    try:
        n = float(s)
    except ValueError:
        raise ParseError(f"'{value}' is not an amount")
    return -n if negative else n


_DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y",
    "%d %b %Y", "%b %d %Y", "%d-%b-%Y", "%d-%b-%y", "%Y%m%d",
]


def _clean_date_text(s: str) -> str:
    s = str(s).strip().replace(".", "").replace(",", "")
    s = re.sub(r"\bSept\b", "Sep", s, flags=re.I)
    return re.sub(r"\s+", " ", s)


def _parse_one(value, fmt: str) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(_clean_date_text(value), fmt).date()
    except (ValueError, TypeError):
        return None


def looks_like_date(value) -> bool:
    return any(_parse_one(value, f) for f in _DATE_FORMATS)


def parse_dates(values: list) -> list[date]:
    """Pick the one format that parses every value in the file (avoids day/month mix-ups)."""
    for fmt in _DATE_FORMATS:
        parsed = [_parse_one(v, fmt) for v in values]
        if all(parsed):
            return parsed
    bad = next(v for v in values if not looks_like_date(v))
    raise ParseError(f"could not understand the date '{bad}'")


# ---------------------------------------------------------------- headers

_SYNONYMS = {
    "date": ["date", "transaction date", "trans date", "posted date", "posting date"],
    "description": ["description", "name", "merchant", "payee", "transaction description", "details"],
    "memo": ["memo", "extended details", "notes"],
    "amount": ["amount", "amount ($)", "amount cad", "transaction amount"],
    "debit": ["debit", "debits", "withdrawal", "withdrawals", "money out"],
    "credit": ["credit", "credits", "deposit", "deposits", "money in"],
}


def _norm(cell) -> str:
    return re.sub(r"\s+", " ", str(cell or "").strip().lower().replace(".", ""))


def find_header(rows: list[list], max_scan: int = 30) -> tuple[int, dict] | None:
    """Locate a header row; return (row_index, {field: column_index})."""
    for i, row in enumerate(rows[:max_scan]):
        cells = [_norm(c) for c in row]
        cols = {}
        for field, names in _SYNONYMS.items():
            for j, c in enumerate(cells):
                if c in names:
                    cols[field] = j
                    break
        if "date" in cols and ("amount" in cols or "debit" in cols or "credit" in cols):
            return i, cols
    return None


# ---------------------------------------------------------------- main entry

def parse_table(rows: list[list], *, sign: int = 1, headerless: dict | None = None) -> pd.DataFrame:
    """Turn raw rows into the canonical frame.

    sign: multiply amounts by this (Amex reports charges as positive, so it uses -1).
    headerless: column positions to use when the file has no header row,
                e.g. {"date": 0, "description": 1, "debit": 2, "credit": 3}.
    """
    found = find_header(rows)
    if found:
        start, cols = found
        data = rows[start + 1:]
    elif headerless:
        cols, data = headerless, rows
    else:
        raise ParseError("could not find the column headings (Date, Description, Amount)")

    def cell(row, field):
        j = cols.get(field)
        return row[j] if j is not None and j < len(row) else None

    # keep only rows whose date cell actually holds a date (drops totals/footers)
    data = [r for r in data if cell(r, "date") not in (None, "") and looks_like_date(cell(r, "date"))]
    if not data:
        raise ParseError("no transactions found in the file")

    dates = parse_dates([cell(r, "date") for r in data])
    out = []
    for r, d in zip(data, dates):
        desc = str(cell(r, "description") or "").strip()
        memo = str(cell(r, "memo") or "").strip()
        if memo and memo.lower() not in desc.lower():
            desc = f"{desc} | {memo}" if desc else memo
        if "amount" in cols:
            amt = parse_amount(cell(r, "amount"))
        else:
            amt = parse_amount(cell(r, "credit")) - parse_amount(cell(r, "debit"))
        out.append((d, desc, round(amt * sign, 2)))
    return pd.DataFrame(out, columns=PARSED_COLUMNS)
