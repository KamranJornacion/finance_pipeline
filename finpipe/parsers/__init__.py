"""Bank parser registry and file sniffing."""
from pathlib import Path

from . import amex, cibc, tangerine, td
from .common import ParseError, _norm, find_header, read_rows

PARSERS = {m.BANK: m.parse for m in (td, cibc, tangerine, amex)}
SUPPORTED_SUFFIXES = {".csv", ".xlsx", ".xls"}


def parse_file(path: Path, bank: str):
    if bank not in PARSERS:
        raise ParseError(f"unknown bank '{bank}' (expected one of {', '.join(PARSERS)})")
    return PARSERS[bank](path)


def sniff_bank(path: Path) -> str | None:
    """Best guess at which bank produced a file, or None if unclear."""
    rows = read_rows(path)
    found = find_header(rows)
    if found:
        cells = {_norm(c) for c in rows[found[0]]}
        if {"transaction", "name", "memo"} <= cells:
            return "Tangerine"
        if "date processed" in cells or "cardmember" in cells or "description" in cells:
            return "Amex"
        return None
    first = str(rows[0][0]).strip() if rows and rows[0] else ""
    if len(first) == 10 and first[4] == "-":
        return "CIBC"
    if "/" in first:
        return "TD"
    return None


__all__ = ["PARSERS", "SUPPORTED_SUFFIXES", "ParseError", "parse_file", "sniff_bank"]
