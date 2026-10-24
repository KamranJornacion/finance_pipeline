"""CIBC CSV: no header row -> date (YYYY-MM-DD), description, debit, credit[, card number]."""
from .common import parse_table, read_rows

BANK = "CIBC"
HEADERLESS = {"date": 0, "description": 1, "debit": 2, "credit": 3}


def parse(path):
    return parse_table(read_rows(path), headerless=HEADERLESS)
