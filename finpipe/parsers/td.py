"""TD (EasyWeb) CSV: no header row -> date, description, debit, credit, balance.

Same layout for chequing, savings and TD credit cards (charges in debit, payments in credit).
"""
from .common import parse_table, read_rows

BANK = "TD"
HEADERLESS = {"date": 0, "description": 1, "debit": 2, "credit": 3}


def parse(path):
    return parse_table(read_rows(path), headerless=HEADERLESS)
