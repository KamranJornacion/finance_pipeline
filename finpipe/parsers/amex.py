"""American Express Canada CSV/Excel: header with Date, Description, Amount.

Amex lists charges as positive and payments/credits as negative, so signs are flipped.
Older headerless CSVs are date, reference, amount, description.
"""
from .common import parse_table, read_rows

BANK = "Amex"
HEADERLESS = {"date": 0, "amount": 2, "description": 3}


def parse(path):
    return parse_table(read_rows(path), sign=-1, headerless=HEADERLESS)
