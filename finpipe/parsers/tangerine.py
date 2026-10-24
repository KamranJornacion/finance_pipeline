"""Tangerine CSV: header Date,Transaction,Name,Memo,Amount. Amounts already signed."""
from .common import parse_table, read_rows

BANK = "Tangerine"


def parse(path):
    return parse_table(read_rows(path))
