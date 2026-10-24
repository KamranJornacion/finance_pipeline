"""Shared data definitions."""
from dataclasses import dataclass

# Columns every parser must return (amount: negative = money out, positive = money in).
PARSED_COLUMNS = ["date", "description", "amount"]

# Columns of the combined ledger after ingest.
LEDGER_COLUMNS = [
    "txn_id", "date", "account", "bank", "account_type",
    "description", "amount", "source_file",
]

BANKS = ["TD", "CIBC", "Tangerine", "Amex"]
ACCOUNT_TYPES = ["Chequing", "Savings", "Credit card"]

UNCATEGORIZED = "Uncategorized"
TRANSFER = "Transfer"


@dataclass
class Account:
    folder: str          # subfolder name inside statements/
    name: str            # friendly name shown in reports
    bank: str            # one of BANKS
    type: str            # one of ACCOUNT_TYPES
    flip_signs: bool = False  # escape hatch if a bank's export has the opposite sign convention


def is_yes(value) -> bool:
    return str(value).strip().lower() in {"yes", "y", "true", "1", "x"}
