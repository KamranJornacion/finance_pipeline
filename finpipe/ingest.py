"""Walk statements/, parse every file with the right bank parser, combine and de-duplicate."""
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .models import LEDGER_COLUMNS, Account
from .parsers import SUPPORTED_SUFFIXES, ParseError, parse_file, sniff_bank


@dataclass
class IngestResult:
    ledger: pd.DataFrame
    files_read: int = 0
    duplicates: int = 0
    problems: list[str] = field(default_factory=list)


def normalize_description(desc: str) -> str:
    return re.sub(r"\s+", " ", str(desc).upper()).strip()


def add_txn_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Stable id per transaction: identical rows in one file get different ids via an
    occurrence counter, while the same row seen again in an overlapping download
    gets the same id and is dropped as a duplicate."""
    df = df.copy()
    key = (
        df["account"] + "|" + df["date"].astype(str) + "|"
        + df["amount"].map(lambda a: f"{a:.2f}") + "|"
        + df["description"].map(normalize_description)
    )
    occurrence = key.groupby(key).cumcount().astype(str)
    df["txn_id"] = (key + "|" + occurrence).map(lambda s: hashlib.sha1(s.encode()).hexdigest()[:12])
    return df


def _files_in(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES and not p.name.startswith("~$")
    )


def _load_one(path: Path, acct: Account, root: Path) -> pd.DataFrame:
    df = parse_file(path, acct.bank)
    if acct.flip_signs:
        df["amount"] = -df["amount"]
    df["account"] = acct.name
    df["bank"] = acct.bank
    df["account_type"] = acct.type
    df["source_file"] = str(path.relative_to(root))
    return add_txn_ids(df)  # ids are per file so occurrence counts match across overlapping files


def ingest(statements_dir: Path, accounts: list[Account]) -> IngestResult:
    root = statements_dir.parent
    by_folder = {a.folder.lower(): a for a in accounts}
    frames, problems, files_read = [], [], 0

    jobs: list[tuple[Path, Account]] = []
    for entry in sorted(statements_dir.iterdir()):
        if entry.is_dir():
            acct = by_folder.get(entry.name.lower())
            files = _files_in(entry)
            if acct is None:
                if files:
                    problems.append(
                        f"Folder 'statements/{entry.name}' isn't listed in Accounts.xlsx, "
                        f"so its {len(files)} file(s) were skipped. Add a row for it.")
                continue
            jobs += [(f, acct) for f in files]
        elif entry.suffix.lower() in SUPPORTED_SUFFIXES:
            # loose file: assign it if the bank is recognisable and only one account matches
            try:
                bank = sniff_bank(entry)
            except ParseError:
                bank = None
            matches = [a for a in accounts if a.bank == bank]
            if bank and len(matches) == 1:
                jobs.append((entry, matches[0]))
            else:
                problems.append(
                    f"'{entry.name}' is loose in the statements folder and it isn't clear which "
                    f"account it belongs to. Move it into that account's folder.")

    for path, acct in jobs:
        try:
            frames.append(_load_one(path, acct, root))
            files_read += 1
        except ParseError as e:
            problems.append(f"Couldn't read '{path.relative_to(root)}' ({acct.bank}): {e}.")
        except Exception as e:  # never let one bad file stop the run
            problems.append(f"Couldn't read '{path.relative_to(root)}': unexpected error {e!r}.")

    if not frames:
        return IngestResult(pd.DataFrame(columns=LEDGER_COLUMNS), files_read, 0, problems)

    combined = pd.concat(frames, ignore_index=True)
    before = len(combined)
    ledger = combined.drop_duplicates("txn_id", keep="first")
    ledger = ledger[LEDGER_COLUMNS].sort_values(["date", "account"], kind="stable").reset_index(drop=True)
    ledger["date"] = pd.to_datetime(ledger["date"])
    return IngestResult(ledger, files_read, before - len(ledger), problems)
