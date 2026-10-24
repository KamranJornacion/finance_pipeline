"""Entry point used by the "Update Finances" launchers: python -m finpipe"""
import argparse
import os
import platform
import subprocess
from pathlib import Path

import pandas as pd

from . import categorize as cat
from .ingest import ingest
from .report import review_items, write_csv, write_workbook
from .setup import ensure_setup, load_accounts

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def open_file(path: Path):
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass  # opening is a convenience; the path is printed anyway


def _month(s: str) -> pd.Period:
    return pd.Period(s, freq="M")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Combine bank statements into Finances.xlsx")
    ap.add_argument("--root", type=Path, default=PROJECT_ROOT, help="project folder (default: this one)")
    ap.add_argument("--from", dest="start", type=_month, help="first month to include, e.g. 2026-01")
    ap.add_argument("--to", dest="end", type=_month, help="last month to include, e.g. 2026-09")
    ap.add_argument("--no-open", action="store_true", help="don't open the workbook when done")
    args = ap.parse_args(argv)

    root = args.root.resolve()
    out_dir = root / "output"
    xlsx = out_dir / "Finances.xlsx"
    csv_path = out_dir / "transactions.csv"
    accounts_path, categories_path = root / "Accounts.xlsx", root / "Categories.xlsx"
    overrides_path = root / "overrides.csv"

    print("Updating your finances...\n")
    first_run = not accounts_path.exists()
    for note in ensure_setup(root):
        print(f"  * {note}")
    if first_run:
        print("\nFirst-time setup done. Next steps:")
        print("  1. Open Accounts.xlsx and list your accounts (one row each), then save.")
        print("  2. Run Update Finances again - a folder appears in 'statements' for each account.")
        print("  3. Download CSV files from your banks into those folders and run Update Finances.")
        if not args.no_open:
            open_file(accounts_path)
        return 0

    # 1. save anything typed into last time's report
    try:
        saved, rules_added = cat.harvest_edits(xlsx, overrides_path, categories_path)
    except PermissionError:
        print("  ! Categories.xlsx is open in another program. Close it and run again.")
        return 1
    if saved:
        print(f"  * Saved {saved} category choice(s) you made in the last report.")
    if rules_added:
        print(f"  * Added {rules_added} new rule(s) to Categories.xlsx.")

    # 2. read statements
    accounts = load_accounts(accounts_path)
    result = ingest(root / "statements", accounts)
    ledger = result.ledger
    if args.start:
        ledger = ledger[ledger["date"].dt.to_period("M") >= args.start]
    if args.end:
        ledger = ledger[ledger["date"].dt.to_period("M") <= args.end]

    if ledger.empty:
        for p in result.problems:
            print(f"  ! {p}")
        print("\nNo transactions found yet. Put CSV downloads into the account folders inside "
              "'statements' (see statements/HOW_TO_DOWNLOAD.txt) and run again.")
        return 0 if not result.problems else 1

    previous_ids = set()
    if csv_path.exists():
        try:
            previous_ids = set(pd.read_csv(csv_path, usecols=["txn_id"], dtype=str)["txn_id"])
        except Exception:
            pass

    # 3. categorize
    rules = cat.load_rules(categories_path)
    df = cat.categorize(ledger, rules, cat.load_overrides(overrides_path))
    budget = cat.load_budget(categories_path)
    categories = sorted(set(budget) | {r.category for r in rules} | set(df["category"]))

    # 4. write outputs
    try:
        write_workbook(xlsx, df, budget, categories, result.problems)
        write_csv(csv_path, df)
    except PermissionError:
        print("\n  ! Finances.xlsx is open in Excel. Close it and run Update Finances again.")
        return 1

    new = len(set(df["txn_id"]) - previous_ids) if previous_ids else len(df)
    to_review = len(review_items(df))
    print(f"  * Read {result.files_read} file(s) from {df['account'].nunique()} account(s).")
    print(f"  * {len(df):,} transactions ({new:,} new"
          + (f", {result.duplicates:,} duplicates from overlapping downloads ignored" if result.duplicates else "")
          + ").")
    print(f"  * {to_review} transaction(s) need a category - see the Review tab." if to_review
          else "  * Everything is categorized.")
    for p in result.problems:
        print(f"  ! {p}")
    print(f"\nDone. Report saved to {xlsx}")
    if not args.no_open:
        open_file(xlsx)
    return 1 if result.problems else 0
