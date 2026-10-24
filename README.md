# Finance Pipeline

Turns bank downloads from **TD, CIBC, Tangerine and American Express** into one Excel
workbook. The workbook shows your monthly income, spending, savings and budget, and covers chequing, savings and credit cards.

---

## Everyday use

1. **Download** your recent transactions as CSV from online banking.
   `statements/HOW_TO_DOWNLOAD.txt` has the steps for each bank.
2. **Save** each file into its account's folder inside `statements/`,
   e.g. `statements/td-chequing/`.
   Overlapping date ranges are fine because duplicates are removed automatically.
3. **Double-click "Update Finances"**:
   - Windows: `Update Finances.bat`
   - Mac: `Update Finances.command`
   - Linux: `Update Finances.desktop`, or run `./update_finances.sh`
4. `output/Finances.xlsx` opens with your updated report.

### Fixing categories

Anything the program couldn't categorize shows up on the **Review** tab.
- Type or pick a category in the yellow **Your category** column.
- Set **Transfer?** to Yes if the money just moved between your own accounts.
- Optionally type a word in **Make a rule from this text** (e.g. `NETFLIX`), so every similar transaction, past and future, is handled automatically.

Then **save and close** the workbook and run Update Finances again. Your choices are kept permanently.
The same yellow columns are on the **Transactions** tab, so you can correct anything.

### Settings files (in this folder)

| File | What it's for |
|---|---|
| `Accounts.xlsx` | Your list of accounts. Add a row and run Update Finances once to get a new folder. |
| `Categories.xlsx` | The rules (description text → category) and your monthly **Budget**. |

Both files are created the first time you run Update Finances.

### Report tabs
**Summary** (overview and charts) · **Review** (needs your input) · **Transactions** ·
**Monthly Cashflow** · **By Category** · **Budget** (vs. your targets) · **By Account**

Transfers between your own accounts, like paying a credit card from chequing, are left out of income and spending so nothing is counted twice.

---

## First-time setup

1. Install **Python 3.10+** from <https://www.python.org/downloads/>.
   On Windows, tick *"Add python.exe to PATH"* during install.
2. Double-click Update Finances. The first run takes about a minute to install what it needs, then creates `Accounts.xlsx`.
3. Edit `Accounts.xlsx` to list your accounts, save it, and run Update Finances again to create the folders.

On Mac, the first time you may need to right-click `Update Finances.command` → **Open**.
On Linux, right-click `Update Finances.desktop` → **Allow launching**.

---

## Privacy

Your statements, reports and personal settings are **never** saved to git or GitHub.
`.gitignore` excludes `statements/`, `output/`, `overrides.csv`, `Accounts.xlsx` and `Categories.xlsx`.
Only the program itself is tracked.

---

## For developers

```
finpipe/
  parsers/      one module per bank -> DataFrame(date, description, amount); amount < 0 = money out
  ingest.py     walk statements/<account folder>/, parse, stable txn_id, de-duplicate
  categorize.py rules (word-start match, first wins), overrides, Review-sheet harvesting, transfer pairing
  report.py     summaries + Finances.xlsx
  setup.py      generates Accounts.xlsx / Categories.xlsx / folders on first run
  cli.py        python -m finpipe [--from YYYY-MM] [--to YYYY-MM] [--no-open] [--root DIR]
```

- **Run tests:** `pip install -r requirements-dev.txt && pytest`.
  Fixtures in `tests/fixtures/` are made-up data.
- **Adding a bank:** add `finpipe/parsers/<bank>.py` with `BANK` and `parse(path)`, register it in
  `parsers/__init__.py`, and add the name to `BANKS` in `models.py`.
- **Dedup:** `txn_id` = hash(account, date, amount, description, occurrence-within-file). The same
  transaction in two overlapping downloads collapses to one, while two identical purchases on the same day both stay.
- `overrides.csv` stores manual choices keyed by `txn_id`. It is maintained automatically from the workbook.
