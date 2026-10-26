"""Summaries and the Finances.xlsx workbook."""
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .categorize import (COL_ID, COL_MAKE_RULE, COL_YOUR_CATEGORY, COL_YOUR_TRANSFER,
                         is_income_category)
from .models import UNCATEGORIZED

MONEY = '"$"#,##0.00;[Red]-"$"#,##0.00'
MONEY0 = '"$"#,##0;[Red]-"$"#,##0'
PCT = "0%"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # yellow = "you can type here"
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
INPUT_COLUMNS = {COL_YOUR_CATEGORY, COL_YOUR_TRANSFER, COL_MAKE_RULE}


# ---------------------------------------------------------------- summaries

def cashflow_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Non-transfer transactions with a month column and an income flag."""
    flows = df[~df["is_transfer"]].copy()
    flows["month"] = flows["date"].dt.to_period("M")
    flows["is_income"] = flows["category"].map(is_income_category)
    return flows


def monthly_cashflow(flows: pd.DataFrame) -> pd.DataFrame:
    g = flows.groupby("month")
    income = g.apply(lambda x: x.loc[x["is_income"], "amount"].sum(), include_groups=False)
    spending = g.apply(lambda x: -x.loc[~x["is_income"], "amount"].sum(), include_groups=False)
    out = pd.DataFrame({"Income": income, "Spending": spending})
    out["Net"] = out["Income"] - out["Spending"]
    out["Savings rate"] = (out["Net"] / out["Income"]).where(out["Income"] > 0)
    out.index = out.index.astype(str)
    out.index.name = "Month"
    return out.reset_index()


def spending_by_category(flows: pd.DataFrame) -> pd.DataFrame:
    spend = flows[~flows["is_income"]]
    if spend.empty:
        return pd.DataFrame(columns=["Category", "Total"])
    p = spend.pivot_table(index="category", columns="month", values="amount",
                          aggfunc="sum", fill_value=0) * -1
    p.columns = [str(c) for c in p.columns]
    p["Total"] = p.sum(axis=1)
    p["Monthly avg"] = p["Total"] / (len(p.columns) - 1)
    p = p.sort_values("Total", ascending=False)
    p.index.name = "Category"
    return p.reset_index()


def budget_vs_actual(flows: pd.DataFrame, budget: dict[str, float], months: int = 6) -> pd.DataFrame:
    spend = flows[~flows["is_income"]]
    recent = sorted(spend["month"].unique())[-months:]
    cats = list(budget) + sorted(set(spend["category"]) - set(budget))
    rows = []
    for cat in cats:
        row = {"Category": cat, "Monthly budget": budget.get(cat)}
        actuals = []
        for m in recent:
            v = -spend.loc[(spend["category"] == cat) & (spend["month"] == m), "amount"].sum()
            row[str(m)] = v
            actuals.append(v)
        row["Average"] = sum(actuals) / len(actuals) if actuals else 0
        row["Average vs budget"] = (row["Monthly budget"] - row["Average"]) if row["Monthly budget"] is not None else None
        if row["Monthly budget"] is not None or row["Average"]:
            rows.append(row)
    return pd.DataFrame(rows)


def by_account(df: pd.DataFrame) -> pd.DataFrame:
    d = df.assign(month=df["date"].dt.to_period("M").astype(str))
    g = d.groupby(["account", "month"])["amount"]
    out = pd.DataFrame({
        "Money in": g.apply(lambda s: s[s > 0].sum()),
        "Money out": g.apply(lambda s: -s[s < 0].sum()),
        "Net": g.sum(),
        "Transactions": g.size(),
    }).reset_index().rename(columns={"account": "Account", "month": "Month"})
    return out


def review_items(df: pd.DataFrame) -> pd.DataFrame:
    needs = df[
        (df["source"] != "You")
        & (((df["category"] == UNCATEGORIZED) & ~df["is_transfer"]) | (df["review_note"] != ""))
    ].copy()
    needs["why"] = needs["review_note"].where(needs["review_note"] != "", "No rule matched")
    return needs.sort_values("date", ascending=False)


# ---------------------------------------------------------------- excel helpers

def _write_table(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1, formats: dict | None = None,
                 widths: dict | None = None, filters: bool = True):
    formats, widths = formats or {}, widths or {}
    for j, col in enumerate(df.columns):
        c = ws.cell(row=start_row, column=start_col + j, value=col)
        c.fill, c.font = (INPUT_FILL, Font(bold=True)) if col in INPUT_COLUMNS else (HEADER_FILL, HEADER_FONT)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, rec in enumerate(df.itertuples(index=False), start=start_row + 1):
        for j, val in enumerate(rec):
            if val is not None and not isinstance(val, str) and pd.isna(val):
                val = None
            elif isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            c = ws.cell(row=i, column=start_col + j, value=val)
            col = df.columns[j]
            if col in formats:
                c.number_format = formats[col]
            if col in INPUT_COLUMNS:
                c.fill = INPUT_FILL
    for j, col in enumerate(df.columns):
        letter = get_column_letter(start_col + j)
        width = widths.get(col) or min(max(len(str(col)), *(len(str(v)) for v in df[col].head(200)), 8) + 2, 50) \
            if len(df) else len(str(col)) + 4
        ws.column_dimensions[letter].width = width
    if filters and len(df):
        ws.auto_filter.ref = f"{get_column_letter(start_col)}{start_row}:" \
                             f"{get_column_letter(start_col + len(df.columns) - 1)}{start_row + len(df)}"
    return start_row + len(df)


def _category_dropdown(wb, ws, categories: list[str], col_letter: str, last_row: int):
    lists = wb["Lists"] if "Lists" in wb.sheetnames else wb.create_sheet("Lists")
    lists.sheet_state = "hidden"
    if lists.max_row <= 1 and lists["A1"].value is None:
        for i, c in enumerate(categories, start=1):
            lists.cell(row=i, column=1, value=c)
    dv = DataValidation(type="list", formula1=f"=Lists!$A$1:$A${max(len(categories), 1)}",
                        allow_blank=True, showErrorMessage=False)  # typing a new category is allowed
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{max(last_row, 2)}")
    yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(yn)
    return yn


# ---------------------------------------------------------------- workbook

def _sheet_title(name: str, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]+", " ", str(name or "Account")).strip()
    base = re.sub(r"\s+", " ", base)[:31] or "Account"
    candidate = base
    suffix = 1
    while candidate in used:
        suffix_text = str(suffix)
        max_base = 31 - len(suffix_text) - 1
        candidate = f"{base[:max_base]} {suffix_text}"
        suffix += 1
    used.add(candidate)
    return candidate


def _month_start(month: str) -> str:
    return pd.Period(month, freq="M").to_timestamp().strftime("DATE(%Y,%m,1)")


def _month_end(month: str) -> str:
    return (pd.Period(month, freq="M") + 1).to_timestamp().strftime("DATE(%Y,%m,1)")


def _ledger_month_formula(month: str, kind: str) -> str:
    start = _month_start(month)
    end = _month_end(month)
    if kind == "income":
        return (
            f'=SUMPRODUCT((Ledger!$A$2:$A$1048576>={start})*(Ledger!$A$2:$A$1048576<{end})*'
            f'(Ledger!$H$2:$H$1048576=FALSE)*(LEFT(Ledger!$G$2:$G$1048576,6)="Income")*(Ledger!$F$2:$F$1048576))'
        )
    if kind == "spending":
        return (
            f'=-SUMPRODUCT((Ledger!$A$2:$A$1048576>={start})*(Ledger!$A$2:$A$1048576<{end})*'
            f'(Ledger!$H$2:$H$1048576=FALSE)*(LEFT(Ledger!$G$2:$G$1048576,6)<>"Income")*(Ledger!$F$2:$F$1048576))'
        )
    raise ValueError(f"Unknown month formula kind: {kind}")


def _ledger_category_formula(category: str) -> str:
    safe = str(category).replace('"', '""')
    return (
        f'=SUMPRODUCT((Ledger!$G$2:$G$1048576="{safe}")*(Ledger!$H$2:$H$1048576=FALSE)*(-Ledger!$F$2:$F$1048576))'
    )


def write_workbook(path: Path, df: pd.DataFrame, budget: dict[str, float], categories: list[str],
                   problems: list[str]):
    flows = cashflow_frame(df)
    monthly = monthly_cashflow(flows)
    cats = spending_by_category(flows)
    review = review_items(df)
    wb = Workbook()

    # ---- Raw data layer
    used = set()
    ws = wb.active
    ws.title = "Summary"
    ledger = df[["date", "account", "bank", "account_type", "description", "amount", "category",
                 "is_transfer", "source", "source_file", "txn_id"]].copy()
    ledger = ledger.sort_values("date", ascending=False).reset_index(drop=True)
    ledger_ws = wb.create_sheet("Ledger")
    _write_table(ledger_ws, ledger, formats={"date": "yyyy-mm-dd", "amount": MONEY},
                 widths={"account": 20, "description": 45, "source_file": 26, "txn_id": 14})
    ledger_ws.freeze_panes = "A2"
    for acct in sorted(df["account"].drop_duplicates()):
        acct_df = ledger[ledger["account"] == acct].copy()
        acct_ws = wb.create_sheet(_sheet_title(acct, used))
        _write_table(acct_ws, acct_df, formats={"date": "yyyy-mm-dd", "amount": MONEY},
                     widths={"account": 20, "description": 45, "source_file": 26, "txn_id": 14})
        acct_ws.freeze_panes = "A2"

    # ---- Summary
    ws = wb["Summary"]
    ws["A1"] = "Cashflow summary"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = (f"Covers {df['date'].min():%b %d, %Y} to {df['date'].max():%b %d, %Y} - "
                f"{len(df):,} transactions across {df['account'].nunique()} account(s). "
                f"Transfers between your accounts are left out.")
    if len(review):
        ws["A3"] = f"{len(review)} transaction(s) need a category - see the Review tab (yellow columns)."
        ws["A3"].font = Font(bold=True, color="C00000")
    last12 = monthly.tail(12)
    ws["A5"] = f"Last {len(last12)} month(s)"
    ws["A5"].font = Font(bold=True)
    kpis = ["Income", "Spending", "Net (saved)", "Savings rate"]
    for j, label in enumerate(kpis):
        ws.cell(row=6, column=1 + j, value=label).font = Font(color="595959")

    monthly_start_row = 2
    ws["B7"] = f"=SUM('Monthly Cashflow'!B{monthly_start_row}:B{monthly_start_row + len(last12) - 1})"
    ws["C7"] = f"=SUM('Monthly Cashflow'!C{monthly_start_row}:C{monthly_start_row + len(last12) - 1})"
    ws["D7"] = "=B7-C7"
    ws["E7"] = "=IF(B7=0,0,D7/B7)"
    for cell in ["B7", "C7", "D7", "E7"]:
        ws[cell].number_format = MONEY0 if cell != "E7" else PCT
        ws[cell].font = Font(bold=True, size=14)

    for idx, month in enumerate(last12["Month"], start=10):
        month_row = idx - 8
        ws.cell(row=idx, column=1, value=f"='Monthly Cashflow'!A{month_row}")
        ws.cell(row=idx, column=2, value=f"='Monthly Cashflow'!B{month_row}")
        ws.cell(row=idx, column=3, value=f"='Monthly Cashflow'!C{month_row}")
        ws.cell(row=idx, column=4, value=f"='Monthly Cashflow'!D{month_row}")
        ws.cell(row=idx, column=5, value=f"='Monthly Cashflow'!E{month_row}")

    ws.cell(row=9, column=1, value="Month")
    ws.cell(row=9, column=2, value="Income")
    ws.cell(row=9, column=3, value="Spending")
    ws.cell(row=9, column=4, value="Net")
    ws.cell(row=9, column=5, value="Savings rate")
    for c in range(1, 6):
        ws.cell(row=9, column=c).fill, ws.cell(row=9, column=c).font = HEADER_FILL, HEADER_FONT
    ws.row_dimensions[9].height = 18
    end = max(10, len(last12) + 9)

    for col in "ABCDE":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["A"].width = 18

    top_rows = []
    for i in range(1, 11):
        top_rows.append({
            "Top spending categories": f"='By Category'!A{i+1}",
            "Total": f"='By Category'!B{i+1}",
            "Per month": f"='By Category'!C{i+1}",
        })
    top_df = pd.DataFrame(top_rows)
    if len(last12):
        chart = BarChart()
        chart.title, chart.y_axis.title, chart.height, chart.width = "Income vs spending", "$", 8, 18
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=9, max_row=end), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=10, max_row=end))
        ws.add_chart(chart, "G5")
    trow = end + 3
    tend = _write_table(ws, top_df, start_row=trow, formats={"Total": MONEY0, "Per month": MONEY0}, filters=False)
    ws.column_dimensions["A"].width = 26
    if len(top_df):
        chart = BarChart()
        chart.type, chart.title, chart.height, chart.width, chart.legend = "bar", "Top categories (per month)", 8, 18, None
        chart.add_data(Reference(ws, min_col=3, min_row=trow + 1, max_row=tend), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=trow + 1, max_row=tend))
        chart.y_axis.scaling.orientation = "minMax"
        chart.x_axis.scaling.orientation = "maxMin"
        ws.add_chart(chart, f"G{trow}")
    if problems:
        prow = tend + 3
        ws.cell(row=prow, column=1, value="Problems found during the last update:").font = Font(bold=True, color="C00000")
        for k, p in enumerate(problems, start=1):
            ws.cell(row=prow + k, column=1, value=p)

    # ---- Review
    rv = wb.create_sheet("Review")
    rdf = pd.DataFrame({
        "Date": review["date"], "Account": review["account"], "Description": review["description"],
        "Amount": review["amount"], "Why it's here": review["why"], COL_YOUR_CATEGORY: "",
        COL_YOUR_TRANSFER: "", COL_MAKE_RULE: "", COL_ID: review["txn_id"],
    })
    last = _write_table(rv, rdf, formats={"Date": "yyyy-mm-dd", "Amount": MONEY},
                        widths={"Description": 45, "Why it's here": 30, COL_YOUR_CATEGORY: 20,
                                COL_YOUR_TRANSFER: 12, COL_MAKE_RULE: 24, COL_ID: 14})
    rv.freeze_panes = "A2"
    yn = _category_dropdown(wb, rv, categories, "F", last)
    yn.add(f"G2:G{max(last, 2)}")
    rv.row_dimensions[1].height = 32

    # ---- Transactions
    tx = wb.create_sheet("Transactions")
    tdf = pd.DataFrame({
        "Date": df["date"], "Account": df["account"], "Description": df["description"],
        "Amount": df["amount"], "Category": df["category"],
        "Transfer?": df["is_transfer"].map({True: "Yes", False: ""}), "Set by": df["source"],
        COL_YOUR_CATEGORY: "", COL_YOUR_TRANSFER: "", COL_MAKE_RULE: "",
        "Source file": df["source_file"], COL_ID: df["txn_id"],
    }).sort_values("Date", ascending=False)
    last = _write_table(tx, tdf, formats={"Date": "yyyy-mm-dd", "Amount": MONEY},
                        widths={"Description": 45, COL_YOUR_CATEGORY: 20, COL_YOUR_TRANSFER: 12,
                                COL_MAKE_RULE: 24, COL_ID: 14})
    tx.freeze_panes = "A2"
    yn = _category_dropdown(wb, tx, categories, "H", last)
    yn.add(f"I2:I{max(last, 2)}")
    tx.row_dimensions[1].height = 32

    # ---- Monthly cashflow
    mc = wb.create_sheet("Monthly Cashflow")
    months = []
    for idx, row in enumerate(monthly.itertuples(index=False), start=2):
        month = str(row[0])
        months.append({
            "Month": month,
            "Income": _ledger_month_formula(month, "income"),
            "Spending": _ledger_month_formula(month, "spending"),
            "Net": f"=B{idx}-C{idx}",
            "Savings rate": f"=IF(B{idx}=0,0,D{idx}/B{idx})",
        })
    mc_df = pd.DataFrame(months)
    _write_table(mc, mc_df, formats={"Income": MONEY, "Spending": MONEY, "Net": MONEY, "Savings rate": PCT},
                 widths={c: 16 for c in mc_df.columns})
    mc.freeze_panes = "B2"
    mc["B2"].value = mc_df.iloc[0, 1]

    # ---- By category
    bc = wb.create_sheet("By Category")
    cat_rows = []
    for idx, row in enumerate(cats.itertuples(index=False), start=2):
        cat_rows.append({
            "Category": row[0],
            "Total": _ledger_category_formula(str(row[0])),
            "Monthly avg": f"=IF(COUNTA(Ledger!$G$2:$G$1048576)=0,0,B{idx}/MAX(1,COUNTIF(Ledger!$G$2:$G$1048576,\"{str(row[0]).replace('"', '""')}\")))",
        })
    bc_df = pd.DataFrame(cat_rows)
    _write_table(bc, bc_df, formats={c: MONEY0 for c in bc_df.columns if c != "Category"},
                 widths={"Category": 24})
    bc.freeze_panes = "B2"

    # ---- Budget
    bd = wb.create_sheet("Budget")
    months = sorted(pd.PeriodIndex(df["date"], freq="M").unique().astype(str))
    if months:
        headers = ["Category", "Monthly budget"] + months + ["Average", "Average vs budget"]
        for r, h in enumerate(headers, start=1):
            bd.cell(row=1, column=r, value=h)
            bd.cell(row=1, column=r).fill, bd.cell(row=1, column=r).font = HEADER_FILL, HEADER_FONT
        for i, cat in enumerate(sorted(set(budget) | set(categories)), start=2):
            bd.cell(row=i, column=1, value=cat)
            if cat in budget:
                bd.cell(row=i, column=2, value=budget[cat])
            for j, month in enumerate(months, start=3):
                start = pd.Period(month, freq="M").to_timestamp().strftime("DATE(%Y,%m,1)")
                end = (pd.Period(month, freq="M") + 1).to_timestamp().strftime("DATE(%Y,%m,1)")
                bd.cell(row=i, column=j, value=f'=SUMPRODUCT((Ledger!$A$2:$A$1048576>={start})*(Ledger!$A$2:$A$1048576<{end})*(Ledger!$H$2:$H$1048576=FALSE)*(Ledger!$G$2:$G$1048576="{cat}")*(-Ledger!$F$2:$F$1048576))')
            avg_col = len(headers) - 1
            bd.cell(row=i, column=avg_col, value=f'=AVERAGE(C{i}:{get_column_letter(len(headers)-2)}{i})')
            bd.cell(row=i, column=len(headers), value=f'=IF(B{i}="",,B{i}-{get_column_letter(avg_col)}{i})')
        vcol = get_column_letter(len(headers))
        rng = f"{vcol}2:{vcol}{bd.max_row}"
        bd.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
        bd.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=GREEN_FILL))
        bd.cell(row=bd.max_row + 2, column=1, value="Average vs budget: green = under budget, red = over. Change budgets in Categories.xlsx > Budget.")
        bd.freeze_panes = "B2"

    # ---- By account
    ba = wb.create_sheet("By Account")
    accounts = sorted(df["account"].drop_duplicates())
    months = sorted(pd.PeriodIndex(df["date"], freq="M").unique().astype(str))
    for r, h in enumerate(["Account", "Month", "Money in", "Money out", "Net", "Transactions"], start=1):
        ba.cell(row=1, column=r, value=h)
        ba.cell(row=1, column=r).fill, ba.cell(row=1, column=r).font = HEADER_FILL, HEADER_FONT
    row = 2
    for acct in accounts:
        for month in months:
            ba.cell(row=row, column=1, value=acct)
            ba.cell(row=row, column=2, value=month)
            start = pd.Period(month, freq="M").to_timestamp().strftime("DATE(%Y,%m,1)")
            end = (pd.Period(month, freq="M") + 1).to_timestamp().strftime("DATE(%Y,%m,1)")
            ba.cell(row=row, column=3, value=f'=SUMPRODUCT((Ledger!$A$2:$A$1048576>={start})*(Ledger!$A$2:$A$1048576<{end})*(Ledger!$B$2:$B$1048576="{acct}")*(Ledger!$F$2:$F$1048576>0)*(Ledger!$F$2:$F$1048576))')
            ba.cell(row=row, column=4, value=f'=SUMPRODUCT((Ledger!$A$2:$A$1048576>={start})*(Ledger!$A$2:$A$1048576<{end})*(Ledger!$B$2:$B$1048576="{acct}")*(Ledger!$F$2:$F$1048576<0)*(-Ledger!$F$2:$F$1048576))')
            ba.cell(row=row, column=5, value=f'=C{row}-D{row}')
            ba.cell(row=row, column=6, value=f'=COUNTIFS(Ledger!$A$2:$A$1048576,">="&{start},Ledger!$A$2:$A$1048576,"<{end}",Ledger!$B$2:$B$1048576,"{acct}")')
            row += 1
    ba.freeze_panes = "A2"

    # ---- How to use
    hu = wb.create_sheet("How to use")
    hu.column_dimensions["A"].width = 110
    for line in [
        "How to use this report",
        "This file is rebuilt every time you run Update Finances. Don't keep notes here - they will be replaced.",
        "Yellow columns are for you. Your entries are saved permanently on the next update:",
        f"  - {COL_YOUR_CATEGORY}: pick or type a category for that one transaction.",
        f"  - {COL_YOUR_TRANSFER}: Yes if it's money moving between your own accounts (it will be left out of "
        "income and spending); No if it was wrongly treated as a transfer.",
        f"  - {COL_MAKE_RULE}: optional. Type a word from the description (e.g. NETFLIX) to give every "
        "similar transaction, past and future, the same category.",
        "Then save this file, close it, and run Update Finances again.",
        "Tabs: Summary (overview), Review (needs your input), Transactions (everything), Monthly Cashflow, "
        "By Category, Budget (vs. your targets), By Account.",
        "Income = categories starting with 'Income'. Spending = everything else that isn't a transfer. "
        "Refunds reduce spending in their category.",
    ]:
        hu.append([line])
    hu["A1"].font = Font(bold=True, size=14)

    wb.move_sheet("Lists", offset=len(wb.sheetnames))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_csv(path: Path, df: pd.DataFrame):
    out = df.drop(columns=["review_note"]).copy()
    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    out.to_csv(path, index=False)
