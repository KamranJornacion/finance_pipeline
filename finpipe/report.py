"""Summaries and the Finances.xlsx workbook."""
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

def write_workbook(path: Path, df: pd.DataFrame, budget: dict[str, float], categories: list[str],
                   problems: list[str]):
    flows = cashflow_frame(df)
    monthly = monthly_cashflow(flows)
    cats = spending_by_category(flows)
    review = review_items(df)
    wb = Workbook()

    # ---- Summary
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Cashflow summary"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = (f"Covers {df['date'].min():%b %d, %Y} to {df['date'].max():%b %d, %Y} - "
                f"{len(df):,} transactions across {df['account'].nunique()} account(s). "
                f"Transfers between your accounts are left out.")
    if len(review):
        ws["A3"] = f"{len(review)} transaction(s) need a category - see the Review tab (yellow columns)."
        ws["A3"].font = Font(bold=True, color="C00000")
    last12 = monthly.tail(12)
    inc, sp = last12["Income"].sum(), last12["Spending"].sum()
    kpis = [("Income", inc, MONEY0), ("Spending", sp, MONEY0), ("Net (saved)", inc - sp, MONEY0),
            ("Savings rate", (inc - sp) / inc if inc > 0 else None, PCT)]
    ws["A5"] = f"Last {len(last12)} month(s)"
    ws["A5"].font = Font(bold=True)
    for j, (label, val, fmt) in enumerate(kpis):
        ws.cell(row=6, column=1 + j, value=label).font = Font(color="595959")
        c = ws.cell(row=7, column=1 + j, value=val)
        c.number_format, c.font = fmt, Font(bold=True, size=14)
    end = _write_table(ws, last12, start_row=9, formats={
        "Income": MONEY0, "Spending": MONEY0, "Net": MONEY0, "Savings rate": PCT}, filters=False)
    for col in "ABCD":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["A"].width = 18
    if len(last12):
        chart = BarChart()
        chart.title, chart.y_axis.title, chart.height, chart.width = "Income vs spending", "$", 8, 18
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=9, max_row=end), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=10, max_row=end))
        ws.add_chart(chart, "G5")
    top = flows[(~flows["is_income"]) & (flows["month"].astype(str).isin(last12["Month"]))]
    top = (-top.groupby("category")["amount"].sum()).sort_values(ascending=False).head(10)
    top_df = pd.DataFrame({"Top spending categories": top.index, "Total": top.values,
                           "Per month": top.values / max(len(last12), 1)})
    trow = end + 3
    tend = _write_table(ws, top_df, start_row=trow, formats={"Total": MONEY0, "Per month": MONEY0}, filters=False)
    ws.column_dimensions["A"].width = 26
    if len(top_df):
        chart = BarChart()
        chart.type, chart.title, chart.height, chart.width, chart.legend = "bar", "Top categories (per month)", 8, 18, None
        chart.add_data(Reference(ws, min_col=3, min_row=trow, max_row=tend), titles_from_data=True)
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
    _write_table(mc, monthly, formats={"Income": MONEY, "Spending": MONEY, "Net": MONEY, "Savings rate": PCT},
                 widths={c: 16 for c in monthly.columns})
    mc.freeze_panes = "B2"

    # ---- By category
    bc = wb.create_sheet("By Category")
    _write_table(bc, cats, formats={c: MONEY0 for c in cats.columns if c != "Category"},
                 widths={"Category": 24})
    bc.freeze_panes = "B2"

    # ---- Budget
    bd = wb.create_sheet("Budget")
    bva = budget_vs_actual(flows, budget)
    if len(bva):
        last = _write_table(bd, bva, formats={c: MONEY0 for c in bva.columns if c != "Category"},
                            widths={"Category": 24, "Average vs budget": 18})
        vcol = get_column_letter(len(bva.columns))
        rng = f"{vcol}2:{vcol}{last}"
        bd.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
        bd.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=GREEN_FILL))
        bd.cell(row=last + 2, column=1, value="Average vs budget: green = under budget, red = over. "
                                              "Change budgets in Categories.xlsx > Budget.")
        bd.freeze_panes = "B2"

    # ---- By account
    ba = wb.create_sheet("By Account")
    acc = by_account(df)
    _write_table(ba, acc, formats={"Money in": MONEY, "Money out": MONEY, "Net": MONEY},
                 widths={"Account": 24, "Month": 10, "Money in": 14, "Money out": 14, "Net": 14})
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
