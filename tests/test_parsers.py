from datetime import date

import pytest

from conftest import FIXTURES
from finpipe.parsers import ParseError, parse_file, sniff_bank
from finpipe.parsers.common import parse_amount, parse_dates


def test_td_headerless_debit_credit():
    df = parse_file(FIXTURES / "td-chequing/accountactivity.csv", "TD")
    assert len(df) == 10
    assert df.iloc[0].tolist() == [date(2026, 9, 1), "PAYROLL ACME CORP", 3200.00]
    assert df.iloc[1]["amount"] == -85.20


def test_cibc_iso_dates_and_card_column():
    df = parse_file(FIXTURES / "cibc-visa/cibc.csv", "CIBC")
    assert df["date"].tolist()[0] == date(2026, 9, 8)
    assert df["amount"].tolist() == [-40.00, 400.00, -55.10, -95.00]


def test_tangerine_signed_amounts_and_memo():
    df = parse_file(FIXTURES / "tangerine-savings/export.csv", "Tangerine")
    assert df.iloc[0]["description"] == "Internet Deposit from TD | Transfer"
    assert df["amount"].tolist() == [500.00, 1.23, 250.00]


def test_amex_charges_become_negative():
    df = parse_file(FIXTURES / "amex-card/activity_sep.csv", "Amex")
    assert df.iloc[0]["date"] == date(2026, 9, 5)
    assert df.iloc[0]["amount"] == -142.35            # charge
    assert df.iloc[4]["amount"] == 850.00             # payment received


def test_header_found_below_preamble(tmp_path):
    f = tmp_path / "x.csv"
    f.write_text("Account summary\nCard ending 1234\n\nDate,Description,Amount\n2026-01-05,COFFEE,4.25\n")
    df = parse_file(f, "Amex")
    assert df.iloc[0]["amount"] == -4.25


@pytest.mark.parametrize("raw,expected", [
    ("1,234.56", 1234.56), ("$12.00", 12.0), ("(5.00)", -5.0), ("7.50-", -7.5), ("", 0.0), (None, 0.0), (3, 3.0),
])
def test_parse_amount(raw, expected):
    assert parse_amount(raw) == expected


def test_day_month_ambiguity_resolved_per_file():
    # 13/02 can only be dd/mm, so the whole file is read as dd/mm
    assert parse_dates(["01/02/2026", "13/02/2026"]) == [date(2026, 2, 1), date(2026, 2, 13)]
    assert parse_dates(["01/02/2026", "02/13/2026"]) == [date(2026, 1, 2), date(2026, 2, 13)]


def test_garbage_file_gives_plain_error(tmp_path):
    f = tmp_path / "junk.csv"
    f.write_text("hello,world\nfoo,bar\n")
    with pytest.raises(ParseError):
        parse_file(f, "Tangerine")


def test_sniff_bank():
    assert sniff_bank(FIXTURES / "td-chequing/accountactivity.csv") == "TD"
    assert sniff_bank(FIXTURES / "cibc-visa/cibc.csv") == "CIBC"
    assert sniff_bank(FIXTURES / "tangerine-savings/export.csv") == "Tangerine"
    assert sniff_bank(FIXTURES / "amex-card/activity_sep.csv") == "Amex"
