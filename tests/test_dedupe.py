from finpipe.ingest import ingest
from finpipe.setup import load_accounts


def test_overlapping_downloads_deduped_but_identical_purchases_kept(project):
    res = ingest(project / "statements", load_accounts(project / "Accounts.xlsx"))
    assert res.files_read == 5
    assert res.duplicates == 2
    coffees = res.ledger[res.ledger["description"] == "TIM HORTONS #44"]
    assert len(coffees) == 2                       # two real coffees on the same day survive
    assert res.ledger["txn_id"].is_unique


def test_txn_ids_stable_across_runs(project):
    accts = load_accounts(project / "Accounts.xlsx")
    a = ingest(project / "statements", accts).ledger["txn_id"].tolist()
    b = ingest(project / "statements", accts).ledger["txn_id"].tolist()
    assert a == b


def test_unknown_folder_and_bad_file_reported_not_fatal(project):
    (project / "statements/mystery-bank").mkdir()
    (project / "statements/mystery-bank/a.csv").write_text("x\n")
    (project / "statements/td-chequing/broken.csv").write_text("not,a,statement\n")
    res = ingest(project / "statements", load_accounts(project / "Accounts.xlsx"))
    assert res.files_read == 5
    assert any("mystery-bank" in p for p in res.problems)
    assert any("broken.csv" in p for p in res.problems)
