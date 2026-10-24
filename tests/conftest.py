import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

FIXTURES = Path(__file__).parent / "fixtures"


@pytest.fixture
def project(tmp_path):
    """A fresh project folder with default setup and the fixture statements dropped in."""
    from finpipe.cli import main
    main(["--root", str(tmp_path), "--no-open"])  # first run: creates workbooks + folders
    shutil.copytree(FIXTURES, tmp_path / "statements", dirs_exist_ok=True)
    return tmp_path
