import os, time
from pathlib import Path
from fredag.retention import apply_retention
from fredag.group_rules import GroupRule

def test_retention_dryrun(tmp_path: Path):
    # lag to filer: én gammel, én ny
    old = tmp_path / "2024/01_Jan"; old.mkdir(parents=True, exist_ok=True)
    f_old = old / "a.txt"; f_old.write_text("x")
    # sett mtime til 60 dager siden
    past = time.time() - 60*24*3600
    os.utime(f_old, (past, past))

    new = tmp_path / "2025/10_Okt"; new.mkdir(parents=True, exist_ok=True)
    f_new = new / "b.txt"; f_new.write_text("y")

    rules = [GroupRule(name="Test", target_dir=str(tmp_path), senders=[], retention_days=30)]
    summary = apply_retention(rules, dry_run=True)
    assert summary["Test"]["deleted"] == 1
    assert summary["Test"]["kept"] >= 1  # den nye filen
