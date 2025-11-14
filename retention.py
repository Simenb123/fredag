from __future__ import annotations
import os
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Tuple

from .group_rules import GroupRule, load_rules

def _iter_files(root: Path) -> Path:
    for p in root.rglob("*"):
        if p.is_file():
            yield p

def _prune_empty_dirs(root: Path, keep: Path):
    # Fjern tomme kataloger under root (men ikke rotmappen 'keep')
    for p in sorted(root.rglob("*"), reverse=True):
        try:
            if p.is_dir() and not any(p.iterdir()) and p != keep:
                p.rmdir()
        except Exception:
            pass

def apply_retention(rules: List[GroupRule], dry_run: bool = False) -> Dict[str, Dict[str, int]]:
    """
    Sletter filer eldre enn 'retention_days' for hver gruppe-mappe (0=behold).
    Returnerer summary per gruppe: {"deleted": x, "kept": y, "errors": z}
    """
    now = datetime.now()
    summary: Dict[str, Dict[str, int]] = {}
    for r in rules:
        days = int(r.retention_days or 0)
        if days <= 0:
            continue
        root = Path(r.target_dir)
        if not root.exists():
            continue
        deleted = kept = errors = 0
        threshold = now - timedelta(days=days)
        for f in _iter_files(root):
            try:
                mtime = datetime.fromtimestamp(f.stat().st_mtime)
                if mtime < threshold:
                    if dry_run:
                        deleted += 1   # ville slettet
                    else:
                        f.unlink(missing_ok=True)
                        deleted += 1
                else:
                    kept += 1
            except Exception:
                errors += 1
        if not dry_run:
            _prune_empty_dirs(root, keep=root)
        summary[r.name] = {"deleted": deleted, "kept": kept, "errors": errors}
    return summary
