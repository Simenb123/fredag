from __future__ import annotations
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional
import json
import fnmatch

def _base_dir() -> Path:
    root = Path(__file__).resolve().parents[1] / ".ragdb"
    root.mkdir(exist_ok=True)
    return root

def default_rules_path() -> Path:
    return _base_dir() / "grupper.json"

@dataclass
class GroupRule:
    name: str
    target_dir: str
    senders: List[str]              # epost, @domene.no, wildcard (*.no)
    note: str = ""
    allowed_exts: List[str] = None  # f.eks. ["pdf","xlsx"]; None/[] = alle
    min_kb: int = 0
    max_kb: int = 0
    category: str = ""
    category_color: str = ""        # "blue", "green", ...
    retention_days: int = 0
    target_template: str = ""       # "{year}/{month2}_{month_abbr}/{domain}/{subject_tag}"
    subject_tag_regex: str = ""     # r"(PRJ-\d+)"
    # Flytt-regler (Outlook-mappe)
    move_to_folder_path: str = ""   # f.eks. r"\\Mailbox - Ola Nordmann\\Arkiv\\KundeX"
    move_mark_read: bool = False

def _norm_exts(exts: Optional[List[str]]) -> List[str]:
    if not exts: return []
    out = []
    for e in exts:
        e = (e or "").strip().lower()
        if not e: continue
        if e.startswith("."): e = e[1:]
        out.append(e)
    return out

def load_rules(path: Optional[Path] = None) -> List[GroupRule]:
    p = path or default_rules_path()
    if not p.exists(): return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        rules = []
        for r in data.get("groups", []):
            rules.append(GroupRule(
                name=r.get("name","").strip(),
                target_dir=r.get("target_dir","").strip(),
                senders=[s.strip().lower() for s in (r.get("senders") or []) if s.strip()],
                note=(r.get("note") or "").strip(),
                allowed_exts=_norm_exts(r.get("allowed_exts")),
                min_kb=int(r.get("min_kb") or 0),
                max_kb=int(r.get("max_kb") or 0),
                category=(r.get("category") or "").strip(),
                category_color=(r.get("category_color") or "").strip(),
                retention_days=int(r.get("retention_days") or 0),
                target_template=(r.get("target_template") or "").strip(),
                subject_tag_regex=(r.get("subject_tag_regex") or "").strip(),
                move_to_folder_path=(r.get("move_to_folder_path") or "").strip(),
                move_mark_read=bool(r.get("move_mark_read") or False),
            ))
        return rules
    except Exception:
        return []

def save_rules(rules: List[GroupRule], path: Optional[Path] = None) -> None:
    p = path or default_rules_path()
    payload = {"version": 6, "groups": [asdict(r) for r in rules]}
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)

def _match_sender(pat: str, smtp: str, name: str) -> bool:
    pat = pat.lower().strip()
    smtp = (smtp or "").lower()
    name = (name or "").lower()
    if not pat: return False
    if pat.startswith("@"):  # domenematch
        return smtp.endswith(pat)
    if any(ch in pat for ch in "*?[]"):  # wildcard
        return fnmatch.fnmatch(smtp, pat) or fnmatch.fnmatch(name, pat)
    return smtp == pat or (pat in name)

def resolve_group(rules: List[GroupRule], smtp: str, name: str) -> Optional[GroupRule]:
    for r in rules:
        for pat in r.senders:
            if _match_sender(pat, smtp, name):
                return r
    return None
