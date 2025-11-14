from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List, Tuple
from pathlib import Path

from .group_rules import GroupRule, save_rules
from .path_template import domain_from_email

@dataclass
class DomainSuggestion:
    domain: str
    count: int
    examples: List[str]

def summarize_unassigned(rows: List[Dict]) -> List[DomainSuggestion]:
    agg: Dict[str, Dict[str, any]] = {}
    for r in rows:
        d = domain_from_email(r.get("from_email") or "")
        if not d:
            d = "(ukjent)"
        x = agg.setdefault(d, {"count": 0, "ex": []})
        x["count"] += 1
        f = r.get("from") or r.get("from_email") or ""
        if len(x["ex"]) < 5 and f not in x["ex"]:
            x["ex"].append(f)
    out = [DomainSuggestion(k, v["count"], v["ex"]) for k, v in agg.items() if k and k != ""]
    out.sort(key=lambda s: s.count, reverse=True)
    return out

def apply_create_groups(existing: List[GroupRule], selected_domains: List[str], base_dir: str) -> List[GroupRule]:
    base = Path(base_dir)
    groups = list(existing)
    existing_names = {g.name.lower() for g in groups}
    for dom in selected_domains:
        name = dom
        i = 2
        while name.lower() in existing_names:
            name = f"{dom} ({i})"; i += 1
        g = GroupRule(
            name=name,
            target_dir=str(base / dom),
            senders=[f"@{dom}"],
        )
        groups.append(g); existing_names.add(name.lower())
    save_rules(groups)
    return groups

def apply_add_to_group(existing: List[GroupRule], selected_domains: List[str], group_name: str) -> List[GroupRule]:
    groups = list(existing)
    target = None
    for g in groups:
        if g.name == group_name:
            target = g; break
    if not target:
        return groups
    have = {s.lower() for s in target.senders}
    for dom in selected_domains:
        pat = f"@{dom}".lower()
        if pat not in have:
            target.senders.append(pat); have.add(pat)
    save_rules(groups)
    return groups
