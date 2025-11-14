from __future__ import annotations
from collections import defaultdict
from typing import Dict, List, Tuple, Optional

from .group_rules import GroupRule, load_rules, resolve_group
from .archiver import archive_messages
from .state_store import was_archived, mark_archived
from .settings import load_settings

Summary = Dict[str, Dict[str, int]]

def _get_item_fn(session):
    def _get(r):
        try:
            return session.GetItemFromID(r.get("eid"), r.get("store"))
        except Exception:
            return None
    return _get

def archive_by_groups(session,
                      results: List[Dict],
                      rules: Optional[List[GroupRule]] = None,
                      dedup: bool = True,
                      dry_run: bool = False) -> Tuple[Summary, List[Dict]]:
    rules = rules or load_rules()
    defaults = load_settings()

    summary: Summary = {}
    buckets: Dict[str, List[Dict]] = defaultdict(list)
    mapping: Dict[str, GroupRule] = {}

    unassigned: List[Dict] = []
    for r in results:
        eid = r.get("eid") or ""
        if not dry_run and (not eid or was_archived(eid)):
            continue
        smtp = (r.get("from_email") or "").lower()
        name = r.get("from") or ""
        g = resolve_group(rules, smtp, name)
        if not g:
            unassigned.append(r); continue
        buckets[g.name].append(r); mapping[g.name] = g

    get_item = _get_item_fn(session)
    for gname, rows in buckets.items():
        rule = mapping[gname]

        exts = rule.allowed_exts or (defaults.get("default_allowed_exts") or [])
        min_kb = int(rule.min_kb or defaults.get("default_min_kb", 0))
        max_kb = int(rule.max_kb or defaults.get("default_max_kb", 0))
        filters = {"exts": [e.lower() for e in exts], "min_kb": min_kb, "max_kb": max_kb}

        category = rule.category or (defaults.get("default_category") or "")
        category_color = rule.category_color or (defaults.get("default_category_color") or "")
        template = rule.target_template or (defaults.get("default_target_template") or "")
        subj_rx  = rule.subject_tag_regex or (defaults.get("default_subject_tag_regex") or "")

        saved, skipped, err = archive_messages(
            session=session, results=rows, get_item=get_item, root_dir=rule.target_dir,
            per_sender=False, dedup=dedup, filters=filters,
            set_category=(category or None), set_category_color=(category_color or None),
            dry_run=dry_run, template=(template or None), subject_regex=(subj_rx or None),
            persist_index=bool(defaults.get("dedup_persist", True)),
            index_ttl_days=int(defaults.get("dedup_ttl_days", 365))
        )
        if not dry_run:
            for r in rows:
                if r.get("eid"): mark_archived(r["eid"])
        summary[gname] = {"saved": saved, "skipped": skipped, "msgs": len(rows)}

    return summary, unassigned
