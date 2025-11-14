from __future__ import annotations
from typing import Dict, List, Tuple, Optional

from .group_rules import GroupRule, load_rules, resolve_group

def _iter_stores(session):
    stores = getattr(session, "Stores", None)
    if stores:
        for i in range(1, int(stores.Count) + 1):
            s = stores.Item(i)
            if s:
                try:
                    yield s.DisplayName, s.GetRootFolder()
                except Exception:
                    continue
    else:
        # Fallback: Session.Folders (eldre Outlook)
        roots = getattr(session, "Folders", None)
        if roots:
            for i in range(1, int(roots.Count) + 1):
                r = roots.Item(i)
                if r:
                    yield getattr(r, "Name", ""), r

def _find_child(parent, name: str):
    subs = getattr(parent, "Folders", None)
    if not subs: return None
    for i in range(1, int(subs.Count) + 1):
        f = subs.Item(i)
        if f and getattr(f, "Name", "").lower() == name.lower():
            return f
    return None

def get_folder_by_path(session, path: str):
    """
    path: r"\\Store DisplayName\\Delmappe\\Under"
    Hvis store‑navn utelates, brukes DefaultStore.
    """
    if not path or not path.strip():
        return None
    p = path.replace("/", "\\").strip()
    while p.startswith("\\"): p = p[1:]
    parts = [x for x in p.split("\\") if x.strip()]
    if not parts:
        return None

    # Finn rot
    root = None
    # Hvis første del matcher en store, bruk den – ellers bruk DefaultStore
    first = parts[0]
    for disp, rf in _iter_stores(session):
        if disp and disp.lower() == first.lower():
            root = rf; parts = parts[1:]; break
    if root is None:
        try:
            root = session.DefaultStore.GetRootFolder()
        except Exception:
            return None

    cur = root
    for name in parts:
        nxt = _find_child(cur, name)
        if nxt is None:
            return None
        cur = nxt
    return cur

def move_by_groups(session,
                   results: List[Dict],
                   rules: Optional[List[GroupRule]] = None,
                   dry_run: bool = False) -> Tuple[Dict[str, Dict[str, int]], List[Dict], List[str]]:
    """
    Flytter meldinger til mappe fra rule.move_to_folder_path.
    Returnerer (summary, unassigned_rows, grupper_uten_dest)
    summary[gname] = {"moved": x, "skipped": y, "errors": z}
    """
    rules = rules or load_rules()
    get_item = lambda r: session.GetItemFromID(r.get("eid"), r.get("store")) if r.get("eid") else None

    # Bucket per gruppe
    buckets: Dict[str, List[Dict]] = {}
    mapping: Dict[str, GroupRule] = {}
    unassigned: List[Dict] = []
    no_dest_groups: List[str] = []

    for r in results:
        smtp = (r.get("from_email") or "").lower()
        name = r.get("from") or ""
        g = resolve_group(rules, smtp, name)
        if not g:
            unassigned.append(r); continue
        if not g.move_to_folder_path:
            if g.name not in no_dest_groups:
                no_dest_groups.append(g.name)
            continue
        buckets.setdefault(g.name, []).append(r)
        mapping[g.name] = g

    summary: Dict[str, Dict[str, int]] = {}
    # Cache mappeobjekter
    dest_cache: Dict[str, object] = {}

    for gname, rows in buckets.items():
        rule = mapping[gname]
        dest_path = rule.move_to_folder_path
        dest = dest_cache.get(dest_path)
        if not dest:
            dest = get_folder_by_path(session, dest_path)
            dest_cache[dest_path] = dest

        moved = skipped = errors = 0
        if not dest:
            # kan ikke flytte – manglende sti
            errors = len(rows)
            summary[gname] = {"moved": 0, "skipped": 0, "errors": errors}
            continue

        for r in rows:
            it = get_item(r)
            if not it:
                errors += 1; continue
            try:
                if dry_run:
                    moved += 1
                else:
                    if rule.move_mark_read:
                        try:
                            if bool(getattr(it, "UnRead", False)):
                                it.UnRead = False; it.Save()
                        except Exception:
                            pass
                    it.Move(dest)
                    moved += 1
            except Exception:
                errors += 1
        summary[gname] = {"moved": moved, "skipped": skipped, "errors": errors}

    return summary, unassigned, no_dest_groups
