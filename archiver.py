from __future__ import annotations
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional

from .path_template import month_abbr as _mabbr, safe_component, extract_subject_tag, render_template, domain_from_email
from .categories import ensure_category
from .dedup_index import load_index, save_index, prune_expired

def _temp_dir() -> Path:
    p = Path(__file__).resolve().parents[1] / ".ragdb" / "tmp"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _hash_file(p: Path) -> str:
    h = hashlib.sha1()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

def _attach_iter(item) -> List:
    atts = getattr(item, "Attachments", None)
    if not atts: return []
    return [atts.Item(i) for i in range(1, int(getattr(atts, "Count", 0)) + 1)]

def _attachment_allowed(att, allowed_exts: List[str], min_kb: int, max_kb: int) -> bool:
    if allowed_exts:
        ext = Path((getattr(att, "FileName", "") or "")).suffix.lower().lstrip(".")
        if ext not in allowed_exts:
            return False
    size = int(getattr(att, "Size", 0) or 0)  # bytes
    if min_kb and size < min_kb * 1024: return False
    if max_kb and size > max_kb * 1024: return False
    return True

def _build_target(root: Path, item, r: Dict, per_sender: bool,
                  template: Optional[str], subject_regex: Optional[str]) -> Path:
    dt = getattr(item, "ReceivedTime", None) or r.get("dt") or datetime.now()
    y = int(getattr(dt, "year", datetime.now().year))
    m = int(getattr(dt, "month", datetime.now().month))
    sender = (r.get("from_email") or r.get("from") or "ukjent_avsender")
    domain = domain_from_email(r.get("from_email") or "")
    sender_safe = safe_component(sender); domain_safe = safe_component(domain)
    subj = r.get("subject") or ""
    tag = extract_subject_tag(subj, subject_regex or "")
    if template:
        meta = {"year": f"{y}", "month2": f"{m:02d}", "month_abbr": _mabbr(m),
                "sender": sender_safe, "domain": domain_safe, "subject_tag": tag}
        rel = render_template(template, meta)
        base = (root / rel)
    else:
        base = root / f"{y}" / f"{m:02d}_{_mabbr(m)}"
        if per_sender:
            base = base / sender_safe
    base.mkdir(parents=True, exist_ok=True)
    return base

def archive_messages(session,
                     results: List[Dict],
                     get_item,
                     root_dir: str,
                     per_sender: bool = False,
                     dedup: bool = True,
                     filters: Optional[Dict] = None,
                     set_category: Optional[str] = None,
                     dry_run: bool = False,
                     template: Optional[str] = None,
                     subject_regex: Optional[str] = None,
                     set_category_color: Optional[str] = None,
                     persist_index: bool = False,
                     index_ttl_days: int = 365) -> Tuple[int, int, str]:
    """
    Arkiverer vedlegg for 'results'
    - filters: {"exts":[...], "min_kb":int, "max_kb":int}
    - set_category(+_color): kategori opprettes ved behov og settes hvis minst ett vedlegg lagres
    - template/subject_regex: sti‑mal + emne‑tag
    - persist_index: vedvarende dedup mot global hash‑indeks (TTL i dager)
    - dry_run: simuler lagring
    Returnerer (saved_count, skipped_count, err_msg)
    """
    root = Path(root_dir); root.mkdir(parents=True, exist_ok=True)
    allowed_exts = [e.lower() for e in (filters or {}).get("exts", []) if e]
    min_kb = int((filters or {}).get("min_kb") or 0)
    max_kb = int((filters or {}).get("max_kb") or 0)

    if set_category and not dry_run:
        try: ensure_category(session, set_category, set_category_color or None)
        except Exception: pass

    saved = skipped = 0
    errors: List[str] = []
    seen_hashes = set()
    tmp_root = _temp_dir()

    # Vedvarende dedup
    idx = {}
    if persist_index:
        idx = load_index()
        try: prune_expired(idx, int(index_ttl_days or 0))
        except Exception: pass

    for r in results:
        it = get_item(r)
        if not it: continue
        base = _build_target(root, it, r, per_sender, template, subject_regex)
        any_saved_here = False

        for att in _attach_iter(it):
            tmp_path = None
            try:
                if not _attachment_allowed(att, allowed_exts, min_kb, max_kb):
                    skipped += 1; continue

                fname = safe_component(getattr(att, "FileName", "") or "vedlegg")
                tmp_path = tmp_root / fname
                att.SaveAsFile(str(tmp_path))

                h = _hash_file(tmp_path)

                # persist dedup først, deretter run‑scope dedup
                if persist_index and h in idx:
                    skipped += 1; continue
                if dedup and h in seen_hashes:
                    skipped += 1; continue

                if dry_run:
                    saved += 1; any_saved_here = True; seen_hashes.add(h)
                else:
                    dest = base / fname
                    if dest.exists():
                        dest = dest.with_name(f"{dest.stem}__{h[:8]}{dest.suffix}")
                    tmp_path.replace(dest)
                    saved += 1; any_saved_here = True; seen_hashes.add(h)
                    if persist_index:
                        idx[h] = datetime.now().timestamp()

            except Exception as e:
                errors.append(str(e))
            finally:
                if tmp_path and tmp_path.exists():
                    try: tmp_path.unlink(missing_ok=True)
                    except Exception: pass

        if set_category and any_saved_here and not dry_run:
            try:
                cats = getattr(it, "Categories", "") or ""
                wanted = set_category.strip()
                parts = [c.strip() for c in cats.split(";") if c.strip()]
                if wanted not in parts:
                    parts.append(wanted); it.Categories = "; ".join(parts); it.Save()
            except Exception:
                pass

    if persist_index and not dry_run:
        try: save_index(idx)
        except Exception: pass

    return saved, skipped, "; ".join(errors)
