from __future__ import annotations
import json, time
from pathlib import Path
from typing import Dict, Tuple

def _path() -> Path:
    root = Path(__file__).resolve().parents[1] / ".ragdb"
    root.mkdir(exist_ok=True)
    return root / "dedup_index.json"

def _now() -> float:
    return time.time()

def load_index() -> Dict[str, float]:
    p = _path()
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        items = data.get("items") if isinstance(data, dict) else data
        return dict(items or {})
    except Exception:
        return {}

def save_index(idx: Dict[str, float]) -> None:
    p = _path()
    payload = {"v": 1, "items": idx}
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)

def prune_expired(idx: Dict[str, float], ttl_days: int) -> int:
    if ttl_days <= 0:  # ikke utløp
        return 0
    cutoff = _now() - ttl_days * 24 * 3600
    keys = [k for k, ts in idx.items() if ts < cutoff]
    for k in keys:
        idx.pop(k, None)
    return len(keys)
