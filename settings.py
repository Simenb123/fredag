from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict, Tuple

_DEFAULTS: Dict[str, Any] = {
    # Søk/ytelse
    "cap_per_folder": 6000,
    "cap_total": 4000,

    # Globale standarder (brukes når gruppefelt mangler)
    "default_allowed_exts": [],        # [] = alle filtyper
    "default_min_kb": 0,               # 0 = ingen min-grense
    "default_max_kb": 0,               # 0 = ingen max-grense
    "default_category": "",            # f.eks. "Arkivert"
    "default_category_color": "",      # f.eks. "blue", "green", "red"
    "default_target_template": "",     # f.eks. "{year}/{month2}_{month_abbr}/{domain}/{subject_tag}"
    "default_subject_tag_regex": "",   # f.eks. r"(PRJ-\d{4,6})"

    # Retention
    "retention_default_days": 0,       # 0 = behold

    # Vedvarende dedup (vedleggs‑hash på tvers av kjøringer)
    "dedup_persist": True,
    "dedup_ttl_days": 365,
}

def _store_dir() -> Path:
    root = Path(__file__).resolve().parents[1] / ".ragdb"
    root.mkdir(exist_ok=True)
    return root

def settings_path() -> Path:
    return _store_dir() / "settings.json"

def load_settings() -> Dict[str, Any]:
    p = settings_path()
    if not p.exists():
        return dict(_DEFAULTS)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        out = dict(_DEFAULTS)
        for k, v in (data or {}).items():
            if k in _DEFAULTS:
                out[k] = v
        return out
    except Exception:
        return dict(_DEFAULTS)

def save_settings(new_values: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        out = dict(_DEFAULTS)
        for k, v in (new_values or {}).items():
            if k in _DEFAULTS:
                out[k] = v
        tmp = settings_path().with_suffix(".tmp")
        tmp.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(settings_path())
        return True, "Innstillinger lagret."
    except Exception as e:
        return False, f"Feil ved lagring: {e}"

def update_settings(partial: Dict[str, Any]) -> Tuple[bool, str]:
    s = load_settings()
    for k, v in (partial or {}).items():
        if k in _DEFAULTS:
            s[k] = v
    return save_settings(s)

def get(key: str, default: Any = None) -> Any:
    return load_settings().get(key, default)
