from __future__ import annotations
import json
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Tuple

from .group_rules import default_rules_path
from .settings import load_settings

def _store_dir() -> Path:
    root = Path(__file__).resolve().parents[1] / ".ragdb"
    root.mkdir(exist_ok=True)
    return root

def export_config(zip_path: str) -> Tuple[bool, str]:
    """
    Pakker grupper.json og settings.json i en ZIP.
    """
    try:
        zp = Path(zip_path)
        zp.parent.mkdir(parents=True, exist_ok=True)
        rules = default_rules_path()
        settings = _store_dir() / "settings.json"
        with zipfile.ZipFile(zp, "w", compression=zipfile.ZIP_DEFLATED) as z:
            if rules.exists():    z.write(rules, arcname="grupper.json")
            if settings.exists(): z.write(settings, arcname="settings.json")
            # liten manifest
            manifest = {
                "exported_at": datetime.now().isoformat(),
                "contains": ["grupper.json" if rules.exists() else None,
                             "settings.json" if settings.exists() else None]
            }
            z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        return True, f"Lagret: {zp}"
    except Exception as e:
        return False, f"Feil under eksport: {e}"

def import_config(zip_path: str, backup_current: bool = True) -> Tuple[bool, str]:
    """
    Leser ZIP og legger tilbake grupper.json/settings.json. Tar backup av nåværende filer om ønskelig.
    """
    try:
        zp = Path(zip_path)
        if not zp.exists():
            return False, "Finner ikke ZIP‑filen."
        target_dir = _store_dir()
        rules_p = default_rules_path()
        settings_p = target_dir / "settings.json"

        # Backup
        if backup_current:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            if rules_p.exists():
                rules_p.rename(rules_p.with_name(f"grupper.{ts}.bak.json"))
            if settings_p.exists():
                settings_p.rename(settings_p.with_name(f"settings.{ts}.bak.json"))

        with zipfile.ZipFile(zp, "r") as z:
            for name in ("grupper.json", "settings.json"):
                if name in z.namelist():
                    tmp = target_dir / (name + ".tmp")
                    tmp.write_bytes(z.read(name))
                    tmp.replace(target_dir / name)

        return True, "Konfig importert."
    except Exception as e:
        return False, f"Feil under import: {e}"
