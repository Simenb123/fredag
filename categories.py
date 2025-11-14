from __future__ import annotations
from typing import Optional, List

# Outlook fargekoder (vanlige) – OlCategoryColor enum
# Kildene varierer mellom versjoner; disse verdiene fungerer på støttede Outlook‑versjoner.
_COLOR_MAP = {
    "none": 0,
    "red": 1,
    "orange": 2,
    "yellow": 4,
    "green": 5,
    "teal": 6,
    "blue": 8,
    "purple": 9,
    "maroon": 10,
    "steel": 11,
    "gray": 13,
    "darkgray": 14,
    "black": 15,
}

def list_category_names(session) -> List[str]:
    try:
        cats = session.Categories
        return [c.Name for c in cats] if cats else []
    except Exception:
        return []

def ensure_category(session, name: str, color_name: Optional[str] = None) -> bool:
    """
    Sørger for at en kategori med 'name' finnes. Lager den ved behov (med farge om mulig).
    Returnerer True hvis finnes/ble opprettet, ellers False.
    """
    if not name:
        return False
    try:
        cats = session.Categories
        # Finn eksisterende
        for i in range(1, int(getattr(cats, "Count", 0)) + 1):
            c = cats.Item(i)
            if c and getattr(c, "Name", "") == name:
                return True
        # Opprett ny
        color = _COLOR_MAP.get((color_name or "").strip().lower(), _COLOR_MAP["none"])
        cats.Add(name, color)
        return True
    except Exception:
        # Hvis Add ikke er tilgjengelig i denne versjonen, bare ignorer – vi kan fortsatt sette tekstkategori.
        return False
