"""
Pytest-konfigurasjon for prosjektet.

Denne filen gjør én viktig ting:
    - Sørger for at foreldremappen til prosjektet (f.eks. F:\\Dokument2\\python)
      legges på sys.path før testene kjøres.

Da vil `import fredag` fungere i testene, og du kan kjøre `pytest`
fra prosjektroten (F:\\Dokument2\\python\\fredag) uten ekstra argumenter.
"""

from __future__ import annotations

import sys
from pathlib import Path


def _ensure_parent_on_syspath() -> None:
    # Sti til denne fila: ...\fredag\conftest.py
    repo_root = Path(__file__).resolve().parent
    parent = repo_root.parent  # ...\ (mappen som inneholder "fredag")

    parent_str = str(parent)
    if parent_str not in sys.path:
        # Legg foreldremappen først, slik at `import fredag` peker på
        # mappen du har åpnet prosjektet fra.
        sys.path.insert(0, parent_str)


_ensure_parent_on_syspath()
