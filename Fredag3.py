"""
Startskript for Fredag-appen.

Denne fila kan kjøres direkte fra prosjektmappa:

    python Fredag3.py

For at importen `from fredag.helgesjekk_app import run_app` skal fungere
når skriptet ligger *inne i* pakken, sørger vi for at foreldremappa
ligger på sys.path (samme prinsipp som i tests/conftest.py).
"""

from __future__ import annotations

import sys
from pathlib import Path


def _ensure_parent_on_syspath() -> None:
    # Sti til denne fila: ...\fredag\Fredag3.py
    repo_root = Path(__file__).resolve().parent
    parent = repo_root.parent  # ...\ (mappen som inneholder "fredag")

    parent_str = str(parent)
    if parent_str not in sys.path:
        # Legg foreldremappen først, slik at `import fredag` peker på
        # riktig prosjektmappe.
        sys.path.insert(0, parent_str)


_ensure_parent_on_syspath()

from fredag.helgesjekk_app import run_app  # type: ignore  # noqa: E402


if __name__ == "__main__":
    run_app()
