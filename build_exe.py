"""
Bygg frittstående .exe av Fredag-appen med PyInstaller.

Bruk:
    1) Aktiver venv
    2) python build_exe.py

Krever at PyInstaller er installert i miljøet:
    pip install pyinstaller
"""

from __future__ import annotations

from pathlib import Path

import PyInstaller.__main__


def main() -> None:
    project_root = Path(__file__).resolve().parent
    entry_script = project_root / "Fredag3.py"

    PyInstaller.__main__.run(
        [
            str(entry_script),
            "--noconfirm",
            "--clean",
            "--windowed",          # ingen console-vindu
            "--onefile",           # én exe-fil
            "--name=Fredag3",      # navn på exe-en
            "--collect-all", "openpyxl",
            "--collect-all", "PIL",         # Pillow (kan fjernes hvis ikke brukt)
            "--collect-submodules", "win32com",
        ]
    )

    dist_path = project_root / "dist" / "Fredag3.exe"
    if dist_path.exists():
        print("\n===========================================")
        print(f"Bygg ferdig! 🎉  Eksporter: {dist_path}")
        print("Du kan nå dobbeltklikke Fredag3.exe for å starte appen.")
        print("===========================================\n")
    else:
        print("\nADVARSEL: Fant ikke dist/Fredag3.exe etter bygg.")
        print("Sjekk PyInstaller-loggen for ERROR/FATAL.\n")


if __name__ == "__main__":
    main()
