from __future__ import annotations
import os
import platform
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from .log_utils import log_path

class DiagnoseWindow(tk.Toplevel):
    """
    En enkel diagnosevisning:
      - Miljøinfo (OS, Python)
      - Outlook DefaultStore/Innboks om session finnes
      - Sti til loggmappe + knapp for å åpne den
    Forventes kalt fra et vindu som har attributtet 'session' (Outlook.Session).
    """
    def __init__(self, master):
        super().__init__(master)
        self.title("Diagnose")
        self.geometry("720x420")
        self.session = getattr(master, "session", None)
        self._build()
        self._fill()

    def _build(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        self.txt = tk.Text(frm, wrap="word")
        self.txt.pack(fill="both", expand=True)

        btns = ttk.Frame(self, padding=(0,8))
        btns.pack(fill="x")
        ttk.Button(btns, text="Åpne loggmappe", command=self._open_logs).pack(side="left")
        ttk.Button(btns, text="Lukk", command=self.destroy).pack(side="right")

    def _open_logs(self):
        try:
            os.startfile(str(log_path().parent))
        except Exception as e:
            messagebox.showerror("Loggmappe", f"Klarte ikke å åpne loggmappe:\n{e}")

    def _fill(self):
        lines = []
        lines.append(f"Tid:        {datetime.now():%Y-%m-%d %H:%M:%S}")
        lines.append(f"Python:     {sys.version.split()[0]} ({sys.executable})")
        lines.append(f"Plattform:  {platform.platform()}")
        lines.append(f"Pakke:      fredag")
        lines.append(f"Loggmappe:  {log_path().parent}")
        lines.append("")

        if self.session:
            try:
                ds = self.session.DefaultStore
                inbox = ds.GetDefaultFolder(6)  # 6 = olFolderInbox
                stores = getattr(self.session, "Stores", None)
                lines.append("Outlook")
                lines.append(f"  DefaultStore: {getattr(ds, 'DisplayName', '?')}")
                lines.append(f"  Innboks:      {getattr(inbox, 'FolderPath', '?')}")
                lines.append(f"  #Stores:      {int(stores.Count) if stores else 'ukjent'}")
            except Exception as e:
                lines.append(f"(Outlook info ikke tilgjengelig: {e})")
        else:
            lines.append("(Ingen Outlook‑session tilgjengelig i denne konteksten.)")

        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(lines))
        self.txt.configure(state="disabled")
