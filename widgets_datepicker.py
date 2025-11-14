import calendar
import tkinter as tk
from tkinter import ttk
from datetime import date, datetime

__all__ = ["DatePicker"]

def _parse_ddmmyyyy(s: str) -> date | None:
    s = (s or "").strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

class DatePicker(tk.Toplevel):
    """Lettvekts kalender‑popup uten eksterne pakker. Skriver valgt dato til en tk.StringVar (dd.mm.yyyy)."""
    def __init__(self, master, target_var: tk.StringVar, init_text: str | None = None):
        super().__init__(master)
        self.title("Velg dato")
        self.resizable(False, False)
        self.transient(master)
        self.target_var = target_var

        init = _parse_ddmmyyyy(init_text or target_var.get()) or date.today()
        self._year, self._month = init.year, init.month

        frm = ttk.Frame(self, padding=8); frm.grid(row=0, column=0)

        nav = ttk.Frame(frm); nav.grid(row=0, column=0, sticky="ew")
        ttk.Button(nav, text="‹", width=3, command=self._prev_month).pack(side="left")
        self._title = ttk.Label(nav, text=self._title_text(), font=("TkDefaultFont", 10, "bold"))
        self._title.pack(side="left", padx=8)
        ttk.Button(nav, text="›", width=3, command=self._next_month).pack(side="left")

        self._grid = ttk.Frame(frm); self._grid.grid(row=1, column=0, pady=(6, 0))
        for i, h in enumerate(("Ma", "Ti", "On", "To", "Fr", "Lø", "Sø")):
            ttk.Label(self._grid, text=h, width=3, anchor="center").grid(row=0, column=i, padx=1, pady=1)

        btns = ttk.Frame(frm); btns.grid(row=2, column=0, pady=(6, 0), sticky="ew")
        ttk.Button(btns, text="I dag", command=self._today).pack(side="left")
        ttk.Button(btns, text="Lukk", command=self.destroy).pack(side="right")

        self._rebuild_days()
        try:
            x = self.winfo_pointerx() - 150
            y = self.winfo_pointery() - 20
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _title_text(self) -> str:
        return date(self._year, self._month, 1).strftime("%B %Y").capitalize()

    def _prev_month(self):
        self._month, self._year = (12, self._year - 1) if self._month == 1 else (self._month - 1, self._year)
        self._title.config(text=self._title_text()); self._rebuild_days()

    def _next_month(self):
        self._month, self._year = (1, self._year + 1) if self._month == 12 else (self._month + 1, self._year)
        self._title.config(text=self._title_text()); self._rebuild_days()

    def _select(self, d: int):
        self.target_var.set(f"{d:02d}.{self._month:02d}.{self._year:04d}"); self.destroy()

    def _today(self):
        t = date.today()
        self._year, self._month = t.year, t.month
        self._title.config(text=self._title_text()); self._rebuild_days()
        self.target_var.set(t.strftime("%d.%m.%Y"))

    def _rebuild_days(self):
        for child in list(self._grid.children.values()):
            if child.grid_info()["row"] != 0:
                child.destroy()
        cal = calendar.Calendar(firstweekday=0)
        row = 1
        for week in cal.monthdayscalendar(self._year, self._month):
            for col, d in enumerate(week):
                if d == 0:
                    ttk.Label(self._grid, text=" ", width=3).grid(row=row, column=col, padx=1, pady=1)
                else:
                    ttk.Button(self._grid, text=f"{d:02d}", width=3,
                               command=lambda dd=d: self._select(dd)).grid(row=row, column=col, padx=1, pady=1)
            row += 1
