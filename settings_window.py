from __future__ import annotations
import os
import tkinter as tk
from tkinter import ttk, messagebox
from .settings import load_settings, save_settings, _DEFAULTS, settings_path

_COLOR_CHOICES = ["", "red","orange","yellow","green","teal","blue","purple","maroon","gray","black"]

class SettingsWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Innstillinger (globale standarder)")
        self.geometry("680x520")
        self._build()
        self._load()

    def _build(self):
        frm = ttk.Frame(self, padding=10); frm.pack(fill="both", expand=True)

        # Filtre
        f1 = ttk.LabelFrame(frm, text="Standard filter (brukes når gruppe mangler egne verdier)")
        f1.pack(fill="x", pady=(0,8))
        ttk.Label(f1, text="Tillatte endelser (komma):").grid(row=0, column=0, sticky="w")
        self.v_exts = tk.StringVar(); ttk.Entry(f1, textvariable=self.v_exts, width=48).grid(row=0, column=1, sticky="w", padx=(6,0))
        ttk.Label(f1, text="Min KB:").grid(row=0, column=2, sticky="e", padx=(12,0))
        self.v_min = tk.StringVar(); ttk.Entry(f1, textvariable=self.v_min, width=8).grid(row=0, column=3, sticky="w")
        ttk.Label(f1, text="Max KB:").grid(row=0, column=4, sticky="e", padx=(12,0))
        self.v_max = tk.StringVar(); ttk.Entry(f1, textvariable=self.v_max, width=8).grid(row=0, column=5, sticky="w")

        # Kategori
        f2 = ttk.LabelFrame(frm, text="Standard kategori")
        f2.pack(fill="x", pady=(0,8))
        ttk.Label(f2, text="Kategori-navn:").grid(row=0, column=0, sticky="w")
        self.v_cat = tk.StringVar(); ttk.Entry(f2, textvariable=self.v_cat, width=30).grid(row=0, column=1, sticky="w", padx=(6,0))
        ttk.Label(f2, text="Farge:").grid(row=0, column=2, sticky="e", padx=(12,0))
        self.v_color = tk.StringVar(value="")
        ttk.Combobox(f2, values=_COLOR_CHOICES, textvariable=self.v_color, width=12, state="readonly").grid(row=0, column=3, sticky="w")

        # Sti-mal / regex
        f3 = ttk.LabelFrame(frm, text="Standard mål‑mal og emne‑regex")
        f3.pack(fill="x", pady=(0,8))
        ttk.Label(f3, text="Mål‑mal:").grid(row=0, column=0, sticky="w")
        self.v_tpl = tk.StringVar(); ttk.Entry(f3, textvariable=self.v_tpl, width=58).grid(row=0, column=1, columnspan=3, sticky="w", padx=(6,0))
        ttk.Label(f3, text="Emne‑regex:").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.v_rx = tk.StringVar(); ttk.Entry(f3, textvariable=self.v_rx, width=58).grid(row=1, column=1, columnspan=3, sticky="w", padx=(6,0), pady=(6,0))
        ttk.Label(f3, text="Plassholdere: {year} {month2} {month_abbr} {sender} {domain} {subject_tag}").grid(row=2, column=0, columnspan=4, sticky="w", pady=(6,4))

        # Retention default
        f4 = ttk.LabelFrame(frm, text="Retention")
        f4.pack(fill="x")
        ttk.Label(f4, text="Standard retention (dager, 0=behold):").grid(row=0, column=0, sticky="w")
        self.v_ret = tk.StringVar(); ttk.Entry(f4, textvariable=self.v_ret, width=10).grid(row=0, column=1, sticky="w", padx=(6,0))

        # Knapperekke
        b = ttk.Frame(frm); b.pack(fill="x", pady=(8,0))
        ttk.Button(b, text="Lagre", command=self._save).pack(side="left")
        ttk.Button(b, text="Tilbakestill til standard", command=self._reset).pack(side="left", padx=(6,0))
        ttk.Button(b, text="Åpne konfig‑mappe", command=lambda: os.startfile(str(settings_path().parent))).pack(side="right")

        self.status = ttk.Label(frm, text=f"Fil: {settings_path()}")
        self.status.pack(fill="x", pady=(8,0))

    def _load(self):
        s = load_settings()
        self.v_exts.set(",".join(s.get("default_allowed_exts") or []))
        self.v_min.set(str(s.get("default_min_kb", 0)))
        self.v_max.set(str(s.get("default_max_kb", 0)))
        self.v_cat.set(s.get("default_category", ""))
        self.v_color.set(s.get("default_category_color", ""))
        self.v_tpl.set(s.get("default_target_template", ""))
        self.v_rx.set(s.get("default_subject_tag_regex", ""))
        self.v_ret.set(str(s.get("retention_default_days", 0)))

    def _reset(self):
        ok, msg = save_settings(_DEFAULTS)
        if ok:
            self._load()
        tk.messagebox.showinfo("Innstillinger", msg)

    def _save(self):
        try:
            min_kb = int(self.v_min.get() or "0")
            max_kb = int(self.v_max.get() or "0")
            ret    = int(self.v_ret.get() or "0")
        except ValueError:
            messagebox.showerror("Innstillinger", "Min/Max KB og Retention må være tall.")
            return
        exts = [e.strip().lower().lstrip(".") for e in (self.v_exts.get() or "").split(",") if e.strip()]
        payload = {
            "default_allowed_exts": exts,
            "default_min_kb": min_kb,
            "default_max_kb": max_kb,
            "default_category": self.v_cat.get().strip(),
            "default_category_color": self.v_color.get().strip(),
            "default_target_template": self.v_tpl.get().strip(),
            "default_subject_tag_regex": self.v_rx.get().strip(),
            "retention_default_days": ret,
            # behold også caps/limits hvis de finnes fra før
            "cap_per_folder": load_settings().get("cap_per_folder", 6000),
            "cap_total": load_settings().get("cap_total", 4000),
        }
        ok, msg = save_settings(payload)
        messagebox.showinfo("Innstillinger", msg)
