from __future__ import annotations
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .group_rules import load_rules
from .group_suggest import summarize_unassigned, apply_create_groups, apply_add_to_group

class SuggestGroupsWindow(tk.Toplevel):
    def __init__(self, master, unassigned_rows: list[dict]):
        super().__init__(master)
        self.title("Foreslå grupper – u‑matchede domener")
        self.geometry("820x560")
        self._rows = unassigned_rows
        self._rules = load_rules()
        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="both", expand=True)
        left = ttk.Frame(top); left.pack(side="left", fill="both", expand=True)

        cols = ("domene","antall")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="extended", height=18)
        self.tree.heading("domene", text="Domene"); self.tree.column("domene", width=480, anchor="w")
        self.tree.heading("antall", text="Antall"); self.tree.column("antall", width=100, anchor="e")
        self.tree.pack(fill="both", expand=True)

        right = ttk.Frame(top, padding=(12,0)); right.pack(side="left", fill="y")
        ttk.Label(right, text="Eksempler:").pack(anchor="w")
        self.lb = tk.Listbox(right, width=32, height=12); self.lb.pack(fill="y", expand=False)

        btns = ttk.Frame(right); btns.pack(fill="x", pady=(10,0))
        ttk.Button(btns, text="Lag grupper…", command=self._create_groups).pack(fill="x")
        ttk.Button(btns, text="Legg i eksisterende…", command=self._add_to_existing).pack(fill="x", pady=(6,0))
        ttk.Button(btns, text="Eksporter CSV…", command=self._export_csv).pack(fill="x", pady=(12,0))

        under = ttk.Frame(self, padding=(10,6)); under.pack(fill="x")
        ttk.Button(under, text="Lukk", command=self.destroy).pack(side="right")
        self._load_data()
        self.tree.bind("<<TreeviewSelect>>", self._on_sel)

    def _load_data(self):
        self.tree.delete(*self.tree.get_children())
        self.suggestions = summarize_unassigned(self._rows)
        for s in self.suggestions:
            self.tree.insert("", "end", values=(s.domain, s.count))

    def _sel_domains(self) -> list[str]:
        out = []
        for iid in self.tree.selection():
            dom = self.tree.item(iid, "values")[0]
            out.append(dom)
        return out

    def _on_sel(self, _):
        self.lb.delete(0, "end")
        sels = self._sel_domains()
        if not sels: return
        dom = sels[0]
        for s in self.suggestions:
            if s.domain == dom:
                for ex in s.examples:
                    self.lb.insert("end", ex)
                break

    def _create_groups(self):
        sels = self._sel_domains()
        if not sels:
            messagebox.showinfo("Foreslå grupper", "Velg ett eller flere domener."); return
        base = filedialog.askdirectory(title="Velg rotmappe for nye grupper")
        if not base: return
        self._rules = apply_create_groups(self._rules, sels, base)
        messagebox.showinfo("Foreslå grupper", f"Laget {len(sels)} gruppe(r).")

    def _add_to_existing(self):
        sels = self._sel_domains()
        if not sels:
            messagebox.showinfo("Foreslå grupper", "Velg ett eller flere domener."); return
        dlg = tk.Toplevel(self); dlg.title("Velg gruppe"); dlg.geometry("320x160")
        frm = ttk.Frame(dlg, padding=10); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Legg domene(r) til gruppe:").pack(anchor="w")
        names = [g.name for g in self._rules]
        var = tk.StringVar(value=names[0] if names else "")
        ttk.Combobox(frm, values=names, textvariable=var, state="readonly").pack(fill="x", pady=(6,10))
        def ok():
            if not var.get(): return
            self._rules = apply_add_to_group(self._rules, sels, var.get())
            messagebox.showinfo("Foreslå grupper", f"Lagt til {len(sels)} domene(r) i «{var.get()}».")
            dlg.destroy()
        ttk.Button(frm, text="OK", command=ok).pack(side="right")
        ttk.Button(frm, text="Avbryt", command=dlg.destroy).pack(side="right", padx=(0,6))

    def _export_csv(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile="umatchede_domener.csv",
            title="Lagre CSV")
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Domene","Antall","Eksempler"])
                for s in self.suggestions:
                    w.writerow([s.domain, s.count, " | ".join(s.examples)])
            messagebox.showinfo("Eksport", f"Lagret CSV:\n{path}")
        except Exception as e:
            messagebox.showerror("Eksport", f"Feil ved skriving av CSV: {e}")
