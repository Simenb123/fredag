import os
import re
import tempfile
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Prøv Pillow for bedre bilde-støtte (JPEG osv.), men fall tilbake til Tk om det ikke finnes.
try:
    from PIL import Image, ImageTk  # type: ignore
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False

_ILLEGAL = r'<>:"/\\|?*'

def _sanitize(name: str) -> str:
    return "".join("_" if c in _ILLEGAL else c for c in (name or ""))

def _unique_path(base_dir: str, filename: str) -> str:
    base, ext = os.path.splitext(filename or "vedlegg")
    cand = os.path.join(base_dir, filename or "vedlegg")
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(base_dir, f"{base} ({i}){ext}")
        i += 1
    return cand

# Enkel HTML -> tekst
_RE_TAG = re.compile(r"<[^>]+>")
def _as_text_from_html(html: str) -> str:
    return _RE_TAG.sub("", html or "").strip()

class AttachmentsWindow(tk.Toplevel):
    """
    Viser alle vedlegg for én MailItem.
    - Søkefelt filtrerer listen live
    - Dobbeltklikk åpner vedlegg (lagres i temp under %TEMP%\\OutlookVedlegg\\<melding>\\)
    - Forhåndsvisning: tekst/HTML vises som tekst, PNG/GIF via Tk, JPEG mm. via Pillow hvis tilgjengelig
    - Lagre valgte / alle, åpne temp-mappe
    """
    def __init__(self, master, mail_item, context_title: str = ""):
        super().__init__(master)
        self.mail = mail_item
        self.title(f"Vedlegg – {context_title or 'melding'}")
        self.geometry("900x560")

        # dedikert temp-mappe per melding
        root = os.path.join(tempfile.gettempdir(), "OutlookVedlegg")
        msg_label = _sanitize(context_title or f"msg_{int(time.time())}")
        self.temp_dir = os.path.join(root, msg_label[:80])
        os.makedirs(self.temp_dir, exist_ok=True)

        # intern liste over vedlegg (metadata)
        # [{'index':1, 'name':'fil.png', 'size':12345, 'ext':'.png', 'temp':None}]
        self._atts = []
        self._row_to_ix = {}     # tree iid -> index i _atts
        self._img_cache = None   # holder på siste PhotoImage så preview ikke blir GC'et

        self._build_ui()
        self._load_attachments()

    # ---------- UI ----------
    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=8)
        wrapper.pack(fill="both", expand=True)

        # Søk + knapper
        top = ttk.Frame(wrapper)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Søk i vedlegg:").pack(side="left")
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.filter_var, width=36)
        ent.pack(side="left", padx=(6, 10))
        ent.bind("<KeyRelease>", lambda e: self._apply_filter())

        ttk.Button(top, text="Åpne", command=self.open_selected).pack(side="left")
        ttk.Button(top, text="Åpne alle", command=self.open_all).pack(side="left", padx=(6, 0))
        ttk.Button(top, text="Lagre valgte…", command=lambda: self.save_selected(False)).pack(side="left", padx=(18, 0))
        ttk.Button(top, text="Lagre alle…", command=lambda: self.save_selected(True)).pack(side="left", padx=(6, 0))
        ttk.Button(top, text="Åpne mappe", command=self.open_dir).pack(side="right")

        # Øvre: liste over vedlegg
        upper = ttk.Frame(wrapper)
        upper.pack(fill="both", expand=True)

        cols = ("fil", "storrelse", "type")
        self.tree = ttk.Treeview(upper, columns=cols, show="headings", height=12)
        self.tree.heading("fil", text="Filnavn")
        self.tree.heading("storrelse", text="Størrelse")
        self.tree.heading("type", text="Type")
        self.tree.column("fil", width=520, anchor="w")
        self.tree.column("storrelse", width=110, anchor="e")
        self.tree.column("type", width=80, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(upper, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Nedre: forhåndsvisning
        prev = ttk.LabelFrame(wrapper, text="Forhåndsvisning")
        prev.pack(fill="both", expand=True, pady=(6, 0))

        self.preview_canvas = tk.Canvas(prev, height=280, bg="#fafafa", highlightthickness=0)
        self.preview_canvas.pack(fill="both", expand=True, padx=4, pady=4)

        # Statuslinje
        self.status = ttk.Label(self, text="Klar.")
        self.status.pack(fill="x", padx=8, pady=(6, 6))

    # ---------- Datakilde ----------
    def _load_attachments(self):
        self._atts.clear()
        try:
            atts = self.mail.Attachments
            count = getattr(atts, "Count", 0)
        except Exception:
            count = 0

        for i in range(1, count + 1):
            try:
                att = atts.Item(i)
                name = getattr(att, "FileName", f"vedlegg_{i}") or f"vedlegg_{i}"
                size = getattr(att, "Size", 0) or 0
                ext = (os.path.splitext(name)[1] or "").lower()
                self._atts.append({"index": i, "name": name, "size": size, "ext": ext, "temp": None})
            except Exception:
                continue

        self._apply_filter(initial=True)

    def _apply_filter(self, initial: bool = False):
        q = (self.filter_var.get() or "").strip().lower()
        self.tree.delete(*self.tree.get_children())
        self._row_to_ix.clear()

        def show_row(a):
            if not q:
                return True
            return (q in (a["name"] or "").lower()) or (q in (a["ext"] or ""))

        shown = 0
        for ix, a in enumerate(self._atts):
            if not show_row(a):
                continue
            size_kb = f"{round((a['size'] or 0) / 1024):,} KB".replace(",", " ")
            ext_disp = (a["ext"] or "").upper().lstrip(".")
            iid = self.tree.insert("", "end", values=(a["name"], size_kb, ext_disp))
            self._row_to_ix[iid] = ix
            shown += 1

        self.status.config(text=f"Fant {shown} vedlegg." if not initial else f"Fant {len(self._atts)} vedlegg.")
        if shown and initial:
            # Velg første rad ved init
            first = self.tree.get_children()
            if first:
                self.tree.selection_set(first[0])
                self.tree.focus(first[0])
                self._on_select(None)
        elif not shown:
            self._show_preview_message("Ingen treff.")

    # ---------- Forhåndsvisning ----------
    def _ensure_saved_temp(self, ix: int) -> str | None:
        a = self._atts[ix]
        if a["temp"] and os.path.exists(a["temp"]):
            return a["temp"]
        try:
            att = self.mail.Attachments.Item(a["index"])
        except Exception:
            return None
        path = _unique_path(self.temp_dir, _sanitize(a["name"]))
        try:
            att.SaveAsFile(path)
        except Exception:
            return None
        a["temp"] = path
        return path

    def _on_select(self, _evt):
        sel = self.tree.selection()
        if not sel:
            return
        ix = self._row_to_ix.get(sel[0])
        if ix is None:
            return
        self._preview_ix(ix)

    def _clear_preview(self):
        self.preview_canvas.delete("all")
        self._img_cache = None

    def _show_preview_message(self, text: str):
        self._clear_preview()
        self.preview_canvas.create_text(
            10, 10, anchor="nw", text=text, font=("Segoe UI", 10), fill="#333", width=self.preview_canvas.winfo_width()-20
        )

    def _preview_ix(self, ix: int):
        path = self._ensure_saved_temp(ix)
        if not path:
            self._show_preview_message("Klarte ikke å hente vedlegget.")
            return

        ext = (os.path.splitext(path)[1] or "").lower()
        # Tekstlige formater
        if ext in (".txt", ".log", ".csv", ".json", ".xml", ".md", ".ini"):
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
                self._show_preview_message(content)
            except Exception:
                self._show_preview_message("Kunne ikke lese som tekst.")
            return

        if ext in (".html", ".htm"):
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
                self._show_preview_message(_as_text_from_html(content))
            except Exception:
                self._show_preview_message("Kunne ikke lese som HTML.")
            return

        # Bilder: prøv PIL (JPG, PNG, etc). Hvis ikke PIL, støtte for PNG/GIF via Tk PhotoImage.
        if ext in (".png", ".gif", ".jpg", ".jpeg", ".bmp", ".webp"):
            self._clear_preview()
            try:
                if _HAS_PIL:
                    im = Image.open(path)
                    # skaler ned til maks 90% av canvas
                    cw = max(100, int(self.preview_canvas.winfo_width() * 0.9))
                    ch = max(100, int(self.preview_canvas.winfo_height() * 0.9))
                    im.thumbnail((cw, ch))
                    self._img_cache = ImageTk.PhotoImage(im)  # hold referanse!
                else:
                    # Tk PhotoImage støtter PNG/GIF (ikke alltid JPEG)
                    if ext not in (".png", ".gif"):
                        self._show_preview_message("Ingen forhåndsvisning (mangler Pillow for JPEG).")
                        return
                    self._img_cache = tk.PhotoImage(file=path)
                # senterer i canvas
                w = self.preview_canvas.winfo_width()
                h = self.preview_canvas.winfo_height()
                self.preview_canvas.create_image(w // 2, h // 2, image=self._img_cache)
            except Exception:
                self._show_preview_message("Klarte ikke å vise bildet.")
            return

        # PDF/Office – tilby åpning
        self._show_preview_message(f"Ingen forhåndsvisning for {os.path.basename(path)}.\n\n"
                                   f"Bruk 'Åpne' for å vise i tilknyttet program.")

    # ---------- Handlinger ----------
    def _selected_indices(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Vedlegg", "Velg minst ett vedlegg.")
            return []
        return [self._row_to_ix[iid] for iid in sel if iid in self._row_to_ix]

    def _save_one(self, ix: int, target: str, open_after: bool = False):
        a = self._atts[ix]
        # For å sikre riktig filnavn uten kollisjon, bruk temp eller COM direkte:
        if a["temp"] and os.path.exists(a["temp"]):
            src = a["temp"]
            fname = os.path.basename(src)
            dst = _unique_path(target, fname)
            try:
                with open(src, "rb") as rf, open(dst, "wb") as wf:
                    wf.write(rf.read())
                if open_after:
                    os.startfile(dst)
                return True
            except Exception:
                return False
        try:
            att = self.mail.Attachments.Item(a["index"])
        except Exception:
            return False
        clean = _sanitize(a["name"] or "vedlegg")
        path = _unique_path(target, clean)
        try:
            att.SaveAsFile(path)
            if open_after:
                os.startfile(path)
            return True
        except Exception:
            return False

    def open_selected(self):
        idxs = self._selected_indices()
        if not idxs:
            return
        ok = 0
        for ix in idxs:
            p = self._ensure_saved_temp(ix)
            if p:
                try:
                    os.startfile(p)
                    ok += 1
                except Exception:
                    pass
        self.status.config(text=f"Åpnet {ok} vedlegg.")

    def open_all(self):
        ok = 0
        for ix in range(len(self._atts)):
            p = self._ensure_saved_temp(ix)
            if p:
                try:
                    os.startfile(p)
                    ok += 1
                except Exception:
                    pass
        self.status.config(text=f"Åpnet {ok} vedlegg.")

    def save_selected(self, all_items: bool):
        if all_items:
            indices = list(range(len(self._atts)))
            if not indices:
                messagebox.showinfo("Vedlegg", "Ingen vedlegg.")
                return
        else:
            indices = self._selected_indices()
            if not indices:
                return

        target = filedialog.askdirectory(title="Lagre vedlegg til mappe")
        if not target:
            return

        ok = 0
        for ix in indices:
            if self._save_one(ix, target, open_after=False):
                ok += 1
        self.status.config(text=f"Lagret {ok} fil(er) → {target}")

    def open_dir(self):
        try:
            os.startfile(self.temp_dir)
        except Exception:
            messagebox.showerror("Mappe", "Klarte ikke å åpne mappen.")
