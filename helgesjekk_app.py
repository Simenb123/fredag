import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta, date, time

from .config import WEEKEND_CUTOFF, DAGNAVN, TOP_N_SENDERS, FALLBACK_EMAIL
from .outlook_core import have_outlook, get_outlook, get_session, default_smtp
from .email_stats import weekly_sender_stats
from .html_email import build_html
from .tools_window import OutlookToolsWindow
from .calendar_window import CalendarWindow


def day_name(d: date) -> str:
    return DAGNAVN[d.weekday()]


def next_friday_cutoff(now: datetime) -> datetime:
    """Neste fredag kl. WEEKEND_CUTOFF."""
    days_to_fri = (4 - now.weekday()) % 7
    cand = (now + timedelta(days=days_to_fri)).replace(
        hour=WEEKEND_CUTOFF.hour,
        minute=WEEKEND_CUTOFF.minute,
        second=0,
        microsecond=0,
    )
    if cand <= now:
        cand += timedelta(days=7)
    return cand


def next_monday_midnight(now: datetime) -> datetime:
    """
    Neste mandag kl. 08:00.

    (Navnet er historisk – tidligere var dette 00:00. Nå definerer vi
    helgen som fredag fra WEEKEND_CUTOFF til mandag kl. 08:00.)
    """
    days_to_mon = (7 - now.weekday()) % 7
    cand = (now + timedelta(days=days_to_mon)).replace(
        hour=8,
        minute=0,
        second=0,
        microsecond=0,
    )
    if cand <= now:
        cand += timedelta(days=7)
    return cand


def is_weekend(now: datetime) -> bool:
    """
    Helg-definisjon:
    - Fredag fra WEEKEND_CUTOFF og utover
    - Hele lørdag og søndag
    - Mandag frem til kl. 08:00
    """
    wd = now.weekday()
    t = now.time()

    # Fredag etter cutoff
    if wd == 4 and t >= WEEKEND_CUTOFF:
        return True

    # Lørdag og søndag
    if wd in (5, 6):
        return True

    # Mandag før arbeidsstart kl. 08
    if wd == 0 and t < time(8, 0):
        return True

    return False


def days_until_friday(now: datetime) -> int:
    return (4 - now.weekday()) % 7


def _send_html_mail(session, subject: str, status_text: str):
    app = get_outlook()
    if not app:
        return False, "Klarte ikke å starte Outlook (klassisk)."
    to_addr = default_smtp(session) or FALLBACK_EMAIL
    if not to_addr:
        return False, "Fant ikke standard e-post i Outlook. Sett FALLBACK_EMAIL i config.py."
    stats = weekly_sender_stats(session, TOP_N_SENDERS)
    html = build_html(subject, status_text, stats)
    m = app.CreateItem(0)
    m.To = to_addr
    m.Subject = subject
    try:
        m.BodyFormat = 2
    except Exception:
        pass
    m.HTMLBody = html
    m.Send()
    return True, f"E-post sendt til {to_addr}."


class HelgesjekkApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Helgesjekk (Outlook – ukesoppsummering, HTML)")
        tk.Label(
            self.root,
            text="Helgesjekk",
            font=("TkDefaultFont", 13, "bold"),
        ).pack(padx=20, pady=(18, 6))

        btns = tk.Frame(self.root)
        btns.pack()
        tk.Button(
            btns,
            text="Sjekk helg",
            width=16,
            command=self.sjekk_helg_og_send,
        ).pack(side="left", padx=6)
        tk.Button(
            btns,
            text="Outlook-verktøy",
            width=16,
            command=self.open_outlook_tools,
        ).pack(side="left", padx=6)
        tk.Button(
            btns,
            text="Min kalender",
            width=16,
            command=self.open_calendar,
        ).pack(side="left", padx=6)

        self.svar_label = tk.Label(
            self.root,
            text="Trykk «Sjekk helg» – sender pen HTML-epost for inneværende uke.",
            font=("TkDefaultFont", 11),
        )
        self.svar_label.pack(pady=(8, 6))
        self.feedback_label = tk.Label(
            self.root,
            text="",
            font=("TkDefaultFont", 9),
        )
        self.feedback_label.pack(pady=(2, 14))
        self._countdown_job = None
        self._last_status_text = ""
        self._tools_window = None
        self._cal_window = None

    def _cancel_countdown(self):
        if self._countdown_job:
            try:
                self.root.after_cancel(self._countdown_job)
            except Exception:
                pass
            self._countdown_job = None

    def _start_countdown(self, target: datetime, prefix: str = ""):
        self._cancel_countdown()
        self._update_countdown(target, prefix)

    def _update_countdown(self, target: datetime, prefix: str):
        now = datetime.now()
        remaining = target - now
        if remaining.total_seconds() <= 0:
            txt = (
                "🎉 GOD HELG! 🎉"
                if is_weekend(datetime.now())
                else "Helgen er over. God mandag!"
            )
            self.svar_label.config(text=txt)
            self._last_status_text = txt
            self._countdown_job = None
            return

        total = int(remaining.total_seconds())
        days, rem = divmod(total, 86400)
        hours, rem = divmod(rem, 3600)
        minutes, seconds = divmod(rem, 60)
        if days > 0:
            txt = f"{prefix}{days}d {hours:02d}:{minutes:02d}:{seconds:02d}"
        else:
            txt = f"{prefix}{hours:02d}:{minutes:02d}:{seconds:02d}"

        self.svar_label.config(text=txt)
        self._last_status_text = txt
        self._countdown_job = self.root.after(
            1000, self._update_countdown, target, prefix
        )

    def sjekk_helg_og_send(self):
        now = datetime.now()
        self._cancel_countdown()

        if is_weekend(now):
            today = day_name(now.date())
            slutt = next_monday_midnight(now)
            self._start_countdown(
                slutt,
                f"Det er {today} – det er helg 🎉  Helgen varer i: ",
            )
        else:
            today = day_name(now.date())
            n = days_until_friday(now)
            if n == 0:
                info = "Det er fredag – nedtelling til helg: "
            elif n == 1:
                info = f"Det er {today} – 1 dag til helg. Nedtelling: "
            else:
                info = f"Det er {today} – {n} dager til helg. Nedtelling: "
            self._start_countdown(next_friday_cutoff(now), info)

        if not have_outlook():
            self.feedback_label.config(
                text="Outlook/pywin32 mangler. Installer pywin32.",
                fg="red",
            )
            return

        session = get_session(get_outlook())
        if not session:
            self.feedback_label.config(
                text="Klarte ikke å hente Outlook-sesjon.",
                fg="red",
            )
            return

        subject = (
            f"Helgesjekk + ukesoppsummering – {datetime.now():%Y-%m-%d %H:%M}"
        )
        ok, msg = _send_html_mail(session, subject, self._last_status_text)
        self.feedback_label.config(text=msg, fg=("green" if ok else "red"))

    def open_outlook_tools(self):
        app = get_outlook()
        if not app:
            messagebox.showerror(
                "Outlook",
                "Klarte ikke å starte Outlook (klassisk).",
            )
            return
        session = get_session(app)
        if not session:
            messagebox.showerror(
                "Outlook",
                "Klarte ikke å hente Outlook-sesjon.",
            )
            return
        if self._tools_window and self._tools_window.winfo_exists():
            self._tools_window.focus_set()
            return
        self._tools_window = OutlookToolsWindow(self.root, session)
        self._tools_window.protocol(
            "WM_DELETE_WINDOW",
            lambda: setattr(self, "_tools_window", None),
        )

    def open_calendar(self):
        app = get_outlook()
        if not app:
            messagebox.showerror(
                "Outlook",
                "Klarte ikke å starte Outlook (klassisk).",
            )
            return
        session = get_session(app)
        if not session:
            messagebox.showerror(
                "Outlook",
                "Klarte ikke å hente Outlook-sesjon.",
            )
            return
        if self._cal_window and self._cal_window.winfo_exists():
            self._cal_window.focus_set()
            return
        self._cal_window = CalendarWindow(self.root, session)
        self._cal_window.protocol(
            "WM_DELETE_WINDOW",
            lambda: setattr(self, "_cal_window", None),
        )


def run_app():
    HelgesjekkApp().root.mainloop()
