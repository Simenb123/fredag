from __future__ import annotations
import argparse
from datetime import datetime, timedelta, date
from typing import Dict, List, Tuple

from .outlook_core import get_session, search_messages, default_smtp
from .group_rules import load_rules
from .group_archiver import archive_by_groups
from .mail_utils import send_html_mail
from .locking import try_acquire_lock
from .retention import apply_retention

LOCK_NAME = "auto_archive_run"

def _from_to_from_days(days: int) -> tuple[date, date]:
    today = datetime.now().date()
    return today - timedelta(days=days), today

def run_archive(session,
                from_date: date,
                to_date: date,
                include_subfolders: bool = True,
                only_attachments: bool = True,
                unread_only: bool = False,
                subject_contains: str = "",
                dry_run: bool = False,
                cap_per_folder: int = 6000,
                cap_total: int = 4000) -> Tuple[Dict, List[Dict]]:
    stop_flag = type("Stop", (), {"is_set": lambda self: False})()
    res, err, aborted = search_messages(
        session=session, sender_query="", subject_contains=subject_contains,
        after_date=from_date, before_date=to_date, include_subfolders=include_subfolders,
        only_unread=unread_only, only_attachments=only_attachments,
        cap_per_folder=cap_per_folder, cap_total=cap_total, stop_evt=stop_flag, progress=None
    )
    if aborted: raise SystemExit("Avbrutt.")
    if err:     raise SystemExit(f"Feil under søk: {err}")

    summary, unassigned = archive_by_groups(session, res, rules=load_rules(), dedup=True, dry_run=dry_run)
    return summary, unassigned

def _html_report(summary: Dict, unassigned_count: int, f: date, t: date, dry: bool) -> str:
    rows = "".join(
        f"<tr><td>{g}</td><td style='text-align:right'>{s['msgs']}</td>"
        f"<td style='text-align:right'>{s['saved']}</td><td style='text-align:right'>{s['skipped']}</td></tr>"
        for g, s in summary.items()
    )
    lbl = "TØRRKJØRING" if dry else "Arkivering"
    return f"""<html><body>
    <h3>{lbl} – rapport</h3>
    <p>Intervall: {f.strftime('%d.%m.%Y')} – {t.strftime('%d.%m.%Y')}</p>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr><th>Gruppe</th><th>Meldinger</th><th>{'Ville lagret' if dry else 'Lagret'}</th><th>Hoppet</th></tr>
      {rows or '<tr><td colspan="4">(Ingen grupper matchet)</td></tr>'}
    </table>
    <p>Uten gruppe: {unassigned_count}</p>
    </body></html>"""

def main():
    ap = argparse.ArgumentParser(description="Arkiver vedlegg etter grupper.")
    ap.add_argument("--from-days", type=int, default=7, help="Antall dager tilbake (default 7)")
    ap.add_argument("--from-date", type=str, help="YYYY-MM-DD (overstyrer --from-days)")
    ap.add_argument("--to-date", type=str, help="YYYY-MM-DD (default i dag)")
    ap.add_argument("--no-subfolders", action="store_true", help="Ikke søk i undermapper")
    ap.add_argument("--only-unread", action="store_true", help="Kun uleste")
    ap.add_argument("--only-attachments", action="store_true", help="Kun med vedlegg")
    ap.add_argument("--subject", type=str, default="", help="Filter: emne inneholder")
    ap.add_argument("--mail-report", action="store_true", help="Send e‑postrapport")
    ap.add_argument("--to", type=str, help="Overstyr mottaker for rapport")
    ap.add_argument("--dry-run", action="store_true", help="Tørrkjøring – lagrer ingenting")
    ap.add_argument("--after-retention", action="store_true", help="Kjør retention etter arkivering")

    args = ap.parse_args()
    if args.from_date:
        f = datetime.strptime(args.from_date, "%Y-%m-%d").date()
        t = datetime.strptime(args.to_date, "%Y-%m-%d").date() if args.to_date else datetime.now().date()
    else:
        f, t = _from_to_from_days(args.from_days)

    lock = try_acquire_lock(LOCK_NAME, timeout_sec=2)
    if not lock:
        print("En annen arkiveringsjobb kjører allerede – avbryter.")
        return
    try:
        session = get_session()
        summary, unassigned = run_archive(
            session=session, from_date=f, to_date=t, include_subfolders=not args.no_subfolders,
            only_attachments=args.only_attachments or True, unread_only=args.only_unread,
            subject_contains=args.subject or "", dry_run=args.dry_run
        )

        print("=== Tørrkjøring pr. gruppe ===" if args.dry_run else "=== Arkivert pr. gruppe ===")
        for g, s in summary.items():
            print(f"- {g}: {('ville lagret' if args.dry_run else 'lagret')} {s['saved']}, hoppet {s['skipped']} (meldinger: {s['msgs']})")
        if unassigned:
            print(f"(Uten gruppe: {len(unassigned)} meldinger – ikke berørt)")

        if args.mail_report:
            to = args.to or (default_smtp(session) or "")
            if to:
                html = _html_report(summary, len(unassigned), f, t, args.dry_run)
                ok, msg = send_html_mail(session, to, "Arkivering – rapport (tørrkjøring)" if args.dry_run else "Arkivering – rapport", html)
                print(f"Rapport: {'OK' if ok else 'FEIL'} – {msg}")
            else:
                print("Ingen standard e‑postadresse – hopper over rapport.")

        if args.after_retention and not args.dry_run:
            rsum = apply_retention(load_rules(), dry_run=False)
            print("=== Retention etter arkivering ===")
            for g, s in rsum.items():
                print(f"- {g}: slettet {s['deleted']}, beholdt {s['kept']}, feil {s['errors']}")

    finally:
        lock.release()

if __name__ == "__main__":
    main()
