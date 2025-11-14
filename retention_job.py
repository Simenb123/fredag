from __future__ import annotations
import argparse
from typing import Dict
from .group_rules import load_rules
from .retention import apply_retention
from .outlook_core import get_session, default_smtp
from .mail_utils import send_html_mail

def _html(summary: Dict[str, Dict[str,int]], dry: bool) -> str:
    rows = "".join(
        f"<tr><td>{g}</td>"
        f"<td style='text-align:right'>{s['deleted']}</td>"
        f"<td style='text-align:right'>{s['kept']}</td>"
        f"<td style='text-align:right'>{s['errors']}</td></tr>"
        for g, s in summary.items()
    )
    title = "Retention – tørrkjøring" if dry else "Retention – opprydding"
    return f"""<html><body>
    <h3>{title}</h3>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr><th>Gruppe</th><th>{'Ville slettet' if dry else 'Slettet'}</th><th>Beholdt</th><th>Feil</th></tr>
      {rows or '<tr><td colspan="4">(Ingen grupper med retention)</td></tr>'}
    </table>
    </body></html>"""

def main():
    ap = argparse.ArgumentParser(description="Kjør retention (opprydding) for alle grupper.")
    ap.add_argument("--dry-run", action="store_true", help="Tørrkjøring – slett ikke")
    ap.add_argument("--mail-report", action="store_true", help="Send rapport på e‑post til deg selv")
    ap.add_argument("--to", type=str, help="Mottaker (overstyr)")
    args = ap.parse_args()

    rules = load_rules()
    summary = apply_retention(rules, dry_run=args.dry_run)

    print("=== Retention ===" + (" (tørrkjøring)" if args.dry_run else ""))
    for g, s in summary.items():
        print(f"- {g}: {('ville slettet' if args.dry_run else 'slettet')} {s['deleted']}, beholdt {s['kept']}, feil {s['errors']}")

    if args.mail_report:
        session = get_session()
        to = args.to or (default_smtp(session) or "")
        if to:
            html = _html(summary, args.dry_run)
            ok, msg = send_html_mail(session, to, "Retention – rapport (tørrkjøring)" if args.dry_run else "Retention – rapport", html)
            print(f"Rapport: {'OK' if ok else 'FEIL'} – {msg}")
        else:
            print("Mangler standard e‑postadresse – hopper over rapport.")

if __name__ == "__main__":
    main()
