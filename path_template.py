from __future__ import annotations
import re
from pathlib import Path
from typing import Dict

_MONTH_ABBR = ["Jan","Feb","Mar","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Des"]

def safe_component(name: str) -> str:
    name = (name or "").strip()
    return "".join("_" if c in r'<>:"/\|?*' else c for c in name) or "_"

def month_abbr(m: int) -> str:
    return _MONTH_ABBR[max(1, min(12, int(m))) - 1]

def extract_subject_tag(subject: str, regex: str) -> str:
    if not regex:
        return ""
    try:
        m = re.search(regex, subject or "", flags=re.IGNORECASE)
        if not m:
            return ""
        g = m.group(1) if m.groups() else m.group(0)
        return safe_component(g)
    except re.error:
        # ugyldig regex – returner tom streng for ikke å feile kjøring
        return ""

class _SafeDict(dict):
    def __missing__(self, key):  # gjør .format_map robust
        return ""

def render_template(template: str, meta: Dict[str, str]) -> Path:
    """
    Bygger relativ sti fra template + meta. Template kan bruke:
      {year} {month2} {month_abbr} {sender} {domain} {subject_tag}
    Eksempel: "{year}/{month2}_{month_abbr}/{domain}/{subject_tag}"
    """
    tpl = (template or "").strip().replace("\\", "/").strip("/")
    if not tpl:
        return Path(".")
    s = tpl.format_map(_SafeDict(meta))
    parts = [safe_component(p) for p in s.split("/")]
    return Path(*[p for p in parts if p and p != "."])

def domain_from_email(smtp: str) -> str:
    smtp = (smtp or "").strip().lower()
    if "@" not in smtp:
        return ""
    return smtp.rsplit("@", 1)[-1]
