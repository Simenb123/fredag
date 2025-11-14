from __future__ import annotations
import subprocess
import sys
from pathlib import Path

def _quote(s: str) -> str:
    return '"' + s.replace('"', '\\"') + '"'

def install_weekly_task(task_name: str,
                        time_hhmm: str,
                        script: Path,
                        args: str = "",
                        day: str = "FRI") -> tuple[bool, str]:
    """
    Oppretter/oppdaterer en ukentlig Windows Scheduled Task (brukerkontekst).
    """
    py = Path(sys.executable)
    tr = f'{_quote(str(py))} {_quote(str(script))} {args}'.strip()
    cmd = [
        "schtasks", "/Create", "/F",
        "/SC", "WEEKLY", "/D", day,
        "/ST", time_hhmm,
        "/TN", task_name,
        "/TR", tr
    ]
    try:
        cp = subprocess.run(cmd, capture_output=True, text=True, check=False)
        ok = cp.returncode == 0
        out = cp.stdout.strip() or cp.stderr.strip()
        return ok, out
    except Exception as e:
        return False, f"Kunne ikke opprette planlagt oppgave: {e}"

def delete_task(task_name: str) -> tuple[bool, str]:
    cmd = ["schtasks", "/Delete", "/F", "/TN", task_name]
    try:
        cp = subprocess.run(cmd, capture_output=True, text=True, check=False)
        ok = cp.returncode == 0
        out = cp.stdout.strip() or cp.stderr.strip()
        return ok, out
    except Exception as e:
        return False, f"Kunne ikke slette oppgave: {e}"
