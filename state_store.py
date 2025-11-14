from __future__ import annotations
import sqlite3
from pathlib import Path
from typing import Optional
from datetime import datetime

_DB = None  # type: Optional[sqlite3.Connection]

def _db_path() -> Path:
    root = Path(__file__).resolve().parents[1] / ".ragdb"
    root.mkdir(exist_ok=True)
    return root / "state.db"

def _conn() -> sqlite3.Connection:
    global _DB
    if _DB is None:
        _DB = sqlite3.connect(str(_db_path()))
        _DB.execute("PRAGMA journal_mode=WAL;")
        _ensure_schema(_DB)
    return _DB

def _ensure_schema(db: sqlite3.Connection) -> None:
    db.executescript("""
    CREATE TABLE IF NOT EXISTS properties (
        k TEXT PRIMARY KEY,
        v TEXT
    );
    CREATE TABLE IF NOT EXISTS archived_messages (
        eid TEXT PRIMARY KEY,
        ts  TEXT NOT NULL
    );
    """)
    db.commit()

# --------- last run -----------
def get_last_run(job: str) -> Optional[datetime]:
    cur = _conn().execute("SELECT v FROM properties WHERE k=?", (f"last_run:{job}",))
    row = cur.fetchone()
    if not row:
        return None
    try:
        return datetime.fromisoformat(row[0])
    except Exception:
        return None

def set_last_run(job: str, ts: Optional[datetime] = None) -> None:
    ts = ts or datetime.now()
    _conn().execute("REPLACE INTO properties(k,v) VALUES (?,?)", (f"last_run:{job}", ts.isoformat()))
    _conn().commit()

# --------- arkiverte meldinger -----------
def was_archived(eid: str) -> bool:
    if not eid:
        return False
    cur = _conn().execute("SELECT 1 FROM archived_messages WHERE eid=?", (eid,))
    return cur.fetchone() is not None

def mark_archived(eid: str) -> None:
    if not eid:
        return
    _conn().execute(
        "INSERT OR IGNORE INTO archived_messages(eid, ts) VALUES (?, ?)",
        (eid, datetime.now().isoformat())
    )
    _conn().commit()
