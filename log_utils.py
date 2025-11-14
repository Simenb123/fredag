from __future__ import annotations
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

def _logs_dir() -> Path:
    """
    Returnerer katalogen for logger: <prosjekt>/.ragdb/logs
    Oppretter mappestrukturen ved behov.
    """
    root = Path(__file__).resolve().parents[1] / ".ragdb" / "logs"
    root.mkdir(parents=True, exist_ok=True)
    return root

def log_path(name: str = "fredag") -> Path:
    """
    Full sti til loggfil (brukes bl.a. av 'Loggmappe'-knappen i GUI).
    """
    safe = name.strip() or "fredag"
    return _logs_dir() / f"{safe}.log"

def setup_file_logger(
    name: str = "fredag",
    level: int = logging.INFO,
    max_bytes: int = 5 * 1024 * 1024,
    backup_count: int = 3,
) -> logging.Logger:
    """
    Oppretter en roterende fil‑logger som skriver til log_path(name).
    Kaller du denne flere ganger får du samme logger (uten duplikate handlers).
    """
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(level)
    handler = RotatingFileHandler(
        log_path(name),
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger

def log_event(msg: str, level: int = logging.INFO, name: str = "fredag") -> None:
    """
    Praktisk hjelpefunksjon for rask logging fra hvor som helst.
    """
    setup_file_logger(name=name).log(level, msg)
