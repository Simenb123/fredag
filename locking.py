from __future__ import annotations
import os, time
from pathlib import Path
from typing import Optional

class Lock:
    def __init__(self, name: str):
        self.name = name
        self.path = Path(__file__).resolve().parents[1] / ".ragdb" / f"{name}.lock"
        self._acquired = False

    def acquire(self, timeout_sec: int = 0) -> bool:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        end = time.time() + max(0, timeout_sec)
        while True:
            try:
                # eksklusiv opprettelse – feiler hvis finnes
                fd = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    f.write(f"pid={os.getpid()}\n")
                self._acquired = True
                return True
            except FileExistsError:
                if time.time() >= end:
                    return False
                time.sleep(0.2)

    def release(self):
        if self._acquired:
            try:
                self.path.unlink(missing_ok=True)
            finally:
                self._acquired = False

def try_acquire_lock(name: str, timeout_sec: int = 0) -> Optional[Lock]:
    lk = Lock(name)
    return lk if lk.acquire(timeout_sec=timeout_sec) else None
