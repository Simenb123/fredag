from datetime import datetime, timedelta
from typing import List, Tuple, Dict

from .config import TOP_N_SENDERS, MAX_PER_FOLDER, FALLBACK_RECENT_N
from .outlook_core import msg_time, normalize_sender, walk_subfolders

def _start_of_week_local() -> datetime:
    t = datetime.now()
    return (t - timedelta(days=t.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)

def _count_senders_in_folder(folder, sow_date, per_folder_limit=MAX_PER_FOLDER) -> Dict[Tuple[str, str], int]:
    counts: Dict[Tuple[str, str], int] = {}
    try:
        items = folder.Items
    except Exception:
        return counts
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    processed = 0
    try:
        it = items.GetFirst()
        seq_ok = True
    except Exception:
        it = None
        seq_ok = False

    if seq_ok:
        while it and processed < per_folder_limit:
            try:
                if getattr(it, "Class", None) == 43:
                    dt = msg_time(it)
                    if dt and dt.date() >= sow_date:
                        name, smtp = normalize_sender(it)
                        key = ((name or "")[:120], (smtp or "")[:200])
                        counts[key] = counts.get(key, 0) + 1
                processed += 1
                it = items.GetNext()
            except Exception:
                break
    else:
        total = getattr(items, "Count", 0)
        upto = min(per_folder_limit, total)
        for idx in range(1, upto + 1):
            try:
                it = items.Item(idx)
                if it and getattr(it, "Class", None) == 43:
                    dt = msg_time(it)
                    if dt and dt.date() >= sow_date:
                        name, smtp = normalize_sender(it)
                        key = ((name or "")[:120], (smtp or "")[:200])
                        counts[key] = counts.get(key, 0) + 1
            except Exception:
                continue
    return counts

def weekly_sender_stats(session, top_n: int = TOP_N_SENDERS) -> List[Tuple[str, str, int]]:
    """Skann Default Innboks + undermapper for inneværende uke. Fallback: N siste i Innboks."""
    sow_date = _start_of_week_local().date()
    counts: Dict[Tuple[str, str], int] = {}

    try:
        inbox = session.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        inbox = None

    if inbox is not None:
        for folder in walk_subfolders(inbox, include_subfolders=True):
            sub_counts = _count_senders_in_folder(folder, sow_date)
            if sub_counts:
                for k, v in sub_counts.items():
                    counts[k] = counts.get(k, 0) + v

    if not counts and inbox is not None:
        sub_counts = _count_senders_in_folder(inbox, sow_date, per_folder_limit=FALLBACK_RECENT_N)
        for k, v in sub_counts.items():
            counts[k] = counts.get(k, 0) + v

    ranked = sorted(((k[0], k[1], v) for k, v in counts.items()),
                    key=lambda x: x[2], reverse=True)
    return ranked[:top_n]
