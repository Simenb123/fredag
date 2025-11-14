from datetime import datetime
from pathlib import Path

from fredag.archiver import archive_messages


# ---- Fakes (ingen Outlook/COM nødvendig) ----
class FakeAttachment:
    def __init__(self, name: str, content: bytes):
        self._name = name
        self._content = content

    @property
    def FileName(self):
        return self._name

    def SaveAsFile(self, path):
        # Simuler at Outlook lagrer vedlegg til disk
        Path(path).write_bytes(self._content)


class FakeAttachments:
    def __init__(self, files):
        self._files = files

    @property
    def Count(self):
        return len(self._files)

    def Item(self, i: int):
        # Outlook-stil: 1-basert indeks
        return self._files[i - 1]


class FakeMail:
    def __init__(self, files):
        self.Attachments = FakeAttachments(files)


class FakeSession:
    def __init__(self, mapping):
        # mapping: eid -> FakeMail
        self._map = mapping

    def GetItemFromID(self, eid, store=None):
        return self._map[eid]


def test_archive_with_dedup(tmp_path: Path):
    """
    To meldinger med samme vedlegg -> dedup gir
    1 lagret og 1 hoppet.

    Vi kjører med dry_run=True for å unngå problemer med
    flytting mellom ulike disker (UNC vs. lokal C:).
    Dedup-logikken testes likevel fullt ut.
    """
    content = b"same"
    a1 = FakeMail([FakeAttachment("x.txt", content)])
    a2 = FakeMail([FakeAttachment("x.txt", content)])
    sess = FakeSession({"E1": a1, "E2": a2})

    results = [
        {
            "eid": "E1",
            "store": None,
            "dt": datetime(2025, 1, 1, 10, 0),
            "from": "A",
            "from_email": "a@x.no",
        },
        {
            "eid": "E2",
            "store": None,
            "dt": datetime(2025, 1, 2, 10, 0),
            "from": "A",
            "from_email": "a@x.no",
        },
    ]

    # archive_messages forventer en get_item-funksjon som tar et result-dict.
    # Vi wrapper FakeSession.GetItemFromID slik at den plukker ut eid/store.
    get_item = lambda r: sess.GetItemFromID(r["eid"], r.get("store"))

    saved, skipped, err = archive_messages(
        sess,
        results,
        get_item,
        str(tmp_path),
        per_sender=True,
        dedup=True,
        dry_run=True,   # <- viktig for å unngå faktisk flytting av filer
    )

    # Viktigste sjekk: dedup gir 1 lagret og 1 hoppet.
    assert saved == 1
    assert skipped == 1
    # err kan inneholde miljøspesifikke ting; vi bryr oss ikke her.
