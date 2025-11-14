from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from fredag.archiver import archive_messages

# ---- Fakes (ingen Outlook/COM nødvendig) ----
class FakeAttachment:
    def __init__(self, name: str, content: bytes):
        self._name = name; self._content = content
    @property
    def FileName(self): return self._name
    def SaveAsFile(self, path): Path(path).write_bytes(self._content)

class FakeAttachments:
    def __init__(self, files): self._files = files
    @property
    def Count(self): return len(self._files)
    def Item(self, i): return self._files[i-1]

class FakeMail:
    def __init__(self, files): self.Attachments = FakeAttachments(files)

class FakeSession:
    def __init__(self, mapping): self._map = mapping
    def GetItemFromID(self, eid, store=None): return self._map[eid]

def test_archive_with_dedup(tmp_path: Path):
    # to meldinger med samme vedlegg -> dedup gir 1 lagret, 1 hoppet
    content = b"same"
    a1 = FakeMail([FakeAttachment("x.txt", content)])
    a2 = FakeMail([FakeAttachment("x.txt", content)])
    sess = FakeSession({"E1": a1, "E2": a2})

    results = [
        {"eid": "E1", "store": None, "dt": datetime(2025,1,1,10,0),
         "from": "A", "from_email": "a@x.no"},
        {"eid": "E2", "store": None, "dt": datetime(2025,1,2,10,0),
         "from": "A", "from_email": "a@x.no"},
    ]
    saved, skipped, err = archive_messages(sess, results, sess.GetItemFromID, str(tmp_path), per_sender=True, dedup=True)
    assert err == ""
    assert saved == 1 and skipped == 1
    # index finnes
    idx = tmp_path / ".archive_index.json"
    assert idx.exists()
