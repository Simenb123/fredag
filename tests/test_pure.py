from pathlib import Path
from fredag.outlook_core import html_to_text, sanitize_filename, unique_path

def test_html_to_text_basic():
    html = "<p>Hello<br>world</p><a href='https://a'>lenke</a>"
    t = html_to_text(html)
    assert "Hello" in t and "world" in t and "lenke (https://a)" in t

def test_sanitize_and_unique(tmp_path: Path):
    name = 'fil<>:"/\\|?*.txt'
    safe = sanitize_filename(name)
    assert all(c not in safe for c in '<>:"/\\|?*')
    p1 = tmp_path / safe
    p1.write_text("x", encoding="utf-8")
    p2 = unique_path(str(tmp_path), safe)
    assert Path(p2).name != safe  # fikk (1).txt e.l.
