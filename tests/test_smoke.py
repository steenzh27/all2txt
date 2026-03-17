from pathlib import Path

from all2txt import decode_file


def test_txt_smoke(tmp_path: Path) -> None:
    sample = tmp_path / "a.txt"
    sample.write_text("hello", encoding="utf-8")
    text = decode_file(sample)
    assert "hello" in text

