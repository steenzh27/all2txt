from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

from .extractors import ExtractorError, extract_result_with_fallbacks, extract_text_with_fallbacks
from .models import DecodeResult


@dataclass
class TextDecoder:
    """Decode text from many document formats with fallback strategies."""

    preferred_tools: Optional[Iterable[str]] = None
    encoding: str = "utf-8"
    fallback_encodings: Optional[Iterable[str]] = None
    output_encoding: str = "utf-8"

    def decode_result(self, file_path: str | Path) -> DecodeResult:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.is_dir():
            raise IsADirectoryError(f"Expected file, got directory: {path}")

        return extract_result_with_fallbacks(
            path,
            preferred_tools=self.preferred_tools,
            default_encoding=self.encoding,
            fallback_encodings=self.fallback_encodings,
        )

    def decode_file(self, file_path: str | Path) -> str:
        return self.decode_result(file_path).text

    def decode_to_txt(self, file_path: str | Path, output_txt: str | Path) -> Path:
        text = self.decode_file(file_path)
        out = Path(output_txt)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding=self.output_encoding, errors="replace")
        return out


def decode_file(file_path: str | Path, preferred_tools: Optional[Iterable[str]] = None) -> str:
    """Convenience function for one-shot text extraction."""
    return TextDecoder(preferred_tools=preferred_tools).decode_file(file_path)


def decode_result(
    file_path: str | Path,
    preferred_tools: Optional[Iterable[str]] = None,
    encoding: str = "utf-8",
    fallback_encodings: Optional[Iterable[str]] = None,
) -> DecodeResult:
    """Convenience function that returns text together with metadata and method details."""
    return TextDecoder(
        preferred_tools=preferred_tools,
        encoding=encoding,
        fallback_encodings=fallback_encodings,
    ).decode_result(file_path)


def decode_to_txt(
    file_path: str | Path,
    output_txt: str | Path,
    preferred_tools: Optional[Iterable[str]] = None,
    encoding: str = "utf-8",
    fallback_encodings: Optional[Iterable[str]] = None,
    output_encoding: str = "utf-8",
) -> Path:
    """Convenience function that writes extracted text to .txt."""
    return TextDecoder(
        preferred_tools=preferred_tools,
        encoding=encoding,
        fallback_encodings=fallback_encodings,
        output_encoding=output_encoding,
    ).decode_to_txt(file_path, output_txt)


__all__ = ["TextDecoder", "DecodeResult", "decode_file", "decode_result", "decode_to_txt", "ExtractorError"]
