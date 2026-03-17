"""all2txt public API."""

from .core import DecodeResult, TextDecoder, decode_file, decode_result, decode_to_txt
from .extractors import ExtractorError, register_extractor

__all__ = [
	"DecodeResult",
	"TextDecoder",
	"decode_file",
	"decode_result",
	"decode_to_txt",
	"ExtractorError",
	"register_extractor",
]

