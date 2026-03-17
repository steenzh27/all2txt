from __future__ import annotations

import csv
import html
import importlib.metadata
import inspect
import json
import plistlib
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET

from .models import DecodeResult, WarningEntry


class ExtractorError(RuntimeError):
    """Raised when text extraction fails for all available strategies."""


ExtractorFunc = Callable[..., str | DecodeResult]
_PLUGIN_EXTRACTORS: dict[str, list[ExtractorFunc]] = {}
_PLUGINS_LOADED = False


def register_extractor(suffixes: str | Iterable[str], extractor: ExtractorFunc) -> None:
    values = [suffixes] if isinstance(suffixes, str) else list(suffixes)
    for suffix in values:
        normalized = suffix.lower()
        if not normalized.startswith("."):
            normalized = f".{normalized}"
        _PLUGIN_EXTRACTORS.setdefault(normalized, []).append(extractor)


def _load_plugin_extractors() -> None:
    global _PLUGINS_LOADED
    if _PLUGINS_LOADED:
        return

    try:
        entry_points = importlib.metadata.entry_points(group="all2txt.extractors")
    except TypeError:
        entry_points = importlib.metadata.entry_points().get("all2txt.extractors", [])

    for entry_point in entry_points:
        try:
            plugin = entry_point.load()
        except Exception:
            continue

        suffixes = getattr(plugin, "suffixes", None)
        if not suffixes:
            continue
        register_extractor(suffixes, plugin)

    _PLUGINS_LOADED = True


def extract_text_with_fallbacks(
    path: Path,
    preferred_tools=None,
    default_encoding: str = "utf-8",
    fallback_encodings: Iterable[str] | None = None,
) -> str:
    return extract_result_with_fallbacks(
        path,
        preferred_tools=preferred_tools,
        default_encoding=default_encoding,
        fallback_encodings=fallback_encodings,
    ).text


def extract_result_with_fallbacks(
    path: Path,
    preferred_tools=None,
    default_encoding: str = "utf-8",
    fallback_encodings: Iterable[str] | None = None,
) -> DecodeResult:
    suffix = path.suffix.lower()
    attempts: list[WarningEntry] = []
    fallback_encodings = tuple(fallback_encodings or ())
    _load_plugin_extractors()

    direct_map = {
        ".txt": _read_text_file,
        ".log": _read_text_file,
        ".ini": _read_text_file,
        ".conf": _read_text_file,
        ".tex": _read_text_file,
        ".bib": _read_text_file,
        ".strings": _read_text_file,
        ".md": _read_text_file,
        ".rst": _read_text_file,
        ".csv": _read_csv_as_text,
        ".tsv": _read_csv_as_text,
        ".json": _read_json_as_text,
        ".plist": _read_plist,
        ".xml": _read_markup_as_text,
        ".html": _read_markup_as_text,
        ".htm": _read_markup_as_text,
        ".mht": _read_mhtml,
        ".mhtml": _read_mhtml,
        ".eml": _read_eml,
        ".msg": _read_msg,
        ".rtf": _read_rtf_as_text,
        ".doc": _read_legacy_doc,
        ".docx": _read_docx,
        ".odt": _read_odt,
        ".ods": _read_ods,
        ".xls": _read_legacy_xls,
        ".xlsx": _read_xlsx,
        ".ppt": _read_legacy_ppt,
        ".pptx": _read_pptx,
        ".pages": _read_iwork,
        ".numbers": _read_iwork,
        ".key": _read_iwork,
        ".djv": _read_djvu,
        ".djvu": _read_djvu,
        ".ps": _read_postscript,
        ".eps": _read_postscript,
        ".chm": _read_chm,
        ".fb2": _read_fb2,
        ".epub": _read_epub,
        ".mobi": _read_mobi,
        ".pdf": _read_pdf,
    }

    for plugin in _PLUGIN_EXTRACTORS.get(suffix, []):
        try:
            result = _run_registered_extractor(
                plugin,
                path,
                default_encoding,
                fallback_encodings,
                used_method=f"plugin:{getattr(plugin, '__name__', plugin.__class__.__name__)}",
            )
            if result.text.strip():
                return result
        except Exception as exc:  # noqa: BLE001
            attempts.append(_build_warning(f"plugin:{getattr(plugin, '__name__', plugin.__class__.__name__)}", exc))

    if suffix in direct_map:
        try:
            result = _run_registered_extractor(
                direct_map[suffix],
                path,
                default_encoding,
                fallback_encodings,
                used_method=f"native:{suffix}",
            )
            if result.text.strip():
                return result
        except Exception as exc:  # noqa: BLE001
            attempts.append(_build_warning(direct_map[suffix].__name__, exc))

    tool_order = list(
        preferred_tools
        or [
            "word",
            "libreoffice",
            "openoffice",
            "antiword",
            "wvtext",
            "catdoc",
            "textutil",
            "calibre",
            "djvutxt",
            "pstotext",
            "chm",
            "ole",
            "strings",
        ]
    )

    for tool in tool_order:
        try:
            if tool == "word":
                text = _extract_via_word_com(path)
            elif tool == "libreoffice":
                text = _extract_via_soffice(path)
            elif tool == "openoffice":
                text = _extract_via_soffice(path, program_hint="openoffice")
            elif tool == "antiword":
                text = _extract_via_antiword(path)
            elif tool == "wvtext":
                text = _extract_via_wvtext(path)
            elif tool == "catdoc":
                text = _extract_via_catdoc_suite(path)
            elif tool == "textutil":
                text = _extract_via_textutil(path)
            elif tool == "calibre":
                text = _extract_via_calibre(path)
            elif tool == "djvutxt":
                text = _extract_via_djvutxt(path)
            elif tool == "pstotext":
                text = _extract_via_pstotext(path)
            elif tool == "chm":
                text = _extract_via_chm(path)
            elif tool == "ole":
                text = _extract_via_ole(path)
            elif tool == "strings":
                text = _extract_via_strings(path)
            else:
                continue

            if text and text.strip():
                return DecodeResult(
                    text=_normalize(text),
                    used_method=tool,
                    source_format=suffix or "",
                    detected_encoding=_detect_text_encoding(path, default_encoding, fallback_encodings),
                    metadata=_collect_metadata(path, text),
                    warnings=attempts.copy(),
                )
        except Exception as exc:  # noqa: BLE001
            attempts.append(_build_warning(tool, exc))

    # Final Python-only fallback: return best-effort text instead of failing hard.
    try:
        text = _extract_via_python_bytes(path)
        if text and text.strip():
            return DecodeResult(
                text=_normalize(text),
                used_method="python-bytes",
                source_format=suffix or "",
                detected_encoding=_detect_text_encoding(path, default_encoding, fallback_encodings),
                metadata=_collect_metadata(path, text),
                warnings=attempts.copy(),
            )
    except Exception as exc:  # noqa: BLE001
        attempts.append(_build_warning("python-bytes", exc))

    raise ExtractorError(
        f"Could not extract text from '{path}'. Attempts: " + " | ".join(
            _warning_to_text(item) for item in (attempts or [{"code": "FALLBACK_USED", "message": "no strategy matched", "source": "core", "install_hint": ""}])
        )
    )


_TOOL_HINTS: dict[str, str] = {
    "word": "Install Microsoft Word and run: pip install all2txt[win]",
    "_extract_via_word_com": "Install Microsoft Word and run: pip install all2txt[win]",
    "libreoffice": "Install LibreOffice and ensure soffice is in PATH",
    "openoffice": "Install OpenOffice and ensure soffice/openoffice4 is in PATH",
    "antiword": "Install antiword and ensure it is in PATH",
    "wvtext": "Install wvText and ensure it is in PATH",
    "catdoc": "Install catdoc suite (catdoc/catppt/xls2csv)",
    "textutil": "On macOS, ensure textutil is available",
    "calibre": "Install Calibre and ensure ebook-convert is in PATH",
    "djvutxt": "Install DjVuLibre and ensure djvutxt/ddjvu are in PATH",
    "pstotext": "Install pstotext or ps2ascii and ensure it is in PATH",
    "chm": "Install extract_chmLib or chm2txt and ensure it is in PATH",
    "ole": "Run: pip install all2txt[ole]",
    "strings": "Install strings utility (or rely on built-in python-bytes fallback)",
    "_read_msg": "Run: pip install all2txt[msg]",
    "_read_pdf": "Run: pip install all2txt[pdf]",
    "_read_mobi": "Run: pip install all2txt[mobi]",
}


def _build_warning(source: str, exc: Exception) -> WarningEntry:
    message = str(exc).strip() or exc.__class__.__name__
    lowered = message.lower()
    code = "FALLBACK_USED"
    install_hint = _TOOL_HINTS.get(source, "")

    if "install optional dependency" in lowered or "pip install all2txt[" in lowered:
        code = "EXTRA_MISSING"
        install_hint = message.split("Install optional dependency:", 1)[-1].strip() if "Install optional dependency:" in message else install_hint
    elif "not found in path" in lowered or "was not found in path" in lowered or "not found" in lowered:
        code = "TOOL_MISSING"

    return {
        "code": code,
        "source": source,
        "message": message,
        "install_hint": install_hint,
    }


def _warning_to_text(item: WarningEntry) -> str:
    hint = item.get("install_hint", "").strip()
    suffix = f" | install_hint: {hint}" if hint else ""
    return f"[{item.get('code', 'FALLBACK_USED')}] {item.get('source', 'unknown')}: {item.get('message', '')}{suffix}"


def _run_registered_extractor(
    extractor: ExtractorFunc,
    path: Path,
    default_encoding: str,
    fallback_encodings: Iterable[str],
    used_method: str,
) -> DecodeResult:
    value = _invoke_extractor(extractor, path, default_encoding, fallback_encodings)

    if isinstance(value, DecodeResult):
        result = value
    else:
        result = DecodeResult(
            text=_normalize(value),
            used_method=used_method,
            source_format=path.suffix.lower(),
            detected_encoding=_detect_text_encoding(path, default_encoding, fallback_encodings),
            metadata=_collect_metadata(path, value),
            warnings=[],
        )

    if not result.metadata:
        result.metadata = _collect_metadata(path, result.text)
    result.source_format = result.source_format or path.suffix.lower()
    result.used_method = result.used_method or used_method
    return result


def _invoke_extractor(
    extractor: ExtractorFunc,
    path: Path,
    default_encoding: str,
    fallback_encodings: Iterable[str],
) -> str | DecodeResult:
    params = inspect.signature(extractor).parameters
    positional = [
        parameter
        for parameter in params.values()
        if parameter.kind in (parameter.POSITIONAL_ONLY, parameter.POSITIONAL_OR_KEYWORD)
    ]
    if len(positional) >= 3:
        return extractor(path, default_encoding, fallback_encodings)
    return extractor(path, default_encoding)


def _normalize(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def _build_encoding_candidates(default_encoding: str, extra_encodings: Iterable[str] | None = None) -> list[str]:
    ordered = [default_encoding, "utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1251", "koi8-r", "cp866", "latin-1"]
    if extra_encodings:
        ordered[1:1] = list(extra_encodings)
    seen: list[str] = []
    for item in ordered:
        if item and item not in seen:
            seen.append(item)
    return seen


def _read_text_file_with_encoding(
    path: Path,
    default_encoding: str,
    extra_encodings: Iterable[str] | None = None,
) -> tuple[str, str | None]:
    for enc in _build_encoding_candidates(default_encoding, extra_encodings):
        try:
            return path.read_text(encoding=enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return path.read_text(encoding=default_encoding, errors="replace"), default_encoding


def _detect_text_encoding(
    path: Path,
    default_encoding: str,
    extra_encodings: Iterable[str] | None = None,
) -> str | None:
    if path.suffix.lower() not in {
        ".txt", ".log", ".ini", ".conf", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm",
        ".mht", ".mhtml", ".eml", ".rtf", ".tex", ".bib", ".strings",
    }:
        return None
    _, detected = _read_text_file_with_encoding(path, default_encoding, extra_encodings)
    return detected


def _guess_language(text: str) -> str | None:
    letters = [char for char in text if char.isalpha()]
    if not letters:
        return None
    cyr = sum(1 for char in letters if "Рђ" <= char <= "СЏ" or char in "РЃС‘")
    lat = sum(1 for char in letters if ("A" <= char <= "Z") or ("a" <= char <= "z"))
    total = len(letters)
    if cyr / total > 0.4:
        return "ru"
    if lat / total > 0.5:
        return "en"
    return None


def _collect_metadata(path: Path, text: str) -> dict[str, object]:
    stat = path.stat()
    metadata: dict[str, object] = {
        "source_path": str(path),
        "source_name": path.name,
        "source_format": path.suffix.lower(),
        "size_bytes": stat.st_size,
        "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
        "language": _guess_language(text),
        "title": path.stem,
    }
    metadata.update(_extract_document_metadata(path))
    return metadata


def _extract_document_metadata(path: Path) -> dict[str, object]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_docx_metadata(path)
    if suffix in {".odt", ".ods"}:
        return _extract_odf_metadata(path)
    if suffix == ".epub":
        return _extract_epub_metadata(path)
    if suffix == ".pdf":
        return _extract_pdf_metadata(path)
    if suffix in {".eml", ".mht", ".mhtml"}:
        return _extract_email_like_metadata(path)
    if suffix == ".plist":
        return _extract_plist_metadata(path)
    return {}


def _extract_docx_metadata(path: Path) -> dict[str, object]:
    result: dict[str, object] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            raw = zf.read("docProps/core.xml")
        root = ET.fromstring(raw)
        ns = {
            "dc": "http://purl.org/dc/elements/1.1/",
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dcterms": "http://purl.org/dc/terms/",
        }
        title = root.findtext("dc:title", default="", namespaces=ns)
        creator = root.findtext("dc:creator", default="", namespaces=ns)
        created = root.findtext("dcterms:created", default="", namespaces=ns)
        if title:
            result["title"] = title
        if creator:
            result["author"] = creator
        if created:
            result["date"] = created
    except Exception:
        return {}
    return result


def _extract_odf_metadata(path: Path) -> dict[str, object]:
    result: dict[str, object] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            raw = zf.read("meta.xml")
        root = ET.fromstring(raw)
        pairs = {
            "{http://purl.org/dc/elements/1.1/}title": "title",
            "{http://purl.org/dc/elements/1.1/}creator": "author",
            "{http://purl.org/dc/elements/1.1/}date": "date",
        }
        for tag, key in pairs.items():
            node = root.find(f".//{tag}")
            if node is not None and node.text:
                result[key] = node.text
    except Exception:
        return {}
    return result


def _extract_epub_metadata(path: Path) -> dict[str, object]:
    result: dict[str, object] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            opf_path = _epub_find_opf(zf)
            if not opf_path:
                return {}
            root = ET.fromstring(zf.read(opf_path))
            title = root.findtext(".//{http://purl.org/dc/elements/1.1/}title")
            creator = root.findtext(".//{http://purl.org/dc/elements/1.1/}creator")
            date = root.findtext(".//{http://purl.org/dc/elements/1.1/}date")
            language = root.findtext(".//{http://purl.org/dc/elements/1.1/}language")
            if title:
                result["title"] = title
            if creator:
                result["author"] = creator
            if date:
                result["date"] = date
            if language:
                result["language"] = language
    except Exception:
        return {}
    return result


def _extract_pdf_metadata(path: Path) -> dict[str, object]:
    try:
        from pypdf import PdfReader
    except ImportError:
        return {}

    try:
        reader = PdfReader(str(path))
    except Exception:
        return {}

    meta = reader.metadata or {}
    result: dict[str, object] = {"page_count": len(reader.pages)}
    if meta.get("/Title"):
        result["title"] = meta["/Title"]
    if meta.get("/Author"):
        result["author"] = meta["/Author"]
    if meta.get("/CreationDate"):
        result["date"] = meta["/CreationDate"]
    return result


def _extract_email_like_metadata(path: Path) -> dict[str, object]:
    try:
        msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    except Exception:
        return {}
    result: dict[str, object] = {}
    for header, target in (("Subject", "subject"), ("From", "from"), ("To", "to"), ("Date", "date")):
        value = msg.get(header)
        if value:
            result[target] = value
    if result.get("subject") and not result.get("title"):
        result["title"] = result["subject"]
    return result


def _extract_plist_metadata(path: Path) -> dict[str, object]:
    try:
        obj = plistlib.loads(path.read_bytes())
    except Exception:
        return {}
    result: dict[str, object] = {}
    for key, target in (("CFBundleName", "title"), ("Author", "author"), ("CFBundleShortVersionString", "version")):
        value = obj.get(key)
        if value:
            result[target] = value
    return result


def _read_text_file(
    path: Path,
    default_encoding: str,
    fallback_encodings: Iterable[str] | None = None,
) -> str:
    text, _ = _read_text_file_with_encoding(path, default_encoding, fallback_encodings)
    return text


def _read_csv_as_text(path: Path, default_encoding: str, fallback_encodings: Iterable[str] | None = None) -> str:
    content = _read_text_file(path, default_encoding, fallback_encodings)
    sep = "\t" if path.suffix.lower() == ".tsv" else ","
    lines = []
    for row in csv.reader(content.splitlines(), delimiter=sep):
        lines.append("\t".join(cell.strip() for cell in row))
    return "\n".join(lines)


def _read_json_as_text(path: Path, default_encoding: str, fallback_encodings: Iterable[str] | None = None) -> str:
    raw = _read_text_file(path, default_encoding, fallback_encodings)
    parsed = json.loads(raw)
    return json.dumps(parsed, ensure_ascii=False, indent=2)


def _read_markup_as_text(path: Path, default_encoding: str, fallback_encodings: Iterable[str] | None = None) -> str:
    raw = _read_text_file(path, default_encoding, fallback_encodings)
    no_script = re.sub(r"<script[^>]*>.*?</script>", " ", raw, flags=re.IGNORECASE | re.DOTALL)
    no_style = re.sub(r"<style[^>]*>.*?</style>", " ", no_script, flags=re.IGNORECASE | re.DOTALL)
    no_tags = re.sub(r"<[^>]+>", " ", no_style)
    return html.unescape(re.sub(r"[ \t]+", " ", no_tags))


def _read_rtf_as_text(path: Path, default_encoding: str, fallback_encodings: Iterable[str] | None = None) -> str:
    raw = _read_text_file(path, default_encoding, fallback_encodings)
    # A lightweight RTF cleanup; for complex RTF consider dedicated parser.
    text = re.sub(r"\\'[0-9a-fA-F]{2}", " ", raw)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)
    text = text.replace("{", "").replace("}", "")
    return text


def _read_plist(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    obj = plistlib.loads(path.read_bytes())
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)


def _read_eml(path: Path, default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    chunks: list[str] = []
    for header in ("Subject", "From", "To", "Date"):
        value = msg.get(header)
        if value:
            chunks.append(f"{header}: {value}")

    plain_parts: list[str] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype.startswith("multipart/"):
                continue
            payload = part.get_payload(decode=True) or b""
            if ctype == "text/plain":
                charset = part.get_content_charset() or default_encoding
                try:
                    plain_parts.append(payload.decode(charset, errors="replace"))
                except LookupError:
                    plain_parts.append(payload.decode(default_encoding, errors="replace"))
            elif ctype in {"text/html", "application/xhtml+xml"}:
                html_parts.append(_strip_html_bytes(payload))
    else:
        payload = msg.get_payload(decode=True) or b""
        if msg.get_content_type() == "text/html":
            html_parts.append(_strip_html_bytes(payload))
        else:
            charset = msg.get_content_charset() or default_encoding
            plain_parts.append(payload.decode(charset, errors="replace"))

    body = plain_parts if plain_parts else html_parts
    chunks.extend(part.strip() for part in body if part and part.strip())
    return "\n\n".join(chunks)


def _read_mhtml(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    text_parts: list[str] = []
    html_parts: list[str] = []
    for part in msg.walk():
        ctype = part.get_content_type()
        payload = part.get_payload(decode=True) or b""
        if ctype == "text/plain":
            charset = part.get_content_charset() or "utf-8"
            text_parts.append(payload.decode(charset, errors="replace"))
        elif ctype in {"text/html", "application/xhtml+xml"}:
            html_parts.append(_strip_html_bytes(payload))
    source = text_parts if text_parts else html_parts
    return "\n\n".join(s for s in source if s and s.strip())


def _read_msg(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    try:
        import extract_msg  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Install optional dependency: pip install all2txt[msg]") from exc

    message = extract_msg.Message(str(path))
    chunks = []
    for label, value in (
        ("Subject", message.subject),
        ("From", message.sender),
        ("To", message.to),
        ("Date", message.date),
    ):
        if value:
            chunks.append(f"{label}: {value}")
    if message.body:
        chunks.append(message.body)
    return "\n\n".join(chunks)


def _read_legacy_doc(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_word_com, _extract_via_soffice, _extract_via_antiword, _extract_via_wvtext, _extract_via_catdoc_suite, _extract_via_ole, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No DOC extractor succeeded")


def _read_legacy_xls(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_soffice, _extract_via_catdoc_suite, _extract_via_ole, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No XLS extractor succeeded")


def _read_legacy_ppt(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_soffice, _extract_via_catdoc_suite, _extract_via_ole, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No PPT extractor succeeded")


def _read_iwork(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    """Best-effort extraction from Apple iWork archives (.pages/.numbers/.key)."""
    try:
        text = _extract_via_textutil(path)
        if text.strip():
            return text
    except Exception:
        pass

    if not zipfile.is_zipfile(path):
        raise RuntimeError("iWork file is expected to be a zip-based package")

    chunks: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        xml_like = [n for n in names if n.lower().endswith((".xml", ".html", ".htm", ".txt"))]
        for name in xml_like:
            try:
                data = zf.read(name)
                if name.lower().endswith((".html", ".htm")):
                    text = _strip_html_bytes(data)
                else:
                    text = _decode_bytes_relaxed(data)
                if text.strip():
                    chunks.append(text)
            except Exception:
                continue

        if not chunks:
            # New iWork stores content in .iwa protobuf chunks; salvage printable text.
            for name in names:
                if not name.lower().endswith(".iwa"):
                    continue
                data = zf.read(name)
                recovered = "\n".join(
                    m.decode("utf-8", errors="ignore")
                    for m in re.findall(rb"[\x20-\x7E\xC0-\xFF]{8,}", data)
                )
                if recovered.strip():
                    chunks.append(recovered)
    return "\n\n".join(chunks)


def _read_djvu(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_djvutxt, _extract_via_calibre, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No DjVu extractor succeeded")


def _read_postscript(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_pstotext, _extract_via_strings, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No PostScript extractor succeeded")


def _read_chm(path: Path, _default_encoding: str, _fallback_encodings: Iterable[str] | None = None) -> str:
    for extractor in (_extract_via_chm, _extract_via_strings, _extract_via_python_bytes):
        try:
            text = extractor(path)
            if text.strip():
                return text
        except Exception:
            continue
    raise RuntimeError("No CHM extractor succeeded")


def _read_docx(path: Path, _default_encoding: str) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    texts = [node.text or "" for node in root.findall(".//w:t", ns)]
    return " ".join(texts)


def _read_odt(path: Path, _default_encoding: str) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("content.xml")
    root = ET.fromstring(xml)
    texts = [node.text for node in root.iter() if node.text]
    return "\n".join(texts)


def _read_pdf(path: Path, _default_encoding: str) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("Install optional dependency: pip install all2txt[pdf]") from exc

    reader = PdfReader(str(path))
    pages = [(page.extract_text() or "") for page in reader.pages]
    return "\n\n".join(pages)


def _extract_via_word_com(path: Path) -> str:
    if path.suffix.lower() not in {".doc", ".docx", ".rtf", ".odt"}:
        raise RuntimeError("Word COM method supports office-like docs only")

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Install optional dependency: pip install all2txt[win]") from exc

    temp_txt = None
    word = None
    doc = None
    try:
        temp_dir = Path(tempfile.mkdtemp(prefix="all2txt_word_"))
        temp_txt = temp_dir / f"{path.stem}.txt"

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(path.resolve()))
        # 2 = wdFormatText
        doc.SaveAs(str(temp_txt), FileFormat=2)
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()

    if not temp_txt or not temp_txt.exists():
        raise RuntimeError("Word did not produce a text file")

    return _read_text_file(temp_txt, "utf-8")


def _extract_via_soffice(path: Path, program_hint: str | None = None) -> str:
    candidates = ["soffice", "libreoffice", "openoffice4", "swriter", "lowriter"]

    if program_hint == "openoffice":
        candidates = ["openoffice4", "soffice", "swriter"]

    exe = None
    for cmd in candidates:
        exe = shutil.which(cmd)
        if exe:
            break

    if not exe:
        raise RuntimeError("LibreOffice/OpenOffice executable was not found in PATH")

    with tempfile.TemporaryDirectory(prefix="all2txt_soffice_") as tmp:
        out_dir = Path(tmp)
        command = [exe, "--headless", "--convert-to", "txt:Text", "--outdir", str(out_dir), str(path)]
        result = subprocess.run(command, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "soffice convert failed")

        txt_path = out_dir / f"{path.stem}.txt"
        if not txt_path.exists():
            generated = list(out_dir.glob("*.txt"))
            if not generated:
                raise RuntimeError("soffice did not create txt output")
            txt_path = generated[0]

        return _read_text_file(txt_path, "utf-8")


def _extract_via_antiword(path: Path) -> str:
    if path.suffix.lower() != ".doc":
        raise RuntimeError("antiword supports mostly .doc")

    exe = shutil.which("antiword")
    if not exe:
        raise RuntimeError("antiword not found in PATH")

    result = subprocess.run([exe, str(path)], capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "antiword conversion failed")
    return result.stdout


def _extract_via_wvtext(path: Path) -> str:
    if path.suffix.lower() != ".doc":
        raise RuntimeError("wvText supports mainly .doc")
    exe = shutil.which("wvText") or shutil.which("wvtext")
    if not exe:
        raise RuntimeError("wvText not found in PATH")
    result = subprocess.run([exe, str(path), "-"], capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "wvText conversion failed")
    return result.stdout


def _extract_via_catdoc_suite(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".doc":
        exe = shutil.which("catdoc")
        cmd = [exe, str(path)] if exe else None
    elif suffix == ".ppt":
        exe = shutil.which("catppt")
        cmd = [exe, str(path)] if exe else None
    elif suffix == ".xls":
        exe = shutil.which("xls2csv")
        cmd = [exe, str(path)] if exe else None
    else:
        raise RuntimeError("catdoc suite supports .doc/.ppt/.xls")

    if not cmd:
        raise RuntimeError("Required catdoc-suite tool was not found in PATH")

    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "catdoc-suite conversion failed")
    return result.stdout


def _extract_via_textutil(path: Path) -> str:
    exe = shutil.which("textutil")
    if not exe:
        raise RuntimeError("textutil not found (macOS only)")
    with tempfile.TemporaryDirectory(prefix="all2txt_textutil_") as tmp:
        out = Path(tmp) / f"{path.stem}.txt"
        result = subprocess.run(
            [exe, "-convert", "txt", "-output", str(out), str(path)],
            capture_output=True, text=True, check=False,
        )
        if result.returncode != 0 or not out.exists():
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "textutil conversion failed")
        return _read_text_file(out, "utf-8")


def _extract_via_ole(path: Path) -> str:
    if path.suffix.lower() not in {".doc", ".xls", ".ppt"}:
        raise RuntimeError("OLE fallback is useful for legacy MS Office binaries")

    try:
        import olefile  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Install optional dependency: pip install all2txt[ole]") from exc

    if not olefile.isOleFile(str(path)):
        raise RuntimeError("Not an OLE compound file")

    chunks = []
    with olefile.OleFileIO(str(path)) as ole:
        for stream_name in ole.listdir(streams=True, storages=False):
            try:
                data = ole.openstream(stream_name).read()
            except Exception:
                continue

            decoded = _decode_bytes_relaxed(data)
            if _looks_like_text(decoded):
                chunks.append(decoded)

    if not chunks:
        raise RuntimeError("No text-like streams found in OLE")

    return "\n\n".join(chunks)


def _decode_bytes_relaxed(data: bytes) -> str:
    for enc in ("utf-8", "cp1251", "utf-16", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _looks_like_text(text: str) -> bool:
    if not text:
        return False
    printable = sum(1 for ch in text if ch.isprintable() or ch in "\n\r\t")
    ratio = printable / max(len(text), 1)
    return ratio > 0.85 and len(text.strip()) > 20


# ---------------------------------------------------------------------------
# NEW FORMAT EXTRACTORS
# ---------------------------------------------------------------------------

def _read_fb2(path: Path, _default_encoding: str) -> str:
    """FictionBook 2.x вЂ“ XML-based, no external dependencies."""
    raw = path.read_bytes()
    root = ET.fromstring(raw)
    ns = "http://www.gribuser.ru/xml/fictionbook/2.0"
    texts: list[str] = []
    # Try namespaced paragraphs first
    for el in root.iter(f"{{{ns}}}p"):
        if el.text and el.text.strip():
            texts.append(el.text.strip())
    if not texts:
        # Fallback: grab all text nodes regardless of namespace
        texts = [el.text.strip() for el in root.iter() if el.text and el.text.strip()]
    return "\n".join(texts)


def _strip_html_bytes(data: bytes) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1251", "latin-1"):
        try:
            raw = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raw = data.decode("latin-1", errors="replace")
    no_script = re.sub(r"<script[^>]*>.*?</script>", " ", raw, flags=re.IGNORECASE | re.DOTALL)
    no_style = re.sub(r"<style[^>]*>.*?</style>", " ", no_script, flags=re.IGNORECASE | re.DOTALL)
    no_tags = re.sub(r"<[^>]+>", " ", no_style)
    return html.unescape(re.sub(r"[ \t]+", " ", no_tags)).strip()


def _epub_find_opf(zf: zipfile.ZipFile) -> str | None:
    try:
        container = ET.fromstring(zf.read("META-INF/container.xml"))
        ns = "urn:oasis:names:tc:opendocument:xmlns:container"
        rootfile = container.find(f".//{{{ns}}}rootfile")
        if rootfile is not None:
            return rootfile.get("full-path")
    except Exception:  # noqa: BLE001
        pass
    for name in zf.namelist():
        if name.endswith(".opf"):
            return name
    return None


def _epub_spine_paths(zf: zipfile.ZipFile, opf_path: str) -> list[str]:
    try:
        opf_dir = "/".join(opf_path.split("/")[:-1])
        opf_root = ET.fromstring(zf.read(opf_path))
        ns_opf = "http://www.idpf.org/2007/opf"
        manifest: dict[str, str] = {}
        for item in opf_root.findall(f".//{{{ns_opf}}}item"):
            item_id = item.get("id")
            href = item.get("href")
            media_type = item.get("media-type", "")
            if item_id and href and "html" in media_type:
                full = (opf_dir + "/" + href).lstrip("/") if opf_dir else href
                manifest[item_id] = full
        result = []
        for itemref in opf_root.findall(f".//{{{ns_opf}}}itemref"):
            idref = itemref.get("idref")
            if idref and idref in manifest:
                result.append(manifest[idref])
        return result
    except Exception:  # noqa: BLE001
        return []


def _read_epub(path: Path, _default_encoding: str) -> str:
    """EPUB 2/3 вЂ“ ZIP of HTML/XHTML, no external dependencies."""
    texts: list[str] = []
    with zipfile.ZipFile(path) as zf:
        opf_path = _epub_find_opf(zf)
        item_paths = _epub_spine_paths(zf, opf_path) if opf_path else []
        if not item_paths:
            item_paths = sorted(n for n in zf.namelist() if n.lower().endswith((".html", ".xhtml", ".htm")))
        for item in item_paths:
            try:
                texts.append(_strip_html_bytes(zf.read(item)))
            except Exception:  # noqa: BLE001
                pass
    return "\n\n".join(t for t in texts if t.strip())


def _read_mobi(path: Path, _default_encoding: str) -> str:
    """MOBI/AZW вЂ“ requires optional 'mobi' package; Calibre is a better fallback."""
    try:
        import mobi  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Install optional dependency: pip install all2txt[mobi]; or install Calibre for ebook-convert fallback") from exc

    with tempfile.TemporaryDirectory(prefix="all2txt_mobi_") as tmp:
        tempdir, extracted_path = mobi.extract(str(path))
        extracted = Path(extracted_path)
        html_files = sorted(extracted.rglob("*.html")) + sorted(extracted.rglob("*.htm"))
        txt_files = sorted(extracted.rglob("*.txt"))
        parts: list[str] = []
        if html_files:
            for f in html_files:
                parts.append(_strip_html_bytes(f.read_bytes()))
        elif txt_files:
            for f in txt_files:
                parts.append(_read_text_file(f, "utf-8"))
        return "\n\n".join(p for p in parts if p.strip())


def _read_xlsx(path: Path, _default_encoding: str) -> str:
    """Excel XLSX вЂ“ ZIP of XML, no external dependencies."""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    shared: list[str] = []
    rows: list[str] = []
    with zipfile.ZipFile(path) as zf:
        if "xl/sharedStrings.xml" in zf.namelist():
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{ns}}}si"):
                parts = [t.text or "" for t in si.findall(f".//{{{ns}}}t")]
                shared.append("".join(parts))
        sheet_names = sorted(
            (n for n in zf.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
            key=lambda x: int(re.search(r"\d+", x).group()),  # type: ignore[union-attr]
        )
        for name in sheet_names:
            sheet_root = ET.fromstring(zf.read(name))
            for row in sheet_root.findall(f".//{{{ns}}}row"):
                cells: list[str] = []
                for c in row.findall(f"{{{ns}}}c"):
                    t = c.get("t")
                    v_el = c.find(f"{{{ns}}}v")
                    if v_el is None or v_el.text is None:
                        cells.append("")
                    elif t == "s":
                        idx = int(v_el.text)
                        cells.append(shared[idx] if idx < len(shared) else "")
                    elif t == "inlineStr":
                        is_el = c.find(f".//{{{ns}}}t")
                        cells.append(is_el.text if is_el is not None else "")
                    else:
                        cells.append(v_el.text)
                line = "\t".join(cells).strip()
                if line:
                    rows.append(line)
    return "\n".join(rows)


def _read_pptx(path: Path, _default_encoding: str) -> str:
    """PowerPoint PPTX вЂ“ ZIP of XML, no external dependencies."""
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    texts: list[str] = []
    with zipfile.ZipFile(path) as zf:
        slide_names = sorted(
            (n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda x: int(re.search(r"\d+", x).group()),  # type: ignore[union-attr]
        )
        for name in slide_names:
            root = ET.fromstring(zf.read(name))
            slide_texts = [el.text for el in root.findall(f".//{{{ns}}}t") if el.text]
            if slide_texts:
                texts.append(" ".join(slide_texts))
    return "\n\n".join(texts)


def _read_ods(path: Path, _default_encoding: str) -> str:
    """OpenDocument Spreadsheet вЂ“ ZIP of XML, no external dependencies."""
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("content.xml")
    root = ET.fromstring(xml)
    ns_t = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    ns_tx = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    rows: list[str] = []
    for row in root.findall(f".//{{{ns_t}}}table-row"):
        cells: list[str] = []
        for cell in row.findall(f"{{{ns_t}}}table-cell"):
            parts = [p.text or "" for p in cell.findall(f".//{{{ns_tx}}}p")]
            cells.append(" ".join(parts).strip())
        line = "\t".join(cells).strip()
        if line:
            rows.append(line)
    return "\n".join(rows)


def _extract_via_calibre(path: Path) -> str:
    """Use Calibre's ebook-convert for EPUB, MOBI, DJVU, AZW and many others."""
    exe = shutil.which("ebook-convert")
    if not exe:
        raise RuntimeError("ebook-convert (Calibre) not found in PATH")
    with tempfile.TemporaryDirectory(prefix="all2txt_calibre_") as tmp:
        out_txt = Path(tmp) / f"{path.stem}.txt"
        result = subprocess.run(
            [exe, str(path), str(out_txt)],
            capture_output=True, text=True, check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "ebook-convert failed")
        if not out_txt.exists():
            generated = list(Path(tmp).glob("*.txt"))
            if not generated:
                raise RuntimeError("Calibre did not produce a txt file")
            out_txt = generated[0]
        return _read_text_file(out_txt, "utf-8")


def _extract_via_djvutxt(path: Path) -> str:
    """Extract text from DjVu via djvutxt CLI (part of DjVuLibre)."""
    if path.suffix.lower() not in {".djvu", ".djv"}:
        raise RuntimeError("djvutxt handles .djvu/.djv files only")
    exe = shutil.which("djvutxt")
    if not exe:
        raise RuntimeError("djvutxt not found in PATH (install DjVuLibre)")
    result = subprocess.run([exe, str(path)], capture_output=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.decode(errors="replace").strip() or "djvutxt failed")
    return _decode_bytes_relaxed(result.stdout)


def _extract_via_pstotext(path: Path) -> str:
    if path.suffix.lower() not in {".ps", ".eps"}:
        raise RuntimeError("pstotext supports .ps/.eps")
    exe = shutil.which("pstotext") or shutil.which("ps2ascii")
    if not exe:
        raise RuntimeError("pstotext/ps2ascii not found in PATH")
    result = subprocess.run([exe, str(path)], capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "PostScript conversion failed")
    return result.stdout


def _extract_via_chm(path: Path) -> str:
    if path.suffix.lower() != ".chm":
        raise RuntimeError("CHM extractor supports .chm only")
    exe = shutil.which("extract_chmLib") or shutil.which("chm2txt")
    if not exe:
        raise RuntimeError("extract_chmLib/chm2txt not found in PATH")

    if Path(exe).name.lower().startswith("chm2txt"):
        result = subprocess.run([exe, str(path)], capture_output=True, text=True, check=False)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "chm2txt failed")
        return result.stdout

    with tempfile.TemporaryDirectory(prefix="all2txt_chm_") as tmp:
        out_dir = Path(tmp)
        result = subprocess.run([exe, str(path), str(out_dir)], capture_output=True, text=True, check=False)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "extract_chmLib failed")
        parts = []
        for html_file in out_dir.rglob("*.htm"):
            try:
                parts.append(_strip_html_bytes(html_file.read_bytes()))
            except Exception:
                continue
        for html_file in out_dir.rglob("*.html"):
            try:
                parts.append(_strip_html_bytes(html_file.read_bytes()))
            except Exception:
                continue
        return "\n\n".join(p for p in parts if p.strip())


def _extract_via_strings(path: Path) -> str:
    exe = shutil.which("strings")
    if not exe:
        return _extract_via_python_bytes(path)
    result = subprocess.run([exe, str(path)], capture_output=True, check=False)
    if result.returncode != 0:
        return _extract_via_python_bytes(path)
    return _decode_bytes_relaxed(result.stdout)


def _extract_via_python_bytes(path: Path) -> str:
    """Pure Python best-effort text recovery from arbitrary binary files."""
    data = path.read_bytes()
    if not data:
        raise RuntimeError("empty file")

    candidates: list[str] = []
    for enc in ("utf-8", "cp1251", "koi8-r", "cp866", "utf-16", "utf-16-le", "utf-16-be", "latin-1"):
        try:
            txt = data.decode(enc, errors="ignore")
        except LookupError:
            continue
        cleaned = _cleanup_recovered_text(txt)
        if cleaned:
            candidates.append(cleaned)

    # Collect printable byte sequences as an additional salvage channel.
    chunks = re.findall(rb"[\x09\x0A\x0D\x20-\x7E\x80-\xFF]{8,}", data)
    if chunks:
        joined = b"\n".join(chunks)
        for enc in ("utf-8", "cp1251", "koi8-r", "cp866", "latin-1"):
            try:
                txt = joined.decode(enc, errors="ignore")
            except LookupError:
                continue
            cleaned = _cleanup_recovered_text(txt)
            if cleaned:
                candidates.append(cleaned)

    if not candidates:
        raise RuntimeError("no recoverable text in binary payload")

    best = max(candidates, key=_recovered_text_score)
    if len(best.strip()) < 20:
        return f"[all2txt] best-effort binary fallback for {path.name}\n{best}"
    return best


def _cleanup_recovered_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[\x01-\x08\x0b\x0c\x0e-\x1f]+", " ", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _recovered_text_score(text: str) -> float:
    if not text:
        return 0.0
    printable = sum(1 for ch in text if ch.isprintable() or ch in "\n\r\t")
    letter_count = sum(1 for ch in text if ch.isalpha())
    ratio_printable = printable / max(len(text), 1)
    ratio_letters = letter_count / max(len(text), 1)
    return ratio_printable * 0.8 + ratio_letters * 0.2

