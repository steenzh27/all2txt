from __future__ import annotations

import argparse
import csv
import importlib
import json
import shutil
from pathlib import Path

from .core import TextDecoder
from .extractors import ExtractorError


_EXTRA_MODULES: dict[str, list[str]] = {
    "pdf": ["pypdf"],
    "win": ["win32com.client"],
    "ole": ["olefile"],
    "mobi": ["mobi"],
    "msg": ["extract_msg"],
    "ocr": ["all2txt_ocr_plugin", "pypdf"],
}

_EXTRA_INSTALL_HINTS: dict[str, str] = {
    "pdf": "pip install all2txt[pdf]",
    "win": "pip install all2txt[win]",
    "ole": "pip install all2txt[ole]",
    "mobi": "pip install all2txt[mobi]",
    "msg": "pip install all2txt[msg]",
    "ocr": "pip install all2txt[ocr]",
}

_TOOL_ALIASES: dict[str, list[str]] = {
    "libreoffice": ["soffice", "libreoffice", "lowriter"],
    "openoffice": ["openoffice4", "soffice", "swriter"],
    "antiword": ["antiword"],
    "wvtext": ["wvText", "wvtext"],
    "catdoc": ["catdoc"],
    "catppt": ["catppt"],
    "xls2csv": ["xls2csv"],
    "textutil": ["textutil"],
    "calibre": ["ebook-convert"],
    "djvu": ["djvutxt", "ddjvu"],
    "postscript": ["pstotext", "ps2ascii"],
    "chm": ["extract_chmLib", "chm2txt"],
    "strings": ["strings"],
    "tesseract": ["tesseract"],
    "ocrmypdf": ["ocrmypdf"],
    "pdftoppm": ["pdftoppm"],
    "magick": ["magick"],
}

_TOOL_INSTALL_HINTS: dict[str, list[str]] = {
    "libreoffice": [
        "winget install TheDocumentFoundation.LibreOffice",
    ],
    "openoffice": [
        "winget install Apache.OpenOffice",
    ],
    "antiword": [
        "choco install antiword",
    ],
    "wvtext": [
        "Install wv package and add wvText to PATH",
    ],
    "catdoc": [
        "choco install catdoc",
    ],
    "catppt": [
        "choco install catdoc",
    ],
    "xls2csv": [
        "choco install catdoc",
    ],
    "textutil": [
        "textutil is built-in on macOS",
    ],
    "calibre": [
        "winget install calibre.calibre",
    ],
    "djvu": [
        "winget install DjVuLibre.DjVuLibre",
    ],
    "postscript": [
        "Install pstotext/ps2ascii and add to PATH",
    ],
    "chm": [
        "Install chm tools (extract_chmLib/chm2txt) and add to PATH",
    ],
    "strings": [
        "Install GNU binutils and add strings.exe to PATH",
    ],
    "tesseract": [
        "winget install UB-Mannheim.TesseractOCR",
    ],
    "ocrmypdf": [
        "pip install ocrmypdf",
    ],
    "pdftoppm": [
        "Install poppler and add pdftoppm to PATH",
    ],
    "magick": [
        "winget install ImageMagick.ImageMagick",
    ],
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="all2txt",
        description="Extract text from many formats and save as TXT",
    )
    parser.add_argument("input", nargs="?", help="Input file or directory")
    parser.add_argument("-o", "--output", help="Output txt file or directory")
    parser.add_argument(
        "--tools",
        nargs="*",
        default=None,
        help="Preferred tools order, e.g. --tools word libreoffice ole",
    )
    parser.add_argument(
        "--method-order",
        nargs="*",
        dest="tools",
        help="Alias for --tools",
    )
    parser.add_argument(
        "--glob",
        default="*",
        help="Glob pattern when input is a directory (default: *)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be processed without writing files")
    parser.add_argument("--report", help="Write CSV report with method, encoding, metadata and status")
    parser.add_argument(
        "--failed-only",
        action="store_true",
        help="Retry mode: process only files that do not yet have output .txt files",
    )
    parser.add_argument(
        "--keep-structure",
        action="store_true",
        help="Preserve input directory structure inside the output directory",
    )
    parser.add_argument("--input-encoding", default="utf-8", help="Preferred input encoding for text-like files")
    parser.add_argument(
        "--fallback-encodings",
        nargs="*",
        default=None,
        help="Additional encodings to try, e.g. cp1251 koi8-r cp866",
    )
    parser.add_argument("--output-encoding", default="utf-8", help="Encoding for written .txt files")
    parser.add_argument(
        "--doctor",
        "--available",
        "--help-env",
        action="store_true",
        dest="doctor",
        help="Show availability report: extras, external tools and format levels",
    )
    return parser


def _has_module(module_name: str) -> bool:
    try:
        importlib.import_module(module_name)
        return True
    except Exception:  # noqa: BLE001
        return False


def _detect_extras() -> tuple[dict[str, bool], dict[str, list[str]]]:
    status: dict[str, bool] = {}
    missing: dict[str, list[str]] = {}
    for extra, modules in _EXTRA_MODULES.items():
        missing_modules = [mod for mod in modules if not _has_module(mod)]
        status[extra] = not missing_modules
        missing[extra] = missing_modules
    return status, missing


def _detect_tools() -> tuple[dict[str, str], dict[str, list[str]]]:
    found: dict[str, str] = {}
    missing: dict[str, list[str]] = {}
    for tool, aliases in _TOOL_ALIASES.items():
        resolved = ""
        for alias in aliases:
            candidate = shutil.which(alias)
            if candidate:
                resolved = candidate
                break
        found[tool] = resolved
        if not resolved:
            missing[tool] = aliases
    return found, missing


def _format_levels(extra_status: dict[str, bool], tool_status: dict[str, str]) -> list[tuple[str, str, str]]:
    has_lo = bool(tool_status.get("libreoffice") or tool_status.get("openoffice"))
    has_cat_suite = bool(tool_status.get("catdoc") and tool_status.get("catppt") and tool_status.get("xls2csv"))
    has_ocr = bool(extra_status.get("ocr") and tool_status.get("tesseract"))

    rows: list[tuple[str, str, str]] = [
        (".txt .log .ini .conf .md .rst .csv .tsv .json .xml .html .htm .plist .tex .bib .strings", "native", "pure python parser"),
        (".docx .odt .ods .xlsx .pptx .fb2 .epub", "native", "zip/xml parser"),
    ]

    if extra_status.get("pdf"):
        rows.append((".pdf", "native+extra", "pypdf available"))
    elif has_ocr:
        rows.append((".pdf", "plugin-ocr", "OCR plugin via tesseract"))
    elif has_lo or tool_status.get("calibre"):
        rows.append((".pdf", "tool", "converter in PATH"))
    else:
        rows.append((".pdf", "fallback", "python-bytes best effort"))

    if extra_status.get("msg"):
        rows.append((".msg", "native+extra", "extract-msg available"))
    else:
        rows.append((".msg", "fallback", "python-bytes best effort"))

    if extra_status.get("mobi") or tool_status.get("calibre"):
        rows.append((".mobi", "native+extra/tool", "mobi package or calibre"))
    else:
        rows.append((".mobi", "fallback", "python-bytes best effort"))

    if tool_status.get("djvu") or tool_status.get("calibre"):
        rows.append((".djvu .djv", "tool", "djvutxt/ddjvu or calibre"))
    elif has_ocr:
        rows.append((".djvu .djv", "plugin-ocr", "OCR plugin"))
    else:
        rows.append((".djvu .djv", "fallback", "python-bytes best effort"))

    if tool_status.get("postscript"):
        rows.append((".ps .eps", "tool", "pstotext/ps2ascii"))
    else:
        rows.append((".ps .eps", "fallback", "python-bytes best effort"))

    if tool_status.get("chm"):
        rows.append((".chm", "tool", "extract_chmLib/chm2txt"))
    else:
        rows.append((".chm", "fallback", "python-bytes best effort"))

    if extra_status.get("win") or has_lo or tool_status.get("antiword") or tool_status.get("wvtext") or has_cat_suite or extra_status.get("ole"):
        rows.append((".doc .xls .ppt", "tool/native+extra", "legacy office chain available"))
    else:
        rows.append((".doc .xls .ppt", "fallback", "python-bytes best effort"))

    return rows


def _print_availability_report() -> int:
    extra_status, extra_missing = _detect_extras()
    tool_status, tool_missing = _detect_tools()

    print("Availability report")
    print("===================")
    print("")

    print("Installed extras")
    for extra in sorted(_EXTRA_MODULES):
        if extra_status.get(extra):
            print(f"  [OK] {extra}")
        else:
            missing_modules = ", ".join(extra_missing.get(extra, [])) or "unknown"
            hint = _EXTRA_INSTALL_HINTS.get(extra, "")
            print(f"  [MISS] {extra} | missing modules: {missing_modules}")
            if hint:
                print(f"         install: {hint}")
    print("")

    print("External tools in PATH")
    for tool in sorted(_TOOL_ALIASES):
        resolved = tool_status.get(tool, "")
        if resolved:
            print(f"  [OK] {tool}: {resolved}")
        else:
            aliases = ", ".join(tool_missing.get(tool, []))
            print(f"  [MISS] {tool}: searched [{aliases}]")
            for cmd in _TOOL_INSTALL_HINTS.get(tool, []):
                print(f"         install: {cmd}")
    print("")

    print("Format availability levels")
    for fmt, level, note in _format_levels(extra_status, tool_status):
        print(f"  - formats: {fmt}")
        print(f"    level: {level}")
        print(f"    note: {note}")

    return 0


def _resolve_output_path(src: Path, input_root: Path, out_root: Path, keep_structure: bool) -> Path:
    if keep_structure:
        relative = src.relative_to(input_root)
        return (out_root / relative).with_suffix(".txt")
    return out_root / f"{src.stem}.txt"


def _record_for_result(src: Path, dst: Path | None, result, error: str | None = None, dry_run: bool = False) -> dict[str, str]:
    metadata = result.metadata if result else {}
    warnings = result.warnings if result else []
    warning_text = " | ".join(
        f"[{w.get('code', 'FALLBACK_USED')}] {w.get('source', 'unknown')}: {w.get('message', '')}"
        for w in warnings
    )
    return {
        "source": str(src),
        "output": str(dst) if dst else "",
        "status": "dry-run" if dry_run and not error else ("failed" if error else "ok"),
        "used_method": result.used_method if result else "",
        "encoding": result.detected_encoding or "",
        "chars": str(len(result.text)) if result else "0",
        "error": error or "",
        "title": str(metadata.get("title", "")),
        "author": str(metadata.get("author", "")),
        "date": str(metadata.get("date", "")),
        "language": str(metadata.get("language", "")),
        "page_count": str(metadata.get("page_count", "")),
        "subject": str(metadata.get("subject", "")),
        "from": str(metadata.get("from", "")),
        "to": str(metadata.get("to", "")),
        "metadata_json": json.dumps(metadata, ensure_ascii=False),
        "warnings": warning_text,
        "warnings_json": json.dumps(warnings, ensure_ascii=False),
    }


def _write_report(report_path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    report_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys())
    with report_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.doctor:
        return _print_availability_report()

    if not args.input:
        parser.error("input is required unless --doctor/--available is used")

    decoder = TextDecoder(
        preferred_tools=args.tools,
        encoding=args.input_encoding,
        fallback_encodings=args.fallback_encodings,
        output_encoding=args.output_encoding,
    )
    src = Path(args.input)
    report_rows: list[dict[str, str]] = []

    try:
        if src.is_file():
            out = Path(args.output) if args.output else src.with_suffix(".txt")
            if args.failed_only and out.exists():
                print(f"SKIP: {src} -> {out} (already exists)")
                return 0
            result = decoder.decode_result(src)
            if args.dry_run:
                print(f"DRY-RUN: {src} -> {out} [{result.used_method}] chars={len(result.text)}")
            else:
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_text(result.text, encoding=args.output_encoding, errors="replace")
                print(f"OK: {src} -> {out} [{result.used_method}]")
            if args.report:
                report_rows.append(_record_for_result(src, out, result, dry_run=args.dry_run))
                _write_report(Path(args.report), report_rows)
            return 0

        if src.is_dir():
            out_dir = Path(args.output) if args.output else src / "decoded_txt"
            converted = 0
            failed = 0
            skipped = 0

            for p in src.rglob(args.glob):
                if not p.is_file():
                    continue
                target = _resolve_output_path(p, src, out_dir, args.keep_structure)
                if args.failed_only and target.exists():
                    skipped += 1
                    continue
                try:
                    result = decoder.decode_result(p)
                    if args.dry_run:
                        converted += 1
                        print(f"DRY-RUN: {p} -> {target} [{result.used_method}] chars={len(result.text)}")
                    else:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        target.write_text(result.text, encoding=args.output_encoding, errors="replace")
                        converted += 1
                        print(f"OK: {p} -> {target} [{result.used_method}]")
                    if args.report:
                        report_rows.append(_record_for_result(p, target, result, dry_run=args.dry_run))
                except Exception as exc:  # noqa: BLE001
                    failed += 1
                    print(f"FAIL: {p} ({exc})")
                    if args.report:
                        report_rows.append(_record_for_result(p, target, None, error=str(exc)))

            if args.report:
                _write_report(Path(args.report), report_rows)

            print(
                f"Done. Converted: {converted}, failed: {failed}, skipped: {skipped}, "
                f"dry-run: {args.dry_run}, output: {out_dir}"
            )
            if args.dry_run:
                return 0
            return 0 if converted > 0 else 2

        print(f"Input does not exist: {src}")
        return 2
    except ExtractorError as exc:
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

