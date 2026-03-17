from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from all2txt.models import DecodeResult

_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
_DOCUMENT_SUFFIXES = {".pdf", ".djvu", ".djv"}
suffixes = sorted(_IMAGE_SUFFIXES | _DOCUMENT_SUFFIXES)


def ocr_plugin(path: Path, default_encoding: str = "utf-8", fallback_encodings=None) -> DecodeResult:
    suffix = path.suffix.lower()
    attempts: list[str] = []

    if suffix == ".pdf" and _pdf_has_text_layer(path):
        raise RuntimeError("PDF already has text layer; native extractor should handle it")

    if suffix in _IMAGE_SUFFIXES:
        for strategy in (_ocr_image_via_tesseract, _ocr_image_via_windows_ocr):
            try:
                text, method = strategy(path)
                if text.strip():
                    return _build_result(path, text, method, attempts)
            except Exception as exc:  # noqa: BLE001
                attempts.append(f"{strategy.__name__}: {exc}")
        return _build_result(path, _fallback_image_text(path, attempts), "plugin:ocr:placeholder", attempts)

    if suffix == ".pdf":
        for strategy in (_ocr_pdf_via_ocrmypdf, _ocr_pdf_via_pdftoppm_tesseract):
            try:
                text, method = strategy(path)
                if text.strip():
                    return _build_result(path, text, method, attempts)
            except Exception as exc:  # noqa: BLE001
                attempts.append(f"{strategy.__name__}: {exc}")

    if suffix in {".djvu", ".djv"}:
        for strategy in (_ocr_djvu_via_ddjvu_tesseract, _ocr_djvu_via_magick_tesseract):
            try:
                text, method = strategy(path)
                if text.strip():
                    return _build_result(path, text, method, attempts)
            except Exception as exc:  # noqa: BLE001
                attempts.append(f"{strategy.__name__}: {exc}")

    raise RuntimeError("OCR plugin unavailable or failed: " + " | ".join(attempts or ["no OCR strategy matched"]))


def _build_result(path: Path, text: str, method: str, attempts: list[str]) -> DecodeResult:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n").strip() + "\n"
    return DecodeResult(
        text=normalized,
        used_method=method,
        source_format=path.suffix.lower(),
        detected_encoding="utf-8",
        metadata={
            "source_path": str(path),
            "source_format": path.suffix.lower(),
            "ocr": True,
            "ocr_engine": method,
        },
        warnings=attempts.copy(),
    )


def _fallback_image_text(path: Path, attempts: list[str]) -> str:
    lines = [
        f"[all2txt] OCR text could not be extracted from image: {path.name}",
        "OCR tools were not found or failed.",
        "Install Tesseract or use Windows OCR for better results.",
    ]
    if attempts:
        lines.append("Attempts:")
        lines.extend(attempts)
    return "\n".join(lines)


def _pdf_has_text_layer(path: Path) -> bool:
    try:
        from pypdf import PdfReader
    except ImportError:
        return False

    try:
        reader = PdfReader(str(path))
    except Exception:
        return False

    sample_pages = reader.pages[: min(3, len(reader.pages))]
    extracted = "\n".join((page.extract_text() or "") for page in sample_pages)
    return len(extracted.strip()) > 50


def _ocr_image_via_tesseract(path: Path) -> tuple[str, str]:
    exe = shutil.which("tesseract")
    if not exe:
        raise RuntimeError("tesseract not found in PATH")
    result = subprocess.run([exe, str(path), "stdout", "-l", "rus+eng"], capture_output=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.decode(errors="replace").strip() or "tesseract failed")
    return result.stdout.decode("utf-8", errors="replace"), "plugin:ocr:tesseract"


def _ocr_image_via_windows_ocr(path: Path) -> tuple[str, str]:
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise RuntimeError("PowerShell was not found")
    script = (
        "$ErrorActionPreference='Stop';"
        "Add-Type -AssemblyName System.Runtime.WindowsRuntime;"
        "$null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime];"
        "$null=[Windows.Graphics.Imaging.SoftwareBitmap,Windows.Foundation,ContentType=WindowsRuntime];"
        "$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime];"
        f"$f=[Windows.Storage.StorageFile]::GetFileFromPathAsync('{str(path).replace("'", "''")}').GetAwaiter().GetResult();"
        "$stream=$f.OpenAsync([Windows.Storage.FileAccessMode]::Read).GetAwaiter().GetResult();"
        "$decoder=[Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream).GetAwaiter().GetResult();"
        "$bmp=$decoder.GetSoftwareBitmapAsync().GetAwaiter().GetResult();"
        "$engine=[Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages();"
        "$result=$engine.RecognizeAsync($bmp).GetAwaiter().GetResult();"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "Write-Output $result.Text"
    )
    result = subprocess.run([powershell, "-NoProfile", "-Command", script], capture_output=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.decode(errors="replace").strip() or "Windows OCR failed")
    text = result.stdout.decode("utf-8", errors="replace")
    if not text.strip():
        raise RuntimeError("Windows OCR returned empty text")
    return text, "plugin:ocr:windows"


def _ocr_pdf_via_ocrmypdf(path: Path) -> tuple[str, str]:
    ocrmypdf = shutil.which("ocrmypdf")
    if not ocrmypdf:
        raise RuntimeError("ocrmypdf not found in PATH")
    with tempfile.TemporaryDirectory(prefix="all2txt_ocrpdf_") as tmp:
        output_pdf = Path(tmp) / f"{path.stem}.ocr.pdf"
        result = subprocess.run(
            [ocrmypdf, "--skip-text", str(path), str(output_pdf)],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "ocrmypdf failed")
        try:
            from pypdf import PdfReader
        except ImportError as exc:
            raise RuntimeError("Install optional dependency: pip install all2txt[pdf]") from exc
        reader = PdfReader(str(output_pdf))
        text = "\n\n".join((page.extract_text() or "") for page in reader.pages)
        return text, "plugin:ocr:ocrmypdf"


def _ocr_pdf_via_pdftoppm_tesseract(path: Path) -> tuple[str, str]:
    pdftoppm = shutil.which("pdftoppm")
    tesseract = shutil.which("tesseract")
    if not pdftoppm or not tesseract:
        raise RuntimeError("pdftoppm and tesseract are required")
    with tempfile.TemporaryDirectory(prefix="all2txt_pdftoppm_") as tmp:
        prefix = Path(tmp) / "page"
        render = subprocess.run([pdftoppm, "-png", str(path), str(prefix)], capture_output=True, text=True, check=False)
        if render.returncode != 0:
            raise RuntimeError(render.stderr.strip() or render.stdout.strip() or "pdftoppm failed")
        pages = sorted(Path(tmp).glob("page-*.png"))
        if not pages:
            raise RuntimeError("pdftoppm produced no pages")
        texts: list[str] = []
        for page in pages:
            ocr = subprocess.run([tesseract, str(page), "stdout", "-l", "rus+eng"], capture_output=True, check=False)
            if ocr.returncode == 0:
                texts.append(ocr.stdout.decode("utf-8", errors="replace"))
        return "\n\n".join(texts), "plugin:ocr:pdftoppm+tesseract"


def _ocr_djvu_via_ddjvu_tesseract(path: Path) -> tuple[str, str]:
    ddjvu = shutil.which("ddjvu")
    tesseract = shutil.which("tesseract")
    if not ddjvu or not tesseract:
        raise RuntimeError("ddjvu and tesseract are required")
    with tempfile.TemporaryDirectory(prefix="all2txt_ddjvu_") as tmp:
        image_path = Path(tmp) / "page.tif"
        render = subprocess.run([ddjvu, "-format=tiff", str(path), str(image_path)], capture_output=True, text=True, check=False)
        if render.returncode != 0 or not image_path.exists():
            raise RuntimeError(render.stderr.strip() or render.stdout.strip() or "ddjvu failed")
        ocr = subprocess.run([tesseract, str(image_path), "stdout", "-l", "rus+eng"], capture_output=True, check=False)
        if ocr.returncode != 0:
            raise RuntimeError(ocr.stderr.decode(errors="replace").strip() or "tesseract failed")
        return ocr.stdout.decode("utf-8", errors="replace"), "plugin:ocr:ddjvu+tesseract"


def _ocr_djvu_via_magick_tesseract(path: Path) -> tuple[str, str]:
    magick = shutil.which("magick")
    tesseract = shutil.which("tesseract")
    if not magick or not tesseract:
        raise RuntimeError("magick and tesseract are required")
    with tempfile.TemporaryDirectory(prefix="all2txt_magick_") as tmp:
        image_path = Path(tmp) / "page.png"
        render = subprocess.run([magick, str(path), str(image_path)], capture_output=True, text=True, check=False)
        if render.returncode != 0 or not image_path.exists():
            raise RuntimeError(render.stderr.strip() or render.stdout.strip() or "magick failed")
        ocr = subprocess.run([tesseract, str(image_path), "stdout", "-l", "rus+eng"], capture_output=True, check=False)
        if ocr.returncode != 0:
            raise RuntimeError(ocr.stderr.decode(errors="replace").strip() or "tesseract failed")
        return ocr.stdout.decode("utf-8", errors="replace"), "plugin:ocr:magick+tesseract"


ocr_plugin.suffixes = suffixes

