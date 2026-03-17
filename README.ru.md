# all2txt (Русский README)

`all2txt` — Python-библиотека и CLI для извлечения текста из разных форматов документов.

English version: [README.md](README.md)

## Что умеет

- Единый API для получения текста из файла
- Сохранение результата в `.txt`
- Расширенный результат (`DecodeResult`): метод извлечения, кодировка, предупреждения, метаданные
- Мягкая деградация: если нужного инструмента нет, библиотека пробует fallback-цепочку

### Форматы, которые работают "из коробки" (без extras)

- Текст/разметка: `.txt`, `.log`, `.ini`, `.conf`, `.md`, `.rst`, `.csv`, `.tsv`, `.json`, `.xml`, `.html`, `.htm`, `.mht`, `.mhtml`, `.eml`, `.plist`, `.tex`, `.bib`, `.strings`, `.rtf`
- OpenXML/ODF/ebook: `.docx`, `.odt`, `.ods`, `.xlsx`, `.pptx`, `.fb2`, `.epub`, `.pages`, `.numbers`, `.key`

### Форматы, где лучше добавить зависимости

- `.pdf` -> `pip install all2txt[pdf]`
- `.mobi` -> `pip install all2txt[mobi]`
- `.msg` -> `pip install all2txt[msg]`

### Внешние инструменты (ставятся отдельно от pip)

Для максимального качества по legacy/сканам могут понадобиться:

- Microsoft Word (COM)
- LibreOffice/OpenOffice
- Calibre
- Tesseract/OCRmyPDF/Poppler
- DjVuLibre
- antiword/wvText/catdoc

## Установка

### Базово

```bash
pip install all2txt
```

### Все опциональные Python-зависимости сразу

```bash
pip install all2txt[all]
```

Это установит Python-пакеты:

- `pypdf`
- `pywin32`
- `olefile`
- `mobi`
- `extract-msg`

Важно: системные утилиты (LibreOffice, Word, Tesseract и т.д.) через эту команду не ставятся.

### Проверить, что доступно в текущем окружении

```bash
all2txt --available
```

Команда покажет:

- какие extras реально доступны
- какие внешние утилиты найдены в `PATH`
- какие форматы доступны на уровнях `native/tool/plugin-ocr/fallback`
- рекомендации установки

## Использование в коде

## Коротко про API

- `decode_file(path)` -> возвращает только текст (`str`)
- `decode_result(path, ...)` -> возвращает `DecodeResult` (текст + служебные поля)
- `decode_to_txt(path, out_path, ...)` -> сохраняет текст в файл
- `TextDecoder(...)` -> переиспользуемый декодер с настройками

### Пример 1: получить только текст

```python
from all2txt import decode_file

text = decode_file("docs/sample.doc")
print(text[:500])
```

### Пример 2: получить текст + метаданные

```python
from all2txt import decode_result

res = decode_result(
    "docs/sample.doc",
    encoding="utf-8",
    fallback_encodings=["cp1251", "koi8-r", "cp866"],
)

print(res.used_method)
print(res.source_format)
print(res.detected_encoding)
print(res.metadata)
print(res.warnings)
```

### Пример 3: пакетная обработка в pandas (для ноутбука)

```python
from pathlib import Path
import pandas as pd
from all2txt import TextDecoder

decoder = TextDecoder(
    preferred_tools=["word", "libreoffice", "ole", "strings"],
    encoding="utf-8",
    fallback_encodings=["cp1251", "koi8-r", "cp866"],
)

rows = []
for p in Path("docs").rglob("*"):
    if not p.is_file():
        continue
    try:
        r = decoder.decode_result(p)
        rows.append({
            "path": str(p),
            "text": r.text,
            "method": r.used_method,
            "encoding": r.detected_encoding,
            "language": r.metadata.get("language"),
            "status": "ok",
        })
    except Exception as exc:
        rows.append({"path": str(p), "text": "", "method": "", "status": "failed", "error": str(exc)})

_df = pd.DataFrame(rows)
```

## CLI

```bash
# Один файл
all2txt input.doc -o output.txt

# Папка
all2txt ./docs -o ./decoded --glob "*.doc*"

# Отчет CSV
all2txt ./docs -o ./decoded --report report.csv

# Только просмотр (без записи)
all2txt ./docs --dry-run
```

Полезные флаги:

- `--available` (`--doctor`, `--help-env`)
- `--report`
- `--failed-only`
- `--keep-structure`
- `--method-order`
- `--input-encoding`, `--fallback-encodings`, `--output-encoding`

## Как добавить новый формат

### Вариант 1: локально в вашем проекте

```python
from pathlib import Path
from all2txt import register_extractor


def my_extractor(path: Path, default_encoding="utf-8", fallback_encodings=None):
    return path.read_text(encoding=default_encoding, errors="replace")


register_extractor(".myfmt", my_extractor)
```

После регистрации `.myfmt` начнет обрабатываться через обычные `decode_file/decode_result`.

### Вариант 2: отдельный plugin-пакет на PyPI

Создайте отдельный пакет (например `all2txt-myformat`) и добавьте entry point:

```toml
[project]
name = "all2txt-myformat"
version = "0.1.0"
dependencies = ["all2txt>=0.1.0"]

[project.entry-points."all2txt.extractors"]
myformat = "all2txt_myformat.plugin:myformat_extractor"
```

В коде plugin-функции укажите расширения:

```python
def myformat_extractor(path, default_encoding="utf-8", fallback_encodings=None):
    ...

myformat_extractor.suffixes = [".myfmt"]
```

Публикация:

```bash
python -m pip install --upgrade build twine
python -m build
python -m twine upload dist/*
```

Пользователь устанавливает расширение:

```bash
pip install all2txt-myformat
```

## Почему это важно

- Не нужно форкать ядро all2txt ради каждого нового формата
- Форматы можно развивать отдельными релизами
- Команды могут поддерживать свои domain-specific плагины независимо
