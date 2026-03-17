from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


WarningEntry = dict[str, str]


@dataclass(slots=True)
class DecodeResult:
    text: str
    used_method: str
    source_format: str
    detected_encoding: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)
    warnings: list[WarningEntry] = field(default_factory=list)
