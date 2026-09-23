from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import quote

INVALID_WINDOWS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_WINDOWS = re.compile(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$", re.I)


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", errors="surrogatepass")).hexdigest()


def sha256_file(path: Path, block_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(block_size), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_name(value: str | None, max_length: int = 72, fallback: str = "unnamed") -> str:
    text = (value or "").replace("\u00a0", " ")
    text = INVALID_WINDOWS.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    if not text:
        text = fallback
    if RESERVED_WINDOWS.match(text):
        text = "_" + text
    if len(text) > max_length:
        suffix = "_" + sha256_text(text)[:8]
        text = text[: max(1, max_length - len(suffix))].rstrip(" .") + suffix
    return text or fallback


def ensure_within(base: Path, candidate: Path) -> Path:
    base_resolved = base.resolve()
    candidate_resolved = candidate.resolve()
    try:
        candidate_resolved.relative_to(base_resolved)
    except ValueError as exc:
        raise ValueError(f"目标路径越过归档目录：{candidate}") from exc
    return candidate_resolved


def relative_link(path: str) -> str:
    return quote(path.replace("\\", "/"), safe="/-_.~")


def iso(value: datetime | None) -> str | None:
    if value is None:
        return None
    return value.astimezone().isoformat(timespec="seconds") if value.tzinfo else value.isoformat(timespec="seconds")


def json_default(value: Any) -> Any:
    if isinstance(value, datetime):
        return iso(value)
    if isinstance(value, Path):
        return str(value)
    raise TypeError(type(value).__name__)


def write_json_atomic(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2, default=json_default) + "\n",
        encoding="utf-8-sig",
    )
    os.replace(temporary, path)


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def unique_actual_name(directory: Path, original: str, index: int, max_length: int = 60) -> str:
    safe = safe_name(original, max_length=max_length, fallback=f"attachment_{index:03d}")
    proposed = safe
    stem, suffix = Path(safe).stem, Path(safe).suffix
    counter = 2
    while (directory / proposed).exists():
        proposed = safe_name(f"{stem} ({counter}){suffix}", max_length=max_length + 6)
        counter += 1
    ensure_within(directory, directory / proposed)
    return proposed


def parse_years(text: str | Iterable[int] | None) -> list[int]:
    if text is None:
        return []
    parts = re.split(r"[,，;；\s]+", text.strip()) if isinstance(text, str) else list(text)
    years = sorted({int(value) for value in parts if str(value).strip()})
    if any(year < 1900 or year > 9999 for year in years):
        raise ValueError("年份必须在 1900 到 9999 之间。")
    return years
