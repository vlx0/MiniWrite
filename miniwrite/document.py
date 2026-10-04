"""Save / load MiniWrite documents (.mw JSON + plain .txt)."""

from __future__ import annotations

import json
from pathlib import Path

FORMAT_VERSION = 2


def save_mw(path: Path, text: str, tags: list[dict]) -> None:
    data = {
        "format": "miniwrite",
        "version": FORMAT_VERSION,
        "text": text,
        "tags": tags,
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_mw(path: Path) -> tuple[str, list[dict]]:
    raw = path.read_text(encoding="utf-8")
    data = json.loads(raw)
    if not isinstance(data, dict) or data.get("format") != "miniwrite":
        raise ValueError("не формат MiniWrite")
    text = data.get("text")
    if text is None:
        text = data.get("plain") or ""
    tags = data.get("tags") or []
    if not isinstance(text, str) or not isinstance(tags, list):
        raise ValueError("битый файл")
    return text, tags


def save_txt(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def load_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8")
