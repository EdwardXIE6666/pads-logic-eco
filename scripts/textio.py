"""Encoding helpers shared by the scripts.

Why: the skill runs inside many hosts (Cursor, 豆包/飞书 workbench, plain cmd on a
Chinese Windows). They disagree on the default text encoding, so a CSV written in
UTF-8 gets read as GBK (or the reverse) and the host reports "csv 编码异常".

    force_utf8_stdio()        make print() emit UTF-8 regardless of console codepage
    read_text_any(path)       UTF-8 (with/without BOM) → GBK → latin-1, never raises
    MAP_ENCODINGS             accepted values for --map-encoding
"""
from __future__ import annotations

import sys
from pathlib import Path

MAP_ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "ascii")


def force_utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def read_text_any(path: str | Path) -> str:
    data = Path(path).read_bytes()
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    for enc in ("utf-8", "gbk"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")
