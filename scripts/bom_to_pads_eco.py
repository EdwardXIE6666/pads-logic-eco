#!/usr/bin/env python3
"""Build a PADS Logic *CHGPART* ECO from a 3-column table.

Input CSV/TSV/whitespace with headers or without:
  RefDes,OldPartType,NewPartType

Rules:
  - exact names only; do not prefix NewPartType with *
  - one space between fields
  - CRLF endings, ASCII
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from textio import force_utf8_stdio, read_text_any

force_utf8_stdio()


def parse_rows(path: Path) -> list[tuple[str, str, str]]:
    text = read_text_any(path)  # UTF-8 / UTF-8-BOM / GBK all accepted
    sample = text[:4096]
    dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    reader = csv.reader(text.splitlines(), dialect)
    rows = [tuple(c.strip() for c in r) for r in reader if any(c.strip() for c in r)]
    if not rows:
        raise SystemExit("empty input")
    start = 0
    head = [c.lower() for c in rows[0]]
    if any("ref" in h or "old" in h or "new" in h or "part" in h for h in head):
        start = 1
    out = []
    for r in rows[start:]:
        if len(r) < 3:
            raise SystemExit(f"need 3 columns, got {r!r}")
        ref, old, new = r[0], r[1], r[2]
        if new.startswith("*") and not new.startswith("*PADS"):
            print(f"warning: stripping leading * from new type {new}", file=sys.stderr)
            new = new.lstrip("*")
        if " " in ref or " " in old or " " in new:
            raise SystemExit(f"token contains space: {ref} {old} {new}")
        out.append((ref, old, new))
    return out


def write_eco(rows: list[tuple[str, str, str]], dest: Path, remark: str) -> None:
    lines = [
        "*PADS-ECO-V3.0-MILS*",
        f"*REMARK* {remark}",
        "*CHGPART*",
    ]
    for ref, old, new in rows:
        lines.append(f"{ref} {old} {new}")
    lines.append("*END*")
    try:
        dest.write_bytes(("\r\n".join(lines) + "\r\n").encode("ascii"))
    except UnicodeEncodeError as e:
        raise SystemExit(f"ECO 必须是纯 ASCII（Logic 不认中文）：{e.object[e.start:e.end]!r} 出现在 remark 或类型名中") from e


def main() -> None:
    ap = argparse.ArgumentParser(description="CSV to PADS *CHGPART* ECO")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--remark", default="generated CHGPART")
    args = ap.parse_args()
    rows = parse_rows(args.input)
    write_eco(rows, args.output, args.remark)
    print(f"wrote {args.output} ({len(rows)} parts)")


if __name__ == "__main__":
    main()
