#!/usr/bin/env python3
"""Verify a PADS Logic *CHGPART* ECO against schematic BOM + component library.

This is the harness enforcement tool.  Every ECO produced by this skill must
pass this checker before delivery.

Checks
------
F0  File exists and is non-empty.
F1  Header: first line is *PADS-ECO-V3.0-MILS*.
F2  *REMARK* line present.
F3  *CHGPART* and/or *SET_ATTRIBUTE* present, *END* present, CHGPART before END.
F4  All lines end with CRLF; no bare LF; no trailing spaces.
F5  *CHGPART* data lines have exactly 3 space-separated tokens; *SET_ATTRIBUTE* is TAB PART RefDes + "Value" TAB @value.
D0  Every refdes in col 1 exists in the schematic BOM.
D1  Col 2 (old type) equals BOM current type, except the second hop of an
    add-pair (where col 2 is the intermediate NC).
D2  Col 3 (new type) exists in the library name set.
D3  A refdes that appears twice is an add-pair: line1.col3 == line2.col2
    and line1.col3 contains '_NC_'.
D4  No duplicate (ref, old, new) lines.
D5  物料删除 is *SET_ATTRIBUTE* PART "Value" = '@'+BOM Value (not CHGPART to NC).
    A single CHGPART whose new type contains '_NC_' and old does not is retired.
D6  Add-pair first hop goes to an NC that exists in the library.
D7  No cross-family add direct-jump (single line where old is an NC of a
    different family than the non-NC new type). Same-family NC -> After is
    the legitimate one-hop add.
D8  No direct modify on a refdes whose BOM Value is '@' (single line, old and
    new both non-NC). PADS Logic applies such a CHGPART to every part of that
    type board-wide; the row must go through NC like an add.

Usage
-----
    python verify_eco.py --eco board.eco --bom schBOM.xlsx --lib library.xlsx

Exit code 0 = all checks passed.  Exit code 1 = one or more failures.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

from bomscan import prefix_at, schematic_maps
from textio import force_utf8_stdio

force_utf8_stdio()

# ---------------------------------------------------------------------------
# Shared helpers (kept in sync with ecn_to_pads_eco.py)
# ---------------------------------------------------------------------------
PKG_MAP = {
    "1005": "01005", "01005": "01005",
    "201": "0201", "0201": "0201",
    "402": "0402", "0402": "0402",
    "603": "0603", "0603": "0603",
    "805": "0805", "0805": "0805",
    "1206": "1206", "1210": "1210",
}
NAME_PREFIX = (("R_F_", "R_F_NC_"), ("C_M_", "C_M_NC_"), ("L_RF_", "L_RF_NC_"))

def norm_code(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    if not s or s.lower() in {"none", "nan"}:
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def split_codes(v):
    s = norm_code(v)
    if not s:
        return []
    return [p.strip() for p in re.split(r"\s*&\s*|\s*,\s*", s) if p.strip()]


def family_nc_prefix(part_type):
    if not part_type:
        return None
    for a, b in NAME_PREFIX:
        if part_type.startswith(a):
            return b
    return None


def load_library(path: Path):
    # Same header rules as generation: 元件名称 is located by name, never by column index.
    from ecn_to_pads_eco import read_library

    audit, _pn_to_names, all_names = read_library(path)
    if audit.name_col is None:
        raise SystemExit(f"器件库缺少元件名称列: {path.name}。至少要有两列: 元件名称、Part Number。")
    return all_names


# BOM scanning lives in bomscan.py (shared with ecn_to_pads_eco.py)


# ---------------------------------------------------------------------------
# ECO parser
# ---------------------------------------------------------------------------
def parse_eco(path: Path):
    """Return (raw, lines, chg_rows, set_value_rows, errors, flags)."""
    errors: list[str] = []
    raw = path.read_bytes()
    text = raw.decode("ascii", errors="replace")
    lines = text.split("\r\n")
    if lines and lines[-1] == "":
        lines = lines[:-1]

    data_rows: list[tuple[str, str, str]] = []
    set_rows: list[tuple[str, str]] = []
    in_chg = False
    in_set = False
    set_cur_ref = ""
    saw_chg = False
    saw_set = False
    saw_end = False
    saw_remark = False

    for idx, line in enumerate(lines, start=1):
        if line.startswith("*REMARK*"):
            saw_remark = True
            continue
        if line == "*CHGPART*":
            in_chg, in_set = True, False
            saw_chg = True
            continue
        if line.startswith("*SET_ATTRIBUTE*"):
            in_chg, in_set = False, True
            saw_set = True
            parts = line.split("\t")
            if len(parts) >= 3 and parts[1] == "PART":
                set_cur_ref = parts[2].strip()
            else:
                errors.append(f"F5 line {idx}: expected *SET_ATTRIBUTE*<TAB>PART<TAB>RefDes, got {line!r}")
                set_cur_ref = ""
            continue
        if line == "*END*":
            in_chg = in_set = False
            set_cur_ref = ""
            saw_end = True
            continue
        if in_chg and line.strip() and not line.startswith("*"):
            tokens = line.split(" ")
            if len(tokens) != 3:
                errors.append(f"F5 line {idx}: expected 3 tokens, got {len(tokens)}: {line!r}")
            else:
                ref, old, new = tokens
                if "*" in ref or "*" in old or "*" in new:
                    errors.append(f"F5 line {idx}: data line contains '*': {line!r}")
                data_rows.append((ref, old, new))
        elif in_set and line.startswith('"Value"'):
            val = line.split("\t", 1)[1] if "\t" in line else ""
            if not set_cur_ref or not val:
                errors.append(f"F5 line {idx}: *SET_ATTRIBUTE* Value line missing ref or value: {line!r}")
            else:
                set_rows.append((set_cur_ref, val))
    return raw, lines, data_rows, set_rows, errors, saw_chg, saw_set, saw_end, saw_remark


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
def run_checks(eco_path: Path, bom_path: Path | None, lib_path: Path, bom_maps=None) -> list[str]:
    errors: list[str] = []

    # F0
    if not eco_path.exists() or eco_path.stat().st_size == 0:
        return [f"F0: ECO file missing or empty: {eco_path}"]

    raw, lines, data_rows, set_rows, parse_err, saw_chg, saw_set, saw_end, saw_remark = parse_eco(eco_path)
    errors.extend(parse_err)

    # F1
    first = repr(lines[0]) if lines else "empty"
    if not lines or lines[0] != "*PADS-ECO-V3.0-MILS*":
        errors.append(f"F1: first line must be *PADS-ECO-V3.0-MILS*, got {first}")

    # F2
    if not saw_remark:
        errors.append("F2: *REMARK* line not found")

    # F3
    if not saw_chg and not saw_set:
        errors.append("F3: neither *CHGPART* nor *SET_ATTRIBUTE* section found")
    if not saw_end:
        errors.append("F3: *END* not found")
    end_idx = next((i for i, l in enumerate(lines) if l == "*END*"), -1)
    if saw_chg and end_idx >= 0:
        chg_idx = next((i for i, l in enumerate(lines) if l == "*CHGPART*"), -1)
        if chg_idx > end_idx:
            errors.append("F3: *CHGPART* appears after *END*")

    # F4: CRLF, no bare LF, no trailing spaces
    if raw.count(b"\r\n") != raw.count(b"\n"):
        errors.append(f"F4: found bare LF (not CRLF); CRLF={raw.count(b'\r\n')} LF={raw.count(b'\n')}")
    trailing = len(re.findall(b" +\r\n", raw))
    if trailing:
        errors.append(f"F4: {trailing} lines have trailing spaces")

    # Stop early if format is broken badly
    if not data_rows and not set_rows:
        errors.append("F5: no data rows in *CHGPART* or *SET_ATTRIBUTE*")
        return errors

    # Load BOM + library. Live maps skip the file.
    bom_type, bom_val = schematic_maps(bom_path, bom_maps)
    bom_at = {r for r, v in bom_val.items() if "@" in v}
    if not bom_type:
        src = "当前原理图" if bom_maps is not None else bom_path.name
        errors.append(f"D0: BOM has no Part Type column ({src}); export the BOM from PADS Logic with Part Type")
        return errors
    all_names = load_library(lib_path)

    # Group rows by refdes
    by_ref: dict[str, list[tuple[int, str, str, str]]] = defaultdict(list)
    for i, (ref, old, new) in enumerate(data_rows):
        by_ref[ref].append((i, old, new))

    # D4: duplicate lines
    seen: set[tuple[str, str, str]] = set()
    for ref, old, new in data_rows:
        key = (ref, old, new)
        if key in seen:
            errors.append(f"D4: duplicate line: {ref} {old} {new}")
        seen.add(key)

    for ref, entries in by_ref.items():
        # D0: refdes in BOM
        if ref not in bom_type:
            errors.append(f"D0: refdes {ref} not found in BOM")
            continue

        bom_cur = bom_type[ref]

        if len(entries) == 1:
            _i, old, new = entries[0]
            # D1: old type matches BOM
            if old != bom_cur:
                errors.append(f"D1: {ref} old={old!r} != BOM={bom_cur!r}")
            # D2: new type in library
            if new not in all_names:
                errors.append(f"D2: {ref} new type {new!r} not in library")
            # D5: CHGPART to NC is the retired delete; 物料删除 now prefixes Value
            if "_NC_" in new and "_NC_" not in old:
                errors.append(f"D5: {ref} CHGPART to NC is retired; 物料删除 prefixes Value with @")
            # D6 for single NC target (legacy / unexpected)
            if "_NC_" in new and new not in all_names:
                errors.append(f"D6: {ref} NC target {new!r} not in library")
            # D7: cross-family add direct-jump (old is NC of another family, new is not NC, single line).
            # A same-family NC -> After single row is the legitimate one-hop add.
            if "_NC_" in old and "_NC_" not in new and family_nc_prefix(old) != family_nc_prefix(new):
                errors.append(f"D7: {ref} cross-family direct-jump {old} -> {new}; should be two-hop via target-family NC")
            # D8: direct modify on a Value='@' part (Logic replaces every part of that type board-wide).
            # Must go through NC like an add; single row is only fine when it is the NC -> After hop.
            if ref in bom_at and "_NC_" not in old and "_NC_" not in new:
                errors.append(f"D8: {ref} BOM Value is '@'; direct modify {old} -> {new} would change the type board-wide; route via NC (two-hop)")

        elif len(entries) == 2:
            # Add-pair
            (i1, old1, new1), (i2, old2, new2) = entries
            # D1 first hop old == BOM
            if old1 != bom_cur:
                errors.append(f"D1: {ref} add-hop1 old={old1!r} != BOM={bom_cur!r}")
            # D3: line1.new == line2.old and line1.new is NC
            if new1 != old2:
                errors.append(f"D3: {ref} add-pair broken: hop1.new={new1!r} != hop2.old={old2!r}")
            if "_NC_" not in new1:
                errors.append(f"D3: {ref} add-pair first hop target {new1!r} is not NC")
            # D2: both new types in library
            if new1 not in all_names:
                errors.append(f"D2: {ref} add-hop1 new {new1!r} not in library")
            if new2 not in all_names:
                errors.append(f"D2: {ref} add-hop2 new {new2!r} not in library")
            # D6: intermediate NC exists
            if new1 not in all_names:
                errors.append(f"D6: {ref} intermediate NC {new1!r} not in library")

        else:
            errors.append(f"D3: {ref} appears {len(entries)} times (expected 1 or 2)")

    seen_set: set[tuple[str, str]] = set()
    for ref, new_val in set_rows:
        if (ref, new_val) in seen_set:
            errors.append(f"D4: duplicate *SET_ATTRIBUTE* Value: {ref} {new_val}")
        seen_set.add((ref, new_val))
        if ref not in bom_type:
            errors.append(f"D0: *SET_ATTRIBUTE* refdes {ref} not found in BOM")
            continue
        expected = prefix_at(bom_val.get(ref, ""))
        if expected is None:
            raw = (bom_val.get(ref) or "").strip()
            if raw.startswith("@"):
                errors.append(f"D5: {ref} BOM Value already starts with @ ({raw!r}); no delete *SET_ATTRIBUTE* needed")
            else:
                errors.append(f"D5: {ref} BOM has no Value to prefix with @")
        elif new_val != expected:
            errors.append(f"D5: {ref} *SET_ATTRIBUTE* Value={new_val!r} != expected {expected!r}")

    return errors


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Verify PADS Logic CHGPART ECO")
    ap.add_argument("--eco", required=True, type=Path)
    ap.add_argument("--bom", required=True, type=Path)
    ap.add_argument("--lib", required=True, type=Path)
    args = ap.parse_args()

    errors = run_checks(args.eco, args.bom, args.lib)
    if errors:
        print(f"FAIL  {len(errors)} issue(s):")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    else:
        print(f"PASS  all checks passed  ({args.eco.name})")
        sys.exit(0)


if __name__ == "__main__":
    main()
