"""Schematic BOM scanner shared by ecn_to_pads_eco.py and verify_eco.py.

scan_bom(path) -> (refdes -> current Part Type, refdes -> Value)

Header-agnostic: the type is any cell whose text before the first comma matches
the library grammar; a refdes is any token matching the designator pattern.

Wrapped records: some BOM exports split one record across two rows when the
refdes list is very long. Row N carries PN / description / a truncated refdes
list and empty type cells; row N+1 starts with the rest of the refdes list
(possibly beginning mid-token, e.g. "7,L2518,...") followed by package / value /
type shifted left. The scanner concatenates the two refdes texts with no
separator and takes the type from row N+1.
"""
from __future__ import annotations

import re
from pathlib import Path

from xlsx_open import load_workbook_any

TYPE_RE = re.compile(r"^(C_M_|L_RF_|R_F_|R_T_)[A-Za-z0-9%._+\-]+")
REF_RE = re.compile(r"^[A-Z]{1,3}\d{1,5}$")
VALUE_HEADER_RE = re.compile(r"^(value|vaule|值|参数值)$", re.I)
SPLIT_RE = re.compile(r"[,，\s]+")


def _refs_in(text: str) -> list[str]:
    return [t for t in SPLIT_RE.split(text) if REF_RE.match(t)]


def prefix_at(value: str | None) -> str | None:
    """Prefix Value with '@'. None if empty or already prefixed (already deleted)."""
    v = (value or "").strip()
    if not v or v.startswith("@"):
        return None
    return "@" + v


def _row_value(cells: list[str], value_col: int | None, type_idx: int | None, continuation: bool) -> str:
    if not continuation and value_col is not None and value_col < len(cells) and cells[value_col]:
        return cells[value_col]
    if type_idx is not None and "," in cells[type_idx]:
        return cells[type_idx].split(",", 1)[1].split(",")[0].strip()
    if continuation and type_idx and type_idx >= 1 and cells[type_idx - 1]:
        cand = cells[type_idx - 1]
        if not REF_RE.match(cand) and not re.fullmatch(r"\d{3,4}", cand):
            return cand
    return ""


def scan_bom(path: Path | str) -> tuple[dict[str, str], dict[str, str]]:
    wb = load_workbook_any(path)
    ws = wb[wb.sheetnames[0]]
    out: dict[str, str] = {}
    values: dict[str, str] = {}
    max_col = min(12, ws.max_column)

    value_col = None
    for row in ws.iter_rows(min_row=1, max_row=min(10, ws.max_row), max_col=max_col):
        for cell in row:
            if cell.value is not None and VALUE_HEADER_RE.match(str(cell.value).strip()):
                value_col = cell.column - 1
                break
        if value_col is not None:
            break

    # refdes text of the previous row when it had designators but no type
    pending_ref_text: str | None = None

    for row in ws.iter_rows(min_row=1, max_col=max_col, values_only=True):
        cells = [str(c).strip() if c is not None else "" for c in row]
        ptype = None
        type_idx = None
        for idx, c in enumerate(cells):
            first = c.split(",")[0].strip()
            if TYPE_RE.match(first):
                ptype, type_idx = first, idx
                break

        ref_cells = [(i, c) for i, c in enumerate(cells) if i != type_idx and _refs_in(c)]

        if not ptype:
            # row with designators but no type: remember it, it may be a wrapped record
            pending_ref_text = max((c for _, c in ref_cells), key=len) if ref_cells else None
            continue

        refs: list[str] = []
        # continuation row: the layout is shifted left, so cell 0 is refdes text
        # (often starting mid-token, "7,L2518,...") instead of a part number
        continuation = (
            pending_ref_text is not None
            and type_idx is not None and type_idx > 0
            and bool(cells[0])
            and not re.fullmatch(r"[\d.\s&]+", cells[0])
            and (bool(_refs_in(cells[0])) or bool(re.match(r"^\d{0,4}[,，]", cells[0])))
        )
        if continuation:
            refs = _refs_in(pending_ref_text + cells[0])
        else:
            for _, c in ref_cells:
                refs.extend(_refs_in(c))
        pending_ref_text = None

        val = _row_value(cells, value_col, type_idx, continuation)

        for ref in refs:
            out[ref] = ptype
            if val:
                values[ref] = val

    wb.close()
    return out, values


def load_bom(path: Path | str) -> dict[str, str]:
    return scan_bom(path)[0]


def schematic_maps(path: Path | str | None, bom_maps=None):
    """File BOM, or maps already read from the open schematic.

    Live maps are returned as-is. A missing file still stops the command line.
    """
    if bom_maps is not None:
        return bom_maps
    if not path:
        raise SystemExit("缺少原理图 BOM (--bom)。当前 Part Type 只能来自原理图导出的 BOM，ECN 不能代替。")
    return scan_bom(path)


def _find_value_header(wb):
    """Return (worksheet, header_row, value_col), 1-based, or None."""
    for name in wb.sheetnames:
        ws = wb[name]
        max_r = min(15, ws.max_row or 1)
        max_c = min(40, ws.max_column or 1)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                raw = ws.cell(r, c).value
                if raw is not None and VALUE_HEADER_RE.match(str(raw).strip()):
                    return ws, r, c
    return None


def ebom_rows_from_bom(src: Path | str) -> tuple[list[list], int, int, int]:
    """Rows for the ECN tool's EBOM sheet. Header is row 0. Source is not modified.

    Drops every data row whose 值 contains '@' or is blank, and every data row
    whose 名称 is blank. The macro aborts — after it has already cleared the
    changelist — if either of those is left in. Title rows above the header are
    omitted so the header lands on EBOM row 1, which is the only row the macro
    searches.

    Returns (rows, removed_at, removed_blank, removed_name).
    """
    src_path = Path(src)
    wb = load_workbook_any(src_path)
    try:
        found = _find_value_header(wb)
        if not found:
            raise ValueError(f"这份 BOM 没有「值」列: {src_path.name}")
        ws, header_row, value_col = found
        max_c = min(40, ws.max_column or 1)
        headers: list[str] = []
        for c in range(1, max_c + 1):
            raw = ws.cell(header_row, c).value
            headers.append("" if raw is None else str(raw).strip())
        while headers and headers[-1] == "":
            headers.pop()
        missing = [name for name in ("下阶料号", "位号", "名称", "封装") if name not in headers]
        if missing or value_col > len(headers):
            seen = " | ".join(h for h in headers if h) or "(空)"
            need = "、".join(missing) if missing else "值"
            raise ValueError(f"这份 BOM 缺少 ECN 工具 EBOM 要的列: {need}。实际表头: {seen}")
        value_idx = value_col - 1
        name_idx = headers.index("名称")
        removed_at = removed_blank = removed_name = 0
        body: list[list] = []
        for r in range(header_row + 1, (ws.max_row or header_row) + 1):
            cells = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
            raw_val = cells[value_idx]
            text = "" if raw_val is None else str(raw_val).strip()
            raw_name = cells[name_idx]
            name = "" if raw_name is None else str(raw_name).strip()
            if "@" in text:
                removed_at += 1
            elif not text:
                removed_blank += 1
            elif not name:
                removed_name += 1
            else:
                body.append(cells)
        return [headers, *body], removed_at, removed_blank, removed_name
    finally:
        wb.close()


def preprocess_reverse_bom(src: Path | str, dest: Path | str) -> tuple[int, int, int]:
    """Drop data rows whose 值 contains '@' or is blank.

    Writes a new workbook. The source file is not modified.
    Returns (removed_at, removed_blank, kept_data_rows).
    """
    src_path, dest_path = Path(src), Path(dest)
    if src_path.resolve() == dest_path.resolve():
        raise ValueError("不会覆盖原始 BOM")
    wb = load_workbook_any(src_path)
    try:
        found = _find_value_header(wb)
        if not found:
            raise ValueError(f"这份 BOM 没有「值」列: {src_path.name}")
        ws, header_row, value_col = found
        removed_at = removed_blank = 0
        for r in range(ws.max_row or header_row, header_row, -1):
            raw = ws.cell(r, value_col).value
            text = "" if raw is None else str(raw).strip()
            if "@" in text:
                ws.delete_rows(r)
                removed_at += 1
            elif not text:
                ws.delete_rows(r)
                removed_blank += 1
        kept = max(0, (ws.max_row or header_row) - header_row)
        wb.save(dest_path)
        return removed_at, removed_blank, kept
    finally:
        wb.close()
