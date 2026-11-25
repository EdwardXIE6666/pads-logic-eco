#!/usr/bin/env python3
"""ECN + library + BOM -> one PADS Logic ECO.

物料修改 *CHGPART* / 物料删除 *SET_ATTRIBUTE* PART "Value" 前加@ / 物料新增经NC到After.
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from textio import MAP_ENCODINGS, force_utf8_stdio, read_text_any
from xlsx_open import load_workbook_any

force_utf8_stdio()

from bomscan import prefix_at, schematic_maps

SKILL_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_LIB = SKILL_ROOT / "library" / "Electronic_components_library.xlsx"

COMMENT_MODIFY = "物料修改"
COMMENT_DELETE = "物料删除"
COMMENT_ADD = "物料新增"

# --map-encoding ascii: hosts that cannot read UTF-8 / GBK CSV get an all-ASCII map
ACTION_ASCII = {COMMENT_MODIFY: "MODIFY", COMMENT_DELETE: "DELETE", COMMENT_ADD: "ADD"}
STATUS_ASCII = {
    "emit(Value@:经NC两跳)": "emit(value_at:via_NC)",
    "emit(Value:@前缀)": "emit(value:at_prefix)",
    "新旧类型相同": "same_type",
    "BOM无此位号": "not_in_BOM",
    "已带@无需删除": "already_at",
    "BOM无Value": "BOM_no_value",
    "Value含非ASCII": "value_non_ascii",
}


def _status_ascii(s: str) -> str:
    pkg_tail = ""
    marker = ";封装变化:"
    if s.startswith("emit(封装变化:") and s.endswith(")"):
        inner = s[len("emit(封装变化:"):-1].replace("→", "->")
        return f"emit(pkg:{inner})"
    if marker in s:
        s, rest = s.split(marker, 1)
        pkg_tail = ";pkg:" + rest.replace("→", "->")
    if s in STATUS_ASCII:
        return STATUS_ASCII[s] + pkg_tail
    for zh, en in (("新料号未入库:", "PN_not_in_lib:"), ("Value@需经NC但", "value_at_needs_NC:"), ("无NC族:", "no_NC_family:"), ("库无NC:", "NC_not_in_lib:")):
        if s.startswith(zh):
            return en + s[len(zh):] + pkg_tail
    return s.encode("ascii", "replace").decode() + pkg_tail


STATUS_ASCII_FN = _status_ascii

def abnormal_skips(rows):
    """Every skip is shown before generation. The ref stays out of the ECO."""
    return list(rows)


def _txt(value) -> str:
    return "" if value is None else str(value).strip()


def other_problem(reason: str) -> str:
    """其它异常的短名称。仍放在同一块里，不按原因再拆块。"""
    if reason == "新旧类型相同":
        return "新旧类型相同"
    if reason == "已带@无需删除":
        return "已带@无需删除"
    if reason == "BOM无此位号":
        return "BOM 缺位号"
    if reason == "BOM无Value":
        return "删除缺 Value"
    if reason == "Value含非ASCII":
        return "Value 含非 ASCII"
    if reason.startswith("库无NC:"):
        return "库无 NC"
    if reason.startswith("封装无NC映射"):
        return "封装对不上 NC"
    if reason.startswith("无NC族"):
        return "无 NC 族"
    if reason.startswith("新增无NC映射"):
        return "新增无 NC"
    if reason.startswith("Value@需经NC但"):
        return "Value 已带 @"
    return reason


def other_current_info(row) -> str:
    """这个位号现在能读到的信息。BOM 缺位号时，当前 Part Type 不在原理图里。"""
    reason = str(row[-1])
    cmt = _txt(row[1]) if len(row) > 1 else ""
    old_pn = _txt(row[2]) if len(row) > 2 else ""
    old_type = _txt(row[3]) if len(row) > 3 else ""
    new_pn = _txt(row[4]) if len(row) > 4 else ""
    new_type = _txt(row[5]) if len(row) > 5 else ""
    if reason.startswith("Value@需经NC但"):
        tail = reason.split("但", 1)[-1]
        inner = other_current_info((row[0], cmt, old_pn, old_type, new_pn, new_type, tail))
        return f"Value 已带 @ · {inner}"
    if reason == "新旧类型相同":
        head = cmt or "物料修改"
        pn = new_pn or old_pn or "—"
        return f"{head} · 料号 {pn} · Part Type {old_type or new_type or '—'}"
    if reason == "已带@无需删除":
        head = cmt or "物料删除"
        return f"{head} · 当前 Part Type {old_type or '—'} · Value {new_type or '—'}"
    if reason == "BOM无此位号":
        parts = [cmt] if cmt else []
        if old_pn:
            parts.append(f"变更前料号 {old_pn}")
        if new_pn:
            parts.append(f"变更后料号 {new_pn}")
        if new_type:
            parts.append(f"目标 {new_type}")
        parts.append("当前 Part Type 不在 BOM")
        return " · ".join(parts)
    if reason == "BOM无Value":
        head = cmt or "物料删除"
        return f"{head} · 当前 Part Type {old_type or '—'} · Value 空"
    if reason == "Value含非ASCII":
        head = cmt or "物料删除"
        return f"{head} · 当前 Part Type {old_type or '—'} · Value {new_type or '—'}"
    hop = f"当前 {old_type or '—'} → 目标 {new_type or '—'}"
    if cmt:
        hop = f"{cmt} · {hop}"
    if reason.startswith("库无NC:"):
        name = reason.split(":", 1)[-1].strip() or "—"
        return f"{hop} · 库里没有 {name}"
    if reason.startswith("封装无NC映射"):
        pkg = reason.split(":", 1)[-1].strip() or "—"
        return f"{hop} · 封装 {pkg}"
    if reason.startswith("无NC族"):
        return f"{hop} · 位号不属于 R/C/L"
    if reason.startswith("新增无NC映射"):
        detail = reason.replace("新增无NC映射", "", 1).strip() or "找不到 NC"
        return f"{hop} · {detail}"
    return reason


def summarize_abnormal(rows) -> dict:
    """未入库按料号汇总。其它异常仍在同一块里，但按位号列出当前信息。"""
    bad = abnormal_skips(rows)
    grouped: dict[tuple[str, str], int] = {}
    order: list[tuple[str, str]] = []
    other_rows: list[dict] = []
    for row in bad:
        reason = str(row[-1])
        if reason.startswith("新料号未入库:"):
            pn = str(row[4] or "").strip() or reason.split(":", 1)[-1].strip() or "(空料号)"
            part_type = str(row[5] or "").strip() or "—"
            key = (pn, part_type)
            if key not in grouped:
                grouped[key] = 0
                order.append(key)
            grouped[key] += 1
        else:
            other_rows.append({
                "ref": _txt(row[0]) or "—",
                "problem": other_problem(reason),
                "current": other_current_info(row),
            })
    library = [{"pn": pn, "part_type": pt, "count": grouped[(pn, pt)]} for pn, pt in order]
    return {
        "total": len(bad),
        "library": library,
        "library_refs": sum(item["count"] for item in library),
        "other": len(other_rows),
        "others": other_rows,
    }


def format_skip_dialog(rows) -> str:
    """弹窗同款摘要，给取消后的日志和命令行。未入库按料号汇总；其它异常按位号列出。"""
    info = summarize_abnormal(rows)
    lines = [
        f"有 {info['total']} 个位号因异常被跳过",
        "导入此版 ECO 将导致这些位号不被修改；其余位号仍可正常生成。",
        "",
    ]
    if info["library"]:
        lines.append(f"未入器件库  跳过 {info['library_refs']} 个位号")
        lines.append("器件库缺少下列料号，无法完成替换")
        lines.append("Part Number    Part Type    位号数")
        for item in info["library"]:
            lines.append(f"{item['pn']}    {item['part_type']}    {item['count']}")
        lines.append("")
    if info["other"]:
        lines.append(f"其他异常  跳过 {info['other']} 个位号")
        lines.append("位号    问题    当前信息")
        for item in info["others"]:
            lines.append(f"{item['ref']}    {item['problem']}    {item['current']}")
        lines.append("")
    lines.append("如需完整修改：请联系入库管理员完成器件入库，并同步更新工具库与原理图库后再重新生成 ECO。")
    return "\n".join(lines).rstrip() + "\n"


@dataclass
class EcoRun:
    """Generator result. skipped rows stay out of the ECO; pkg_changes are still emitted."""
    skipped: list
    pkg_changes: list


PKG_HINT = "请先打开 ECN 变更表，核对下列封装变化是否正确。确认无误后再继续生成。"


def format_pkg_dialog(rows, ecn=None) -> str:
    """弹窗同款摘要：请先核对这些封装变化，确认后再继续生成。"""
    hint = PKG_HINT
    if ecn:
        hint = (
            f"请先打开 ECN 变更表（{Path(ecn).name}），"
            "核对下列封装变化是否正确。确认无误后再继续生成。"
        )
    lines = [
        f"有 {len(rows)} 个位号封装发生变化",
        hint,
        "",
        "位号    变更类型    变更前封装    变更后封装",
    ]
    for row in rows:
        ref = _txt(row[0]) or "—"
        cmt = _txt(row[1]) if len(row) > 1 else ""
        before = _txt(row[2]) if len(row) > 2 else ""
        after = _txt(row[3]) if len(row) > 3 else ""
        lines.append(f"{ref}    {cmt or '—'}    {before or '—'}    {after or '—'}")
    return "\n".join(lines).rstrip() + "\n"

PKG_MAP = {
    "1005": "01005",
    "01005": "01005",
    "201": "0201",
    "0201": "0201",
    "402": "0402",
    "0402": "0402",
    "603": "0603",
    "0603": "0603",
    "805": "0805",
    "0805": "0805",
    "1206": "1206",
    "1210": "1210",
}
NC_PREFIX = {"R": "R_F_NC_", "C": "C_M_NC_", "L": "L_RF_NC_"}
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


_AUDIT_SAMPLE = 8
_HEADER_SCAN_ROWS = 8


def norm_name(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() in {"none", "nan"}:
        return None
    return s


def _is_name_header(h: str) -> bool:
    return bool(h) and (
        h in {"元件名称", "元件类型", "Part Type", "PART NAME"} or "元件名称" in h
    )


def _is_pn_header(h: str) -> bool:
    return bool(h) and (h.lower() == "part number" or h in {"物料编码", "Material Code"})


def _match_library_headers(headers: list[str]):
    name_col = pn_col = None
    name_header = pn_header = ""
    for i, h in enumerate(headers, start=1):
        if name_col is None and _is_name_header(h):
            name_col, name_header = i, h
        if pn_col is None and _is_pn_header(h):
            pn_col, pn_header = i, h
    return name_col, pn_col, name_header, pn_header


@dataclass
class LibraryAudit:
    path: Path
    sheet: str
    header_row: int
    headers: list[str]
    name_col: int | None
    pn_col: int | None
    name_header: str
    pn_header: str
    n_both: int
    n_name_only: int
    n_pn_only: int
    name_only_samples: list[tuple[int, str]]
    pn_only_samples: list[tuple[int, str]]

    @property
    def missing_columns(self) -> list[str]:
        missing = []
        if self.name_col is None:
            missing.append("元件名称")
        if self.pn_col is None:
            missing.append("Part Number")
        return missing


def library_block_reason(audit: LibraryAudit) -> str | None:
    """Stop generation when the library cannot be queried by name and Part Number."""
    if audit.missing_columns:
        seen = " | ".join(h for h in audit.headers if h) or "(空)"
        return (
            f"器件库缺少必需列: {'、'.join(audit.missing_columns)}。"
            f"至少要有两列: 元件名称、Part Number。"
            f" 实际表头: {seen}"
        )
    if audit.n_both == 0:
        return "器件库没有同时填写元件名称和 Part Number 的行，无法按料号查元件名称。"
    return None


def library_audit_brief(audit: LibraryAudit) -> str:
    if audit.missing_columns:
        return "缺列: " + "、".join(audit.missing_columns)
    if audit.n_both == 0:
        return "没有齐全行"
    parts = [f"齐全 {audit.n_both}"]
    if audit.n_name_only:
        parts.append(f"缺料号 {audit.n_name_only}")
    if audit.n_pn_only:
        parts.append(f"缺名称 {audit.n_pn_only}")
    return " · ".join(parts)


def _gap_lines(title: str, n: int, samples: list[tuple[int, str]], note: str) -> list[str]:
    if not n:
        return []
    lines = [f"  {title} {n} 行（{note}）"]
    for row, text in samples:
        lines.append(f"    行{row} {text}")
    rest = n - len(samples)
    if rest > 0:
        lines.append(f"    … 另有 {rest} 行")
    return lines


def format_library_audit(audit: LibraryAudit) -> str:
    lines = [f"库自查: {audit.path.name}  [{audit.sheet}]"]
    if audit.name_col:
        lines.append(f"  元件名称: 第{audit.name_col}列「{audit.name_header}」")
    else:
        lines.append("  元件名称: 未找到")
    if audit.pn_col:
        lines.append(f"  Part Number: 第{audit.pn_col}列「{audit.pn_header}」")
    else:
        lines.append("  Part Number: 未找到")
    if audit.missing_columns:
        seen = " | ".join(h for h in audit.headers if h) or "(空)"
        lines.append(f"  表头: {seen}")
        return "\n".join(lines)
    lines.append(f"  齐全 {audit.n_both} 行（元件名称和 Part Number 都有）")
    lines.extend(_gap_lines("缺 Part Number", audit.n_name_only, audit.name_only_samples, "有名称、无料号，不能按料号选用"))
    lines.extend(_gap_lines("缺元件名称", audit.n_pn_only, audit.pn_only_samples, "有料号、无名称"))
    return "\n".join(lines)


def _header_cells(values) -> list[str]:
    return [str(v or "").strip() for v in values]


def _library_from_rows(path: Path, sheet: str, rows):
    """rows yields (1-based row number, sequence of cell values)."""
    buffered = []
    chosen = None
    partial = None
    first_headers: list[str] = []
    for r, values in rows:
        buffered.append((r, values))
        if r > _HEADER_SCAN_ROWS and chosen is None and partial is not None:
            break
        headers = _header_cells(values)
        if r == 1:
            first_headers = headers
        name_col, pn_col, name_header, pn_header = _match_library_headers(headers)
        if name_col and pn_col:
            chosen = (r, headers, name_col, pn_col, name_header, pn_header)
            break
        if partial is None and (name_col or pn_col) and r <= _HEADER_SCAN_ROWS:
            partial = (r, headers, name_col, pn_col, name_header, pn_header)
        if r >= _HEADER_SCAN_ROWS and chosen is None:
            break
    if chosen is None:
        chosen = partial or (1, first_headers, None, None, "", "")
    header_row, headers, name_col, pn_col, name_header, pn_header = chosen

    pn_to_names: dict[str, list[str]] = defaultdict(list)
    all_names: set[str] = set()
    n_both = n_name_only = n_pn_only = 0
    name_only_samples: list[tuple[int, str]] = []
    pn_only_samples: list[tuple[int, str]] = []

    def consume(r, values):
        nonlocal n_both, n_name_only, n_pn_only
        if r <= header_row or not (name_col or pn_col):
            return
        name = norm_name(values[name_col - 1]) if name_col and name_col <= len(values) else None
        pn_val = values[pn_col - 1] if pn_col and pn_col <= len(values) else None
        codes = split_codes(pn_val) if pn_col else []
        if not name and not codes:
            return
        if name and codes:
            n_both += 1
            all_names.add(name)
            for code in codes:
                if name not in pn_to_names[code]:
                    pn_to_names[code].append(name)
        elif name:
            n_name_only += 1
            all_names.add(name)
            if len(name_only_samples) < _AUDIT_SAMPLE:
                name_only_samples.append((r, name))
        else:
            n_pn_only += 1
            if len(pn_only_samples) < _AUDIT_SAMPLE:
                pn_only_samples.append((r, ", ".join(codes)))

    for r, values in buffered:
        consume(r, values)
    for r, values in rows:
        consume(r, values)

    audit = LibraryAudit(
        path=path,
        sheet=sheet,
        header_row=header_row,
        headers=headers,
        name_col=name_col,
        pn_col=pn_col,
        name_header=name_header,
        pn_header=pn_header,
        n_both=n_both,
        n_name_only=n_name_only,
        n_pn_only=n_pn_only,
        name_only_samples=name_only_samples,
        pn_only_samples=pn_only_samples,
    )
    return audit, pn_to_names, all_names


def _read_library_csv(path: Path):
    text = read_text_any(path)
    first = text.splitlines()[0] if text else ""
    delim = "\t" if first.count("\t") > first.count(",") else ","
    rows = ((r, row) for r, row in enumerate(csv.reader(io.StringIO(text), delimiter=delim), start=1))
    return _library_from_rows(path, "CSV", rows)


def _read_library_xlsx(path: Path):
    wb = load_workbook_any(path)
    try:
        ws = wb[wb.sheetnames[0]]
        max_c = min(ws.max_column or 1, 40)

        def rows():
            for r in range(1, (ws.max_row or 1) + 1):
                yield r, [ws.cell(r, c).value for c in range(1, max_c + 1)]

        return _library_from_rows(path, ws.title, rows())
    finally:
        wb.close()


def _is_text_library(path: Path) -> bool:
    """CSV, or a text table saved under another suffix. Not a zip and not 亿赛通."""
    if path.suffix.lower() == ".csv":
        return True
    with path.open("rb") as f:
        head = f.read(512)
    if head[:2] == b"PK" or b"E-SafeNet" in head or b"ESafeNet" in head:
        return False
    return b"," in head or b"\t" in head


_LIB_FOLDER = "Electronic component library"


def _csv_headers_found(rows) -> bool:
    for values in rows[:_HEADER_SCAN_ROWS]:
        name_col, pn_col, _name_header, _pn_header = _match_library_headers(_header_cells(values))
        if name_col and pn_col:
            return True
    return False


def _csv_table(path: Path) -> list[list[str]]:
    """CSV rows. Plain text stays in-process. 亿赛通 ciphertext is read through Excel."""
    with path.open("rb") as f:
        head = f.read(512)
    if b"E-SafeNet" in head or b"ESafeNet" in head:
        from xlsx_open import read_csv_via_excel

        return read_csv_via_excel(path, _csv_headers_found)
    text = read_text_any(path)
    if not text.strip():
        raise ValueError("CSV 是空文件。")
    first = text.splitlines()[0]
    delim = "\t" if first.count("\t") > first.count(",") else ","
    return list(csv.reader(io.StringIO(text), delimiter=delim))


def import_library_csv(src: Path, dest: Path) -> int:
    """Read a CSV and write 元件类型 / Part Number into the tool's library xlsx.

    The xlsx columns stay 库 | 元件名称 | Part Number. A CSV header of 元件类型
    or 元件名称 is the name column. dest is not replaced when the CSV cannot be used.
    Returns the number of data rows written.
    """
    table = _csv_table(src)
    chosen = None
    partial = None
    first_headers: list[str] = []
    for i, values in enumerate(table[:_HEADER_SCAN_ROWS], start=1):
        headers = _header_cells(values)
        if i == 1:
            first_headers = headers
        name_col, pn_col, _name_header, _pn_header = _match_library_headers(headers)
        if name_col and pn_col:
            chosen = (i, name_col, pn_col, headers)
            break
        if partial is None and (name_col or pn_col):
            partial = (i, name_col, pn_col, headers)
    if chosen is None:
        chosen = partial or (1, None, None, first_headers)
    header_row, name_col, pn_col, headers = chosen
    if not name_col or not pn_col:
        seen = " | ".join(h for h in headers if h) or "(空)"
        missing = []
        if not name_col:
            missing.append("元件类型")
        if not pn_col:
            missing.append("Part Number")
        raise ValueError(
            f"CSV 缺少必需列: {'、'.join(missing)}。"
            f"至少要有两列: 元件类型（或元件名称）、Part Number。实际表头: {seen}"
        )

    pairs: list[tuple[str | None, str | None]] = []
    for values in table[header_row:]:
        name = norm_name(values[name_col - 1]) if name_col <= len(values) else None
        pn = norm_code(values[pn_col - 1]) if pn_col <= len(values) else None
        if not name and not pn:
            continue
        pairs.append((name, pn))
    if not any(name and pn for name, pn in pairs):
        raise ValueError("CSV 里没有同时填写元件类型和 Part Number 的行，器件库未改。")

    import openpyxl

    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["库", "元件名称", "Part Number"])
        for name, pn in pairs:
            ws.append([_LIB_FOLDER, name, pn])
        for cell in ws["C"][1:]:
            if cell.value is not None:
                cell.number_format = "@"
                cell.value = str(cell.value)
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_name(f".{dest.name}.{os.getpid()}.tmp.xlsx")
        try:
            wb.save(tmp)
            os.replace(tmp, dest)
        finally:
            if tmp.exists():
                tmp.unlink()
    finally:
        wb.close()
    return len(pairs)


def read_library(path: Path):
    """Read the component library. Columns are found by header, never by position.

    A CSV (or a text table that is not a real xlsx) is read as UTF-8/GBK text and
    does not go through Excel. A real .xlsx still uses the workbook reader,
    including 亿赛通 files via Excel.
    Returns (audit, pn_to_names, all_names). Names without a Part Number stay in
    all_names so NC lookup still works; they are not keyed by PN.
    """
    if _is_text_library(path):
        return _read_library_csv(path)
    return _read_library_xlsx(path)


def pick_name(names, hint=None):
    if not names:
        return None
    if len(names) == 1:
        return names[0]
    if hint:
        hint_s = str(hint).strip()
        for n in names:
            if hint_s and hint_s in n:
                return n
        tail = hint_s.split("_")[-1] if hint_s else ""
        if tail:
            for n in names:
                if n.endswith(tail) or tail in n:
                    return n
    return names[0]


def family_nc_prefix(part_type: str | None) -> str | None:
    if not part_type:
        return None
    for a, b in NAME_PREFIX:
        if part_type.startswith(a):
            return b
    return None


def is_nc(part_type: str | None) -> bool:
    return bool(part_type) and "_NC_" in part_type


def pkg_from_type(part_type: str | None) -> str | None:
    if not part_type:
        return None
    for tok in part_type.split("_"):
        if tok in PKG_MAP.values():
            return tok
    return None


def norm_pkg(pkg) -> str | None:
    """Normalize ECN/BOM package text. 1005→01005, 201→0201; unknown packages keep their text."""
    s = (norm_code(pkg) or "").upper().replace(" ", "")
    if not s:
        return None
    return PKG_MAP.get(s) or PKG_MAP.get(s.lower()) or s


def pkg_of_type(part_type: str | None) -> str | None:
    """Chip-package token in a library name, mapped through PKG_MAP (1005 and 01005 both work)."""
    if not part_type:
        return None
    for tok in part_type.split("_"):
        mapped = PKG_MAP.get(tok) or PKG_MAP.get(tok.upper()) or PKG_MAP.get(tok.lower())
        if mapped:
            return mapped
    return None


def detect_pkg_change(before_pkg, after_pkg, old_type=None, new_type=None):
    """Return (before, after) when the ECN packages or Part Type packages differ."""
    ecn_b, ecn_a = norm_pkg(before_pkg), norm_pkg(after_pkg)
    typ_b, typ_a = pkg_of_type(old_type), pkg_of_type(new_type)
    if ecn_b and ecn_a and ecn_b != ecn_a:
        return ecn_b, ecn_a
    if typ_b and typ_a and typ_b != typ_a:
        return typ_b, typ_a
    before = ecn_b or typ_b
    after = ecn_a or typ_a
    if before and after and before != after:
        return before, after
    return None


def nc_target(old_type, ref, pkg, all_names):
    pkg_s = (norm_code(pkg) or "").upper().replace(" ", "")
    suffix = PKG_MAP.get(pkg_s) or PKG_MAP.get(pkg_s.lower())
    if not suffix:
        return None, f"封装无NC映射:{pkg}"
    pref = family_nc_prefix(old_type)
    if pref is None and ref:
        pref = NC_PREFIX.get(ref[0].upper())
    if not pref:
        return None, f"无NC族:{ref}"
    name = pref + suffix
    if name not in all_names:
        return None, f"库无NC:{name}"
    return name, None


# BOM scanning lives in bomscan.py (shared with verify_eco.py)


def add_hops(old: str, new: str, pkg, all_names):
    # 物料新增 may retarget a part onto its own type: self → NC → self.
    # The only same-type no-op is a part that is already the target NC.
    if old == new and is_nc(old):
        return [], "新旧类型相同"
    if is_nc(old):
        # one jump only when current NC is same family as target
        if family_nc_prefix(old) == family_nc_prefix(new):
            return [(old, new)], None
        # cross-family NC falls through to two-jump via target family NC
    dest_pref = family_nc_prefix(new) or family_nc_prefix(old)
    raw = (norm_code(pkg) or "").upper().replace(" ", "")
    suffix = PKG_MAP.get(raw) or PKG_MAP.get(raw.lower()) or pkg_from_type(new) or pkg_from_type(old)
    if not dest_pref or not suffix:
        return [], f"新增无NC映射 pkg={pkg}"
    mid = dest_pref + suffix
    if mid not in all_names:
        return [], f"库无NC:{mid}"
    if mid == old or mid == new:
        return [(old, new)], None
    return [(old, mid), (mid, new)], None


def find_sheet(wb):
    for name in wb.sheetnames:
        if name.strip().lower() == "rf_ecn_generation":
            return wb[name]
    for name in wb.sheetnames:
        ws = wb[name]
        blob = " ".join(
            str(c.value)
            for r in ws.iter_rows(min_row=1, max_row=8, max_col=20)
            for c in r
            if c.value
        )
        if "物料修改" in blob or "物料删除" in blob or "物料新增" in blob or "Ref.Des" in blob:
            return ws
    return wb[wb.sheetnames[0]]


def find_header(ws):
    """Locate ECN header columns.

    Supports both single-row headers and two-row merged headers where the
    parent row carries Ref.Des / Before / After / Change Comments and the
    child row carries Value / Package / Material Code / Part Type.

    Returns (header_row, col_mapping, data_start_row).
    """
    aliases = {
        "ref": ["ref.des", "refdes", "位号"],
        "before_pn": ["before material code", "before-material code", "before material", "变更前物料"],
        "after_pn": ["after material code", "after-material code", "after material", "变更后物料"],
        "after_pt": ["part type", "after part type"],
        "before_pkg": ["before package", "变更前封装"],
        "after_pkg": ["after package", "变更后封装"],
        "comment": ["change comments", "change comment", "变更说明"],
    }
    max_scan = min(8, ws.max_row + 1)

    # 1. locate the row carrying Ref.Des
    ref_row = ref_col = None
    for r in range(1, max_scan):
        for c in range(1, ws.max_column + 1):
            raw = ws.cell(r, c).value
            if raw is None:
                continue
            key = str(raw).strip().lower().replace("\n", " ")
            if any(a in key for a in aliases["ref"]):
                ref_row, ref_col = r, c
                break
        if ref_row:
            break
    if not ref_row:
        raise SystemExit("Cannot find Ref.Des header")

    # 2. parent header at ref_row, with merged-cell propagation (left-fill)
    def parent_at(col):
        for cc in range(col, 0, -1):
            v = ws.cell(ref_row, cc).value
            if v is not None:
                return str(v).strip().lower().replace("\n", " ")
        return ""

    # 3. detect whether the next row is a sub-header (Value/Package/Material Code…)
    child_row = ref_row + 1
    sub_header_markers = {"value", "vaule", "package", "material code", "part type"}
    has_child = False
    if child_row <= ws.max_row:
        for c in range(1, ws.max_column + 1):
            v = ws.cell(child_row, c).value
            if v and str(v).strip().lower() in sub_header_markers:
                has_child = True
                break

    # 4. match columns using parent (+ child when present)
    mapping = {"ref": ref_col}
    for c in range(1, ws.max_column + 1):
        parent = parent_at(c)
        child = ""
        if has_child and child_row <= ws.max_row:
            cv = ws.cell(child_row, c).value
            if cv is not None:
                child = str(cv).strip().lower().replace("\n", " ")
        combined = f"{parent} {child}".strip()
        for role, names in aliases.items():
            if role in mapping:
                continue
            for cand in (combined, parent, child):
                if cand and any(a in cand for a in names):
                    mapping[role] = c
                    break

    data_start = (ref_row + 2) if has_child else (ref_row + 1)
    return ref_row, mapping, data_start


def write_eco(path: Path, remark: str, chg_rows, set_value_rows):
    lines = ["*PADS-ECO-V3.0-MILS*", f"*REMARK* {remark}"]
    if chg_rows:
        lines.append("*CHGPART*")
        for ref, old_t, new_t in chg_rows:
            lines.append(f"{ref} {old_t} {new_t}")
    if set_value_rows:
        # PADS Logic / ECOGEN: *SET_ATTRIBUTE* <TAB> PART <TAB> RefDes
        # then "Value" <TAB> @oldValue   (see TEST.eco from ECOGEN 6.4v)
        for ref, new_val in set_value_rows:
            lines.append(f"*SET_ATTRIBUTE*\tPART\t{ref}")
            lines.append(f'"Value"\t{new_val}')
    lines.append("*END*")
    try:
        path.write_bytes(("\r\n".join(lines) + "\r\n").encode("ascii"))
    except UnicodeEncodeError as e:
        raise SystemExit(f"ECO 含非 ASCII 字符，无法写出: {e}") from e


def run(ecn: Path, lib: Path, out_mod: Path, bom: Path | None = None, map_encoding: str = "utf-8-sig", bom_maps=None):
    audit, pn_to_names, all_names = read_library(lib)
    print(format_library_audit(audit))
    reason = library_block_reason(audit)
    if reason:
        raise SystemExit(reason)
    bom_type, bom_val = schematic_maps(bom, bom_maps)
    bom_at = {r for r, v in bom_val.items() if "@" in v}
    if not bom_type:
        if bom_maps is not None:
            from pads_logic_live import LOGIC_NO_TYPE
            raise SystemExit(LOGIC_NO_TYPE)
        raise SystemExit(
            f"BOM 没有 Part Type 列: {bom.name}\n"
            "  未找到任何 C_M_/L_RF_/R_F_/R_T_ 开头的元件名称。这可能是采购 BOM / EBOM / 下阶料号表，\n"
            "  请从 PADS Logic 原理图导出带 Part Type 的 BOM 后重试。"
        )
    wb = load_workbook_any(ecn)
    ws = find_sheet(wb)
    header_row, cols, data_start = find_header(ws)

    if "before_pkg" not in cols:
        for c in range(1, ws.max_column + 1):
            raw1 = ws.cell(header_row, c).value
            raw2 = ws.cell(header_row + 1, c).value if header_row + 1 <= ws.max_row else None
            raw = raw1 or raw2
            if raw and "package" in str(raw).lower() and c not in cols.values():
                cols["before_pkg"] = c
                break

    emitted_mod, emitted_mod_at, emitted_del, emitted_add, skipped = [], [], [], [], []
    pkg_changes = []

    def note_pkg(ref, cmt, before_pkg, after_pkg, old_type, new_type):
        pair = detect_pkg_change(before_pkg, after_pkg, old_type, new_type)
        if pair:
            pkg_changes.append((ref, cmt, pair[0], pair[1], old_type or "", new_type or ""))

    for r in range(data_start, ws.max_row + 1):
        cmt = str(ws.cell(r, cols["comment"]).value or "").strip() if "comment" in cols else ""
        if cmt not in (COMMENT_MODIFY, COMMENT_DELETE, COMMENT_ADD):
            continue
        ref = norm_code(ws.cell(r, cols["ref"]).value)
        if not ref:
            continue
        refs = [x.strip() for x in re.split(r"[,，\s]+", ref) if x.strip()]
        old_pn = norm_code(ws.cell(r, cols["before_pn"]).value) if "before_pn" in cols else None
        new_pn = norm_code(ws.cell(r, cols["after_pn"]).value) if "after_pn" in cols else None
        short = ws.cell(r, cols["after_pt"]).value if "after_pt" in cols else None
        type_hint = "" if short is None else str(short).strip()
        pkg = ws.cell(r, cols["before_pkg"]).value if "before_pkg" in cols else None
        apkg = ws.cell(r, cols["after_pkg"]).value if "after_pkg" in cols else None

        if cmt == COMMENT_ADD:
            news = pn_to_names.get(new_pn or "", [])
            if not news:
                for x in refs:
                    skipped.append((x, cmt, "", bom_type.get(x, ""), new_pn, type_hint, f"新料号未入库:{new_pn}"))
                continue
            new_type = pick_name(news, short)
            for x in refs:
                old_type = bom_type.get(x)
                if not old_type:
                    skipped.append((x, cmt, "", "", new_pn, new_type, "BOM无此位号"))
                    continue
                hops, err = add_hops(old_type, new_type, apkg or pkg, all_names)
                if err:
                    skipped.append((x, cmt, "", old_type, new_pn, new_type, err))
                    continue
                for a, b in hops:
                    emitted_add.append((x, a, b))
                note_pkg(x, cmt, pkg, apkg, old_type, new_type)
            continue

        # old_type always comes from schematic BOM, never from ECN Before PN
        for x in refs:
            old_type = bom_type.get(x)
            if not old_type:
                skipped.append((x, cmt, old_pn, "", new_pn, "", "BOM无此位号"))
                continue

            if cmt == COMMENT_DELETE:
                new_val = prefix_at(bom_val.get(x, ""))
                if new_val is None:
                    raw = (bom_val.get(x) or "").strip()
                    reason = "已带@无需删除" if raw.startswith("@") else "BOM无Value"
                    skipped.append((x, cmt, old_pn, old_type, "", raw, reason))
                    continue
                try:
                    new_val.encode("ascii")
                except UnicodeEncodeError:
                    skipped.append((x, cmt, old_pn, old_type, "", new_val, "Value含非ASCII"))
                    continue
                emitted_del.append((x, new_val))
                continue

            # modify
            new_names = pn_to_names.get(new_pn or "", [])
            if not new_names:
                skipped.append((x, cmt, old_pn, old_type, new_pn, type_hint, f"新料号未入库:{new_pn}"))
                continue
            new_type = pick_name(new_names, short)
            if old_type == new_type:
                skipped.append((x, cmt, old_pn, old_type, new_pn, new_type, "新旧类型相同"))
                continue
            if x in bom_at:
                # PADS Logic bug: a direct CHGPART on a part whose Value is '@'
                # replaces every part of that type board-wide. Route through NC
                # exactly like 物料新增, which is known to stay per-refdes.
                hops, err = add_hops(old_type, new_type, apkg or pkg, all_names)
                if err:
                    skipped.append((x, cmt, old_pn, old_type, new_pn, new_type, f"Value@需经NC但{err}"))
                    continue
                for a, b in hops:
                    emitted_mod_at.append((x, a, b))
                note_pkg(x, cmt, pkg, apkg, old_type, new_type)
                continue
            emitted_mod.append((x, old_type, new_type))
            note_pkg(x, cmt, pkg, apkg, old_type, new_type)

    chg_rows = emitted_mod + emitted_mod_at + emitted_add
    write_eco(out_mod, f"{ecn.name} modify+delAt+addViaNC", chg_rows, emitted_del)
    csv_path = out_mod.with_suffix(".map.csv")
    ascii_map = map_encoding == "ascii"
    act = (lambda s: ACTION_ASCII.get(s, s)) if ascii_map else (lambda s: s)
    st = STATUS_ASCII_FN if ascii_map else (lambda s: s)
    enc = "ascii" if ascii_map else map_encoding
    pkg_note = {row[0]: f"{row[2]}→{row[3]}" for row in pkg_changes}

    def emit_status(ref, base):
        note = pkg_note.get(ref)
        if not note:
            return base
        if base == "emit":
            return f"emit(封装变化:{note})"
        return f"{base};封装变化:{note}"

    with csv_path.open("w", newline="", encoding=enc, errors="replace") as f:
        w = csv.writer(f)
        w.writerow(["RefDes", "Action", "BeforePN", "OldType", "AfterPN", "NewType", "Status"])
        for ref, old_t, new_t in emitted_mod:
            w.writerow([ref, act(COMMENT_MODIFY), "", old_t, "", new_t, st(emit_status(ref, "emit"))])
        for ref, old_t, new_t in emitted_mod_at:
            w.writerow([ref, act(COMMENT_MODIFY), "", old_t, "", new_t, st(emit_status(ref, "emit(Value@:经NC两跳)"))])
        for ref, new_val in emitted_del:
            w.writerow([ref, act(COMMENT_DELETE), "", bom_type.get(ref, ""), "", new_val, st("emit(Value:@前缀)")])
        for ref, old_t, new_t in emitted_add:
            w.writerow([ref, act(COMMENT_ADD), "", old_t, "", new_t, st(emit_status(ref, "emit"))])
        for row in skipped:
            row = list(row)
            row[1] = act(row[1]); row[-1] = st(row[-1])
            w.writerow(row)
    print(f"modify {len(emitted_mod)} modify-via-NC {len(emitted_mod_at)} delete {len(emitted_del)} add {len(emitted_add)} -> {out_mod}")
    print(f"skip   {len(skipped)} -> {csv_path}")
    print(f"pkg-change {len(pkg_changes)}")
    missing = sum(1 for row in skipped if row[-1] == "BOM无此位号")
    total_refs = len(
        {row[0] for row in skipped}
        | {r for r, _, _ in emitted_mod + emitted_mod_at + emitted_add}
        | {r for r, _ in emitted_del}
    )
    if total_refs and missing / total_refs > 1 / 3:
        print(
            f"WARNING: {missing}/{total_refs} 个 ECN 位号在 BOM 中找不到。BOM 可能不完整（部分导出、折行记录未合并、"
            "或不是这块板的原理图 BOM）。请先核对 BOM 再导入。"
        )
    return EcoRun(skipped, pkg_changes)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ecn")
    ap.add_argument("--lib", default=str(DEFAULT_LIB))
    ap.add_argument("--bom", required=True, help="schematic BOM exported from PADS Logic, with Part Type column")
    ap.add_argument("-o", "--output", default="")
    ap.add_argument(
        "--map-encoding", choices=MAP_ENCODINGS, default="utf-8-sig",
        help="encoding of <out>.map.csv (default utf-8-sig for Excel; gbk for Chinese-Windows hosts; ascii = English tokens only)",
    )
    args = ap.parse_args()
    ecn = Path(args.ecn)
    out = Path(args.output) if args.output else ecn.with_suffix(".eco")
    run(ecn, Path(args.lib), out, Path(args.bom) if args.bom else None, map_encoding=args.map_encoding)


if __name__ == "__main__":
    main()
