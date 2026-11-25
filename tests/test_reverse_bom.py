#!/usr/bin/env python3
"""Reverse check: drop BOM rows whose 值 contains @ or is blank."""
import os
import sys
import tempfile
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from bomscan import ebom_rows_from_bom, preprocess_reverse_bom  # noqa: E402
from ecn_reverse import ecn_working_copy, interpret  # noqa: E402


def save(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    fd, name = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    path = Path(name)
    wb.save(path)
    wb.close()
    return path


def values(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = [tuple(c.value for c in row) for row in wb.active.iter_rows()]
    wb.close()
    return rows


def main() -> None:
    src = save([
        ["标题", None],
        ["位号", "值", "PART-TYPE"],
        ["C100", "0.3pF", "C_M_0P3"],
        ["L100", "@1.0nH", "L_RF_1N0"],
        ["R100", "@", "R_F_NC_0201"],
        ["C101", None, "C_M_NC_0201"],
        ["C102", "  ", "C_M_1N0"],
        ["L200", "2.0nH", "L_RF_2N0"],
    ])
    dest = src.with_name(src.stem + ".out.xlsx")
    try:
        removed_at, removed_blank, kept = preprocess_reverse_bom(src, dest)
        assert (removed_at, removed_blank, kept) == (2, 2, 2), (removed_at, removed_blank, kept)
        assert values(src)[2][1] == "0.3pF"
        out = values(dest)
        assert out[0][0] == "标题"
        assert out[1][0] == "位号"
        assert [r[0] for r in out[2:]] == ["C100", "L200"]
        assert out[2][1] == "0.3pF" and out[3][1] == "2.0nH"

        kept_path = src.with_name(src.stem + ".keep.xlsx")
        kept_path.write_bytes(b"KEEP")
        bad = save([["位号", "封装"], ["C100", "0201"]])
        try:
            preprocess_reverse_bom(bad, kept_path)
        except ValueError as e:
            assert "值" in str(e)
        else:
            raise SystemExit("missing 值 column should fail")
        assert kept_path.read_bytes() == b"KEEP"
        bad.unlink()
        kept_path.unlink()
    finally:
        src.unlink(missing_ok=True)
        dest.unlink(missing_ok=True)

    table = save([
        ["封面", None, None, None, None],
        ["下阶料号", "名称", "位号", "封装", "值"],
        ["0310401000562", "MLCC,0.3pF", "C100", "0201", "0.3pF"],
        ["111", "gone", "L100", "0201", "@1.0nH"],
        ["222", "blank", "R100", "0201", None],
        ["333", None, "C101", "0201", "1pF"],
        ["444", "keep", "C102", "0402", "10nF"],
    ])
    try:
        rows, at, blank, name = ebom_rows_from_bom(table)
        assert (at, blank, name, len(rows) - 1) == (1, 1, 1, 2), (at, blank, name, rows)
        assert rows[0] == ["下阶料号", "名称", "位号", "封装", "值"]
        assert rows[1][0] == "0310401000562" and rows[1][2] == "C100"
        assert rows[2][0] == "444" and rows[1][1] == "MLCC,0.3pF"
        assert values(table)[0][0] == "封面"
    finally:
        table.unlink(missing_ok=True)

    missing = save([["位号", "值"], ["C100", "1pF"]])
    try:
        ebom_rows_from_bom(missing)
    except ValueError as e:
        assert "下阶料号" in str(e)
    else:
        raise SystemExit("missing EBOM columns should fail")
    finally:
        missing.unlink(missing_ok=True)

    assert interpret(True, [], ["执行完毕"]) == (True, "ECN 为空")
    assert interpret(True, ["C100"], ["执行完毕"])[0] is False
    assert "仍有 1 行" in interpret(True, ["C100"], [])[1]
    assert interpret(False, [], ["工具过期，请获取最新工具"]) == (False, "宏中途退出，ECN 未重写")
    assert interpret(False, [], [])[0] is False
    assert interpret(True, [], ["执行完毕", "位号输入格式有误"])[0] is False

    folder = Path(tempfile.mkdtemp())
    try:
        original = folder / "FG201.xlsm"
        original.write_bytes(b"PK-original")
        backup = ecn_working_copy(original)
        assert backup.name == "FG201.反向检查.xlsm"
        assert backup.read_bytes() == b"PK-original"
        assert original.read_bytes() == b"PK-original"
        original.write_bytes(b"PK-second")
        backup.write_bytes(b"stale")
        again = ecn_working_copy(original)
        assert again.read_bytes() == b"PK-second"
        assert original.read_bytes() == b"PK-second"
        assert ecn_working_copy(again) == again.resolve()
        assert again.read_bytes() == b"PK-second"
    finally:
        for child in folder.iterdir():
            child.unlink(missing_ok=True)
        folder.rmdir()

    print("PASS reverse bom")


if __name__ == "__main__":
    main()
