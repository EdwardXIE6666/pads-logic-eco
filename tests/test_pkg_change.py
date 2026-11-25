#!/usr/bin/env python3
"""封装变化要进弹窗。请先核对这些封装，确认后再继续；用户取消时不留下文件。"""
import os
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tool"))
sys.path.insert(0, str(ROOT / "scripts"))

import ecn_to_pads_eco  # noqa: E402
import pads_eco_app  # noqa: E402

PKG = ("C100", "物料修改", "0201", "0402", "C_M_0P3_50V_0201_X", "C_M_0P5_50V_0402_Y")


def test_user_gen_log() -> None:
    raw = (
        "库自查: lib.xlsx  [Sheet1]\n"
        "  元件名称: 第2列「元件名称」\n"
        "  Part Number: 第3列「Part Number」\n"
        "  齐全 10 行（元件名称和 Part Number 都有）\n"
        "  缺 Part Number 2 行（有名称、无料号，不能按料号选用）\n"
        "    行4 C_M_NC_0201\n"
        "    … 另有 1 行\n"
        "modify 2 modify-via-NC 2 delete 1 add 4 -> x.eco\n"
        "skip   0 -> x.map.csv\n"
        "pkg-change 0\n"
        "WARNING: 3/9 个 ECN 位号在 BOM 中找不到。\n"
    )
    text = pads_eco_app._user_gen_log(raw)
    assert "库自查" not in text
    assert "元件名称" not in text
    assert "C_M_NC_0201" not in text
    assert "修改 4，删除 1，新增 4" in text
    assert "跳过" not in text
    assert "封装变化" not in text
    assert "WARNING: 3/9" in text
    skipped = pads_eco_app._user_gen_log("skip   3 -> x.map.csv\npkg-change 2\n")
    assert "跳过 3 个位号，原因见对照表" in skipped
    assert "封装变化 2 个位号" in skipped


def test_detect() -> None:
    d = ecn_to_pads_eco.detect_pkg_change
    assert d("0201", "0201") is None
    assert d("201", "0201") is None
    assert d("1005", "01005") is None
    assert d("0201", "0402") == ("0201", "0402")
    assert d("201", "402") == ("0201", "0402")
    assert d(None, None, "C_M_0P3_50V_0201_X", "C_M_0P5_50V_0201_Y") is None
    assert d(None, None, "C_M_0P3_50V_0201_X", "C_M_0P5_50V_0402_Y") == ("0201", "0402")
    assert d("0201", "0201", "C_M_A_0201_X", "C_M_B_0402_Y") == ("0201", "0402")
    assert d("0201", None, "C_M_A_0201_X", "C_M_B_0402_Y") == ("0201", "0402")
    assert d("SOT23", "QFN24") == ("SOT23", "QFN24")
    assert d(None, "0402", None, None) is None


def test_format_and_line() -> None:
    text = ecn_to_pads_eco.format_pkg_dialog([PKG])
    assert "有 1 个位号封装发生变化" in text
    assert "请先打开 ECN 变更表" in text
    assert "确认无误后再继续生成" in text
    assert "仍会写入 ECO" not in text
    assert "C100" in text
    assert "0201" in text
    assert "0402" in text
    named = ecn_to_pads_eco.format_pkg_dialog([PKG], Path("D:/boards/Bom&ECN-changelist.xlsx"))
    assert "Bom&ECN-changelist.xlsx" in named
    pads_eco_app._open_ecn_file(None)
    pads_eco_app._open_ecn_file(Path("C:/no/such/ecn.xlsx"))
    line = pads_eco_app._pkg_modal_line(PKG)
    assert line == {
        "ref": "C100",
        "action": "物料修改",
        "before": "0201",
        "after": "0402",
        "pair": "0201 → 0402",
    }


def _patch(pkg_rows, skipped=None):
    real = ecn_to_pads_eco.run

    def fake_run(ecn, lib, out, bom, map_encoding="utf-8-sig", bom_maps=None):
        out.write_bytes(b"*PADS-ECO-V3.0-MILS*\r\n*END*\r\n")
        out.with_suffix(".map.csv").write_text("x\n", encoding="utf-8")
        print("modify 1 modify-via-NC 0 delete 0 add 0 -> x")
        print("skip   0 -> x")
        print(f"pkg-change {len(pkg_rows)}")
        return ecn_to_pads_eco.EcoRun(list(skipped or []), list(pkg_rows))

    ecn_to_pads_eco.run = fake_run
    return real


def test_cancel_discards() -> None:
    real = _patch([PKG])
    asked = []
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []
            ok = pads_eco_app.do_gen(
                ROOT / "tests" / "test_ecn.xlsx",
                ROOT / "tests" / "test_bom.xlsx",
                ROOT / "library" / "Electronic_components_library.xlsx",
                out,
                lines.append,
                confirm_pkg=lambda rows: asked.append(rows) or False,
            )
            blob = "\n".join(lines)
            assert ok is False
            assert asked == [[PKG]]
            assert out.read_text(encoding="ascii") == "KEEP\n"
            assert not out.with_suffix(".map.csv").exists()
            assert "已取消。未生成 ECO。" in blob
            assert "C100" in blob
            assert "封装发生变化" in blob
    finally:
        ecn_to_pads_eco.run = real


def test_continue_keeps_going() -> None:
    real = _patch([PKG])
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []
            ok = pads_eco_app.do_gen(
                ROOT / "tests" / "test_ecn.xlsx",
                ROOT / "tests" / "test_bom.xlsx",
                ROOT / "library" / "Electronic_components_library.xlsx",
                out,
                lines.append,
                confirm_pkg=lambda rows: True,
            )
            blob = "\n".join(lines)
            assert "已选择继续" in blob
            assert "封装变化 1 个位号" in blob
            assert "已取消" not in blob
            # mocked eco fails verify; destination stays KEEP
            assert ok is False
            assert out.read_text(encoding="ascii") == "KEEP\n"
    finally:
        ecn_to_pads_eco.run = real


def test_cli_logs_and_continues() -> None:
    real = _patch([PKG])
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []
            pads_eco_app.do_gen(
                ROOT / "tests" / "test_ecn.xlsx",
                ROOT / "tests" / "test_bom.xlsx",
                ROOT / "library" / "Electronic_components_library.xlsx",
                out,
                lines.append,
            )
            blob = "\n".join(lines)
            assert "有 1 个位号封装发生变化" in blob
            assert "请先打开 ECN 变更表" in blob
            assert "test_ecn.xlsx" in blob
            assert "C100" in blob
            assert "已取消" not in blob
    finally:
        ecn_to_pads_eco.run = real


def _save(rows, suffix=".xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    fd, name = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    path = Path(name)
    wb.save(path)
    wb.close()
    return path


def test_run_collects_pkg_change() -> None:
    lib = _save([
        ["元件名称", "Part Number"],
        ["C_M_0P3_50V_0201_OLD", "310401000562"],
        ["C_M_0P5_50V_0402_NEW", "310401000559"],
        ["C_M_NC_0201", None],
        ["C_M_NC_0402", None],
    ])
    bom = _save([
        ["下阶料号", "名称", "位号", "封装", "值", "PART-TYPE"],
        [310401000562, "MLCC", "C100", "0201", "0.3pF", "C_M_0P3_50V_0201_OLD,0.3pF"],
        [310401000562, "MLCC", "C200", "0201", "0.3pF", "C_M_0P3_50V_0201_OLD,0.3pF"],
    ])
    ecn = _save([
        ["TEST-BOARD"],
        ["Ref.Des", "Before", None, None, "After", None, None, None, None, "Change Comments"],
        [None, "Vaule", "Package", "Material Code", "Vaule", "Package", "Material Code", "Part Type"],
        ["C100", "0.3pF", "0201", 310401000562, "0.5pF", "0402", 310401000559, "50V_0402_NEW", None, "物料修改"],
        ["C200", "0.3pF", "0201", 310401000562, "0.5pF", "0201", 310401000559, "50V_0402_NEW", None, "物料修改"],
    ])
    fd, name = tempfile.mkstemp(suffix=".eco")
    os.close(fd)
    out = Path(name)
    try:
        result = ecn_to_pads_eco.run(ecn, lib, out, bom)
        refs = [row[0] for row in result.pkg_changes]
        assert "C100" in refs
        assert result.pkg_changes[refs.index("C100")][2:4] == ("0201", "0402")
        # C200: ECN packages same, but After Part Type is 0402 → still a change
        assert "C200" in refs
        text = out.with_suffix(".map.csv").read_text(encoding="utf-8-sig")
        assert "封装变化:0201→0402" in text
        eco = out.read_text(encoding="ascii")
        assert "C100 C_M_0P3_50V_0201_OLD C_M_0P5_50V_0402_NEW" in eco
    finally:
        for p in (lib, bom, ecn, out, out.with_suffix(".map.csv")):
            try:
                p.unlink()
            except OSError:
                pass


def test_same_pkg_silent() -> None:
    lib = _save([
        ["元件名称", "Part Number"],
        ["C_M_0P3_50V_0201_OLD", "310401000562"],
        ["C_M_0P5_50V_0201_NEW", "310401000559"],
        ["C_M_NC_0201", None],
    ])
    bom = _save([
        ["下阶料号", "名称", "位号", "封装", "值", "PART-TYPE"],
        [310401000562, "MLCC", "C100", "0201", "0.3pF", "C_M_0P3_50V_0201_OLD,0.3pF"],
    ])
    ecn = _save([
        ["TEST-BOARD"],
        ["Ref.Des", "Before", None, None, "After", None, None, None, None, "Change Comments"],
        [None, "Vaule", "Package", "Material Code", "Vaule", "Package", "Material Code", "Part Type"],
        ["C100", "0.3pF", "0201", 310401000562, "0.5pF", "0201", 310401000559, "50V_0201_NEW", None, "物料修改"],
    ])
    fd, name = tempfile.mkstemp(suffix=".eco")
    os.close(fd)
    out = Path(name)
    try:
        result = ecn_to_pads_eco.run(ecn, lib, out, bom)
        assert result.pkg_changes == []
        assert result.skipped == []
    finally:
        for p in (lib, bom, ecn, out, out.with_suffix(".map.csv")):
            try:
                p.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    test_detect()
    test_format_and_line()
    test_user_gen_log()
    test_cancel_discards()
    test_continue_keeps_going()
    test_cli_logs_and_continues()
    test_run_collects_pkg_change()
    test_same_pkg_silent()
    print("PASS pkg change")
