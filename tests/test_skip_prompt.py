#!/usr/bin/env python3
"""每一次跳过都要进弹窗。用户取消时不留下 ECO。"""
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tool"))
sys.path.insert(0, str(ROOT / "scripts"))

import ecn_to_pads_eco  # noqa: E402
import pads_eco_app  # noqa: E402

ECN = ROOT / "tests" / "test_ecn.xlsx"
BOM = ROOT / "tests" / "test_bom.xlsx"
LIB = ROOT / "library" / "Electronic_components_library.xlsx"

ABNORMAL = ("C101", "物料修改", "111", "C_M_OLD", "222333", "C_M_10P_0402", "新料号未入库:222333")
SAME_PN = ("C109", "物料修改", "111", "C_M_OLD", "222333", "C_M_10P_0402", "新料号未入库:222333")
OTHER = ("L201", "物料新增", "", "", "444", "L_RF_X", "BOM无此位号")
OTHER_2 = ("R301", "物料删除", "555", "R_F_10K", "", "", "BOM无Value")
OTHER_3 = ("R302", "物料删除", "555", "R_F_10K", "", "", "BOM无Value")
BENIGN = ("C102", "物料修改", "", "C_M_10P_0402", "", "C_M_10P_0402", "新旧类型相同")
ALREADY = ("R1", "物料删除", "", "R_F_10K_0402", "", "@10k", "已带@无需删除")


def test_classify() -> None:
    bad = ecn_to_pads_eco.abnormal_skips([ABNORMAL, SAME_PN, OTHER, OTHER_2, OTHER_3, BENIGN, ALREADY])
    assert [row[0] for row in bad] == ["C101", "C109", "L201", "R301", "R302", "C102", "R1"]
    info = ecn_to_pads_eco.summarize_abnormal(bad)
    assert info["total"] == 7
    assert info["library_refs"] == 2
    assert info["library"] == [{"pn": "222333", "part_type": "C_M_10P_0402", "count": 2}]
    assert info["other"] == 5
    assert info["others"] == [
        {
            "ref": "L201",
            "problem": "BOM 缺位号",
            "current": "物料新增 · 变更后料号 444 · 目标 L_RF_X · 当前 Part Type 不在 BOM",
        },
        {
            "ref": "R301",
            "problem": "删除缺 Value",
            "current": "物料删除 · 当前 Part Type R_F_10K · Value 空",
        },
        {
            "ref": "R302",
            "problem": "删除缺 Value",
            "current": "物料删除 · 当前 Part Type R_F_10K · Value 空",
        },
        {
            "ref": "C102",
            "problem": "新旧类型相同",
            "current": "物料修改 · 料号 — · Part Type C_M_10P_0402",
        },
        {
            "ref": "R1",
            "problem": "已带@无需删除",
            "current": "物料删除 · 当前 Part Type R_F_10K_0402 · Value @10k",
        },
    ]
    same = pads_eco_app._skip_modal_line(BENIGN)
    assert same == {"ref": "C102", "pn": "—", "pt": "C_M_10P_0402", "reason": "新旧类型相同"}
    already = pads_eco_app._skip_modal_line(ALREADY)
    assert already["pt"] == "R_F_10K_0402"
    assert already["reason"] == "已带@无需删除"
    assert "@10k" not in already["pt"]
    text = ecn_to_pads_eco.format_skip_dialog(bad)
    assert text.count("其他异常") == 1
    assert "BOM 缺位号" in text
    assert "当前 Part Type 不在 BOM" in text
    assert "Value 空" in text
    assert "不逐条展开" not in text
    assert "C101" not in text
    assert "L201" in text
    assert "R301" in text
    assert "C102" in text
    assert "新旧类型相同" in text
    assert "R1" in text
    assert "已带@无需删除" in text


def _patch(rows):
    real = ecn_to_pads_eco.run

    def fake_run(ecn, lib, out, bom, map_encoding="utf-8-sig"):
        out.write_bytes(b"*END*\r\n")
        out.with_suffix(".map.csv").write_text("x\n", encoding="utf-8")
        print("modify 0 modify-via-NC 0 delete 0 add 0 -> x")
        print(f"skip   {len(rows)} -> x")
        return list(rows)

    ecn_to_pads_eco.run = fake_run
    return real


def test_cancel_discards() -> None:
    real = _patch([ABNORMAL])
    asked = []
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []

            def confirm(rows):
                asked.append(rows)
                return False

            ok = pads_eco_app.do_gen(ECN, BOM, LIB, out, lines.append, confirm)
            blob = "\n".join(lines)
            assert ok is False
            assert asked == [[ABNORMAL]]
            assert out.read_text(encoding="ascii") == "KEEP\n"
            assert not out.with_suffix(".map.csv").exists()
            assert "已取消。未生成 ECO。" in blob
            assert "未入器件库" in blob
            assert "222333" in blob
            assert "无法完成替换" in blob
            assert "C101" not in blob
    finally:
        ecn_to_pads_eco.run = real


def test_every_skip_asks() -> None:
    real = _patch([BENIGN, ALREADY])
    asked = []
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []
            ok = pads_eco_app.do_gen(ECN, BOM, LIB, out, lines.append, lambda rows: asked.append(rows) or False)
            assert asked == [[BENIGN, ALREADY]]
            assert ok is False
            assert out.read_text(encoding="ascii") == "KEEP\n"
            assert "已取消。未生成 ECO。" in "\n".join(lines)
            assert "新旧类型相同" in "\n".join(lines)
            assert "已带@无需删除" in "\n".join(lines)
    finally:
        ecn_to_pads_eco.run = real


def test_continue_keeps_going() -> None:
    real = _patch([ABNORMAL])
    try:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "board.eco"
            out.write_text("KEEP\n", encoding="ascii")
            lines: list[str] = []
            ok = pads_eco_app.do_gen(ECN, BOM, LIB, out, lines.append, lambda rows: True)
            blob = "\n".join(lines)
            assert ok is False
            assert "已选择继续" in blob
            assert "已取消" not in blob
            assert out.read_text(encoding="ascii") == "KEEP\n"
    finally:
        ecn_to_pads_eco.run = real


if __name__ == "__main__":
    test_classify()
    test_cancel_discards()
    test_every_skip_asks()
    test_continue_keeps_going()
    print("PASS skip prompt")
