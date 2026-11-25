#!/usr/bin/env python3
"""Library self-check: 元件名称 + Part Number are required before generation."""
import os
import sys
import tempfile
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from ecn_to_pads_eco import (  # noqa: E402
    format_library_audit,
    import_library_csv,
    library_block_reason,
    read_library,
    run,
)

LIB = Path(__file__).resolve().parents[1] / "library" / "Electronic_components_library.xlsx"


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


def main() -> None:
    audit, pn_to_names, names = read_library(LIB)
    assert audit.missing_columns == [], audit
    assert audit.name_header == "元件名称"
    assert audit.pn_header == "Part Number"
    assert audit.n_both > 0
    assert audit.n_name_only > 0  # NC / holes / test pads have no PN
    assert library_block_reason(audit) is None
    assert "C_M_NC_0201" in names
    assert "C_M_NC_0201" not in {n for ns in pn_to_names.values() for n in ns}
    text = format_library_audit(audit)
    assert text.startswith("库自查:")
    assert "元件名称: 第2列" in text
    assert "Part Number: 第3列" in text

    missing_pn = save([
        ["元件名称", "描述"],
        ["C_M_NC_0201", "nc"],
    ])
    audit, pn_to_names, names = read_library(missing_pn)
    assert audit.missing_columns == ["Part Number"]
    assert library_block_reason(audit)
    assert pn_to_names == {}
    assert names == {"C_M_NC_0201"}

    swapped = save([
        ["备注", "Part Number", "元件名称"],
        ["x", 310401000562, "C_M_0P3_50V_0201_V0R3A0201C0G500NAT"],
        ["y", None, "C_M_NC_0201"],
        ["z", 310401000999, None],
    ])
    audit, pn_to_names, names = read_library(swapped)
    assert audit.name_col == 3 and audit.pn_col == 2
    assert audit.n_both == 1 and audit.n_name_only == 1 and audit.n_pn_only == 1
    assert library_block_reason(audit) is None
    assert pn_to_names["310401000562"] == ["C_M_0P3_50V_0201_V0R3A0201C0G500NAT"]
    assert "C_M_NC_0201" in names

    # No header match: do not guess column 1/2.
    positional = save([
        ["AAA", "BBB"],
        ["C_M_NC_0201", "310401000562"],
    ])
    audit, _pn, names = read_library(positional)
    assert set(audit.missing_columns) == {"元件名称", "Part Number"}
    assert names == set()

    alias = save([
        ["元件名称", "物料编码"],
        ["R_F_10K_5%_0402_X", "310501000001"],
    ])
    audit, pn_to_names, _names = read_library(alias)
    assert audit.missing_columns == []
    assert pn_to_names["310501000001"] == ["R_F_10K_5%_0402_X"]

    empty = save([
        ["元件名称", "Part Number"],
        ["C_M_NC_0201", None],
    ])
    audit, _pn, _names = read_library(empty)
    assert audit.missing_columns == []
    assert "没有同时填写" in library_block_reason(audit)

    # Generation stops before writing an ECO when the library has no PN column.
    ecn = Path(__file__).resolve().parent / "test_ecn.xlsx"
    bom = Path(__file__).resolve().parent / "test_bom.xlsx"
    fd, name = tempfile.mkstemp(suffix=".eco")
    os.close(fd)
    out = Path(name)
    out.unlink()
    try:
        run(ecn, missing_pn, out, bom)
    except SystemExit as e:
        assert "Part Number" in str(e)
    else:
        raise SystemExit("run() should stop when Part Number column is missing")
    assert not out.exists()

    fd, name = tempfile.mkstemp(suffix=".csv")
    os.close(fd)
    csv_path = Path(name)
    csv_path.write_bytes(
        "元件名称,Part Number\r\n"
        "C_M_0P3_50V_0201_V0R3A0201C0G500NAT,310401000562\r\n"
        "C_M_NC_0201,\r\n".encode("gbk")
    )
    audit, pn_to_names, names = read_library(csv_path)
    assert audit.sheet == "CSV"
    assert audit.missing_columns == []
    assert audit.n_both == 1 and audit.n_name_only == 1
    assert pn_to_names["310401000562"] == ["C_M_0P3_50V_0201_V0R3A0201C0G500NAT"]
    assert "C_M_NC_0201" in names
    assert library_block_reason(audit) is None

    posed = Path(__file__).resolve().parents[1] / "dist" / "pads-eco" / "library" / "Electronic_components_library.xlsx"
    if posed.exists() and posed.read_bytes()[:2] != b"PK":
        audit, pn_to_names, _names = read_library(posed)
        assert audit.missing_columns == [], format_library_audit(audit)
        assert audit.n_both > 0
        assert "331300000179" in pn_to_names

    # Update path: CSV in, tool library xlsx out. A bad CSV must not replace the xlsx.
    fd, name = tempfile.mkstemp(suffix=".csv")
    os.close(fd)
    src_csv = Path(name)
    src_csv.write_bytes(
        "元件类型,Part Number\r\n"
        "C_M_0P3_50V_0201_V0R3A0201C0G500NAT,310401000562\r\n"
        "C_M_NC_0201,\r\n".encode("gbk")
    )
    fd, name = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    dest = Path(name)
    dest.write_bytes(b"KEEP")
    try:
        n = import_library_csv(src_csv, dest)
        assert n == 2
        assert dest.read_bytes()[:2] == b"PK"
        wb = openpyxl.load_workbook(dest)
        ws = wb.active
        assert [c.value for c in ws[1]] == ["库", "元件名称", "Part Number"]
        assert ws["B2"].value == "C_M_0P3_50V_0201_V0R3A0201C0G500NAT"
        assert ws["C2"].value == "310401000562"
        assert ws["B3"].value == "C_M_NC_0201"
        assert ws["C3"].value is None
        wb.close()
        audit, pn_to_names, names = read_library(dest)
        assert audit.name_col == 2 and audit.pn_col == 3
        assert pn_to_names["310401000562"] == ["C_M_0P3_50V_0201_V0R3A0201C0G500NAT"]
        assert "C_M_NC_0201" in names
        assert library_block_reason(audit) is None

        kept = dest.with_name(dest.name + ".keep.xlsx")
        kept.write_bytes(b"KEEP")
        bad = src_csv.with_name(src_csv.name + ".bad.csv")
        bad.write_bytes("只有名称\r\nC_M_NC_0201\r\n".encode("gbk"))
        try:
            import_library_csv(bad, kept)
        except ValueError as e:
            assert "Part Number" in str(e)
        else:
            raise SystemExit("import should reject a CSV without Part Number")
        assert kept.read_bytes() == b"KEEP"

        # Ciphertext is not rejected: Excel (trusted by 亿赛通) supplies the rows.
        locked = src_csv.with_name(src_csv.name + ".lock.csv")
        locked.write_bytes(b"E-SafeNet LOCK\r\nnot,a,table\r\n")
        import xlsx_open
        from ecn_to_pads_eco import _csv_headers_found

        def fake_excel(path, accept):
            rows = [
                ["元件类型", "Part Number"],
                ["C_M_FROM_EXCEL", "0310401000562"],
            ]
            assert accept(rows)
            assert Path(path).name == locked.name
            return rows

        orig = xlsx_open.read_csv_via_excel
        xlsx_open.read_csv_via_excel = fake_excel
        try:
            n = import_library_csv(locked, kept)
        finally:
            xlsx_open.read_csv_via_excel = orig
        assert n == 1
        assert kept.read_bytes()[:2] == b"PK"
        wb = openpyxl.load_workbook(kept)
        assert wb.active["B2"].value == "C_M_FROM_EXCEL"
        assert wb.active["C2"].value == "0310401000562"
        wb.close()

        live = src_csv.with_name(src_csv.name + ".live.csv")
        live.write_bytes(
            "元件类型,Part Number\r\n"
            "CONN_BTB_100P_FX8-100P-SV1(92),330509000078\r\n"
            "CONN_BATT_M0641-LF,0310401000562\r\n".encode("gbk")
        )
        rows = orig(live, _csv_headers_found)
        assert rows[1] == ["CONN_BTB_100P_FX8-100P-SV1(92)", "330509000078"]
        assert rows[2] == ["CONN_BATT_M0641-LF", "0310401000562"]
    finally:
        for p in (
            src_csv,
            dest,
            dest.with_name(dest.name + ".keep.xlsx"),
            src_csv.with_name(src_csv.name + ".bad.csv"),
            src_csv.with_name(src_csv.name + ".lock.csv"),
            src_csv.with_name(src_csv.name + ".live.csv"),
        ):
            if p.exists():
                p.unlink()

    print("PASS library audit")
    print(format_library_audit(read_library(LIB)[0]))


if __name__ == "__main__":
    main()
