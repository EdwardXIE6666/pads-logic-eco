#!/usr/bin/env python3
"""A failed verify must not leave an .eco or a .FAIL.eco, and must print the reasons."""
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tool"))
sys.path.insert(0, str(ROOT / "scripts"))

import pads_eco_app  # noqa: E402
import verify_eco  # noqa: E402

ECN = ROOT / "tests" / "test_ecn.xlsx"
BOM = ROOT / "tests" / "test_bom.xlsx"
LIB = ROOT / "library" / "Electronic_components_library.xlsx"


def main() -> None:
    real_checks = verify_eco.run_checks
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "board.eco"
        out.write_text("KEEP\n", encoding="ascii")
        lines: list[str] = []

        def fail_checks(_eco, _bom, _lib, bom_maps=None):
            return ["D2: C100 new type 'NOPE' not in library", "D8: C101 direct modify"]

        verify_eco.run_checks = fail_checks
        try:
            ok = pads_eco_app.do_gen(ECN, BOM, LIB, out, lines.append)
        finally:
            verify_eco.run_checks = real_checks

        assert ok is False
        assert out.read_text(encoding="ascii") == "KEEP\n"
        assert not out.with_suffix(".map.csv").exists()
        assert not out.with_name(out.stem + ".FAIL.eco").exists()
        assert not list(Path(td).glob("*.tmp.eco"))
        assert not list(Path(td).glob(".*.tmp.eco"))
        text = "\n".join(lines)
        assert "FAIL  2 项未通过" in text
        assert "D2: C100 new type 'NOPE' not in library" in text
        assert "D8: C101 direct modify" in text
        assert "未生成 ECO" in text

        lines.clear()
        ok = pads_eco_app.do_gen(ECN, BOM, LIB, out, lines.append)
        assert ok is True, text
        assert out.exists() and out.read_text(encoding="ascii") != "KEEP\n"
        assert out.with_suffix(".map.csv").exists()
        assert "PASS" in "\n".join(lines)
        assert not out.with_name(out.stem + ".FAIL.eco").exists()

    print("PASS no .FAIL.eco; reasons printed; passed run writes .eco")


if __name__ == "__main__":
    main()
