#!/usr/bin/env python3
"""Generate synthetic test fixtures for pads-logic-eco regression.

Creates tests/test_ecn.xlsx, tests/test_bom.xlsx, tests/test_golden.eco
covering all three action types: modify, delete=prefix Value with @, add-via-NC (cross-family),
plus a modify on a Value='@' part (must be routed via NC).
"""
import openpyxl
from pathlib import Path

TEST_DIR = Path(__file__).resolve().parent

# --- test BOM ---
bom_wb = openpyxl.Workbook()
bws = bom_wb.active
bws.title = "Sheet1"
bws.append(["下阶料号", "名称", "位号", "封装", "值", "PART-TYPE"])
bws.append([310401000562, "MLCC 0.3pF", "C100", "0201", "0.3pF",
            "C_M_0P3_50V_0201_V0R3A0201C0G500NAT,0.3pF,±0.05pF"])
bws.append([311001000253, "RF Inductor 1.0nH", "L100", "0201", "1.0nH",
            "L_RF_1N0_750M_0201_SDCL0603Q1N0BT02B03,1.0nH"])
bws.append([None, "NC placeholder", "R100", "0201", "@",
            "C_M_NC_0201,@"])
# C101: real part whose Value is '@' -> a modify must be routed via NC (Value @ guard)
bws.append([310401000562, "MLCC 0.3pF", "C101", "0201", "@",
            "C_M_0P3_50V_0201_V0R3A0201C0G500NAT,@"])
# Wrapped record (as produced by the real export when the refdes list is long):
# row N = PN / desc / truncated list, no type; row N+1 = rest of list (mid-token) + pkg/value/type shifted left
bws.append([311001000253, "RF Inductor 1.0nH", "L200,L201,L20", None, None, None])
bws.append(["2,L203", "0201", "1.0nH", "L_RF_1N0_750M_0201_SDCL0603Q1N0BT02B03,1.0nH", None, None])
# L301: already the After library name; 物料新增 still two-hops through NC
bws.append([311001000228, "RF Inductor 12nH", "L301", "01005", "@",
            "L_RF_12N_140M_01005_MLG0402P12NHT000,@"])
bom_wb.save(TEST_DIR / "test_bom.xlsx")

# --- test ECN (two-row header like real FG201 format) ---
ecn_wb = openpyxl.Workbook()
ews = ecn_wb.active
ews.title = "RF_ECN_Generation"
ews.append(["TEST-BOARD", None, "ECN Date:", "2026-01-01"])
ews.append(["Ref.Des", "Before", None, None, "After", None, None, None, None, "Change Comments", "Remark"])
ews.append([None, "Vaule", "Package", "Material Code", "Vaule", "Package", "Material Code", "Part Type"])
# C100: modify C_M_0P3 -> C_M_0P5
ews.append(["C100", "0.3pF", "0201", 310401000562, "0.5pF", "0201", 310401000559,
            "50V_0201_V0R5A0201C0G500NAT", None, "物料修改", "test modify"])
# C101: same modify, but BOM Value is '@' -> expect two-hop via C_M_NC_0201
ews.append(["C101", "0.3pF", "0201", 310401000562, "0.5pF", "0201", 310401000559,
            "50V_0201_V0R5A0201C0G500NAT", None, "物料修改", "test modify with Value @"])
# L100: delete -> prefix BOM Value with @
ews.append(["L100", "1.0nH", "0201", 311001000253, None, None, None, None, None, "物料删除", "test delete"])
# L202: sits on the wrapped BOM record, split mid-token ("L20" + "2") -> modify to 2.0nH
ews.append(["L202", "1.0nH", "0201", 311001000253, "2.0nH", "0201", 311001000302,
            "600M_0201_SDCL0603Q2N0BT02B03", None, "物料修改", "test wrapped BOM record"])
# R100: add cross-family C_M_NC -> L_RF_2N0
ews.append(["R100", None, "0201", None, "2.0nH", "0201", 311001000302,
            "600M_0201_SDCL0603Q2N0BT02B03", None, "物料新增", "test add cross-family"])
# L301: 物料新增 onto the part's own type -> self → NC → self
ews.append(["L301", None, None, None, "12nH", "01005", 311001000228,
            "140M_01005_MLG0402P12NHT000", None, "物料新增", "test add same type via NC"])
ecn_wb.save(TEST_DIR / "test_ecn.xlsx")

# --- golden ECO ---
golden_lines = [
    "*PADS-ECO-V3.0-MILS*",
    "*REMARK* TEST-BOARD modify+delAt+addViaNC  current-from-sch-BOM",
    "*CHGPART*",
    "C100 C_M_0P3_50V_0201_V0R3A0201C0G500NAT C_M_0P5_50V_0201_V0R5A0201C0G500NAT",
    "L202 L_RF_1N0_750M_0201_SDCL0603Q1N0BT02B03 L_RF_2N0_600M_0201_SDCL0603Q2N0BT02B03",
    "C101 C_M_0P3_50V_0201_V0R3A0201C0G500NAT C_M_NC_0201",
    "C101 C_M_NC_0201 C_M_0P5_50V_0201_V0R5A0201C0G500NAT",
    "R100 C_M_NC_0201 L_RF_NC_0201",
    "R100 L_RF_NC_0201 L_RF_2N0_600M_0201_SDCL0603Q2N0BT02B03",
    "L301 L_RF_12N_140M_01005_MLG0402P12NHT000 L_RF_NC_01005",
    "L301 L_RF_NC_01005 L_RF_12N_140M_01005_MLG0402P12NHT000",
    "*SET_ATTRIBUTE*\tPART\tL100",
    '"Value"\t@1.0nH',
    "*END*",
]
(TEST_DIR / "test_golden.eco").write_bytes(
    ("\r\n".join(golden_lines) + "\r\n").encode("ascii")
)

print("Test fixtures created in", TEST_DIR)
for f in sorted(TEST_DIR.glob("test_*")):
    print(f"  {f.name}  ({f.stat().st_size} bytes)")
