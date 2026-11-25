# Regression test fixtures

Small synthetic ECN + BOM + golden ECO covering all three action types.

## Files

| File | Purpose |
|---|---|
| `test_ecn.xlsx` | Two-row-header ECN with 1 modify + 1 delete + 1 add (cross-family) |
| `test_bom.xlsx` | Schematic BOM: C100, C101 (Value `@`), L100, R100, plus a wrapped record L200–L203 |
| `test_golden.eco` | Expected ECO output |
| `make_fixtures.py` | Regenerate the xlsx fixtures if needed |

## Run

```bash
# Generate from test ECN+BOM
python scripts/ecn_to_pads_eco.py tests/test_ecn.xlsx \
  --bom tests/test_bom.xlsx \
  --lib library/Electronic_components_library.xlsx \
  -o tests/test_output.eco

# Verify golden
python scripts/verify_eco.py \
  --eco tests/test_golden.eco \
  --bom tests/test_bom.xlsx \
  --lib library/Electronic_components_library.xlsx

# Verify generated output matches golden (data lines)
python scripts/verify_eco.py \
  --eco tests/test_output.eco \
  --bom tests/test_bom.xlsx \
  --lib library/Electronic_components_library.xlsx
```

Expected: `modify 2 modify-via-NC 2 delete 1 add 4` (row counts), both verify runs print `PASS`.
L301 is a 物料新增 onto its own Part Type: `L_RF_12N_… → L_RF_NC_01005 → L_RF_12N_…`.
C101 is the Value `@` guard case: a 物料修改 that must come out as two rows via `C_M_NC_0201`.
L100 is the delete case: `*SET_ATTRIBUTE*` PART L100 `"Value"` `@1.0nH` (BOM Value prefixed with `@`, Part Type unchanged).
L202 sits on a wrapped BOM record (refdes list split mid-token across two rows, type on the second row).
