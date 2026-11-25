---
name: pads-logic-eco
description: Generate a PADS Logic `*CHGPART*` ECO from an ECN changelist (`Bom&ECN-changelist`) plus a schematic BOM and the component library. Use when the user asks for a PADS Logic ECO or CHGPART, mentions 物料变更 / 物料修改 / 物料删除 / 物料新增 on a schematic, uploads an ECN together with a 原理图BOM, or pastes a PADS Logic ECO import error.
metadata:
  type: workflow
  version: "3.5"
  product: PADS Logic
  library: library/Electronic_components_library.xlsx
---

# PADS Logic ECO

物料修改 and 物料新增 are `*CHGPART*` rows: `<RefDes> <current Part Type from the schematic BOM> <library full name>`. 物料删除 is not a type change: it writes `*SET_ATTRIBUTE*` TAB `PART` TAB `<RefDes>` then `"Value"` TAB `@<existing Value>` (ECOGEN / Logic syntax; `*SET*` is not a section and Logic treats it as still inside `*CHGPART*`). Addition is a **two-hop** through NC. Nothing ships until the **closed loop** (Step 5) passes.

Chinese strings in this file are literals from the Excel sheets and PADS logs; match them verbatim.

## Inputs

| # | File | Role | How to find it |
|---|---|---|---|
| 1 | ECN | RefDes, action, After PN / package | filename contains `Bom&ECN-changelist`; sheet `RF_ECN_Generation` (or any sheet whose `Change Comments` column holds the three actions) |
| 2 | Schematic BOM | **sole source of current Part Type** | uploaded separately by the user |
| 3 | Component library | 物料编码 → 元件名称 | user-uploaded `*component*library*.xlsx` this turn, else bundled `library/Electronic_components_library.xlsx` |

**Schematic BOM columns**

| Meaning | Column names seen |
|---|---|
| RefDes | `Ref.Des` `RefDes` `位号` `Designator` — may hold a comma-separated list; split it, each RefDes gets the same type |
| Current Part Type | `PART NAME` `Part Type` `PartType` `PART-TYPE` `元件名称` `元件类型` — take the full name before the first comma |
| Value | `Value` `Vaule` `值` — source for 物料删除 (`@` prefix) and the Value `@` guard |

**Wrapped records** — when a RefDes list is very long the export splits one record over two rows: row N has PN / description / a truncated list and empty type cells; row N+1 starts with the rest of the list (often mid-token, `7,L2518,…`) and the package / value / type shifted left. Join the two RefDes texts with no separator, then split; the type comes from row N+1. `bomscan.py` does this; in Manual mode do it by hand before building the BOM map.

**ECN columns** (header may span one or two merged rows; locate by name)

| Meaning | Column names seen |
|---|---|
| RefDes | `Ref.Des` `RefDes` `位号` — split on comma |
| Before package | `Before Package` |
| Before PN | `Before Material Code` — cross-check only |
| After PN | `After Material Code` — strip Excel `.0` |
| After package | `After Package` |
| After short type | `Part Type` `After Part Type` — used to disambiguate library names, never written to the ECO |
| Action | `Change Comments` ∈ {物料修改, 物料删除, 物料新增} |

## Step 0 — Route

| Environment | Mode |
|---|---|
| Can run Python | **Script** — `ecn_to_pads_eco.py` composes, `verify_eco.py` checks |
| Chat-only, no shell | **Manual** — compose rows by hand, self-check against the Step 5 table |

Done when: the mode is stated in one sentence.

## Step 1 — Confirm inputs

Done when: all three files are identified by name and the library self-check has passed (both columns present, at least one row with both a name and a Part Number), **or** the turn ends with a request for the missing item or a stop that names the missing library column. A missing schematic BOM ends the turn: "请提供原理图导出的 BOM" — the ECN workbook is never a substitute for it.

**亿赛通 (E-SafeNet) encrypted xlsx** — the scripts detect the `E-SafeNet` header and read the workbook through Excel COM (`scripts/xlsx_open.py`), so an encrypted ECN / BOM / library needs no manual decryption on a machine with Excel. If the error says Excel is missing, ask the user to decrypt the file first.

**BOM without Part Type** — a BOM is usable only if it carries the current Part Type (cells matching `C_M_… / L_RF_… / R_F_… / R_T_…`). A file with RefDes, PN, description and value but no such column (a purchasing BOM, an EBOM sheet, a 下阶料号 list) cannot drive the ECO. End the turn and say which file was checked and what is missing: "这份 BOM 没有 Part Type 列，请从 PADS Logic 原理图导出带 Part Type 的 BOM" — a PN → library lookup is never used as a substitute, because the schematic may differ from the purchasing data.

**Library self-check — before any ECO row is written.** The library sheet must contain both `元件名称` and `Part Number`. A PN column headed `物料编码` or `Material Code` counts as Part Number. Find the columns by header text in the first rows, never by column position. If either column is missing, or no data row has both a name and a Part Number, stop and do not write an ECO. Say which file was checked, which column is missing, and the headers actually seen: "器件库缺少必需列: …。至少要有两列: 元件名称、Part Number。" Rows with a name and an empty Part Number are listed (they cannot be chosen by After PN; the name stays available for NC lookup) and are not a stop. Script mode prints `库自查:` before `modify N …`.

## Step 2 — Build two maps

| Map | Key → Value | Note |
|---|---|---|
| Library | 物料编码 → 元件名称 | one PN may map to several names; pick the one that contains the ECN short type |
| BOM | RefDes → current Part Type | full name before the first comma; Description / Value columns are never the type |
| BOM Value | RefDes → current Value | 物料删除 prefixes this with `@` |
| BOM `@` set | RefDes whose Value contains `@` | drives the Value `@` guard; also skip 物料删除 (already marked) |

Done when: both maps are non-empty and every ECN RefDes either hits the BOM map or is logged as skip `BOM无此位号`. An empty BOM map means the file has no Part Type column → back to Step 1. If more than a third of the ECN RefDes miss the BOM, say so before delivering: the BOM is probably a partial export or a wrapped record was not merged.

## Step 3 — Compose rows

Script mode: run the generator (see Scripts) and read its `.map.csv`; the rules below are what it implements. Manual mode: apply them row by row.

| Action | What is written | Rows |
|---|---|---|
| 物料修改 | `*CHGPART*` BOM type → After library full name | 1 (2 under the Value `@` guard) |
| 物料删除 | `*SET_ATTRIBUTE*` PART `"Value"` `@<BOM Value>` | 1 |
| 物料新增 | `*CHGPART*` via NC to After full name (two-hop) | 1 or 2 |

**Value `@` guard** — PADS Logic bug: a direct `*CHGPART*` on a part whose Value is `@` replaces every part of that type board-wide, not just the RefDes. So a 物料修改 whose RefDes is in the BOM `@` set is routed through NC exactly like a 物料新增 (two-hop table below) and marked `emit(Value@:经NC两跳)` in the mapping table. Tell the user which rows took this path.

**Two-hop rule for 物料新增 (and `@`-guarded 物料修改)** — a same-family part never jumps straight to After:

| Current type | Rows |
|---|---|
| already the target family's NC | `NC → After` |
| same type as After, not NC | `current → same-family NC`, `NC → current` |
| same family, not NC | `current → same-family NC`, `NC → After` |
| different family | `current → target-family NC`, `NC → After` |

The two rows of a hop are adjacent; row 1 column 3 equals row 2 column 2. Each 物料新增 targets its own After PN; a shared test part is used only when the user says 可行性测试.

**物料删除** — keep the Part Type; prefix the BOM Value with `@` (`1.0nH` → `@1.0nH`). Already starting with `@` → skip `已带@无需删除`. Empty Value → skip `BOM无Value`. Do not `*CHGPART*` to NC and do not use `*DELPART*`.

**NC target** (物料新增 and `@`-guarded 物料修改 only) — family from the BOM current-type prefix (else RefDes initial): `R_F_`/`R` → `R_F_NC_`, `C_M_`/`C` → `C_M_NC_`, `L_RF_`/`L` → `L_RF_NC_`. Package from ECN `After Package` (add) or `Before Package`, else parsed from the current type; normalise `1005→01005 201→0201 402→0402 603→0603 805→0805 1206 1210`. The NC name must exist in the library.

**Skip (log the reason, emit nothing)**: RefDes absent from BOM · After PN not in library · 物料修改 with old type == new type · 物料新增 already on the target NC · 物料删除 with empty Value or Value already starting with `@` · add/`@`-guarded modify with no matching NC in library. A 物料新增 whose current type already equals After is not a skip: it still two-hops `current → NC → current`.

**Scope**: every component in the ECN whose family has library names (chip R / C, RF and power L, tantalum, NTC…). ICs, modules, connectors, crystals only when the user names them. First 6 digits of a PN are a sanity check, not a skip condition.

**Package change popup** — for every emitted 物料修改 / 物料新增, compare ECN `Before Package` vs `After Package` and the package token in old vs new Part Type. Normalize `1005→01005 201→0201 402→0402 603→0603 805→0805`. If they differ, the GUI lists 位号 / 变更类型 / 变更前封装 / 变更后封装. The subtitle is 请先打开 ECN 变更表，核对下列封装变化是否正确。确认无误后再继续生成。 `ECN 变更表` is a link that opens the selected ECN file. Cancel discards the ECO. Continue still emits those `*CHGPART*` rows. CLI prints the same list (with the ECN filename) and continues. 物料删除 is not a package change. The mapping-table Status becomes `emit(封装变化:0201→0402)`.

Done when: every ECN RefDes appears exactly once in either the ECO rows or the skip list with a reason.

## Step 4 — Write the ECO and the mapping table

```
*PADS-ECO-V3.0-MILS*
*REMARK* <board> modify+delAt+addViaNC  current-from-sch-BOM
*CHGPART*
<RefDes> <BomOld> <LibNew>
*SET_ATTRIBUTE*<TAB>PART<TAB><RefDes>
"Value"<TAB>@<BomValue>
*END*
```

Omit `*CHGPART*` when there are no type changes; omit `*SET_ATTRIBUTE*` when there are no deletes. `*CHGPART*` rows: single space between fields. `*SET_ATTRIBUTE*` rows: TAB between `*SET_ATTRIBUTE*` / `PART` / RefDes and between `"Value"` / `@…` (this is the ECOGEN 6.4v / Logic form). No trailing spaces · data rows carry no `*` · CRLF · ASCII. Never write a `*SET*` section — Logic does not close `*CHGPART*` on that token. The mapping table lists every emitted row and every skipped RefDes with its reason.

Done when: one `.eco` and one mapping table exist.

## Step 5 — Closed loop

Script mode: `python scripts/verify_eco.py --eco <out.eco> --bom <BOM.xlsx> --lib <lib.xlsx>`.
Manual mode: walk the table and tick each check in the mapping table.

| Check | Passes when |
|---|---|
| F0–F5 | file exists; header/footer markers; CRLF; `*CHGPART*` single spaces; `*SET_ATTRIBUTE*` TABs; no trailing space; no `*` in data rows |
| D0 | every RefDes is in the BOM |
| D1 | column 2 == BOM current type (except the second row of a two-hop) |
| D2 | column 3 exists in the library |
| D3 | each 物料新增 RefDes has exactly two rows; row 1 col 3 == row 2 col 2 and is an NC |
| D4 | no duplicate rows |
| D5 | each `*SET_ATTRIBUTE*` `"Value"` equals `@` + BOM Value; no `*CHGPART*` to NC for delete |
| D6 | intermediate NC exists in the library |
| D7 | no single-row jump from an NC of one family to a non-NC part of another (same-family `NC → After` is the valid one-hop) |
| D8 | no single-row non-NC → non-NC change on a RefDes in the BOM `@` set (Value `@` guard) |

PASS → deliver `.eco` + mapping table. FAIL → print each failing check, leave no `.eco`, fix only the failing rows, return to Step 3, re-run. The app (`pads_eco_app.py` GUI / `gen`) verifies before the destination file is replaced; it does not rename a failed file to `.FAIL.eco`.

Done when: verify prints `PASS`, or every manual check is ticked.

## Scripts

```
python scripts/ecn_to_pads_eco.py <ECN.xlsx> --bom <BOM.xlsx> --lib <lib.xlsx> -o <out.eco>
    # writes <out.eco> and <out.map.csv>; prints "modify N modify-via-NC N delete N add N" and skip count
python scripts/verify_eco.py --eco <out.eco> --bom <BOM.xlsx> --lib <lib.xlsx>
```

**Encoding** — the `.eco` is pure ASCII and never has an encoding problem. The `.map.csv` is UTF-8 with BOM by default (opens in Excel). If the host reports `csv 编码异常` / mojibake when it reads the map, do not hand-fix the file; regenerate with `--map-encoding gbk` (Chinese-Windows hosts) or `--map-encoding ascii` (Action/Status become English tokens: `MODIFY / DELETE / ADD`, `not_in_BOM`, `same_type`, `already_at`…) and read that one. Read the map yourself with `encoding="utf-8-sig"`; scripts print in UTF-8 regardless of console codepage. A `rows.csv` handed to `bom_to_pads_eco.py` may be UTF-8, UTF-8-BOM or GBK — it is detected. `--remark` must be ASCII.

Fallback when the generator cannot locate the ECN headers: compose a `RefDes,OldPartType,NewPartType` CSV by the Step 3 rules, then
`python scripts/bom_to_pads_eco.py rows.csv <out.eco> --remark "<board> modify+addViaNC"` and verify as above. That fallback is `*CHGPART*` only — 物料删除 still belongs in `*SET_ATTRIBUTE*` PART `"Value"`.

## After delivery

Import in Logic: back up the `.sch`, disconnect Layout, disable online DRC, 文件 → 导入 → ECO Files. Spot-check one delete (Value now starts with `@`, Part Type unchanged) and one add (After full name); cross-family two-hops deserve a symbol-placement check. If a previous import already applied `*CHGPART*` and only the delete section failed, import a delete-only `.eco` (`*SET_ATTRIBUTE*` blocks, no `*CHGPART*`).

- Logic reports an import error → open [references/eco-errors.md](references/eco-errors.md).
- A library name cannot be resolved from the PN → open [references/library-naming.md](references/library-naming.md).
