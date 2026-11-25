# PADS Logic ECO import errors

Opened from SKILL.md when Logic rejects an ECO. The log appears in the import dialog and is usually also written next to the `.eco`. Logic applies none of the `*CHGPART*` rows it lists; rows absent from the log succeeded — re-emit only the listed rows.

## `行 N:错误: 元件 <RefDes> 的类型为 <A>，而不是 <B>`

Column 2 (`B`) does not match the schematic. `A` is the true current Part Type.

Fix: set column 2 to `A` for that RefDes, regenerate only the failed rows, re-run verify. If several rows fail this way, the schematic BOM the user supplied is stale — say so and ask for a fresh export.

## `行 N:错误: 找不到元件类型项目 < NAME >`

`NAME` is not a Part Type under that exact string in any library on Logic's search path. Spaces around `NAME` inside `< >` are log decoration.

Causes, most common first:
1. ECN short type written instead of the library full name → rebuild from 物料编码 → 元件名称.
2. A `*` copied from the Logic search box into the ECO → strip it.
3. Part not yet created in the company library → ask the user; if confirmed absent, drop the row and deliver the rest.
4. Company library not on Logic's search path → put it first (文件 → 库 → 管理库列表).

To confirm a name in Logic: 文件 → 库 → 元件, search `*keyword*`, copy the name from the **list**, not from the ECN.

## `行 N:错误: 读取不到旧/新的元件类型名称`

Logic is still parsing `*CHGPART*`. This happens when the file used a fake section `*SET*` / `Value` / `<RefDes> @…` — that token is not a section closer. Use the ECOGEN form:

```
*SET_ATTRIBUTE*<TAB>PART<TAB><RefDes>
"Value"<TAB>@<BomValue>
```

If the log starts at the first `*SET*` line and no `*CHGPART*` rows are listed, those CHGPART rows already applied. Re-emit **only** the delete `*SET_ATTRIBUTE*` blocks.

## `发现输入错误`

Summary line after one or more row failures. Fix every listed row before re-importing.

## After a clean import

Run ECO to PCB so Layout Part Types follow the schematic.
