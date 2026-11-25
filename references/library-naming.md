# Company library Part Type names

Opened from SKILL.md when a 物料编码 cannot be resolved to a 元件名称, or a resolved name is rejected by Logic. Source of record: `library/Electronic_components_library.xlsx` (columns 元件名称 | Part Number). Names are looked up by PN; the grammar below is for sanity-checking a hit, not for constructing names.

## Grammar

```
<class>_<value>_<rating>_<package>_<MPN>
```

| Class | Prefix | NC part | Rating examples | Packages |
|---|---|---|---|---|
| MLCC | `C_M_` | `C_M_NC_<pkg>` | `25V` `16V` `10V` `6V3` `50V` | `01005` `0201` `0402` `0603` |
| RF inductor | `L_RF_` | `L_RF_NC_<pkg>` | `100M` … `350M` (self-resonant, MHz) | `01005` `0201` `0402` |
| Thick-film resistor | `R_F_` | `R_F_NC_<pkg>` | `1%` `5%` | `01005` `0201` `0402` |
| NTC | `R_T_` | — | | |

No spaces anywhere. Resistor rating keeps the `%`.

Examples of real names:

```
C_M_0P5_25V_01005_V0R5B0105C0G250NZT
L_RF_18N_200M_0201_SDCL0603Q18NHT02B03
R_F_10K_5%_01005_RC0100JR-0710KL
```

## ECN short type vs library full name

ECN sheets carry the tail only, e.g. `25V_01005_C0105C0G0R7A250NTZ` or `200M_0201_SDCL0603Q18NHT02B03`. Two things differ from the library:

- The library prepends `<class>_<value>_`.
- Capacitor MPNs in the library use the vendor form `V0R7A0105C0G250NZT`; the ECN form `C0105C0G0R7A250NTZ` is not a Part Type.

So the short type is only a disambiguator when one PN maps to several names — pick the name that contains it (or its value/package tokens). The ECO always gets the full name from the xlsx.

## When one PN maps to several names

Sibling names differ by one vendor-revision letter inside the MPN (`…V0R7A…` / `…V0R7B…`, `…Q56NHT…` / `…Q56NJT…`). Prefer the one whose tokens match the ECN short type; if still tied, ask the user which revision the board uses.

## When a resolved name is rejected by Logic

A name that exists in the xlsx can still fail import if its library record is broken (symbol attached to the wrong body). Treat `找不到元件类型项目` on a name you resolved correctly as a library defect: report it, and offer the sibling-revision name as a candidate for the user to confirm.
