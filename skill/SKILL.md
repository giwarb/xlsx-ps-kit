---
name: xlsx-ps
description: "Use this skill any time a spreadsheet file is the primary input or output and the environment has PowerShell + Excel desktop but no Python. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .xltx, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path — even casually — and wants something done to it or produced from it. The deliverable must be a spreadsheet file. Do NOT trigger when the primary deliverable is a Word document, HTML report, standalone script, or Graph API / SharePoint integration, even if tabular data is involved."
---

<!-- DRAFT — T-15 で確定。【要充填】は実装メモ・レビュー結果から埋める -->

# XLSX creation, editing, and analysis (PowerShell + Excel COM)

| Task | Approach |
|---|---|
| **Create** or **edit** with formulas/formatting | Excel COM inside `Invoke-XlsSession { param($app,$wb) ... }` — see gotchas below |
| **Bulk data** in or out | `Get-XlsRange` / `Set-XlsRange` (2-D array ⇄ CSV/JSON). Never loop over cells |
| **Quick look** at a sheet | `Get-XlsOverview file.xlsx` — `## SheetName` per sheet; reads `.xlsm` too. No cell coordinates, so don't plan edits from it |
| **Read** a model (formulas *and* values) | `Get-XlsModel file.xlsx -FormulasOnly` — one pass gives both |

> `Import-Module scripts/XlsAgent.psm1` first. Nothing to install. Every COM call goes inside `Invoke-XlsSession`; never `New-Object -ComObject` yourself and never `GetActiveObject`.

> Script paths below are relative to this skill's directory.

## Requirements for every output

- **Professional font** (Arial, Times New Roman) throughout, unless the user says otherwise.
- **Zero formula errors.** Never ship while `Test-XlsFormulas.ps1` reports `errors_found`. If you think an error predates you, prove it: run `Get-XlsModel` on the *original* and look at that cell. An error you introduced looks exactly like one you inherited.
- **Use formulas, never hardcoded results.** Write `$ws.Range("B10").Formula = "=SUM(B2:B9)"`, not the PowerShell-computed total. The sheet must recalculate when its inputs change.
- **Follow the user's spec literally.** Exact tab names, exact column headers, and the formula they spelled out. A redesign that computes something else fails, however elegant.
- **Document every assumption and hardcoded number** where the reader will see it — a cell comment, or an adjacent cell at a table's end. Cite a real source when one exists; when the number came from the user, say so plainly.
- **A workbook *you create* for someone to fill in** needs a short legend naming which cells to edit, and one example row of realistic values showing the expected format. Never add such a row to a file you were asked to edit.
- **Editing an existing file: match its conventions exactly.** They override every guideline here. Find its designated input cells first — a distinct font color, fill, or shading marks them — write only there, and leave every existing formula untouched.

## Recalculate (mandatory whenever the file contains formulas)

`Invoke-XlsSession` opens Excel in **manual calculation** for speed. Formulas you write show stale or empty `Value2` until something recalculates. `Save-XlsWorkbook` recalculates and restores automatic mode before saving, but you must still verify:

```powershell
pwsh -File scripts/Test-XlsFormulas.ps1 output.xlsx [timeout_seconds]   # default 30
```

Excel computes every formula and you get JSON: `status` (`success` | `errors_found` | `refused`), `total_formulas`, `total_errors`, and an `error_summary` naming up to 100 cells per error type (`locations_truncated` says how many it withheld — trust `total_errors`, not the length of the list). Fix what it names and run it again. **JSON with an `error` key instead of a `status` means nothing was recalculated**, and only that case exits non-zero — `errors_found` exits 0, so never treat a clean exit as a clean workbook.

**A green result proves your formulas *evaluate*, not that they are *right*.** An off-by-one range or a reference to the wrong row yields a clean, error-free file with wrong numbers. Write 2–3 formulas first and check they pull the values you expect (`Get-XlsModel -Range`), before building out a grid.

**A workbook that links to another file** whose source is missing gets `status: refused` — recalculating would turn those cells into `#REF!` and Excel would drop the links. Copy the linked cells' values out first (`Get-XlsRange` on the original), or pass `-Force` and accept the loss.

## Choosing formulas

Excel itself evaluates, so there is no compatibility list to worry about. Two rules only:

- Write `.Formula` in **English function names with US separators** (`=SUMIFS(C:C,A:A,"East")`). Never `.FormulaLocal`.
- Dynamic-array functions (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`, `LET`, `LAMBDA`) go through `.Formula2`, not `.Formula` — the latter inserts an implicit-intersection `@` and breaks the spill.

## COM gotchas

【要充填 — T-03〜T-12 の実装メモとレビュー結果から。以下は候補】

- **Use `.Value2`.** `.Value` returns dates as `DateTime` and rounds currency; `.Text` is the displayed string.
- **1-based.** `Cells(1,1)` is A1. PowerShell arrays are 0-based — off-by-one lives here.
- **Merged cells: write the top-left anchor only** (`$r.MergeArea.Cells(1,1)`).
- **Bulk writes take `[object[,]]`.** A jagged `@(@(...),@(...))` fills one row. Use `Set-XlsRange`.
- **A sheet name containing a space must be quoted** in a formula: `='Assumptions Inputs'!$B$5`.
- **Protected sheets:** `Unprotect` before writing, restore on exit.
- **`.xlsm` keeps its macros only through `Save-XlsWorkbook`** (it picks `FileFormat` 52).
- **Close the file in Excel first.** `Invoke-XlsSession` throws if the workbook opens read-only because another process holds it.
- **`SpecialCells` throws when nothing matches** — wrap in `try`.

## Financial models

Unless the user says otherwise, or the existing file already does something else.

**Color:** blue text (`0,0,255`) for hardcoded inputs and scenario levers · black for formulas ·
green (`0,128,0`) for links to another sheet · red (`255,0,0`) for links to another file ·
yellow fill (`255,255,0`) for key assumptions and cells the user should fill in.

**Numbers:** currency `$#,##0`, with the unit named in the header (`Revenue ($mm)`) · zeros
render as `-`, including in percentages (`$#,##0;($#,##0);-`) · negatives in parentheses ·
percentages `0.0%`, **stored as fractions** (`0.15` renders `15.0%`; storing `15` renders
`1500.0%`) · valuation multiples `0.0x` · years as text (`"2024"`, never `2,024`).

**Structure:** every assumption in its own labeled cell, referenced by the formulas that use it
(`=B5*(1+$B$6)`, never `=B5*1.05`) · formulas consistent across every projection period, since a
lone edited cell mid-row is the commonest silent error · guard denominators that can be zero.

## Dependencies

Excel desktop (Microsoft 365 Apps) · PowerShell 5.1 or later · nothing else. See `reference/com-constants.md` for constants and `reference/patterns.md` for tested snippets (headers, tables, comments, conditional formatting, validation, charts).
