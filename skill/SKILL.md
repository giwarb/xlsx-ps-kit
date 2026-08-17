---
name: xlsx-ps
description: "Use this skill any time a spreadsheet file is the primary input or output and the environment has PowerShell + Excel desktop but no Python. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .xltx, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path — even casually — and wants something done to it or produced from it. The deliverable must be a spreadsheet file. Do NOT trigger when the primary deliverable is a Word document, HTML report, standalone script, or Graph API / SharePoint integration, even if tabular data is involved."
---

# XLSX creation, editing, and analysis (PowerShell + Excel COM)

| Task | Approach |
|---|---|
| **Create** or **edit** with formulas/formatting | Excel COM inside `Invoke-XlsSession -Path file.xlsx -ScriptBlock { param($app,$wb) ... }` — see gotchas below |
| **Bulk data** in or out | `Get-XlsRange` / `Set-XlsRange` (2-D array ⇄ CSV/JSON). Never loop over cells |
| **Quick look** at a sheet | `Get-XlsOverview -Path file.xlsx` — `## SheetName` per sheet; reads `.xlsm` too. No cell coordinates, so don't plan edits from it |
| **Read** a model (formulas *and* values) | `Get-XlsModel -Path file.xlsx -FormulasOnly` — one pass gives both |

> `Import-Module scripts/XlsAgent.psm1` first. Nothing to install. Every COM call goes inside `Invoke-XlsSession`; never `New-Object -ComObject` yourself and never `GetActiveObject`. If a prior run crashed and left an `EXCEL.EXE` behind, run `Clear-XlsOrphans` — it only stops processes this module itself started (tracked by PID + start time), never a live user session. Don't reach for `Get-Process EXCEL | Stop-Process`.

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

No `pwsh` on this machine? `powershell -File scripts/Test-XlsFormulas.ps1 output.xlsx` works the same way (Windows PowerShell 5.1).

Excel computes every formula and you get JSON in one of two distinct shapes:

- **`success` or `errors_found`:** `status`, `total_formulas`, `total_errors`, and an `error_summary` naming up to 100 cells per error type (`locations_truncated` says how many it withheld — trust `total_errors`, not the length of the list). Fix what it names and run it again.
- **`refused`:** only `status: "refused"`, `reason`, and `links` (every external reference the workbook has, resolved or not). It does **not** contain `total_formulas`, `total_errors`, or `error_summary` — nothing was recalculated, so there is nothing to total. See below for why.

**JSON with an `error` key instead of a `status` means nothing was recalculated**, and only that case exits non-zero — `errors_found` (and `refused`) exit 0, so never treat a clean exit as a clean workbook.

**A green result proves your formulas *evaluate*, not that they are *right*.** An off-by-one range or a reference to the wrong row yields a clean, error-free file with wrong numbers. Write 2–3 formulas first and check they pull the values you expect (`Get-XlsModel -Range`), before building out a grid.

**A workbook that links to another file** whose source is missing gets the `refused` shape above — recalculating would turn those cells into `#REF!` and Excel would drop the links. Copy the linked cells' values out first (`Get-XlsRange` on the original), or pass `-Force` to skip the check and recalculate anyway; the linked cells become `#REF!` and show up as `errors_found` (the normal shape, not `refused`), which is the expected outcome for a genuinely broken link, not a silent loss.

## Choosing formulas

Excel itself evaluates, so there is no compatibility list to worry about. Two rules only:

- Write `.Formula` in **English function names with US separators** (`=SUMIFS(C:C,A:A,"East")`). Never `.FormulaLocal`.
- Dynamic-array functions (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`, `LET`, `LAMBDA`) go through `.Formula2`, not `.Formula` — the latter inserts an implicit-intersection `@` and breaks the spill.

## COM gotchas

These are traps you can hit *inside* the `Invoke-XlsSession` script block. Things the session
wrapper already handles for you — forcing a new process, detecting a workbook opened read-only
elsewhere, STA, and cleanup — aren't listed; you can't get them wrong from the outside.

- **Use `.Value2`, not `.Value`/`.Text`, and mind the 1-based/scalar quirks.** In PowerShell, bare `$cell.Value` returns a `PSParameterizedProperty` descriptor, not the contents (call `.Value()`) — and even called right, it converts dates to `System.DateTime` instead of the raw `[double]` serial `.Value2` gives you. `.Text` is a display string (rounds, and shows `#` if the column's too narrow). Indexing is 1-based, but it's asymmetric: a multi-cell `Value2` you *read* comes back 1-based (`GetLowerBound(0) == 1`); the `[object[,]]` you *write* must be 0-based. And a *single-cell* range returns `Value2`/`Formula`/`NumberFormat` as a bare scalar, not a 1x1 array. `Get-XlsRange`/`Get-XlsModel` absorb all of this for their own use — these only bite when you read/write `Value2` yourself.
- **`NumberFormat`, `Text`, `Font.Color`, and `HasFormula` collapse to one shared value across a multi-cell range only if every cell agrees** — otherwise you get `[DBNull]::Value`, not `$null` and not an array. `Workbook.LinkSources(xlExcelLinks)` is the odd one out: it never collapses, even with exactly one link — always a `string[]` or `$null`.
- **Merged cells: write the top-left anchor only** (`$r.MergeArea.Cells(1,1)`). Writing directly to any other cell in the merge is a silent no-op — no exception, but the value never lands anywhere.
- **Bulk writes take `[object[,]]`, cast right before you store each value (`[double]`, etc.)** — an untyped value that crosses a function/`ScriptBlock` boundary before landing in `Range.Value2 =` can corrupt the write (`COMException 0x800A03EC`). Building a 2-D array by hand has its own parser and flattening traps; see `reference/patterns.md`'s misc-gotchas section. Prefer `Set-XlsRange`, which handles all of this.
- **Opening/saving workbooks has a few sharp edges.** `Workbooks.Open`'s arguments are positional only — PowerShell's late-bound COM calls don't understand VBA-style named arguments (`UpdateLinks:=0`). Opening a `.xltx`/`.xltm` template creates a new, unsaved workbook based on it, not an edit of the template file itself. And `.xlsm` keeps its macros only if you save through `Save-XlsWorkbook` (it picks `FileFormat` 52).
- **`NumberFormat` for the default format is locale-dependent** — on a Japanese-locale Excel it reads `"G/標準"`, not `"General"`. Test for date/time tokens instead of string-matching the default format's name.
- **Dates in `Value2` are OLE Automation serials, not `DateTime`.** Write `[DateTime]::ToOADate()`; a serial you read is just a number unless you also check `NumberFormat`. `Get-XlsRange`/`Set-XlsRange` do this conversion (1900/1904 leap-year quirks included) for you.
- **A sheet name containing a space or apostrophe must be quoted in a formula — and getting it wrong doesn't throw.** `.Formula = '=My Data!A1'` (unquoted, sheet name has a space) is accepted silently and evaluates to `#NAME?`; only reading `Value2` back or running `Test-XlsFormulas` catches it. Quote it (`='My Data'!A1`) and double any apostrophe inside the name. (`Names.Add`'s `RefersTo` has its own quoting-normalization quirk — see `reference/patterns.md`'s Named ranges section.)
- **Protected sheets:** writing throws immediately (a localized "the cell is on a protected sheet" error). `Unprotect` before writing, restore on exit.
- **`SpecialCells` throws (HRESULT `0x800A03EC`) when nothing matches.** Catch that specific error as "zero results" — don't swallow every exception, or a real COM failure looks like an empty range.
- **A cell's `.Formula` isn't empty just because the cell isn't a formula** — a non-formula cell returns its value's string form instead. Don't test with a leading `=`; a cell forced to text (`'=5`) has `Formula = "=5"` and `HasFormula = $false`. Use `SpecialCells(xlCellTypeFormulas)`. (Reading `.Formula` from a *multi-cell* range also repeats the text once per cell — read a single cell if you want one string.)
- **Two things belong in `reference/com-constants.md`, not here:** the raw `Int32` codes `Value2` marshals error cells to (not VBA's `XlCVError` enum — don't feed one to a COM API expecting that), and the BGR-long formula for `Font.Color`/`Interior.Color` (there's no COM `RGB()` helper).
- **A function returning a `HashSet`, array, or other enumerable must `return , $value`** (unary comma), or PowerShell's output stream unwraps/flattens it — a zero- or one-element collection silently turns into `$null` or a bare scalar. Just always use it; it's harmless even on the rare type (`Hashtable`) that doesn't actually need it.
- **CSV does not preserve the original cell type metadata.** `Set-XlsRange -FromCsv` infers empty→`$null`, exact `True`/`False`→bool, invariant-culture numbers→`[double]`, exact ISO dates (the format `Get-XlsRange` emits)→date — everything else stays a string, so a numeric-looking string (`"42"`) can silently round-trip as a number. `-Header` has no effect with `-FromCsv`. JSON keeps real types, but integers arrive as `[long]` — cast to `[double]` before a raw `Value2` write.
- **Give recalculation and saves room to finish.** `Test-XlsFormulas.ps1`'s `timeoutSec` doesn't start counting until Excel's COM activation finishes — that alone has taken 4–61s on one dev machine, so don't set it near the low end. Separately, a burst of fast writes immediately followed by `Save-XlsWorkbook` can — rarely — drop the last write or two; read it back before trusting the file.
- **COM methods with an optional argument you want to skip take `[Type]::Missing`, not `$null` or `0`** (e.g. `ListObjects.Add(1, $range, [Type]::Missing, 1, [Type]::Missing, [Type]::Missing)`).

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

Excel desktop (Microsoft 365 Apps) · PowerShell 5.1 or later · nothing else. See `reference/com-constants.md` for constants and `reference/patterns.md` for tested snippets (header formatting, tables, column auto-fit, comments, conditional formatting, data validation, charts, named ranges, sheet add/copy/delete) plus a misc-gotchas section for lower-frequency PowerShell/COM traps that didn't make the cut above.
