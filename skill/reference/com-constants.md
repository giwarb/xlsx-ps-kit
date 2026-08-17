# COM constants (Excel)

<!-- T-13 で実測充填。「確認」列は実行して確認した日付を入れる。未確認の行を残さない（G-12） -->

`XlsAgent.psm1` の `$script:Xl` と同期させる。

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlCellTypeFormulas | -4123 | SpecialCells | 2026-08-17（T-08 Get-XlsModel、実機。`Range.SpecialCells(-4123)` で数式セルのみの Range を取得、ヒットなしは HRESULT 0x800A03EC 例外になることを確認） |
| xlErrors | 16 | SpecialCells 第 2 引数 | |
| xlCalculationManual | -4135 | Application.Calculation | 2026-08-17（T-03 Invoke-XlsSession、実機） |
| xlCalculationAutomatic | -4105 | Application.Calculation | 2026-08-17（T-04 Save-XlsWorkbook、実機。Automatic のまま保存すると xl/workbook.xml の calcPr から calcMode 属性が省略されることまで確認） |
| xlOpenXMLWorkbook | 51 | SaveAs .xlsx | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認） |
| xlOpenXMLWorkbookMacroEnabled | 52 | SaveAs .xlsm | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後・再オープン後とも FileFormat=52 を確認） |
| xlOpenXMLTemplateMacroEnabled | 53 | SaveAs .xltm | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認） |
| xlOpenXMLTemplate | 54 | SaveAs .xltx | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認。Workbooks.Open で開き直すとテンプレートから新規ブックが生成されるため、reopen では確認できない罠あり） |
| xlCSV | 6 | SaveAs .csv | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat とファイル内容で確認） |
| xlExcelLinks | 1 | LinkSources | 2026-08-17（T-08 Get-XlsModel、実機。`Workbook.LinkSources(1)` が未解決の外部リンク先パスを 1-based string[] で返す（リンクなしは $null）ことを確認） |
| xlSrcRange | 1 | ListObjects.Add | |
| xlYes | 1 | ListObjects.Add HasHeaders | |
| xlCenter | -4108 | HorizontalAlignment | |
| xlLeft | -4131 | HorizontalAlignment | |
| xlRight | -4152 | HorizontalAlignment | |
| xlContinuous | 1 | Borders.LineStyle | |
| xlThin | 2 | Borders.Weight | |
| xlEdgeBottom | 9 | Borders(index) | |
| xlInsideHorizontal | 12 | Borders(index) | |
| xlValidateList | 3 | Validation.Add | |
| xlCellValue | 1 | FormatConditions.Add | |
| xlGreater | 5 | FormatConditions.Add Operator | |
| xlColumnClustered | 51 | Chart.ChartType | |
| xlLine | 4 | Chart.ChartType | |
| xlShiftDown | -4121 | Range.Insert | |
| xlToRight | -4161 | Range.Insert | |
| xlErrNull | -2146826288 | Range.Value2（エラーセルの生値、#NULL!） | 2026-08-17（T-06 Get-XlsRange、実機。`=SUM(1:1 2:2)` で確認） |
| xlErrDiv0 | -2146826281 | Range.Value2（エラーセルの生値、#DIV/0!） | 2026-08-17（T-06 Get-XlsRange、実機。`=1/0` で確認） |
| xlErrValue | -2146826273 | Range.Value2（エラーセルの生値、#VALUE!） | 2026-08-17（T-06 Get-XlsRange、実機。`=1/""` で確認） |
| xlErrRef | -2146826265 | Range.Value2（エラーセルの生値、#REF!） | 2026-08-17（T-06 Get-XlsRange、実機。`=#REF!` で確認） |
| xlErrName | -2146826259 | Range.Value2（エラーセルの生値、#NAME?） | 2026-08-17（T-06 Get-XlsRange、実機。`=NoSuchName` で確認） |
| xlErrNum | -2146826252 | Range.Value2（エラーセルの生値、#NUM!） | 2026-08-17（T-06 Get-XlsRange、実機。`=SQRT(-1)` で確認） |
| xlErrNA | -2146826246 | Range.Value2（エラーセルの生値、#N/A） | 2026-08-17（T-06 Get-XlsRange、実機。`=NA()` で確認） |
