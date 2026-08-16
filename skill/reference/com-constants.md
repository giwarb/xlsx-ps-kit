# COM constants (Excel)

<!-- T-13 で実測充填。「確認」列は実行して確認した日付を入れる。未確認の行を残さない（G-12） -->

`XlsAgent.psm1` の `$script:Xl` と同期させる。

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlCellTypeFormulas | -4123 | SpecialCells | |
| xlErrors | 16 | SpecialCells 第 2 引数 | |
| xlCalculationManual | -4135 | Application.Calculation | |
| xlCalculationAutomatic | -4105 | Application.Calculation | |
| xlOpenXMLWorkbook | 51 | SaveAs .xlsx | |
| xlOpenXMLWorkbookMacroEnabled | 52 | SaveAs .xlsm | |
| xlOpenXMLTemplateMacroEnabled | 53 | SaveAs .xltm | |
| xlOpenXMLTemplate | 54 | SaveAs .xltx | |
| xlCSV | 6 | SaveAs .csv | |
| xlExcelLinks | 1 | LinkSources | |
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
| xlErrValue / xlErrRef 等 | 【要確認】 | CVErr 判定 | |
