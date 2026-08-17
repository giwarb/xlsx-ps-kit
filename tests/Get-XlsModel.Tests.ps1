#Requires -Version 5.1

<#
tests/Get-XlsModel.Tests.ps1

T-08: Get-XlsModel（Formula/Value2 を 1 回の読み取りでまとめて返すモデルダンプ。01-design.md §3.4）と、
内部関数 ConvertTo-XlsColumnLetter / Get-XlsFormulaCellKeySet / ConvertTo-XlsModelCellList の受け入れテスト。
harness/state/tasks/T-08.md の受け入れテスト要点:
  - 数式セルで formula が '=' 始まり、value2 が計算結果。
  - 名前定義・外部リンクが取れる（リンクは存在しないパスへのリンクを持つ fixture を
    Invoke-XlsSession 内で作成し、保存は Save-XlsWorkbook 経由）。
  - tables（ListObjects）、protected が取れる。
  - -Range / -FormulasOnly の絞り込みが効く。
  - -AsJson が recalc 可能な JSON を返す（パースして構造確認）。
  - 異常系: 存在しないシート名で次の一手が分かる例外。

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中
（Get-XlsModel 自体を含む）でのみ触る。フィクスチャ作成は Common.ps1 の New-XlsObjectArray2D /
New-XlsSingleCellArray で作った [object[,]]（G-08）を ScriptBlock 内で Range.Value2 に代入する形と、
Range.Formula への文字列代入（COM プロパティ代入であり Value2 ではないため G-08 の対象外）で行う。

純粋関数（ConvertTo-XlsColumnLetter）は COM を一切使わないため、他の *.Tests.ps1 と同じ
`& $script:ModuleRef { ... }` パターンでモジュールスコープに入って直接テストする。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

function Get-XlsModelCellByAddress {
    param($Cells, [string]$Address)
    return @($Cells | Where-Object { $_.address -eq $Address }) | Select-Object -First 1
}

Describe 'Get-XlsModel (T-08)' {

    BeforeEach {
        Import-Module $modulePath -Force
        $script:ModuleRef = Get-Module $moduleName
        $script:BaselinePids = Get-ExcelPids
    }

    AfterEach {
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }

    AfterAll {
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }

    Context 'ConvertTo-XlsColumnLetter (pure function, no COM)' {

        It 'converts 1-based column numbers to Excel column letters' {
            $cases = @{
                1   = 'A'
                2   = 'B'
                26  = 'Z'
                27  = 'AA'
                28  = 'AB'
                52  = 'AZ'
                53  = 'BA'
                702 = 'ZZ'
                703 = 'AAA'
            }
            foreach ($col in $cases.Keys) {
                $result = & $script:ModuleRef { param($c) ConvertTo-XlsColumnLetter -Column $c } $col
                $result | Should Be $cases[$col]
            }
        }

        It 'throws a next-step error for a column number less than 1' {
            $errorSeen = $null
            try {
                & $script:ModuleRef { ConvertTo-XlsColumnLetter -Column 0 }
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Column must be 1 or greater'
        }
    }

    Context 'Get-XlsModel integration (real Excel via Invoke-XlsSession)' {

        It 'reports formula cells with a leading-= formula and the calculated value2, and excludes empty cells' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 2.0)
                $ws.Range('B1').Formula = '=A1*3'
                # C1 は空セル（値も数式もない）のまま残す。
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path

            $model.sheets.Count | Should Be 1
            $sheet1 = $model.sheets[0]
            $sheet1.name | Should Be 'Sheet1'

            $a1 = Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'A1'
            $a1 | Should Not Be $null
            $a1.formula | Should Be $null
            $a1.value2 | Should Be 2.0

            $b1 = Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'B1'
            $b1 | Should Not Be $null
            ($b1.formula.StartsWith('=')) | Should Be $true
            $b1.formula | Should Be '=A1*3'
            $b1.value2 | Should Be 6.0

            $c1 = Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'C1'
            $c1 | Should Be $null
        }

        It 'does not misclassify an apostrophe-forced text cell that looks like a formula as a formula cell' {
            # 罠（実装メモ参照）: 先頭アポストロフィで強制テキスト化したセルは Value2/Formula とも
            # "=5" を返すが HasFormula は False。'=' 始まり判定ではなく SpecialCells で判定すること。
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value "'=5")
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $a1 = Get-XlsModelCellByAddress -Cells $model.sheets[0].cells -Address 'A1'

            $a1 | Should Not Be $null
            $a1.formula | Should Be $null
            $a1.value2 | Should Be '=5'
        }

        It 'converts a date-formatted numeric cell to an ISO 8601 string via the reused ConvertFrom-XlsValue2 rules' {
            $path = New-TempXlsxPath
            $expectedDate = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 45000.0)
                $ws.Range('A1').NumberFormat = 'm/d/yyyy'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $a1 = Get-XlsModelCellByAddress -Cells $model.sheets[0].cells -Address 'A1'

            $a1.value2 | Should Be $expectedDate
            $a1.numberFormat | Should Be 'm/d/yyyy'
        }

        It 'converts a formula error result to its display string (e.g. #DIV/0!)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = '=1/0'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $a1 = Get-XlsModelCellByAddress -Cells $model.sheets[0].cells -Address 'A1'

            $a1.formula | Should Be '=1/0'
            $a1.value2 | Should Be '#DIV/0!'
        }

        It 'reads workbook-level defined names (name, refersTo) and excludes sheet-scoped names' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                $wb.Names.Add('WorkbookLevelName', '=Sheet1!$A$1')
                $ws.Names.Add('SheetScopedName', '=Sheet1!$A$1')
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path

            $wbName = @($model.workbook.names | Where-Object { $_.name -eq 'WorkbookLevelName' })
            $wbName.Count | Should Be 1
            $wbName[0].refersTo | Should Be '=Sheet1!$A$1'

            $sheetScoped = @($model.workbook.names | Where-Object { $_.name -match 'SheetScopedName' })
            $sheetScoped.Count | Should Be 0
        }

        It 'reads unresolved external link sources recorded on the workbook' {
            # タスクカード指定: 存在しないパスへのリンクを持つ fixture を Invoke-XlsSession 内で作成し、
            # 保存は Save-XlsWorkbook 経由で行う。
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = "='C:\NoSuchFolderXYZ_T08\[Other.xlsx]Sheet1'!A1"
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path

            $model.workbook.links.Count | Should BeGreaterThan 0
            ($model.workbook.links -join ';') | Should Match ([regex]::Escape('NoSuchFolderXYZ_T08'))
        }

        It 'reports an empty links array when the workbook has no external links' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            , $model.workbook.links | Should Not Be $null
            $model.workbook.links.Count | Should Be 0
        }

        It 'reads ListObjects (tables) name and range, and sheet protection state' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('H1', 'H2'),
                    @(1.0, 2.0),
                    @(3.0, 4.0)
                ))
                $lo = $ws.ListObjects.Add(1, $ws.Range('A1:B3'), $null, 1)
                $lo.Name = 'MyTable'
                $ws.Protect()
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $sheet1 = $model.sheets[0]

            $sheet1.tables.Count | Should Be 1
            $sheet1.tables[0].name | Should Be 'MyTable'
            $sheet1.tables[0].range | Should Be 'A1:B3'
            $sheet1.protected | Should Be $true
            $sheet1.usedRange | Should Be 'A1:B3'
        }

        It 'reports protected = $false for an unprotected sheet' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $model.sheets[0].protected | Should Be $false
        }

        It '-Range restricts cells to the given address without changing the reported usedRange' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:D2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(1.0, 2.0, 3.0, 4.0),
                    @(5.0, 6.0, 7.0, 8.0)
                ))
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path -Range 'A1:B1'
            $sheet1 = $model.sheets[0]

            $sheet1.cells.Count | Should Be 2
            (Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'A1') | Should Not Be $null
            (Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'B1') | Should Not Be $null
            (Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'C1') | Should Be $null
            # usedRange はシート全体を報告する（-Range の絞り込みの影響を受けない、実装判断）。
            $sheet1.usedRange | Should Be 'A1:D2'
        }

        It '-FormulasOnly limits cells to formula cells only, leaving sheet metadata untouched' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 10.0)
                $ws.Range('B1').Formula = '=A1*2'
                $ws.Range('C1').Value2 = (New-XlsSingleCellArray -Value 'plain text')
                $ws.Range('D1').Formula = '=A1+1'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path -FormulasOnly
            $sheet1 = $model.sheets[0]

            $sheet1.cells.Count | Should Be 2
            foreach ($cell in $sheet1.cells) {
                ($cell.formula.StartsWith('=')) | Should Be $true
            }
            (Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'A1') | Should Be $null
            (Get-XlsModelCellByAddress -Cells $sheet1.cells -Address 'C1') | Should Be $null
            $sheet1.usedRange | Should Be 'A1:D1'
        }

        It '-Sheet filters the sheets array and preserves the requested order, while workbook.sheets always lists all sheets' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws1 = $wb.Worksheets.Item(1)
                $ws1.Name = 'First'
                $ws1.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                $ws2 = $wb.Worksheets.Add()
                $ws2.Name = 'Second'
                $ws2.Range('A1').Value2 = (New-XlsSingleCellArray -Value 2.0)
                $ws3 = $wb.Worksheets.Add()
                $ws3.Name = 'Third'
                $ws3.Range('A1').Value2 = (New-XlsSingleCellArray -Value 3.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path -Sheet 'Third', 'First'

            $model.sheets.Count | Should Be 2
            $model.sheets[0].name | Should Be 'Third'
            $model.sheets[1].name | Should Be 'First'

            @($model.workbook.sheets).Count | Should Be 3
            ($model.workbook.sheets -contains 'Second') | Should Be $true
        }

        It 'throws a next-step error listing available sheets for a nonexistent sheet name' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $errorSeen = $null
            try {
                Get-XlsModel -Path $path -Sheet 'NoSuchSheet'
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'not found'
            $errorSeen.Exception.Message | Should Match 'Available sheets'
        }

        It 'throws a next-step error for an invalid -Range address' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $errorSeen = $null
            try {
                Get-XlsModel -Path $path -Range 'NotAValidAddr$$'
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Invalid range address'
        }

        It 'throws a next-step error and leaves no orphan Excel process when the file does not exist' {
            # レビュー指摘対応（T-08 round 1 should-fix）: 公開関数 Get-XlsModel 経由でのファイル不在
            # （-ReadOnly で Excel 起動後に throw、後始末される経路）を固定する。AfterEach の
            # Assert-NoOrphanExcel がプロセス残留も検証する。
            $path = New-TempXlsxPath
            $errorSeen = $null

            try {
                Get-XlsModel -Path $path
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'File not found'
        }

        It 'reports version and build as non-empty strings' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $model.workbook.path | Should Be $path
            [string]::IsNullOrEmpty($model.workbook.version) | Should Be $false
            [string]::IsNullOrEmpty($model.workbook.build) | Should Be $false
        }

        It '-AsJson returns a JSON string that parses to the same structure, even with single-element arrays' {
            # 罠（実装メモ参照）: 要素数 1 の配列が ConvertTo-Json / パイプラインでアンラップされないことを
            # 1 シート・1 セル・1 テーブル・1 名前という縮退ケースで固定する回帰テスト。
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 7.0)
                $wb.Names.Add('OnlyName', '=Sheet1!$A$1')
                $lo = $ws.ListObjects.Add(1, $ws.Range('A1:A1'), $null, 1)
                $lo.Name = 'OnlyTable'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $json = Get-XlsModel -Path $path -AsJson
            $json | Should BeOfType ([string])

            $parsed = ConvertFrom-Json $json

            # レビュー指摘対応（T-08 round 1 should-fix）: `@(...).Count` は JSON が配列ではなく単一
            # オブジェクト/文字列にスカラー化されていても 1 を返してしまい（PowerShell はスカラーへの
            # `@()` ラップも `[0]` アクセスも許すため）、単一要素配列がアンラップされていないことを
            # 固定できない。`-is [System.Array]` で配列性そのものを先に固定してから、要素数・中身を
            # 確認する。
            ($parsed.workbook.sheets -is [System.Array]) | Should Be $true
            ($parsed.workbook.names -is [System.Array]) | Should Be $true
            ($parsed.workbook.links -is [System.Array]) | Should Be $true
            ($parsed.sheets -is [System.Array]) | Should Be $true
            ($parsed.sheets[0].tables -is [System.Array]) | Should Be $true
            ($parsed.sheets[0].cells -is [System.Array]) | Should Be $true

            $parsed.workbook.path | Should Be $path
            $parsed.workbook.sheets.Count | Should Be 1
            $parsed.workbook.names.Count | Should Be 1
            $parsed.workbook.names[0].name | Should Be 'OnlyName'
            $parsed.workbook.links.Count | Should Be 0

            $parsed.sheets.Count | Should Be 1
            $parsed.sheets[0].tables.Count | Should Be 1
            $parsed.sheets[0].tables[0].name | Should Be 'OnlyTable'

            $parsed.sheets[0].cells.Count | Should Be 1
            $parsed.sheets[0].cells[0].address | Should Be 'A1'
            $parsed.sheets[0].cells[0].value2 | Should Be 7
        }

        It 'handles multiple discontiguous formula cells within one range (multi-area SpecialCells)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)
                $ws.Range('A2').Formula = '=A1+1'
                $ws.Range('A3').Value2 = (New-XlsSingleCellArray -Value 3.0)
                $ws.Range('A4').Formula = '=A3+1'
                $ws.Range('A5').Value2 = (New-XlsSingleCellArray -Value 5.0)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $model = Get-XlsModel -Path $path
            $cells = $model.sheets[0].cells

            (Get-XlsModelCellByAddress -Cells $cells -Address 'A2').formula | Should Be '=A1+1'
            (Get-XlsModelCellByAddress -Cells $cells -Address 'A4').formula | Should Be '=A3+1'
            (Get-XlsModelCellByAddress -Cells $cells -Address 'A1').formula | Should Be $null
            (Get-XlsModelCellByAddress -Cells $cells -Address 'A3').formula | Should Be $null
            (Get-XlsModelCellByAddress -Cells $cells -Address 'A5').formula | Should Be $null
        }
    }
}
