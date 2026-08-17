#Requires -Version 5.1

<#
tests/Get-XlsOverview.Tests.ps1

T-07: Get-XlsOverview（markitdown 相当のシート概観ダンプ。01-design.md §3.3）と、Markdown 生成に
使う内部関数 ConvertTo-XlsMarkdownTableCell の受け入れテスト。
harness/state/tasks/T-07.md の受け入れテスト要点:
  - `## SheetName` 見出しが出る（複数シート、-Sheet フィルタ）。
  - 上限: MaxRows/MaxCols を超えるデータで `... (N rows x M cols total)` の省略表記。
  - 表示文字列: 日付セルが表示どおりの文字列で出る（Value2 の生シリアルではない）。
  - 座標（A1 等）が出力に含まれない。
  - 異常系: 存在しないシート名指定で次の一手が分かる例外。

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中でのみ触る。
このテストファイルも例外ではない。テストフィクスチャ（セルへの書き込み）は T-09（Set-XlsRange）が
まだ実装されていないため、Common.ps1 の New-XlsObjectArray2D / New-XlsSingleCellArray で作った
[object[,]]（G-08）を ScriptBlock 内で直接 Range.Value2 に代入して用意する。

純粋関数（ConvertTo-XlsMarkdownTableCell）は COM を一切使わないため、Get-XlsRange.Tests.ps1 と同じ
`& $script:ModuleRef { ... }` パターンでモジュールスコープに入って直接テストする。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

Describe 'Get-XlsOverview (T-07)' {

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

    Context 'ConvertTo-XlsMarkdownTableCell (pure function, no COM)' {

        It 'renders $null and empty string as an empty cell' {
            (& $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text $null }) | Should Be ''
            (& $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text '' }) | Should Be ''
        }

        It 'escapes a pipe character so it does not break the table column separator' {
            $result = & $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text 'a|b' }
            $result | Should Be 'a\|b'
        }

        It 'escapes a backslash before escaping pipes/newlines, so it is not double-escaped' {
            $result = & $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text 'a\b' }
            $result | Should Be 'a\\b'
        }

        It 'replaces CRLF, CR-only, and LF-only line breaks with a literal \n so rows stay on one line' {
            (& $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text "a`r`nb" }) | Should Be 'a\nb'
            (& $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text "a`rb" }) | Should Be 'a\nb'
            (& $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text "a`nb" }) | Should Be 'a\nb'
        }

        It 'combines backslash, pipe, and newline escaping correctly in one value' {
            $result = & $script:ModuleRef { ConvertTo-XlsMarkdownTableCell -Text "pipe|char`r`nnewline\back" }
            $result | Should Be 'pipe\|char\nnewline\\back'
        }
    }

    Context 'Get-XlsOverview integration (real Excel via Invoke-XlsSession)' {

        It 'renders a sheet as a ## heading + Markdown table using display text (not raw Value2)' {
            $path = New-TempXlsxPath
            $expectedDate = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Name = 'Data'
                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Name', 'Score'),
                    @('Alice', 10.0),
                    @('Bob', 20.0)
                ))
                $ws.Range('A4').Value2 = (New-XlsSingleCellArray -Value 45000.0)
                $ws.Range('A4').NumberFormat = 'yyyy-mm-dd'
                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $out = Get-XlsOverview -Path $path

            $out | Should Match ([regex]::Escape('## Data'))
            $out | Should Match ([regex]::Escape('| Name | Score |'))
            $out | Should Match ([regex]::Escape('| --- | --- |'))
            $out | Should Match ([regex]::Escape('| Alice | 10 |'))
            $out | Should Match ([regex]::Escape('| Bob | 20 |'))

            # 表示文字列ベース: 日付は ISO 風の表示文字列であり、Value2 の生シリアル値 (45000) ではない。
            $out | Should Match ([regex]::Escape($expectedDate))
            $out | Should Not Match '45000'
        }

        It 'does not include any A1-style cell coordinates in the output' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Name = 'Coords'
                $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Item', 'Qty'),
                    @('Widget', 3.0)
                ))
                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $out = Get-XlsOverview -Path $path

            # 見出し文字列自体（データ由来ではない部分）に A1 形式の座標が出ていないことを確認する。
            $out | Should Not Match '\$[A-Z]{1,3}\$?\d+'
        }

        It 'appends the "... (N rows x M cols total)" note only when a limit is exceeded, using the actual overflow character' {
            $path = New-TempXlsxPath
            $timesSign = [char]0x00D7

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Name = 'Wide'
                $ws.Range('A1:D4').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('r1c1', 'r1c2', 'r1c3', 'r1c4'),
                    @('r2c1', 'r2c2', 'r2c3', 'r2c4'),
                    @('r3c1', 'r3c2', 'r3c3', 'r3c4'),
                    @('r4c1', 'r4c2', 'r4c3', 'r4c4')
                ))
                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $clipped = Get-XlsOverview -Path $path -Sheet 'Wide' -MaxRows 2 -MaxCols 2
            $expectedNote = "... (4 rows $timesSign 4 cols total)"
            $clipped | Should Match ([regex]::Escape($expectedNote))
            $clipped | Should Match ([regex]::Escape('| r1c1 | r1c2 |'))
            $clipped | Should Not Match 'r3c1'
            $clipped | Should Not Match 'r1c3'

            $full = Get-XlsOverview -Path $path -Sheet 'Wide' -MaxRows 50 -MaxCols 30
            $full | Should Not Match 'total\)'
        }

        It 'renders "(empty)" for a sheet whose UsedRange is a single blank cell' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws1 = $wb.Worksheets.Item(1)
                $ws1.Name = 'HasData'
                $ws1.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'x')

                $ws2 = $wb.Worksheets.Add()
                $ws2.Name = 'Blank'

                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $out = Get-XlsOverview -Path $path -Sheet 'Blank'
            $out | Should Match ([regex]::Escape('## Blank'))
            $out | Should Match '\(empty\)'
        }

        It '-Sheet filters to the requested sheets and preserves the requested order' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws1 = $wb.Worksheets.Item(1)
                $ws1.Name = 'Data'
                $ws1.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'd')

                $ws2 = $wb.Worksheets.Add()
                $ws2.Name = 'Summary'
                $ws2.Range('A1').Value2 = (New-XlsSingleCellArray -Value 's')

                $ws3 = $wb.Worksheets.Add()
                $ws3.Name = 'Extra'
                $ws3.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'e')

                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $out = Get-XlsOverview -Path $path -Sheet 'Data', 'Summary'

            $out | Should Match ([regex]::Escape('## Data'))
            $out | Should Match ([regex]::Escape('## Summary'))
            $out | Should Not Match ([regex]::Escape('## Extra'))

            $indexData = $out.IndexOf('## Data')
            $indexSummary = $out.IndexOf('## Summary')
            ($indexData -lt $indexSummary) | Should Be $true
        }

        It 'omits -Sheet to include every worksheet' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws1 = $wb.Worksheets.Item(1)
                $ws1.Name = 'One'
                $ws1.Range('A1').Value2 = (New-XlsSingleCellArray -Value 1.0)

                $ws2 = $wb.Worksheets.Add()
                $ws2.Name = 'Two'
                $ws2.Range('A1').Value2 = (New-XlsSingleCellArray -Value 2.0)

                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $out = Get-XlsOverview -Path $path
            $out | Should Match ([regex]::Escape('## One'))
            $out | Should Match ([regex]::Escape('## Two'))
        }

        It 'throws a next-step error listing available sheets for a nonexistent sheet name' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Name = 'RealSheet'
                Save-XlsWorkbook -Workbook $wb -Path $path
            } | Out-Null

            $errorSeen = $null
            try {
                Get-XlsOverview -Path $path -Sheet 'DoesNotExist'
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'not found'
            $errorSeen.Exception.Message | Should Match ([regex]::Escape("'DoesNotExist'"))
            $errorSeen.Exception.Message | Should Match ([regex]::Escape('RealSheet'))
        }

        It 'throws a next-step error for MaxRows or MaxCols less than 1 without opening Excel' {
            $path = New-TempXlsxPath
            $script:BaselinePidsLocal = Get-ExcelPids

            $errorSeen = $null
            try {
                Get-XlsOverview -Path $path -MaxRows 0
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'MaxRows'

            $errorSeen2 = $null
            try {
                Get-XlsOverview -Path $path -MaxCols -1
            }
            catch {
                $errorSeen2 = $_
            }
            $errorSeen2 | Should Not Be $null
            $errorSeen2.Exception.Message | Should Match 'MaxCols'

            # ファイルが存在しない状態でも、パラメーター検証で弾かれ Excel は一切起動しないはず。
            Assert-NoOrphanExcel -Baseline $script:BaselinePidsLocal
        }
    }
}
