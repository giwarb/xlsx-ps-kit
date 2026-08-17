#Requires -Version 5.1
<#
.SYNOPSIS
    T-16 ベンチマーク（harness/benchmark/tasks.md）用の fixtures を生成する。
.DESCRIPTION
    fixtures/ 直下に次の 5 ファイルを作る（既存ファイルは削除してから作り直すため、
    再実行しても同じ内容で上書きされる）:
      - model.xlsx  : Assumptions（青文字の入力セル）+ Model（数式）の 5 年財務モデル
      - sales.csv   : 1,000 行 + ヘッダー（Date,Region,Amount）、UTF-8 BOM
      - broken.xlsx : #REF! 3 箇所・#DIV/0! 2 箇所・#NAME? 1 箇所（合計・比率の意図が読み取れる構造）
      - macro.xlsm  : マクロ実体なしの .xlsm（Save-XlsWorkbook の FileFormat 52 保存を確認する縮退版）
      - linked.xlsx : C:\nonexistent\source.xlsx への外部リンク数式を持つセルを含む

    xlsx-ps スキル自身の流儀（skill/SKILL.md）に従う: COM はすべて Invoke-XlsSession の
    ScriptBlock 内、保存は Save-XlsWorkbook のみ、Value2 への一括書き込みは Set-XlsRange
    （内部は [object[,]] を 1 回だけ代入）。数式セルは 1 セルずつ .Formula に書き込む必要がある
    （複数セル範囲への .Formula 代入は同じ文字列がセル数分ブロードキャストされるだけで、
    列ごとに異なる数式を書く用途には使えないため）。書き込み直後に読み戻して一致を確認し、
    不一致なら再試行する Set-XlsFormulaVerified は tests/Test-XlsFormulas.Tests.ps1 で確立した
    パターンをそのまま踏襲した（実機で「連続書き込み直後の Save-XlsWorkbook が最後の 1 セルを
    落とすことがある」罠への対処。SKILL.md 本体の COM gotchas 章を参照）。
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File fixtures\Make-Fixtures.ps1
.NOTES
    harness/state/tasks/T-16.md 手順 1 / harness/benchmark/tasks.md「fixtures の作り方」節に対応。
    EXCEL.EXE を残さない（G-10）: 実行前後の EXCEL.EXE PID 集合を比較し、差分があれば例外にする。
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$fixturesDir = $PSScriptRoot
$repoRoot = Split-Path -Parent $fixturesDir
$modulePath = Join-Path $repoRoot 'skill\scripts\XlsAgent.psm1'

Import-Module -Name $modulePath -Force

# --- ローカルヘルパー（このスクリプト専用。モジュールは export しない。G-02: 高水準ラッパーを
#     本体モジュールに追加しない方針のため、fixture 生成専用のこの補助関数はここに閉じる） ---

function Set-XlsFormulaVerified {
    <#
    .SYNOPSIS
        単一セル Range.Formula に書き込み、直後に読み戻して一致することを確認する。一致しなければ
        短い待機を挟んで再試行する（最大 10 回）。tests/Test-XlsFormulas.Tests.ps1 で確立した
        パターンと同一（実機で踏んだ「連続書き込み直後の保存が最後の 1 セルを落とすことがある」
        罠への対処）。COM は Range オブジェクトの読み書きのみ（呼び出し元の Invoke-XlsSession
        ScriptBlock 内から呼ばれる前提。G-06）。
    .PARAMETER Range
        書き込み先の Range COM オブジェクト（単一セルを想定）。
    .PARAMETER Formula
        書き込む数式文字列（例 '=1/0'）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Range,

        [Parameter(Mandatory = $true)]
        [string]$Formula
    )

    $firstCell = $Range.Cells.Item(1)
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $Range.Formula = $Formula
        if ([string]$firstCell.Formula -eq $Formula) {
            return
        }
        Start-Sleep -Milliseconds 200
    }

    throw "Failed to persist formula '$Formula' on '$($Range.Address(0, 0))' after 10 attempts (last read back: '$($firstCell.Formula)')."
}

function Initialize-XlsSingleSheet {
    <#
    .SYNOPSIS
        新規ワークブックを「名前を付けた 1 枚のシートだけ」の状態にする（既定シート数が
        環境によって 1 枚とは限らないための防御）。COM は Worksheet/Workbook の操作のみ
        （呼び出し元の Invoke-XlsSession ScriptBlock 内から呼ばれる前提。G-06）。
    .PARAMETER Workbook
        Invoke-XlsSession の ScriptBlock 内で受け取った Workbook COM オブジェクト。
    .PARAMETER Name
        残す 1 枚のシートに付ける名前。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    while ($Workbook.Worksheets.Count -gt 1) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }
    $ws = $Workbook.Worksheets.Item(1)
    $ws.Name = $Name
    return $ws
}

function Remove-XlsFixtureFile {
    <#
    .SYNOPSIS
        再実行可能性のため、生成前に既存 fixture を削除する（COM は触らない）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
}

# --- 各 fixture の生成 ---

function New-ModelFixture {
    <#
    .SYNOPSIS
        model.xlsx を作る。Assumptions シート（成長率・粗利率・営業費用率・初年度売上を
        青文字 Font.Color=16711680 の入力セルとして持つ）と、それを参照する数式だけで
        できた 5 年（2024〜2028）の Model シート。B-2（既存ブックの入力セル更新）の題材。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Remove-XlsFixtureFile -Path $Path

    Invoke-XlsSession -Path $Path -ScriptBlock {
        param($app, $wb)

        $wsA = $wb.Worksheets.Item(1)
        $wsA.Name = 'Assumptions'
        while ($wb.Worksheets.Count -gt 1) {
            $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
        }
        $wsM = $wb.Worksheets.Add()
        $wsM.Name = 'Model'
        # Add() はアクティブシートの直前に挿入される（skill/reference/patterns.md）。
        # 挿入位置に依存せず、常に Assumptions を 1 枚目に矯正する。
        if ($wb.Worksheets.Item(1).Name -ne 'Assumptions') {
            $wb.Worksheets.Item('Assumptions').Move($wb.Worksheets.Item(1))
        }

        # --- Assumptions シート（入力セルは青文字 = ハードコード入力。com-constants.md で実測済み） ---
        $assumptionRows = @(
            @('5-Year Financial Model - Assumptions', $null),
            @($null, $null),
            @('Revenue growth rate', 0.05),
            @('Gross margin rate', 0.35),
            @('Operating expense rate', 0.2),
            @('Base revenue - 2024 ($)', 1000000.0)
        )
        Set-XlsRange -Worksheet $wsA -Range 'A1' -Data $assumptionRows

        $wsA.Range('A1').Font.Bold = $true
        $wsA.Range('B3:B6').Font.Color = 16711680   # 青、ハードコード入力（com-constants.md T-13 実測）
        $wsA.Range('B3:B5').NumberFormat = '0.0%'
        $wsA.Range('B6').NumberFormat = '$#,##0'
        $wsA.Columns.Item(1).ColumnWidth = 28

        # --- Model シート（すべて Assumptions を参照する数式。ハードコード値は書かない） ---
        Set-XlsRange -Worksheet $wsM -Range 'A1' -Data (, @('5-Year Financial Model'))
        $wsM.Range('A1').Font.Bold = $true

        # 年は文字列として書く（"2024" であって 2,024 という数値ではない。SKILL.md 財務モデル節）。
        # 罠（実機確認）: 既定（General）書式のセルに Range.Value2 で "2024" のような数値に見える
        # 文字列を書き込むと、Excel 側が自動的に数値（[double]）へ変換してしまい、文字列のまま
        # 保持されない（Set-XlsRange 自身は NumberFormat を変更しないため防げない）。書き込み前に
        # 対象セルへ NumberFormat='@'（テキスト書式）を設定しておくことで、この自動変換を回避できる
        # （回避しない場合との比較・NumberFormat='@' 適用後に Value2 の .NET 型が System.String の
        # ままであることを実機で確認済み）。
        $wsM.Range('A3:F3').NumberFormat = '@'
        Set-XlsRange -Worksheet $wsM -Range 'A3' -Data (, @('Year', '2024', '2025', '2026', '2027', '2028'))
        $wsM.Range('A3:F3').Font.Bold = $true

        Set-XlsRange -Worksheet $wsM -Range 'A4' -Data @(
            @('Revenue'),
            @('Gross Profit'),
            @('Operating Expenses'),
            @('Operating Income')
        )

        $cols = @('B', 'C', 'D', 'E', 'F')

        # Revenue: 初年度は Assumptions の初年度売上、以降は前年 * (1 + 成長率)。
        # シート名 "Assumptions" にはスペース/アポストロフィがなく引用符が不要なため、
        # あえて付けない（付けると Excel が自動的に取り除き、書き込み直後の読み戻し
        # 一致検証 Set-XlsFormulaVerified が常に不一致になってしまう。実機で確認済み。
        # skill/reference/patterns.md「Named ranges」の同種の実機確認と整合）。
        Set-XlsFormulaVerified -Range $wsM.Range('B4') -Formula '=Assumptions!$B$6'
        for ($i = 1; $i -lt $cols.Count; $i++) {
            $cur = $cols[$i]
            $prev = $cols[$i - 1]
            $formula = '={0}4*(1+Assumptions!$B$3)' -f $prev
            Set-XlsFormulaVerified -Range $wsM.Range(($cur + '4')) -Formula $formula
        }

        # Gross Profit / Operating Expenses / Operating Income は各年とも同じ式（列だけが違う）。
        foreach ($c in $cols) {
            Set-XlsFormulaVerified -Range $wsM.Range(($c + '5')) -Formula ('={0}4*Assumptions!$B$4' -f $c)
            Set-XlsFormulaVerified -Range $wsM.Range(($c + '6')) -Formula ('={0}4*Assumptions!$B$5' -f $c)
            Set-XlsFormulaVerified -Range $wsM.Range(($c + '7')) -Formula ('={0}5-{0}6' -f $c)
        }

        $wsM.Range('B4:F7').NumberFormat = '$#,##0;($#,##0);-'
        $wsM.Columns.Item(1).ColumnWidth = 22

        $wsA.UsedRange.Font.Name = 'Arial'
        $wsM.UsedRange.Font.Name = 'Arial'

        Save-XlsWorkbook -Workbook $wb -Path $Path
    } | Out-Null
}

function New-SalesCsvFixture {
    <#
    .SYNOPSIS
        sales.csv を作る（1,000 行 + ヘッダー Date,Region,Amount、日付は 2024 年、
        Region は East/West/North/South、UTF-8 BOM）。COM は使わない（表計算ではなく
        テキスト I/O のため。S-01: -Encoding UTF8 を明示。Windows PowerShell 5.1 の
        Export-Csv -Encoding UTF8 は BOM 付きで書き出す）。B-3 の題材。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Remove-XlsFixtureFile -Path $Path

    $regions = @('East', 'West', 'North', 'South')
    # 固定シードで再実行しても内容が安定する（再実行可能性の要件）。
    $rng = New-Object System.Random(20240817)
    $startDate = [DateTime]::new(2024, 1, 1)

    $rows = for ($i = 0; $i -lt 1000; $i++) {
        $offsetDays = $rng.Next(0, 366)   # 2024 年はうるう年（1/1〜12/31 の 366 日）
        $date = $startDate.AddDays($offsetDays)
        $region = $regions[$rng.Next(0, $regions.Count)]
        $amount = [Math]::Round(($rng.NextDouble() * 4900.0) + 100.0, 2)

        [pscustomobject][ordered]@{
            Date   = $date.ToString('yyyy-MM-dd')
            Region = $region
            Amount = $amount
        }
    }

    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function New-BrokenFixture {
    <#
    .SYNOPSIS
        broken.xlsx を作る。地域別（East/West/North/South）の四半期売上・合計・ノルマ比率という
        「合計」「比率」の意図がはっきり読み取れる表に、#REF! 3 箇所・#DIV/0! 2 箇所・#NAME? 1 箇所を
        仕込む。各エラーセルは互いに独立させ（エラーセルを参照する別の数式を作らない）、
        Test-XlsFormulas の種別ごとの件数が仕込んだ数と正確に一致するようにする。B-4 の題材。
    .NOTES
        エラーの仕込み方（すべて実機で確認済みの手法。skill/reference/com-constants.md,
        tests/Test-XlsFormulas.Tests.ps1 と同じ）:
          - #REF!   : 数式内にエラーリテラル '#REF!' を直接書く（'=#REF!' 単体、または
                      '=SUM(...)+#REF!' のように正常な式へ加える）。
          - #DIV/0! : ノルマ（Quota）が 0 の行で '=(Q1+Q2)/Quota' を計算する（本物のゼロ除算）。
          - #NAME?  : 未定義の名前（GrowthRate）を数式内で参照する。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Remove-XlsFixtureFile -Path $Path

    Invoke-XlsSession -Path $Path -ScriptBlock {
        param($app, $wb)

        $ws = Initialize-XlsSingleSheet -Workbook $wb -Name 'Sales'

        $data = @(
            @('Region', 'Q1', 'Q2', 'Total', 'Quota', '% of Quota'),
            @('East', 100.0, 150.0, $null, 200.0, $null),
            @('West', 120.0, 130.0, $null, 220.0, $null),
            @('North', 90.0, 110.0, $null, 0.0, $null),
            @('South', 80.0, 95.0, $null, 0.0, $null),
            @('Total', $null, $null, $null, $null, $null)
        )
        Set-XlsRange -Worksheet $ws -Range 'A1' -Data $data

        # Total 列（D）: East/West は正しい合計。North/South は #REF!（3 箇所のうち 2 箇所）。
        Set-XlsFormulaVerified -Range $ws.Range('D2') -Formula '=SUM(B2:C2)'
        Set-XlsFormulaVerified -Range $ws.Range('D3') -Formula '=SUM(B3:C3)'
        Set-XlsFormulaVerified -Range $ws.Range('D4') -Formula '=SUM(B4:C4)+#REF!'   # #REF! (1/3): 列削除で壊れた合計を模す
        Set-XlsFormulaVerified -Range $ws.Range('D5') -Formula '=#REF!'              # #REF! (2/3): 参照先ごと消えた合計を模す

        # % of Quota 列（F）: East/West は正しい比率。North/South はノルマが 0 のため #DIV/0!（2 箇所）。
        Set-XlsFormulaVerified -Range $ws.Range('F2') -Formula '=D2/E2'
        Set-XlsFormulaVerified -Range $ws.Range('F3') -Formula '=D3/E3'
        Set-XlsFormulaVerified -Range $ws.Range('F4') -Formula '=(B4+C4)/E4'   # #DIV/0! (1/2): E4(Quota)=0
        Set-XlsFormulaVerified -Range $ws.Range('F5') -Formula '=(B5+C5)/E5'   # #DIV/0! (2/2): E5(Quota)=0

        # Total 行（6）: B/C/E は正しい合計。D は #REF!（3/3、D2:D3 だけを参照し North/South の
        # #REF! を巻き込まないようにしてある）。F は未定義の名前 GrowthRate を参照した #NAME?。
        Set-XlsFormulaVerified -Range $ws.Range('B6') -Formula '=SUM(B2:B5)'
        Set-XlsFormulaVerified -Range $ws.Range('C6') -Formula '=SUM(C2:C5)'
        Set-XlsFormulaVerified -Range $ws.Range('D6') -Formula '=SUM(D2:D3)+#REF!'   # #REF! (3/3)
        Set-XlsFormulaVerified -Range $ws.Range('E6') -Formula '=SUM(E2:E5)'
        Set-XlsFormulaVerified -Range $ws.Range('F6') -Formula '=E6*GrowthRate'      # #NAME? (1/1): GrowthRate は未定義の名前

        $ws.Range('A1:F1').Font.Bold = $true
        $ws.Columns.Item(1).ColumnWidth = 12
        $ws.UsedRange.Font.Name = 'Arial'

        Save-XlsWorkbook -Workbook $wb -Path $Path
    } | Out-Null
}

function New-MacroFixture {
    <#
    .SYNOPSIS
        macro.xlsm を作る。T-16 の環境制約（Trust Center 設定変更が必要な VBA プロジェクトへの
        コード注入は行わない）による縮退版で、マクロの実体は持たない。目的は
        「Save-XlsWorkbook が .xlsm 拡張子から FileFormat 52（xlOpenXMLWorkbookMacroEnabled）で
        保存すること」の確認であり、B-5 では Summary シート追加の題材として使う。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Remove-XlsFixtureFile -Path $Path

    Invoke-XlsSession -Path $Path -ScriptBlock {
        param($app, $wb)

        $ws = Initialize-XlsSingleSheet -Workbook $wb -Name 'Data'

        $data = @(
            @('Item', 'Category', 'Amount'),
            @('Widget A', 'Hardware', 120.5),
            @('Widget B', 'Hardware', 75.0),
            @('Service Plan', 'Services', 300.0),
            @('Widget C', 'Hardware', 45.25),
            @('Consulting', 'Services', 500.0),
            @('Widget D', 'Hardware', 60.0),
            @('Support Contract', 'Services', 220.0),
            @('Widget E', 'Hardware', 90.75),
            @('Training', 'Services', 150.0),
            @('Widget F', 'Hardware', 33.1)
        )
        Set-XlsRange -Worksheet $ws -Range 'A1' -Data $data

        $ws.Range('A1:C1').Font.Bold = $true
        $ws.Columns.Item(1).ColumnWidth = 18
        $ws.UsedRange.Font.Name = 'Arial'

        Save-XlsWorkbook -Workbook $wb -Path $Path
    } | Out-Null
}

function New-LinkedFixture {
    <#
    .SYNOPSIS
        linked.xlsx を作る。C:\nonexistent\source.xlsx への外部リンク数式
        （'C:\nonexistent\[source.xlsx]Sheet1'!A<n> 形式、T-08/T-11 のテストで確立した手法）を
        D 列（Last Year Sales）の 12 行すべてに持つ。B 列はわざと空けてあり、B-6
        （「B 列に前年比を追加」）が書き込む場所として使う。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Remove-XlsFixtureFile -Path $Path

    $months = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

    Invoke-XlsSession -Path $Path -ScriptBlock {
        param($app, $wb)

        $ws = Initialize-XlsSingleSheet -Workbook $wb -Name 'Sales'

        Set-XlsRange -Worksheet $ws -Range 'A1' -Data (, @('Month', $null, 'This Year Sales', 'Last Year Sales (Linked)'))
        $ws.Range('A1:D1').Font.Bold = $true

        for ($i = 0; $i -lt $months.Count; $i++) {
            $r = $i + 2
            $thisYearSales = [double](100 + $i * 15)
            Set-XlsRange -Worksheet $ws -Range ('A' + $r) -Data (, @($months[$i], $null, $thisYearSales, $null))

            # 外部リンク数式（存在しないブックへのリンク）。T-08/T-11 と同じ
            # 'パス\[ファイル名]シート名'!セル 形式。
            $formula = '=''C:\nonexistent\[source.xlsx]Sheet1''!A{0}' -f $r
            Set-XlsFormulaVerified -Range $ws.Range(('D' + $r)) -Formula $formula
        }

        $ws.Columns.Item(1).ColumnWidth = 10
        $ws.Columns.Item(3).ColumnWidth = 16
        $ws.Columns.Item(4).ColumnWidth = 26
        $ws.UsedRange.Font.Name = 'Arial'

        Save-XlsWorkbook -Workbook $wb -Path $Path
    } | Out-Null
}

# --- メイン ---

function Get-XlsFixtureExcelPids {
    [CmdletBinding()]
    param()
    return , @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

$baselinePids = Get-XlsFixtureExcelPids

$modelPath = Join-Path $fixturesDir 'model.xlsx'
$salesCsvPath = Join-Path $fixturesDir 'sales.csv'
$brokenPath = Join-Path $fixturesDir 'broken.xlsx'
$macroPath = Join-Path $fixturesDir 'macro.xlsm'
$linkedPath = Join-Path $fixturesDir 'linked.xlsx'

Write-Output "Generating $modelPath ..."
New-ModelFixture -Path $modelPath

Write-Output "Generating $salesCsvPath ..."
New-SalesCsvFixture -Path $salesCsvPath

Write-Output "Generating $brokenPath ..."
New-BrokenFixture -Path $brokenPath

Write-Output "Generating $macroPath ..."
New-MacroFixture -Path $macroPath

Write-Output "Generating $linkedPath ..."
New-LinkedFixture -Path $linkedPath

$currentPids = Get-XlsFixtureExcelPids
$orphans = @($currentPids | Where-Object { $baselinePids -notcontains $_ })
if ($orphans.Count -gt 0) {
    throw "Orphan EXCEL.EXE process(es) left running after fixture generation: $($orphans -join ', ') (baseline: $($baselinePids -join ', '))"
}

Write-Output 'Done. Generated fixtures:'
Get-ChildItem -LiteralPath $fixturesDir -File |
    Where-Object { $_.Name -match '^(model\.xlsx|sales\.csv|broken\.xlsx|macro\.xlsm|linked\.xlsx)$' } |
    Sort-Object Name |
    ForEach-Object { Write-Output ("  {0,-14} {1,10:N0} bytes" -f $_.Name, $_.Length) }
