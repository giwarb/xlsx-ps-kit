#Requires -Version 5.1
Set-StrictMode -Version Latest

<#
XlsAgent.psm1

Excel COM オートメーションを PowerShell 5.1 から安全に使うための薄いラッパー。
公開関数は docs/01-design.md §3 の 7 つ + Clear-XlsOrphans（G-10 / T-05）の計 8 つ。
書式・罫線・グラフ等の高水準ラッパーは追加しない（G-02）。

COM は Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中でのみ触ること（G-06）。
このファイルは T-01 時点では骨組みのみ。各関数の実装は担当タスクカード（T-03 以降）で行う。
#>

# COM 定数表の器。値は実測した上で T-13 で充填し、skill/reference/com-constants.md にも追記する（G-12）。
# 例: $script:Xl.xlCellTypeFormulas, $script:Xl.xlCalculationManual ...
$script:Xl = @{}

function Invoke-XlsSession {
    <#
    .SYNOPSIS
        新規 Excel.Application を起動し、ワークブックを開いた状態で ScriptBlock を実行し、後始末をする。
    .PARAMETER Path
        開く（または新規作成する）ワークブックのパス。
    .PARAMETER ReadOnly
        読み取り専用で開く。省略時、開いた結果が読み取り専用になっていた場合は例外にする
        （他プロセスがファイルを開いている可能性があるため）。
    .PARAMETER Visible
        Excel ウィンドウを表示する（既定は非表示）。
    .PARAMETER ScriptBlock
        `param($app, $wb) ... ` の形で COM オブジェクトを受け取り、処理を行うスクリプトブロック。
        戻り値はそのまま Invoke-XlsSession の戻り値になる。
    .EXAMPLE
        Invoke-XlsSession -Path 'C:\tmp\book.xlsx' -ScriptBlock { param($app, $wb) $wb.Sheets.Count }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$ReadOnly,

        [switch]$Visible,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    throw 'Invoke-XlsSession は未実装です（T-03 で実装予定）。'
}

function Save-XlsWorkbook {
    <#
    .SYNOPSIS
        ワークブックを再計算した上で保存する（保存はこの関数からのみ行う）。
    .PARAMETER Workbook
        Invoke-XlsSession の ScriptBlock 内で受け取った Workbook COM オブジェクト。
    .PARAMETER Path
        保存先パス。拡張子から FileFormat を決定する。
    .EXAMPLE
        Invoke-XlsSession -Path $p -ScriptBlock { param($app, $wb) Save-XlsWorkbook -Workbook $wb -Path $p }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    throw 'Save-XlsWorkbook は未実装です（T-04 で実装予定）。'
}

function Get-XlsOverview {
    <#
    .SYNOPSIS
        シートごとに Markdown 表で内容を概観する（座標なし、編集計画には使わない）。
    .PARAMETER Path
        対象ワークブックのパス。
    .PARAMETER Sheet
        対象シート名（省略時は全シート）。
    .PARAMETER MaxRows
        シートごとに表示する最大行数（既定 50）。
    .PARAMETER MaxCols
        シートごとに表示する最大列数（既定 30）。
    .EXAMPLE
        Get-XlsOverview -Path 'C:\tmp\book.xlsx'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Sheet,

        [int]$MaxRows = 50,

        [int]$MaxCols = 30
    )

    throw 'Get-XlsOverview は未実装です（T-07 で実装予定）。'
}

function Get-XlsModel {
    <#
    .SYNOPSIS
        数式（Formula）と値（Value2）を 1 回の読み取りでまとめて返す。
    .PARAMETER Path
        対象ワークブックのパス。
    .PARAMETER Sheet
        対象シート名（省略時は全シート）。
    .PARAMETER Range
        対象範囲（省略時はシート全体の UsedRange）。
    .PARAMETER FormulasOnly
        数式セルのみを返す。
    .PARAMETER AsJson
        オブジェクトではなく JSON 文字列で返す。
    .EXAMPLE
        Get-XlsModel -Path 'C:\tmp\book.xlsx' -FormulasOnly
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Sheet,

        [string]$Range,

        [switch]$FormulasOnly,

        [switch]$AsJson
    )

    throw 'Get-XlsModel は未実装です（T-08 で実装予定）。'
}

function Get-XlsRange {
    <#
    .SYNOPSIS
        Range.Value2 を 2 次元配列で一括取得する（セル単位ループの代替）。
    .PARAMETER Worksheet
        対象の Worksheet COM オブジェクト。
    .PARAMETER Range
        取得する範囲（A1 形式）。
    .PARAMETER AsCsv
        指定した場合、結果を UTF-8 (BOM 付き) CSV として書き出すパス。
    .PARAMETER AsJson
        指定した場合、結果を JSON として書き出すパス。
    .PARAMETER Header
        1 行目をヘッダーとして扱う。
    .EXAMPLE
        Get-XlsRange -Worksheet $ws -Range 'A1:C10'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$Range,

        [string]$AsCsv,

        [string]$AsJson,

        [switch]$Header
    )

    throw 'Get-XlsRange は未実装です（T-06 で実装予定）。'
}

function Set-XlsRange {
    <#
    .SYNOPSIS
        2 次元配列（または CSV/JSON）を Range.Value2 へ一括書き込みする（セル単位ループ禁止）。
    .PARAMETER Worksheet
        対象の Worksheet COM オブジェクト。
    .PARAMETER Range
        書き込み開始範囲（左上セルのみでも可。データ寸法で自動拡張）。
    .PARAMETER Data
        書き込む [object[,]] 2 次元配列。
    .PARAMETER FromCsv
        書き込むデータの入った CSV ファイルパス。
    .PARAMETER FromJson
        書き込むデータの入った JSON ファイルパス。
    .PARAMETER Header
        1 行目をヘッダーとして書き込む。
    .EXAMPLE
        Set-XlsRange -Worksheet $ws -Range 'A1' -FromCsv 'C:\tmp\data.csv' -Header
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$Range,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByData')]
        [object[,]]$Data,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByCsv')]
        [string]$FromCsv,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByJson')]
        [string]$FromJson,

        [switch]$Header
    )

    throw 'Set-XlsRange は未実装です（T-09 で実装予定）。'
}

function Test-XlsFormulas {
    <#
    .SYNOPSIS
        全シートを再計算し、数式エラーの有無を recalc.py 互換の JSON 契約で返す。
    .PARAMETER Path
        対象ワークブックのパス。
    .PARAMETER TimeoutSec
        再計算のタイムアウト秒数（既定 30）。
    .PARAMETER Force
        外部リンク先が見つからない場合でも再計算を続行する。
    .EXAMPLE
        Test-XlsFormulas -Path 'C:\tmp\book.xlsx'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$TimeoutSec = 30,

        [switch]$Force
    )

    throw 'Test-XlsFormulas は未実装です（T-10〜T-12 で実装予定）。'
}

function Clear-XlsOrphans {
    <#
    .SYNOPSIS
        本モジュールが起動して後始末に失敗した EXCEL.EXE プロセスだけを回収する。
        `$env:TEMP\xlsagent\<pid>.marker` に記録された PID のみを対象にし、
        ユーザーが自分で開いている Excel には触らない（G-10）。
    .EXAMPLE
        Clear-XlsOrphans
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    throw 'Clear-XlsOrphans は未実装です（T-05 で実装予定）。'
}

Export-ModuleMember -Function @(
    'Invoke-XlsSession',
    'Save-XlsWorkbook',
    'Get-XlsOverview',
    'Get-XlsModel',
    'Get-XlsRange',
    'Set-XlsRange',
    'Test-XlsFormulas',
    'Clear-XlsOrphans'
)
