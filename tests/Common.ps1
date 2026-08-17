#Requires -Version 5.1
Set-StrictMode -Version Latest

<#
tests/Common.ps1

Pester テスト（Pester 3.4）から dot-source して使う共通ヘルパー。
tests/test-plan.md の「共通（tests/Common.ps1）」節に対応する。
COM は触らない（EXCEL.EXE プロセスの有無を見るだけ）。
#>

function Get-ExcelPids {
    <#
    .SYNOPSIS
        現在動いている EXCEL.EXE プロセスの PID 一覧を返す。
    .EXAMPLE
        $pids = Get-ExcelPids
    #>
    [CmdletBinding()]
    param()

    # 単項 , でラップしないと、ヒット 0 件のとき PowerShell が配列でなく $null を返してしまう
    # （呼び出し側で -Baseline に $null が渡り ParameterBindingValidationException になる）。
    return ,@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Assert-NoOrphanExcel {
    <#
    .SYNOPSIS
        テスト開始前に取得した EXCEL.EXE PID 集合（Baseline）と、現在の PID 集合を比較し、
        差分（残骸プロセス）があれば throw する。各 Tests.ps1 の AfterEach から呼ぶこと（G-07）。
    .PARAMETER Baseline
        テスト開始前に Get-ExcelPids で取得した PID 配列。
    .EXAMPLE
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [int[]]$Baseline
    )

    # 注意: Get-ExcelPids は空集合でも $null にならないよう ,@() で単一オブジェクトとして
    # 配列を返している。ここでさらに @() を被せると二重にラップされてしまうので、そのまま受ける。
    $current = Get-ExcelPids
    $orphans = @($current | Where-Object { $Baseline -notcontains $_ })

    if ($orphans.Count -gt 0) {
        throw "Orphan EXCEL.EXE process(es) left running: $($orphans -join ', ') (baseline: $($Baseline -join ', '))"
    }
}

function New-XlsSingleCellArray {
    <#
    .SYNOPSIS
        単一値を 1x1 の [object[,]] にラップする。COM は触らない。
    .PARAMETER Value
        Range.Value2 に代入したい単一の値。
    .EXAMPLE
        $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'hello')
    .NOTES
        レビュー指摘対応（T-04 round 1 blocking）: G-08「Range.Value2 への代入は [object[,]] のみ」に
        単一セルの例外はない。テストコードでも scalar を直接代入せず、この関数で必ず [object[,]] に
        包んでから Value2 へ渡す。T-06 で Save-XlsWorkbook.Tests.ps1 の同名ヘルパーを Common.ps1 に
        集約した（タスク指示により共通化許可）。Save-XlsWorkbook.Tests.ps1 は独自にローカル定義した
        ものを引き続き使っており（別タスクのファイルを不必要に触らない方針）、こちらは T-06 以降の
        新規テストが使う共通版。

        罠（T-06 round 2 レビュー対応中に発見、実機確認）: `$Value` を型指定なし（untyped）の
        パラメーターで受け取ったまま `$cell[0, 0] = $Value` に入れ、その `[object[,]]` を
        `return ,` で関数境界を、さらに `& $ScriptBlock` でスクリプトブロック境界をまたいで
        `Range.Value2 = ...` に渡すと、値が [double]（例: `5.0`）のときだけ COMException
        (0x800A03EC) になることを実機で確認した。[string]/[bool] の値では再現しない。
        `[Parameter()]` に型を付けず `param($Value)` のまま関数をまたぐと、.NET 側では
        型が正しく [Double] のまま見えるにもかかわらず、COM 相互運用の SAFEARRAY 生成時の
        マーシャリングが崩れるらしい（根本原因は特定できていないが、影響と回避策は実機で
        再現・確認済み）。回避策: 格納する直前に `[double]$Value` へ明示的に再キャストすると
        再現しなくなる（同じく実機確認）。`New-XlsObjectArray2D`（`[array]$Rows` という型付き
        パラメーターで値を受け取り、配列インデックスアクセス経由で格納する）はこの罠を踏まない
        ことも確認済み（実装メモに詳細を記録。SKILL.md gotchas 候補）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    [object[,]]$cell = [Array]::CreateInstance([object], 1, 1)
    if ($Value -is [double]) {
        # 罠（.NOTES 参照）: untyped パラメーターに入った [double] をそのまま格納すると、
        # 関数境界＋スクリプトブロック境界を越えた先の Range.Value2 代入で COMException になる。
        # 明示的に再キャストして格納する。
        $cell[0, 0] = [double]$Value
    }
    else {
        $cell[0, 0] = $Value
    }
    return , $cell
}

function New-XlsObjectArray2D {
    <#
    .SYNOPSIS
        ジャグ配列（行の配列の配列）から Range.Value2 に代入できる 0-based の [object[,]] を作る。
        COM は触らない。
    .PARAMETER Rows
        各要素が 1 行分の値配列であるジャグ配列。すべての行が同じ列数であること。
    .EXAMPLE
        $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(@(1, 2), @(3, 4)))
    .NOTES
        G-08: Range.Value2 への代入は [object[,]] のみ。[Array]::CreateInstance の既定は 0-based で、
        書き込み用途ではこれで問題なく受け付けられることを T-04/実機で確認済み（読み取り側の
        Range.Value2 が 1-based で marshaling されるのとは非対称。Get-XlsRange 実装メモ参照）。
        T-09（Set-XlsRange）が実装されるまで、テストのフィクスチャ作成はこの関数で直接 COM に書き込む
        （Set-XlsRange 未実装を理由に Value2 への素の scalar 代入や別経路を使わない）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Rows
    )

    $rowCount = $Rows.Count
    $colCount = if ($rowCount -gt 0) { @($Rows[0]).Count } else { 0 }

    [object[,]]$arr = [Array]::CreateInstance([object], $rowCount, $colCount)
    for ($r = 0; $r -lt $rowCount; $r++) {
        $rowValues = @($Rows[$r])
        for ($c = 0; $c -lt $colCount; $c++) {
            $arr[$r, $c] = $rowValues[$c]
        }
    }

    return , $arr
}

function New-TempXlsxPath {
    <#
    .SYNOPSIS
        一時テスト用の .xlsx パスを発行する（$env:TEMP\xlsagent-tests\<guid>.xlsx）。
        ファイル自体は作らない。呼び出し側が Invoke-XlsSession 等で作成/保存する。
    .EXAMPLE
        $path = New-TempXlsxPath
    #>
    [CmdletBinding()]
    param()

    $dir = Join-Path $env:TEMP 'xlsagent-tests'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    Join-Path $dir ("{0}.xlsx" -f [guid]::NewGuid().ToString())
}
