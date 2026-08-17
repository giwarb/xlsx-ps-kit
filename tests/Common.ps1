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
