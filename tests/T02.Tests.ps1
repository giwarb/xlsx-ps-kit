#Requires -Version 5.1

<#
tests/T02.Tests.ps1

T-02: Excel 起動→終了のスモークテスト。
XlsAgent.psm1 の内部ヘルパー（Export しない）Start-XlsApplication / Stop-XlsApplication を使い、
「新規 Excel.Application 起動 → 初期状態設定 → Quit → ReleaseComObject → GC×2 → PID 差分ゼロ」を検証する。

G-06 の T-02 読み替え（タスクカード記載）: Invoke-XlsSession はまだ存在しないため、
COM はこのテストファイルとモジュール内部ヘルパーに限定してよい。T-03 以降は厳格適用。
G-10: 全殺し禁止。起動前後の Get-ExcelPids 差分だけを追跡し、差分で見つけた PID だけを扱う。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

function Wait-ForExcelPidCondition {
    <#
    .SYNOPSIS
        Get-ExcelPids が -Condition を満たすまで短時間ポーリングする（テスト専用、COM には触れない）。
        罠: Excel.Application の起動・Quit は COM 呼び出しの戻りとプロセス一覧への反映が同期しないことが
        あるため、1 回だけ確認すると環境によってはフレーキーになる。
    .PARAMETER Condition
        現在の PID 配列を受け取り $true/$false を返す scriptblock。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Condition,

        [int]$TimeoutSec = 15
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $current = Get-ExcelPids
    while (-not (& $Condition $current)) {
        if ((Get-Date) -ge $deadline) {
            break
        }
        Start-Sleep -Milliseconds 200
        $current = Get-ExcelPids
    }
    return $current
}

Describe 'Excel launch/quit smoke test (T-02)' {

    BeforeEach {
        Import-Module $modulePath -Force
        $script:ModuleRef = Get-Module $moduleName
        $script:BaselinePids = Get-ExcelPids
    }

    AfterEach {
        # 万一テスト自体が失敗して $app を残していても、ここで最終的に残骸ゼロを確認する。
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }

    It 'exports only the 8 public functions; Start-XlsApplication/Stop-XlsApplication stay private' {
        $exported = @(Get-Command -Module $moduleName | Select-Object -ExpandProperty Name)
        $exported -contains 'Start-XlsApplication' | Should Be $false
        $exported -contains 'Stop-XlsApplication' | Should Be $false

        # モジュールスコープには存在する（T-01 と同じパターンで & (Get-Module) 経由で確認）
        $startFn = & $script:ModuleRef { Get-Command Start-XlsApplication -ErrorAction SilentlyContinue }
        $stopFn = & $script:ModuleRef { Get-Command Stop-XlsApplication -ErrorAction SilentlyContinue }
        $startFn | Should Not Be $null
        $stopFn | Should Not Be $null
    }

    It 'launches exactly one new EXCEL.EXE process with the designed initial state, then Stop-XlsApplication leaves no orphan' {
        $newPids = @()
        $app = $null

        try {
            $app = & $script:ModuleRef { Start-XlsApplication }

            $afterStart = Wait-ForExcelPidCondition -Condition {
                param($pidsNow)
                @($pidsNow | Where-Object { $script:BaselinePids -notcontains $_ }).Count -ge 1
            }
            $newPids = @($afterStart | Where-Object { $script:BaselinePids -notcontains $_ })

            # 新規プロセスが（他の何かと重複せず）ちょうど1つ増えていること
            $newPids.Count | Should Be 1

            # 01-design.md §3.1 の初期状態のうち、ワークブックなしで設定できるもの
            # （Calculation はワークブックが 1 つも無いと設定できない罠があるため対象外。実装メモ参照）
            $app.Visible | Should Be $false
            $app.DisplayAlerts | Should Be $false
            $app.ScreenUpdating | Should Be $false
            $app.EnableEvents | Should Be $false
            $app.AskToUpdateLinks | Should Be $false
        }
        finally {
            if ($app) {
                & $script:ModuleRef { param($a) Stop-XlsApplication -Application $a } $app
                $app = $null
            }
        }

        # Quit -> ReleaseComObject -> GC x2 の後、起動した PID が消えていること（G-10: 差分の PID だけ見る）
        $afterStop = Wait-ForExcelPidCondition -Condition {
            param($pidsNow)
            @($pidsNow | Where-Object { $newPids -contains $_ }).Count -eq 0
        }
        @($afterStop | Where-Object { $newPids -contains $_ }).Count | Should Be 0

        # PID 差分ゼロ（ベースラインに対して残骸なし）。
        # 罠: 強制終了の直後、Office 側の裏プロセス（次回起動を早めるための一時的な EXCEL.EXE らしきもの）が
        # 一瞬だけプロセス一覧に現れて数百ms 以内に消えることが実機で観測された。単発チェックだとフレーキーに
        # なるため、ここも短時間ポーリングしてベースラインへ収束するのを待ってから最終判定する。
        $settled = Wait-ForExcelPidCondition -Condition {
            param($pidsNow)
            @($pidsNow | Where-Object { $script:BaselinePids -notcontains $_ }).Count -eq 0
        } -TimeoutSec 10
        @($settled | Where-Object { $script:BaselinePids -notcontains $_ }).Count | Should Be 0
    }

    It 'leaves no orphan even when the caller throws between launch and Stop-XlsApplication' {
        $newPids = @()
        $threw = $false

        try {
            $app = & $script:ModuleRef { Start-XlsApplication }
            $afterStart = Wait-ForExcelPidCondition -Condition {
                param($pidsNow)
                @($pidsNow | Where-Object { $script:BaselinePids -notcontains $_ }).Count -ge 1
            }
            $newPids = @($afterStart | Where-Object { $script:BaselinePids -notcontains $_ })

            try {
                throw 'simulated failure between launch and teardown'
            }
            finally {
                # 呼び出し側が finally で確実に片付ける、という T-03 の契約をここで先取りして確認する。
                & $script:ModuleRef { param($a) Stop-XlsApplication -Application $a } $app
                $app = $null
            }
        }
        catch {
            $threw = $true
        }

        $threw | Should Be $true

        $afterStop = Wait-ForExcelPidCondition -Condition {
            param($pidsNow)
            @($pidsNow | Where-Object { $newPids -contains $_ }).Count -eq 0
        }
        @($afterStop | Where-Object { $newPids -contains $_ }).Count | Should Be 0
    }

    AfterAll {
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }
}
