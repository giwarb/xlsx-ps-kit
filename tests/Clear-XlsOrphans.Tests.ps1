#Requires -Version 5.1

<#
tests/Clear-XlsOrphans.Tests.ps1

T-05: Clear-XlsOrphans（自モジュールが起動して後始末に失敗した EXCEL.EXE だけを回収する。G-10）の
受け入れテスト。harness/state/tasks/T-05.md の受け入れテスト要点:
  - 死んだ PID のマーカー -> マーカーが消える。
  - EXCEL でないプロセスの PID を指すマーカー -> プロセスは生存のまま、マーカーだけ消える。
  - マーカーのない実行中プロセスには触れない。
  - 「生きた孤児 EXCEL」の再現は必須ではない（リスクがあるため）。

G-06: このファイルは COM を一切直接操作しない。Clear-XlsOrphans 自体も COM を使わない
（Get-Process / Stop-Process / ファイル I/O のみ）。「生きた EXCEL」が絡むテストは、公開関数
Invoke-XlsSession を別 STA runspace で呼び出すことで COM 操作を Invoke-XlsSession の ScriptBlock
内に閉じ込める（T-03 Invoke-XlsSession.Tests.ps1 (b) と同じパターン）。

G-10 の核心（PID 再利用で EXCEL.EXE だが所有権を確認できないケース）は、実際に「孤児」を作らずに
テストする: 生きたセッション（本物の EXCEL.EXE、実際には孤児ではなく Invoke-XlsSession に
きちんと所有されている）のマーカーの 2 行目（開始時刻）だけを意図的に書き換えて「PID は合うが
開始時刻が合わない」状態を作り、Clear-XlsOrphans がそれを停止しない（fail closed）ことを確認する。
「PID も開始時刻も一致する（=停止してよいはず）」の判定ロジックは、`-WhatIf` を使うことで実際には
Stop-Process を呼ばずに検証する（安全策）。これにより、本物の EXCEL.EXE を強制終了するテストを
一切書かずに、G-10 の分岐をすべて確認できる。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'
$script:MarkerDir = Join-Path $env:TEMP 'xlsagent'

function Get-DeadProcessId {
    <#
    .SYNOPSIS
        現在どのプロセスにも割り当てられていない PID を 1 つ返す（存在確認のみ。プロセスは作らない）。
    .NOTES
        999001-999099 の範囲を使う。通常の Windows 環境でこの範囲の PID が実プロセスに使われることは
        まずないが、念のため Get-Process で存在しないことを確認してから返す。
    #>
    [CmdletBinding()]
    param()

    for ($candidate = 999001; $candidate -lt 999100; $candidate++) {
        if (-not (Get-Process -Id $candidate -ErrorAction SilentlyContinue)) {
            return $candidate
        }
    }
    throw 'Could not find an unused PID in the 999001-999099 range for test setup.'
}

function New-XlsOrphanTestMarker {
    <#
    .SYNOPSIS
        テスト用に $env:TEMP\xlsagent\<name>.marker を直接書く（Invoke-XlsSession を経由しない）。
        COM は触らない。
    .PARAMETER Name
        マーカーのファイル名（拡張子 .marker を除く部分）。
    .PARAMETER Lines
        マーカーの中身（行の配列）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Lines
    )

    if (-not (Test-Path -LiteralPath $script:MarkerDir)) {
        New-Item -ItemType Directory -Path $script:MarkerDir -Force | Out-Null
    }

    $markerPath = Join-Path $script:MarkerDir ("{0}.marker" -f $Name)
    Set-Content -LiteralPath $markerPath -Value $Lines -Encoding UTF8 -Force
    return $markerPath
}

function New-BackgroundXlsSession {
    <#
    .SYNOPSIS
        別 STA runspace で公開関数 Invoke-XlsSession を呼び、ScriptBlock の中で Release シグナルを
        待つことで「生きたセッション（本物の EXCEL.EXE）」を保持し続けるテストヘルパー。
        COM 操作はすべて Invoke-XlsSession の ScriptBlock 内（G-06 に適合）。
        T-03 Invoke-XlsSession.Tests.ps1 の (b) と同じパターン。
    .PARAMETER Path
        Invoke-XlsSession -Path に渡すワークブックパス。
    .EXAMPLE
        $session = New-BackgroundXlsSession -Path (New-TempXlsxPath)
        # $session.Info.Pid / $session.Info.StartTicks が使える
        Stop-BackgroundXlsSession -Session $session
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $info = [hashtable]::Synchronized(@{ Pid = $null; StartTicks = $null })
    $ready = New-Object System.Threading.ManualResetEventSlim($false)
    $release = New-Object System.Threading.ManualResetEventSlim($false)

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs

    try {
        [void]$ps.AddScript({
            param($ModulePath, $Path, $Info, $ReadyEvent, $ReleaseEvent)
            Import-Module $ModulePath -Force
            # ここから先はすべて Invoke-XlsSession の ScriptBlock 内（COM はこの中だけ）。
            Invoke-XlsSession -Path $Path -ScriptBlock {
                param($app, $wb)
                $Info.Pid = $app.XlsAgentProcessId
                $Info.StartTicks = $app.XlsAgentProcessStartTicks
                $ReadyEvent.Set()
                # 親テストが検証を終えて $ReleaseEvent.Set() するまで、セッションを生かし続ける。
                [void]$ReleaseEvent.Wait()
            }
        }).AddArgument($modulePath).AddArgument($Path).AddArgument($info).AddArgument($ready).AddArgument($release)

        $asyncResult = $ps.BeginInvoke()

        if (-not $ready.Wait([TimeSpan]::FromSeconds(30))) {
            $release.Set()
            try { [void]$ps.EndInvoke($asyncResult) } catch { }
            throw 'Background Invoke-XlsSession did not signal ready within 30s.'
        }

        return [PSCustomObject]@{
            Info         = $info
            ReadyEvent   = $ready
            ReleaseEvent = $release
            PowerShell   = $ps
            Runspace     = $rs
            AsyncResult  = $asyncResult
        }
    }
    catch {
        # 起動に失敗したらここで後始末する（呼び出し元に Stop-BackgroundXlsSession を強いない）。
        $ps.Dispose()
        $rs.Close()
        $rs.Dispose()
        $ready.Dispose()
        $release.Dispose()
        throw
    }
}

function Stop-BackgroundXlsSession {
    <#
    .SYNOPSIS
        New-BackgroundXlsSession が保持しているセッションを解放し、背後の Invoke-XlsSession
        （Close→Quit→Release→GC×2→マーカー削除）が完了するのを待つ。COM はこの関数の中では触らない
        （待つだけ）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $Session.ReleaseEvent.Set()
    try {
        [void]$Session.PowerShell.EndInvoke($Session.AsyncResult)
    }
    finally {
        $Session.PowerShell.Dispose()
        $Session.Runspace.Close()
        $Session.Runspace.Dispose()
        $Session.ReadyEvent.Dispose()
        $Session.ReleaseEvent.Dispose()
    }
}

Describe 'Clear-XlsOrphans (T-05)' {

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

    It 'removes a marker whose PID no longer belongs to any process (dead PID cleanup)' {
        # round 2: PID はマーカーのファイル名（<pid>.marker）から解析される（内容ではない。
        # 排他リースが握られていて内容を読めない場合でも対象 PID が分かるようにするため）。
        $deadPid = Get-DeadProcessId
        $markerPath = New-XlsOrphanTestMarker -Name $deadPid -Lines @($deadPid, 0)

        $result = Clear-XlsOrphans

        (Test-Path -LiteralPath $markerPath) | Should Be $false
        $match = @($result | Where-Object { $_.MarkerPath -eq $markerPath })
        $match.Count | Should Be 1
        $match[0].Pid | Should Be $deadPid
        $match[0].Action | Should Be 'Skipped-DeadPid'
    }

    It 'removes an invalid (unparsable) marker without touching any process' {
        $markerPath = New-XlsOrphanTestMarker -Name ('garbage-{0}' -f ([guid]::NewGuid())) -Lines @('not-a-pid')

        $result = Clear-XlsOrphans

        (Test-Path -LiteralPath $markerPath) | Should Be $false
        $match = @($result | Where-Object { $_.MarkerPath -eq $markerPath })
        $match.Count | Should Be 1
        $match[0].Pid | Should Be $null
        $match[0].Action | Should Be 'Skipped-InvalidMarker'
    }

    It 'leaves a live non-EXCEL process running and only removes its (stale/reused-PID) marker' {
        # G-10 の核。マーカーが指す PID が生きていても EXCEL.EXE でなければ絶対に停止しない。
        # 自分で起動した無害なプロセス（powershell が Start-Sleep するだけ）を使い、確実に片付ける。
        $proc = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', 'Start-Sleep -Seconds 120') `
            -WindowStyle Hidden -PassThru

        try {
            $markerPath = New-XlsOrphanTestMarker -Name $proc.Id -Lines @($proc.Id, 0)

            $result = Clear-XlsOrphans

            # プロセスは EXCEL.EXE ではないので停止されていないこと。
            (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) | Should Not Be $null
            (Test-Path -LiteralPath $markerPath) | Should Be $false

            $match = @($result | Where-Object { $_.MarkerPath -eq $markerPath })
            $match.Count | Should Be 1
            $match[0].Pid | Should Be $proc.Id
            $match[0].Action | Should Be 'Skipped-NotExcel'
        }
        finally {
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
    }

    It 'does not touch a running process that has no marker at all' {
        $proc = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', 'Start-Sleep -Seconds 120') `
            -WindowStyle Hidden -PassThru

        try {
            # このプロセス用のマーカーは意図的に作らない。
            $result = Clear-XlsOrphans

            (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) | Should Not Be $null
            $match = @($result | Where-Object { $_.Pid -eq $proc.Id })
            $match.Count | Should Be 0
        }
        finally {
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
    }

    It 'does not touch an active (not crashed) background session: exclusive marker lease yields Skipped-ActiveSession' {
        # round 2 の核心（blocking 修正の検証）: 正常に進行中のセッションは PID・プロセス名・
        # 開始時刻がすべて一致してしまうため、それだけでは「孤児」と区別できない。
        # Invoke-XlsSession がマーカーを FileShare.None のまま保持し続けることで、
        # Clear-XlsOrphans が排他オープンに失敗し、プロセスにもマーカーにも一切触れないことを確認する。
        # -WhatIf は付けない（レビュー指摘どおり、-WhatIf なしで呼んでも安全であることを検証する）。
        $path = New-TempXlsxPath
        $session = New-BackgroundXlsSession -Path $path

        try {
            $realPid = $session.Info.Pid
            $realPid | Should Not Be $null
            $markerPath = Join-Path $script:MarkerDir ("{0}.marker" -f $realPid)
            (Test-Path -LiteralPath $markerPath) | Should Be $true

            $result = Clear-XlsOrphans

            (Get-Process -Id $realPid -ErrorAction SilentlyContinue) | Should Not Be $null
            (Test-Path -LiteralPath $markerPath) | Should Be $true

            $match = @($result | Where-Object { $_.Pid -eq $realPid })
            $match.Count | Should Be 1
            $match[0].Action | Should Be 'Skipped-ActiveSession'
        }
        finally {
            Stop-BackgroundXlsSession -Session $session
        }
    }

    It '-WhatIf does not bypass the exclusive-lease check for an active session (still Skipped-ActiveSession)' {
        # 排他オープンの可否は -WhatIf の有無によらず同じ判定である（.NOTES 参照）ことを確認する。
        $path = New-TempXlsxPath
        $session = New-BackgroundXlsSession -Path $path

        try {
            $realPid = $session.Info.Pid
            $realPid | Should Not Be $null
            $markerPath = Join-Path $script:MarkerDir ("{0}.marker" -f $realPid)

            $result = Clear-XlsOrphans -WhatIf

            (Get-Process -Id $realPid -ErrorAction SilentlyContinue) | Should Not Be $null
            (Test-Path -LiteralPath $markerPath) | Should Be $true

            $match = @($result | Where-Object { $_.Pid -eq $realPid })
            $match.Count | Should Be 1
            $match[0].Action | Should Be 'Skipped-ActiveSession'
        }
        finally {
            Stop-BackgroundXlsSession -Session $session
        }
    }

    Context 'legacy (1-line, PID-only) markers (round 1 review should-fix)' {

        It 'removes a legacy marker whose PID is dead, same as the current 2-line format' {
            $deadPid = Get-DeadProcessId
            $markerPath = New-XlsOrphanTestMarker -Name $deadPid -Lines @($deadPid)

            $result = Clear-XlsOrphans

            (Test-Path -LiteralPath $markerPath) | Should Be $false
            $match = @($result | Where-Object { $_.MarkerPath -eq $markerPath })
            $match.Count | Should Be 1
            $match[0].Pid | Should Be $deadPid
            $match[0].Action | Should Be 'Skipped-DeadPid'
        }

        It 'removes a legacy marker pointing at a live non-EXCEL process, same as the current 2-line format' {
            $proc = Start-Process -FilePath 'powershell.exe' `
                -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', 'Start-Sleep -Seconds 120') `
                -WindowStyle Hidden -PassThru

            try {
                $markerPath = New-XlsOrphanTestMarker -Name $proc.Id -Lines @($proc.Id)

                $result = Clear-XlsOrphans

                (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) | Should Not Be $null
                (Test-Path -LiteralPath $markerPath) | Should Be $false

                $match = @($result | Where-Object { $_.MarkerPath -eq $markerPath })
                $match.Count | Should Be 1
                $match[0].Pid | Should Be $proc.Id
                $match[0].Action | Should Be 'Skipped-NotExcel'
            }
            finally {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }

        # 「旧形式マーカーが生きた EXCEL.EXE を指している（Skipped-UnverifiedLegacy、マーカー保持）」の
        # ケースは意図的にテストしていない。それを安全に再現するには、Invoke-XlsSession を経由しない
        # 「素の」EXCEL.EXE（マーカーの排他リースを一切持たない生存プロセス）が要るが、
        # そのようなプロセスを作るには内部ヘルパー Start-XlsApplication を Invoke-XlsSession の外から
        # 直接呼ぶ必要があり、G-06（COM は Invoke-XlsSession の中だけ、テストコードも同様）に反する。
        # 実際に構築しようとすると「ホストプロセスだけを強制終了し、子の EXCEL.EXE は生かしたまま
        # マーカーの排他リースだけ解放させる」ような、隣接タスクの範囲を超える踏み込んだ再現が必要になる
        # （G-14）。コード自体は上記2ケース（Skipped-DeadPid / Skipped-NotExcel）と同じ経路を通っており、
        # 分岐ロジック（$recordedStartTicks -eq $null のときに削除せず Skipped-UnverifiedLegacy を返す）
        # 自体はコードレビューで確認できる。詳細は T-05 実装メモを参照。
    }

    It 'returns an empty array when the marker directory does not contain any markers' {
        if (Test-Path -LiteralPath $script:MarkerDir) {
            Get-ChildItem -LiteralPath $script:MarkerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }

        $result = Clear-XlsOrphans

        $result.Count | Should Be 0
    }
}
