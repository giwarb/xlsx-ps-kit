#Requires -Version 5.1

<#
tests/Invoke-XlsSession.Tests.ps1

T-03: Invoke-XlsSession（COM セッション管理の中核）の受け入れテスト。
02-implementation-plan.md §4 の受け入れ要点:
  (a) 存在しないパスで Workbooks.Add になる
  (b) 別プロセスで開いたファイルを渡すと ReadOnly 例外
  (c) ScriptBlock 内で throw しても Quit されプロセスが残らない
  (d) 戻り値が透過する
  追加: STA 検査、マーカーの作成と削除、-ReadOnly での読み取り専用オープン

G-06（T-03 round 1 レビュー指摘対応）: T-02 の「Invoke-XlsSession 未実装期間の一時的な読み替え」は
Invoke-XlsSession が実装された T-03 では使えない。このファイルは COM を一切直接操作しない。
「別プロセスにロックを保持させる」のは、別 STA runspace で公開関数 Invoke-XlsSession を呼び、
その ScriptBlock の中を同期オブジェクト（ManualResetEventSlim）で「準備完了まで待つ／解放指示まで
留まる」ことで実現する（ScriptBlock の中は Invoke-XlsSession の中なので G-06 上問題ない）。

G-09（T-03 round 2 レビュー指摘対応）: 「事前に既存ファイルを用意する」ために ScriptBlock 内で
$wb.SaveAs を直接呼んでいたが、保存経路は Save-XlsWorkbook（T-04）だけに限定するという G-09 に反する
指摘を受けた。T-04 未実装であることは保存を迂回する理由にならないため、既存ファイルが要る
テストケースは、事前に1回だけ生成して tests/fixtures/blank.xlsx にコミットした静的 fixture を使う
方式に変更した（生成方法は harness/state/tasks/T-03.md の実装メモを参照。生成スクリプト自体は
リポジトリに含めていない — 成果物コードではない使い捨てスクリプトのため）。これにより、
このテストファイルから SaveAs 呼び出しが完全になくなった。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

# 静的fixture: 保存済みの空ワークブック。ReadOnly オープン・別プロセスロックのテストで使う既存ファイル。
# COM を使わずに用意できるものはこれで十分（内容そのものは検証しない。存在確認のみ下記 BeforeAll で行う）。
$script:FixturePath = Join-Path $here 'fixtures\blank.xlsx'

Describe 'Invoke-XlsSession (T-03)' {

    BeforeAll {
        if (-not (Test-Path -LiteralPath $script:FixturePath)) {
            throw "Test fixture not found: '$script:FixturePath'. Regenerate tests/fixtures/blank.xlsx (see harness/state/tasks/T-03.md implementation notes for how it was created)."
        }
    }

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

    It '(a) opens a non-existent path via Workbooks.Add (new, unsaved workbook)' {
        $path = New-TempXlsxPath

        $wbPathSeen = Invoke-XlsSession -Path $path -ScriptBlock {
            param($app, $wb)
            # 新規未保存ワークブックは .Path が空文字列（保存済みなら実フォルダパスが入る）。
            $wb.Path
        }

        $wbPathSeen | Should Be ''
    }

    It '(b) throws a read-only error when another process already has the file open' {
        $path = $script:FixturePath

        # 別プロセス（別 Excel.Application インスタンス）に非読み取り専用でロックを保持させる。
        # 公開関数 Invoke-XlsSession だけを、別の STA runspace 上で呼び出す（レビュー指摘の修正案）。
        $ready = New-Object System.Threading.ManualResetEventSlim($false)
        $release = New-Object System.Threading.ManualResetEventSlim($false)

        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs

        try {
            [void]$ps.AddScript({
                param($ModulePath, $Path, $ReadyEvent, $ReleaseEvent)
                Import-Module $ModulePath -Force
                # ここから先はすべて Invoke-XlsSession の ScriptBlock 内（COM はこの中だけ）。
                Invoke-XlsSession -Path $Path -ScriptBlock {
                    param($app, $wb)
                    $ReadyEvent.Set()
                    # 親テストが検証を終えて $ReleaseEvent.Set() するまで、ここでロックを保持し続ける。
                    [void]$ReleaseEvent.Wait()
                }
            }).AddArgument($modulePath).AddArgument($path).AddArgument($ready).AddArgument($release)

            $asyncResult = $ps.BeginInvoke()

            try {
                if (-not $ready.Wait([TimeSpan]::FromSeconds(30))) {
                    throw 'holder runspace did not signal ready within 30s'
                }

                $errorSeen = $null
                try {
                    Invoke-XlsSession -Path $path -ScriptBlock { param($app, $wb) $wb.Name }
                }
                catch {
                    $errorSeen = $_
                }

                $errorSeen | Should Not Be $null
                $errorSeen.Exception.Message | Should Match 'read-only'
            }
            finally {
                # ロック保持側を解放し、そちらの Invoke-XlsSession（Close/Quit/Release/GCx2）が
                # 完了するのを EndInvoke で待つ。ここが終わるまで次のテストへ進まない。
                $release.Set()
                [void]$ps.EndInvoke($asyncResult)
            }
        }
        finally {
            $ps.Dispose()
            $rs.Close()
            $rs.Dispose()
            $ready.Dispose()
            $release.Dispose()
        }
    }

    It '(c) leaves no orphan EXCEL.EXE process when the ScriptBlock throws' {
        $path = New-TempXlsxPath
        $threw = $false

        try {
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                throw 'simulated failure inside ScriptBlock'
            }
        }
        catch {
            $threw = $true
            $_.Exception.Message | Should Match 'simulated failure inside ScriptBlock'
        }

        $threw | Should Be $true
        # AfterEach の Assert-NoOrphanExcel が最終確認。ここでも明示的に確認する。
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
    }

    It '(d) passes the ScriptBlock return value through unchanged' {
        $path = New-TempXlsxPath
        $marker = 'xlsagent-{0}' -f ([guid]::NewGuid())

        $result = Invoke-XlsSession -Path $path -ScriptBlock {
            param($app, $wb)
            $marker
        }

        $result | Should Be $marker
    }

    It 'throws before launching Excel when the current thread is not STA' {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
        $rs.Open()

        try {
            $ps = [PowerShell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript({
                param($ModulePath)
                Import-Module $ModulePath -Force
                try {
                    Invoke-XlsSession -Path 'C:\does-not-matter.xlsx' -ScriptBlock { param($app, $wb) $wb }
                    'no-throw'
                }
                catch {
                    "threw: $($_.Exception.Message)"
                }
            }).AddArgument($modulePath)

            $output = $ps.Invoke()
            # 罠: 内側の try/catch で例外を捕まえていても、PowerShell エンジンは内部的に
            # $ps.HadErrors を $true にする（Streams.Error は空のまま）。これは Invoke-XlsSession の
            # 挙動ではなくホスト API 側の記録なので、HadErrors ではなく出力メッセージで判定する。
            $output[0] | Should Match 'threw:'
            $output[0] | Should Match 'STA'
        }
        finally {
            $rs.Close()
            $rs.Dispose()
        }

        # STA 検査は Excel を起動する前に行われるはず。新規プロセスが増えていないことを確認する。
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
    }

    It 'writes a PID marker under $env:TEMP\xlsagent while the ScriptBlock runs, then removes it' {
        $path = New-TempXlsxPath
        $markerDir = Join-Path $env:TEMP 'xlsagent'

        $capturedPid = Invoke-XlsSession -Path $path -ScriptBlock {
            param($app, $wb)
            $ownPid = $app.XlsAgentProcessId
            $markerPath = Join-Path (Join-Path $env:TEMP 'xlsagent') ("{0}.marker" -f $ownPid)
            (Test-Path -LiteralPath $markerPath) | Should Be $true
            $ownPid
        }

        $markerPath = Join-Path $markerDir ("{0}.marker" -f $capturedPid)
        (Test-Path -LiteralPath $markerPath) | Should Be $false
    }

    It 'opens an existing file read-only without throwing when -ReadOnly is requested explicitly' {
        $path = $script:FixturePath

        $readOnlySeen = Invoke-XlsSession -Path $path -ReadOnly -ScriptBlock {
            param($app, $wb)
            $wb.ReadOnly
        }

        $readOnlySeen | Should Be $true
    }

    It 'sets Application.Calculation to xlCalculationManual after Open/Add and before the ScriptBlock runs' {
        $path = New-TempXlsxPath
        $expected = & $script:ModuleRef { $script:Xl.xlCalculationManual }

        $calcSeen = Invoke-XlsSession -Path $path -ScriptBlock {
            param($app, $wb)
            $app.Calculation
        }

        $calcSeen | Should Be $expected
    }
}
