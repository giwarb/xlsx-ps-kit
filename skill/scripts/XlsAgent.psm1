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

if (-not ('XlsAgent.NativeMethods' -as [type])) {
    # レビュー指摘対応（T-02 round 1 blocking）: PID の所有権を Hwnd から解決するための P/Invoke。
    # Get-Process の前後差分だけでは「同時刻にユーザーが起動した別の Excel」を誤って
    # 自分のものと判定しうるため、Application.Hwnd -> GetWindowThreadProcessId で確実に PID を求める。
    Add-Type -Namespace XlsAgent -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(
    System.IntPtr hWnd,
    out uint processId
);
'@
}

function Start-XlsApplication {
    <#
    .SYNOPSIS
        新規 Excel.Application を起動し、ワークブックを開かずに設定できる範囲の初期状態
        （Visible/DisplayAlerts/ScreenUpdating/EnableEvents/AskToUpdateLinks）を設定して返す。
        Export しない内部ヘルパー（T-02）。T-03 の Invoke-XlsSession から使う想定の土台。
    .PARAMETER Visible
        Excel ウィンドウを表示する（既定は非表示）。
    .EXAMPLE
        $app = Start-XlsApplication
    .NOTES
        01-design.md §3.1 は初期状態の一項目として Calculation=xlCalculationManual を挙げているが、
        ここではあえて設定しない（罠、実装メモ参照）。Workbooks.Count が 0 の状態で
        Application.Calculation に代入すると HRESULT 0x800A03EC で失敗することを実機で確認した。
        Calculation の設定は Workbooks.Open/Add の後（T-03 の Invoke-XlsSession 内）で行うこと。

        戻り値の COM オブジェクトには、`Application.Hwnd` から `GetWindowThreadProcessId` で解決した PID を
        `XlsAgentProcessId`、その時点のプロセス開始時刻（UTC Ticks）を `XlsAgentProcessStartTicks` として
        NoteProperty で付与する（Stop-XlsApplication が強制終了フォールバックの所有権確認に使う。
        レビュー指摘: PID 差分だけでは同時刻に起動された別の Excel を誤検出しうるため、
        PID 単独ではなく「PID + 起動時刻」の組で照合する）。
    #>
    [CmdletBinding()]
    param(
        [switch]$Visible
    )

    $app = $null
    try {
        # G-05: New-Object -ComObject 以外（GetActiveObject 等）で既存プロセスを掴まない。必ず新規プロセス。
        $app = New-Object -ComObject Excel.Application

        $app.Visible = [bool]$Visible
        $app.DisplayAlerts = $false
        $app.ScreenUpdating = $false
        $app.EnableEvents = $false
        $app.AskToUpdateLinks = $false

        # 罠: Hwnd はプロセス生成直後は 0 のことがあるため、確定するまで短時間ポーリングする。
        [uint32]$ownerPid = 0
        $deadline = (Get-Date).AddSeconds(5)
        while ($ownerPid -eq 0 -and (Get-Date) -lt $deadline) {
            if ($app.Hwnd -ne 0) {
                [void][XlsAgent.NativeMethods]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$ownerPid)
            }
            if ($ownerPid -eq 0) {
                Start-Sleep -Milliseconds 100
            }
        }

        if ($ownerPid -eq 0) {
            throw 'Application.Hwnd から PID を解決できませんでした（5 秒待っても Hwnd が確定しない）。'
        }

        $process = Get-Process -Id $ownerPid -ErrorAction Stop
        if ($process.ProcessName -ne 'EXCEL') {
            throw "Hwnd から解決した PID $ownerPid は EXCEL.EXE ではありません（$($process.ProcessName)）。所有権を確認できないため中断します。"
        }

        # COM オブジェクトも PowerShell 上では PSObject でラップされているため Add-Member で拡張できる。
        $app | Add-Member -NotePropertyName XlsAgentProcessId -NotePropertyValue ([int]$ownerPid) -Force
        $app | Add-Member -NotePropertyName XlsAgentProcessStartTicks -NotePropertyValue $process.StartTime.ToUniversalTime().Ticks -Force

        return $app
    }
    catch {
        # レビュー指摘対応（T-02 round 1 should-fix）: 初期化途中で例外になっても
        # 起動済みの COM オブジェクトを取りこぼさず、Quit -> Release -> GC x2 まで後始末を試みる。
        if ($null -ne $app) {
            try { $app.Quit() } catch {}
            if ([Runtime.InteropServices.Marshal]::IsComObject($app)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        throw "Excel の起動・初期化に失敗したため後始末を試みました: $($_.Exception.Message)"
    }
}

function Stop-XlsApplication {
    <#
    .SYNOPSIS
        Start-XlsApplication で起動した Excel.Application を Quit → ReleaseComObject → GC ×2 の順で
        後始末する。Export しない内部ヘルパー（T-02）。
    .PARAMETER Application
        後始末する Excel.Application COM オブジェクト（Start-XlsApplication の戻り値）。
    .PARAMETER GraceSec
        Quit 後、プロセスが自然終了するのを待つ猶予秒数（既定 5 秒）。
    .EXAMPLE
        Stop-XlsApplication -Application $app
    .NOTES
        罠（実機で確認）: Application.Quit() → Marshal.ReleaseComObject（戻り値 0 = 参照は完全に解放済み）
        → GC.Collect()/WaitForPendingFinalizers ×2 まで行っても、EXCEL.EXE プロセス自体が数十秒〜1 分以上
        居座るケースを確認した。クライアント側の COM 参照は残っていないため、これはこちら側の参照リークではなく、
        Excel 側の終了処理が非同期・低優先度で行われるためと見られる。そのため「起動時に特定したその PID」
        だけを対象に、猶予後は強制終了するフォールバックを持たせている（G-10: 全殺し禁止の対象外。
        Get-Process EXCEL | Stop-Process のような無差別停止ではなく、自分が起動した PID 1 つに限定）。

        レビュー指摘対応（T-02 round 1 blocking）: 強制終了の直前に、Start-XlsApplication が記録した
        PID とプロセス開始時刻の**両方**が現在の Get-Process の結果と一致するかを確認する。PID が
        再利用され別プロセス（ユーザーの Excel を含む）に割り当てられていた場合、開始時刻が一致しないため
        fail closed（強制終了しない）で throw する。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Application,

        [int]$GraceSec = 5
    )

    $targetPid = $null
    $targetStartTicks = $null
    if ($Application.PSObject.Properties['XlsAgentProcessId']) {
        $targetPid = $Application.XlsAgentProcessId
    }
    if ($Application.PSObject.Properties['XlsAgentProcessStartTicks']) {
        $targetStartTicks = $Application.XlsAgentProcessStartTicks
    }

    try {
        # ワークブックを開いていない前提（T-02 の範囲）。開いている場合の Close は Invoke-XlsSession（T-03）が担う。
        $Application.Quit()
    }
    finally {
        # レビュー指摘対応（T-02 round 1 should-fix）: Quit() が例外を投げても
        # Release と GC x2 まで必ず到達させる（try/finally）。
        if ([Runtime.InteropServices.Marshal]::IsComObject($Application)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Application)
        }

        # 罠: GC を 1 回だけだと RCW の解放が終わる前に呼び出し元が PID を確認してしまい、
        # プロセスがまだ残って見えることがある（ファイナライザのタイミング）。2 回連続で確実に潰す。
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (-not $targetPid) {
        return
    }

    # 罠（.NOTES 参照）: ここまでやってもプロセスが自然に消えるまで数十秒〜1 分以上かかることがある。
    # テスト・Invoke-XlsSession の呼び出し元を長時間ブロックしないよう、猶予秒数だけ待って、
    # まだ残っていれば「起動時に特定したその PID」だけを対象に強制終了する。
    $deadline = (Get-Date).AddSeconds($GraceSec)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process -Id $targetPid -ErrorAction SilentlyContinue)) {
            return
        }
        Start-Sleep -Milliseconds 250
    }

    $target = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
    if (-not $target) {
        return
    }

    # fail closed: PID + 起動時刻が一致しなければ強制終了しない（PID 再利用でユーザーの Excel を
    # 誤って殺さないため。G-10）。
    if (-not $targetStartTicks -or $target.StartTime.ToUniversalTime().Ticks -ne $targetStartTicks) {
        throw "Owned Excel process PID $targetPid の所有権を確認できなかったため強制終了を中止しました（プロセス開始時刻が一致せず、PID 再利用の可能性があります。安全のため何もしません。G-10）。"
    }

    # レビュー指摘対応（T-02 round 1 should-fix）: 強制終了失敗を握りつぶさず、消滅を確認してから返す。
    Stop-Process -InputObject $target -Force -ErrorAction Stop

    $confirmDeadline = (Get-Date).AddSeconds(5)
    while ((Get-Date) -lt $confirmDeadline -and (Get-Process -Id $targetPid -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 100
    }

    if (Get-Process -Id $targetPid -ErrorAction SilentlyContinue) {
        throw "Owned Excel process PID $targetPid did not terminate after Stop-Process."
    }
}

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
