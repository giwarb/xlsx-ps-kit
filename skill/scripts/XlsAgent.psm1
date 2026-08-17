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

# COM 定数表。値は実測した上で使う関数ができ次第ここに追加し、skill/reference/com-constants.md にも
# 追記する（G-12）。全量充填は T-13。ここでは実際に使う関数が実装され次第、増分で埋める。
$script:Xl = @{
    # T-03 Invoke-XlsSession: Workbooks.Open/Add の後、ScriptBlock 実行前に設定する
    # （Workbooks.Count が 0 の間は設定できない罠。T-02 実装メモ参照）。実機で -4135 を確認済み。
    xlCalculationManual = -4135

    # T-04 Save-XlsWorkbook: 保存前に Automatic へ戻す。実機で確認済み（実装メモ参照）:
    # Application.Calculation が Automatic のまま保存すると xl/workbook.xml の <calcPr> から
    # calcMode 属性自体が省略される（Excel の既定値が automatic のため）。Manual で保存すると
    # calcMode="manual" が明示的に書き込まれる。
    xlCalculationAutomatic = -4105

    # T-04 Save-XlsWorkbook: 拡張子 -> SaveAs の FileFormat。実機で SaveAs 直後の Workbook.FileFormat
    # を読み戻して確認済み（.xltx/.xltm は Workbooks.Open で開き直すとテンプレートから新規ブックが
    # 生成される罠があるため、SaveAs 直後のその場で確認した。実装メモ参照）。
    xlOpenXMLWorkbook = 51
    xlOpenXMLWorkbookMacroEnabled = 52
    xlOpenXMLTemplate = 54
    xlOpenXMLTemplateMacroEnabled = 53
    xlCSV = 6
}

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
    .PARAMETER Workbook
        （T-03 で追加）Application.Quit() の後、Application 自身より先に ReleaseComObject したい
        Workbook COM オブジェクト。省略可（T-02 のようにワークブックを扱わない呼び出し元は指定不要）。
        指定した場合、解放順は Quit -> Release(Workbook) -> Release(Application) -> GC x2 になる
        （01-design.md §3.1「ReleaseComObject を逆順」＝子から親、を満たすため）。
    .PARAMETER GraceSec
        Quit 後、プロセスが自然終了するのを待つ猶予秒数（既定 5 秒）。
    .EXAMPLE
        Stop-XlsApplication -Application $app
    .EXAMPLE
        Stop-XlsApplication -Application $app -Workbook $wb
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

        レビュー指摘対応（T-03 round 1 should-fix）: 以前は Invoke-XlsSession 側で Workbook を
        Quit より先に Release していたため、設計どおりの Close -> Quit -> Release（Workbook→Application
        の逆順）-> GC x2 になっていなかった。Workbook の解放をこの関数に取り込み、Quit の直後・
        Application の Release より先に行うようにした。各段を入れ子の finally にし、途中の例外が
        後続の解放・GC を妨げないようにしている。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Application,

        $Workbook,

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
        $Application.Quit()
    }
    finally {
        # レビュー指摘対応（T-03 round 1 should-fix）: Quit() の直後、Application を Release する前に
        # Workbook を Release する（子から親への逆順）。Quit() や各 Release が例外を投げても、
        # 後続の解放・GC x2 に必ず到達するよう入れ子の finally にする。
        try {
            if ($null -ne $Workbook -and [Runtime.InteropServices.Marshal]::IsComObject($Workbook)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
            }
        }
        finally {
            try {
                if ([Runtime.InteropServices.Marshal]::IsComObject($Application)) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Application)
                }
            }
            finally {
                # 罠: GC を 1 回だけだと RCW の解放が終わる前に呼び出し元が PID を確認してしまい、
                # プロセスがまだ残って見えることがある（ファイナライザのタイミング）。2 回連続で確実に潰す。
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }
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
    .NOTES
        すべての COM 操作はこの関数の中（または内部ヘルパー Start-XlsApplication/Stop-XlsApplication）に
        限定すること（G-06）。ScriptBlock の外へ Application/Workbook を持ち出して使い続けてはならない。

        後始末は必ず finally で行う: Workbook.Close(SaveChanges:=$false) -> Stop-XlsApplication
        （内部で Application.Quit -> Release(Workbook) -> Release(Application) -> GC x2）-> マーカー
        リース解放 -> マーカー削除。
        ReleaseComObject は Quit の**後**に子（Workbook）から親（Application）の順で行う
        （01-design.md §3.1 の「Close → Quit → Release 逆順 → GC×2」。T-03 round 1 レビューで
        Workbook を Quit より先に Release していた点を指摘され、Stop-XlsApplication に -Workbook を
        渡す形に修正した）。保存したい場合は ScriptBlock の中で明示的に Save-XlsWorkbook を呼ぶこと
        （Close は常に SaveChanges:=$false。G-09 の「DisplayAlerts=$false 中に Close(SaveChanges:=$true) を
        呼ばない」を、そもそも呼ばないことで満たす）。

        マーカー（$env:TEMP\xlsagent\<pid>.marker）は T-05 round 1 レビュー指摘対応で「OS レベルの
        排他リース」に変更した。PID・起動時刻を書き込んだ `FileStream` を `FileShare.None` のまま
        finally まで**開いたまま保持**する。理由: PID・プロセス名・起動時刻だけでは「進行中の正常な
        セッション」と「クラッシュで取り残された孤児」を区別できない（両者とも一致してしまう）。
        排他リースにより、`Clear-XlsOrphans` はマーカーの排他オープンに成功した場合だけを孤児候補として
        扱い、共有違反（＝いま誰かがこのマーカーを保持している＝進行中のセッション）なら一切手を
        触れない。PowerShell ホストがクラッシュすれば OS がハンドルを自動的に解放するため、真の孤児
        だけが回収可能になる（詳細は Clear-XlsOrphans の .NOTES と T-05 実装メモを参照）。
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

    # 罠（設計どおり）: Excel.Application を STA でない COM アパートメントから作ると、
    # ダイアログや一部プロパティアクセスが不安定になる／ハングしうる。Excel を起動する前に検査する。
    $apartmentState = [Threading.Thread]::CurrentThread.ApartmentState
    if ($apartmentState -ne [Threading.ApartmentState]::STA) {
        throw "Invoke-XlsSession requires an STA thread (current ApartmentState: $apartmentState). Run powershell.exe (defaults to STA) or create a runspace with ApartmentState = STA before calling this function."
    }

    $markerDir = Join-Path $env:TEMP 'xlsagent'
    if (-not (Test-Path -LiteralPath $markerDir)) {
        New-Item -ItemType Directory -Path $markerDir -Force | Out-Null
    }

    $app = $null
    $wb = $null
    $markerPath = $null
    $markerLease = $null

    try {
        # G-05: New-Object -ComObject 以外（GetActiveObject 等）で既存プロセスを掴まない。T-02 の
        # Start-XlsApplication が新規プロセスの起動と初期状態設定（Calculation を除く）を担う。
        $app = Start-XlsApplication -Visible:$Visible

        # 起動した PID を marker として記録する（Clear-XlsOrphans が自モジュール起動分だけを回収するための証跡。T-05）。
        # マーカー書式: 1 行目 PID、2 行目 プロセス開始時刻（UTC Ticks）。
        # T-05 round 1 レビュー指摘対応（blocking）: マーカーを書いた後もファイルを閉じてしまうと、
        # 「進行中の正常なセッション」と「クラッシュで取り残された孤児」が PID・プロセス名・開始時刻の
        # 一致だけでは区別できず、Clear-XlsOrphans が進行中セッションを孤児と誤認しうる。これを防ぐため、
        # マーカーを書き込んだ FileStream を FileShare.None のまま finally まで開いたまま保持する
        # （OS レベルの排他リース。詳細は関数 .NOTES と Clear-XlsOrphans の .NOTES を参照）。
        $markerPath = Join-Path $markerDir ("{0}.marker" -f $app.XlsAgentProcessId)
        $markerLease = [IO.File]::Open($markerPath, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $markerText = "{0}`r`n{1}`r`n" -f $app.XlsAgentProcessId, $app.XlsAgentProcessStartTicks
        [byte[]]$markerBytes = [Text.Encoding]::UTF8.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($markerText)
        $markerLease.Write($markerBytes, 0, $markerBytes.Length)
        $markerLease.Flush()

        $fileExists = Test-Path -LiteralPath $Path -PathType Leaf

        if ($fileExists) {
            # UpdateLinks:=0 でリンク更新の確認を抑止（AskToUpdateLinks=$false と二重の防御）。
            $wb = $app.Workbooks.Open($Path, 0, [bool]$ReadOnly)
        }
        elseif ($ReadOnly) {
            throw "File not found: '$Path'. Cannot open a non-existent file with -ReadOnly (there is nothing to read). Omit -ReadOnly to create a new workbook at that path, or check the path."
        }
        else {
            $wb = $app.Workbooks.Add()
        }

        # 開いた直後に ReadOnly を確認する。-ReadOnly を指定していないのに読み取り専用になった場合は
        # 他プロセスが既に開いている可能性が高い（例: 別の Excel インスタンスが排他ロックを保持）。
        if (-not $ReadOnly -and $wb.ReadOnly) {
            throw "Workbook '$Path' was opened read-only by Excel even though -ReadOnly was not requested; another process likely has it open. Close it elsewhere, or pass -ReadOnly if read-only access is what you intended."
        }

        # 罠（T-02 実装メモ）: Application.Calculation は Workbooks.Count=0 の間は設定できない
        # （HRESULT 0x800A03EC）。Open/Add で最低 1 つワークブックが存在するようになった、
        # かつ ScriptBlock を呼ぶ前のこのタイミングで設定する（Codex レビュー承認済みの順序）。
        $app.Calculation = $script:Xl.xlCalculationManual

        return & $ScriptBlock $app $wb
    }
    finally {
        # Close -> Quit -> Release(Workbook -> Application の逆順) -> GC x2 の順（01-design.md §3.1）。
        # Release と Quit・GC は Stop-XlsApplication -Workbook にまとめて任せる（T-03 round 1 レビュー
        # 指摘対応: 以前は Workbook を Quit より先に Release していたため順序が設計と食い違っていた）。
        try {
            if ($null -ne $wb) {
                try { $wb.Close($false) } catch { }
            }
        }
        finally {
            try {
                if ($null -ne $app) {
                    Stop-XlsApplication -Application $app -Workbook $wb
                }
            }
            finally {
                # リースを解放してからでないとマーカーファイルを削除できない（自分自身のハンドルが
                # FileShare.None で開いたままだと Remove-Item も共有違反になる）。
                if ($null -ne $markerLease) {
                    try { $markerLease.Dispose() } catch { }
                }
                if ($markerPath -and (Test-Path -LiteralPath $markerPath)) {
                    Remove-Item -LiteralPath $markerPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

function Get-XlsSaveFileFormat {
    <#
    .SYNOPSIS
        保存先パスの拡張子から SaveAs に渡す FileFormat 定数を決める。COM は触らない純粋関数。
        Export しない内部ヘルパー（T-04）。
    .PARAMETER Path
        保存先パス（拡張子だけを見る。ファイルの存在は問わない）。
    .EXAMPLE
        Get-XlsSaveFileFormat -Path 'C:\tmp\book.xlsm'
    .NOTES
        01-design.md §3.2 の対応表: .xlsx->51 .xlsm->52 .xltx->54 .xltm->53 .csv->6。
        いずれも実機で SaveAs 直後の Workbook.FileFormat を読み戻して確認済み（実装メモ参照）。
        対応表にない拡張子は、次の一手が分かる例外にする（S-04）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $ext = [IO.Path]::GetExtension($Path)
    switch ($ext.ToLowerInvariant()) {
        '.xlsx' { return $script:Xl.xlOpenXMLWorkbook }
        '.xlsm' { return $script:Xl.xlOpenXMLWorkbookMacroEnabled }
        '.xltx' { return $script:Xl.xlOpenXMLTemplate }
        '.xltm' { return $script:Xl.xlOpenXMLTemplateMacroEnabled }
        '.csv' { return $script:Xl.xlCSV }
        default {
            throw "Save-XlsWorkbook does not know the FileFormat for extension '$ext' (path: '$Path'). Supported extensions: .xlsx, .xlsm, .xltx, .xltm, .csv."
        }
    }
}

function Test-XlsSamePath {
    <#
    .SYNOPSIS
        2 つのパス文字列が同一ファイルを指すかを FullName 相当の正規化で比較する。
        COM は一切触らない純粋関数。Export しない内部ヘルパー（T-04）。
    .PARAMETER PathA
        比較するパス（空文字列・$null なら「同一ではない」を返す。未保存ワークブックの
        Workbook.Path が空文字列であることに対応するため）。
    .PARAMETER PathB
        比較するもう一方のパス。
    .EXAMPLE
        Test-XlsSamePath -PathA $Workbook.FullName -PathB $Path
    .NOTES
        [IO.FileInfo]::FullName で相対パス・大文字小文字・区切り文字の違いを正規化してから比較する
        （Windows のファイルパスは大文字小文字を区別しないため OrdinalIgnoreCase）。
        呼び出し側（Save-XlsWorkbook）が Workbook.Path / Workbook.FullName を読む COM アクセス自体を
        行い、ここには文字列だけを渡すことで、このヘルパーを COM なしに単体テストできるようにしている。
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$PathA,

        [AllowEmptyString()]
        [string]$PathB
    )

    if ([string]::IsNullOrEmpty($PathA) -or [string]::IsNullOrEmpty($PathB)) {
        return $false
    }

    $fullA = (New-Object IO.FileInfo($PathA)).FullName
    $fullB = (New-Object IO.FileInfo($PathB)).FullName
    return [string]::Equals($fullA, $fullB, [StringComparison]::OrdinalIgnoreCase)
}

function Set-XlsCalculationAutomatic {
    <#
    .SYNOPSIS
        Application.Calculation を xlCalculationAutomatic に戻し CalculateFullRebuild を実行する。
        Save-XlsWorkbook（T-04）と Test-XlsFormulas（T-10、02-implementation-plan.md §5）が共有する
        内部ヘルパー。Export しない。
    .PARAMETER Application
        対象の Excel.Application COM オブジェクト。Invoke-XlsSession の ScriptBlock 内
        （またはそこから得た Workbook.Application）から渡すこと（G-06）。
    .EXAMPLE
        Set-XlsCalculationAutomatic -Application $Workbook.Application
    .NOTES
        罠（実機確認、実装メモ参照）: Application.Calculation は「今このプロパティに入っている値」が
        SaveAs/Save の瞬間に xl/workbook.xml の <calcPr calcMode="..."/> へそのまま焼き込まれる。
        Manual のまま保存すると calcMode="manual" が明示され、Automatic で保存すると calcMode 属性
        自体が省略される（Excel の既定が automatic のため）。CalculateFullRebuild は Manual 中に
        変更された数式の依存先が古い Value2 のままになる（実機確認）ことへの対策で、保存前に必ず
        呼ぶ。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Application
    )

    $Application.Calculation = $script:Xl.xlCalculationAutomatic
    $Application.CalculateFullRebuild()
}

function Save-XlsWorkbook {
    <#
    .SYNOPSIS
        ワークブックを再計算した上で保存する（保存はこの関数からのみ行う。G-09）。
    .PARAMETER Workbook
        Invoke-XlsSession の ScriptBlock 内で受け取った Workbook COM オブジェクト。
    .PARAMETER Path
        保存先パス。拡張子（.xlsx/.xlsm/.xltx/.xltm/.csv）から FileFormat を決定する。
        Workbook の現在の保存先（FullName）とこの Path が同一ファイルを指す場合は Save()、
        それ以外は SaveAs() を使う。
    .EXAMPLE
        Invoke-XlsSession -Path $p -ScriptBlock { param($app, $wb) Save-XlsWorkbook -Workbook $wb -Path $p }
    .NOTES
        保存前に Application.Calculation を xlCalculationAutomatic に戻し CalculateFullRebuild を
        実行する（Invoke-XlsSession はセッション開始時に Manual にするため、手動のまま保存すると
        開いた人の環境で古い値が見える。01-design.md §3.2、Set-XlsCalculationAutomatic 参照）。

        保存後は Application.Calculation を xlCalculationManual に戻す。01-design.md §3.2 は
        保存後の扱いを明記していないが、Automatic で保存した時点で xl/workbook.xml に
        calcMode 属性省略（= automatic）が焼き込まれているため、「保存後に開き直すと Automatic」
        という受け入れ要点（T-04(b)）は保存の瞬間に既に満たされている。したがって保存後は、
        Invoke-XlsSession がセッション開始時に確立した「セッション中は Manual」という前提を
        ScriptBlock の残りの処理のために壊さないよう、ライブの Application オブジェクト側は
        Manual に戻す判断にした（実装メモ参照）。

        DisplayAlerts=$false は Invoke-XlsSession が起動時に設定済みという前提で SaveAs する
        （G-09 が禁じるのは DisplayAlerts=$false の状態で Workbook.Close(SaveChanges:=$true) を
        呼ぶことであり、Save/SaveAs 自体を禁じてはいない。保存の唯一の経路はこの関数であり、
        ここでは Close を呼ばない）。

        レビュー指摘対応（T-04 round 1 should-fix）: 以前は Set-XlsCalculationAutomatic の呼び出しと
        パス判定が try の外にあり、CalculateFullRebuild や COM プロパティ取得が例外を投げると
        Manual に戻らないまま関数を抜けてしまう可能性があった。Automatic 切替から Save/SaveAs までを
        1 つの try に入れ、catch で生の COM 例外（HRESULT のみ）を次の一手が分かるメッセージに
        包み直し、finally で必ず Manual に戻す（S-04・G-16）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fileFormat = Get-XlsSaveFileFormat -Path $Path
    $app = $Workbook.Application

    try {
        Set-XlsCalculationAutomatic -Application $app

        $currentFullName = $null
        if ($Workbook.Path) {
            $currentFullName = $Workbook.FullName
        }
        $sameFile = Test-XlsSamePath -PathA $currentFullName -PathB $Path

        if ($sameFile) {
            $Workbook.Save()
        }
        else {
            $targetFullName = (New-Object IO.FileInfo($Path)).FullName
            $Workbook.SaveAs($targetFullName, $fileFormat)
        }
    }
    catch {
        throw "Failed to recalculate or save workbook to '$Path'. Check the parent directory, permissions, file locks, and extension. Excel reported: $($_.Exception.Message)"
    }
    finally {
        # 保存の成否によらず、セッションの Manual 前提へ戻す（.NOTES 参照）。
        $app.Calculation = $script:Xl.xlCalculationManual
    }
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

function Remove-XlsOrphanMarker {
    <#
    .SYNOPSIS
        マーカーファイルを削除し、削除できたことを確認する（Export しない内部ヘルパー、T-05 round 2）。
        Clear-XlsOrphans の「掃除したマーカーの一覧を返す」という戻り値契約を、
        `-ErrorAction SilentlyContinue` で握りつぶして黙って成功扱いにしないためのもの。
    .NOTES
        T-05 round 1 レビュー指摘対応（should-fix）: 以前はすべての `Remove-Item` が
        `-ErrorAction SilentlyContinue` のみで、削除後の存在確認がなかった。アクセス拒否や共有違反
        でも `Skipped-*`/`Stopped` を返してしまい、「掃除した」という戻り値の実態と食い違っていた。
        ここで削除後に `Test-Path` で確認し、まだ残っていれば次の一手が分かる例外にする（S-04）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MarkerPath
    )

    Remove-Item -LiteralPath $MarkerPath -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $MarkerPath) {
        throw "Marker '$MarkerPath' could not be removed. Check file permissions or whether another process still holds an exclusive lease on it."
    }
}

function Clear-XlsOrphans {
    <#
    .SYNOPSIS
        本モジュールが起動して後始末に失敗した EXCEL.EXE プロセスだけを回収する。
        `$env:TEMP\xlsagent\<pid>.marker` に記録された PID のみを対象にし、
        ユーザーが自分で開いている Excel には触らない（G-10）。進行中の正常なセッションには触れない。
    .OUTPUTS
        PSCustomObject の配列（マーカーが 1 件もなければ空配列）。各要素:
          Pid        マーカーのファイル名から解析した PID（ファイル名から解析できない場合は $null）
          MarkerPath マーカーファイルのフルパス
          Action     'Stopped'（EXCEL.EXE と確認できたので停止しマーカー削除） /
                     'WhatIf-WouldStop'（-WhatIf 指定時、停止せずマーカーも残す） /
                     'Skipped-ActiveSession'（マーカーを排他オープンできなかった＝別プロセスが
                     いま保持中の進行中セッション。停止もマーカー削除もしない） /
                     'Skipped-DeadPid'（PID のプロセスが既に存在せずマーカーのみ削除） /
                     'Skipped-NotExcel'（PID は生存しているが EXCEL.EXE ではない。PID 再利用。マーカーのみ削除） /
                     'Skipped-Unverified'（PID は EXCEL.EXE だが記録した開始時刻と一致せず所有権を
                     確認できない。fail closed で停止せずマーカーのみ削除） /
                     'Skipped-UnverifiedLegacy'（2 行目（開始時刻）がない旧形式マーカーが生存中の
                     EXCEL.EXE を指している。所有権を確認できないため停止しないが、将来の回収の
                     余地を残すためマーカーも削除しない） /
                     'Skipped-InvalidMarker'（マーカーのファイル名から PID を解析できない。マーカーのみ削除） /
                     'Skipped-CannotAccess'（権限不足など、上記以外の理由でマーカーを開けない。
                     何もしない）
    .EXAMPLE
        Clear-XlsOrphans
    .EXAMPLE
        Clear-XlsOrphans -WhatIf
    .NOTES
        マーカーは Invoke-XlsSession が起動直後に書き、正常終了時（finally）に削除する
        （01-design.md §3.1）。

        T-05 round 1 レビュー指摘対応（blocking）: PID・プロセス名・開始時刻がすべて一致していても、
        それだけでは「クラッシュで取り残された孤児」と「いま実行中の正常なセッション」を区別できない
        （両者ともまったく同じ値になる）。そこで Invoke-XlsSession は、マーカーに書き込んだ
        `FileStream` を `FileShare.None` のまま finally まで**開いたまま保持する**ようにした
        （OS レベルの排他リース）。Clear-XlsOrphans は各マーカーをまず `[IO.File]::Open(...,
        FileShare.None)` で排他オープンしようと試みる:
          - 共有違反（`IOException`、`FileNotFoundException` を除く）-> 別プロセスが今そのマーカーを
            保持している＝進行中の正常なセッション。`Skipped-ActiveSession`。停止もマーカー削除も
            しない（fail closed）。
          - `FileNotFoundException` -> ちょうど今、正常終了した Invoke-XlsSession がマーカーを削除した
            直後（走査と処理の間のレース）。対象が既にないので結果に含めず読み飛ばす。
          - オープンに成功 -> 現在このマーカーを保持しているプロセスはいない（＝ホストがクラッシュして
            OS がハンドルを解放したか、正常終了直後にファイルが残っている異常ケース）。孤児候補として
            中身（PID・開始時刻）を読み、以下の判定に進む。

        判定順序（排他オープンに成功した場合のみ。fail closed。上から順に最初に該当したもので確定）:
          1. マーカーのファイル名（`<pid>.marker` の `<pid>` 部分）から PID を解析できない ->
             停止せずマーカーのみ削除（`Skipped-InvalidMarker`）。ファイル名は Invoke-XlsSession が
             唯一の書き込み元であり、常に `<pid>.marker` の形式になる（内容が読めなくても対象 PID が
             分かるように、内容ではなくファイル名を正とする）。
          2. その PID のプロセスが存在しない -> 停止せず（対象がないので）マーカーのみ削除
             （`Skipped-DeadPid`）。
          3. プロセスは存在するが EXCEL.EXE ではない（PID 再利用で別プロセス） -> 停止せず
             マーカーのみ削除（`Skipped-NotExcel`）。
          4. プロセスは EXCEL.EXE だが、マーカーに開始時刻が記録されていない（2 行目がない旧形式
             マーカー） -> 所有権を確認できないため停止しない。ただし、生きた EXCEL.EXE を指す
             手がかりを失わないよう、マーカー自体は削除しない（`Skipped-UnverifiedLegacy`。
             round 1 レビュー should-fix 対応）。
          5. プロセスは EXCEL.EXE だが、マーカーの開始時刻が実際のプロセス開始時刻と一致しない
             （PID 再利用で別の EXCEL.EXE。ユーザーが新しく開いた Excel の可能性がある） -> 停止せず
             マーカーのみ削除（`Skipped-Unverified`。もう有効な追跡情報ではないので削除してよい）。
          6. プロセスが EXCEL.EXE で、開始時刻も一致する（自モジュール起動の本人だと確認できた） ->
             `$PSCmdlet.ShouldProcess` を通ったうえで停止しマーカー削除（`Stopped`）。

        T-05 round 1 レビュー指摘対応（blocking、TOCTOU）: 6. の判定に使った `Process` オブジェクトを
        そのまま停止に使う。以前は `Stop-Process -Id $markerPid` で PID を再解決していたため、
        照合〜停止（`-Confirm` があればユーザー応答待ちの間も含む）の間に対象プロセスが終了して
        PID が別プロセス（ユーザーの Excel を含む）に再利用されると、その新しいプロセスを誤って
        停止しうる状態だった。修正: `ShouldProcess` を通過した**直後**に同じ `Process` オブジェクトへ
        `Refresh()` を呼び、`HasExited`/`ProcessName`/`StartTime` を再検証する。再検証に失敗したら
        （＝停止しようとした一瞬の間に対象が変わった）停止せずマーカーのみ削除して安全側に倒す。
        再検証を通過した場合のみ `Stop-Process -InputObject $process`（PID の再解決なし）で停止し、
        終了確認は PID の再検索ではなく `Process.WaitForExit(5000)` を使う。

        `-WhatIf` を渡すと、6. に該当する場合でも `Stop-Process` を呼ばず、マーカーも削除しない
        （何が起きるかだけを確認できるようにする）。`-WhatIf` は排他オープン自体はスキップしない
        （進行中セッションかどうかの判定は `-WhatIf` の有無によらず同じ）ため、進行中セッションに対して
        `-WhatIf` を付けても付けなくても結果は `Skipped-ActiveSession` のまま変わらない。
        1〜5 のマーカー掃除はプロセスに触れない低リスク操作のため `-WhatIf` の対象にしていない
        （プロセスへの作用のみを `$PSCmdlet.ShouldProcess` でガードする）。
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $markerDir = Join-Path $env:TEMP 'xlsagent'
    $results = @()

    if (-not (Test-Path -LiteralPath $markerDir)) {
        return , $results
    }

    $markerFiles = @(Get-ChildItem -LiteralPath $markerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue)

    foreach ($markerFile in $markerFiles) {
        $markerPath = $markerFile.FullName

        # ファイル名（<pid>.marker）が唯一の書き込み元（Invoke-XlsSession）と一致する正の PID の
        # 出どころ。中身が読めなくても対象 PID が分かるよう、ここではファイル名だけを見る。
        [int]$markerPid = 0
        $pidParsed = [int]::TryParse($markerFile.BaseName, [ref]$markerPid)

        if (-not $pidParsed -or $markerPid -le 0) {
            Remove-XlsOrphanMarker -MarkerPath $markerPath
            $results += [PSCustomObject]@{ Pid = $null; MarkerPath = $markerPath; Action = 'Skipped-InvalidMarker' }
            continue
        }

        # 排他オープンを試みる。成功する = 現在このマーカーを保持しているプロセスはいない
        # （進行中の正常なセッションではない）。G-10 の核心（blocking 修正）。
        $orphanLease = $null
        try {
            $orphanLease = [IO.File]::Open($markerPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        }
        catch [IO.FileNotFoundException] {
            # 走査と処理の間に正常終了した Invoke-XlsSession がちょうど削除した（レース）。対象なし。
            continue
        }
        catch [IO.IOException] {
            # 共有違反 = 別プロセス（進行中の Invoke-XlsSession）がいま排他的に保持している。
            # fail closed: 停止もマーカー削除もしない。
            $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = 'Skipped-ActiveSession' }
            continue
        }
        catch {
            # 権限不足などその他の理由。所有権を確認できないため、ここでも何もしない（fail closed）。
            $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = 'Skipped-CannotAccess' }
            continue
        }

        $recordedStartTicks = $null
        $action = $null
        $process = $null

        try {
            $orphanLease.Position = 0
            $reader = New-Object IO.StreamReader($orphanLease, [Text.Encoding]::UTF8, $true, 1024, $true)
            try {
                $text = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }

            $lines = @($text -split "`r`n|`n" | Where-Object { $_ -ne '' })
            if ($lines.Count -ge 2) {
                [long]$ticksParsed = 0
                if ([long]::TryParse($lines[1], [ref]$ticksParsed)) {
                    $recordedStartTicks = $ticksParsed
                }
            }

            $process = Get-Process -Id $markerPid -ErrorAction SilentlyContinue

            if (-not $process) {
                $action = 'Skipped-DeadPid'
            }
            elseif ($process.ProcessName -ne 'EXCEL') {
                $action = 'Skipped-NotExcel'
            }
            elseif ($null -eq $recordedStartTicks) {
                # 旧形式（1 行だけ）マーカー。生きた EXCEL.EXE を指しているが所有権を確認できない。
                # round 1 レビュー should-fix 対応: 削除して回収の手がかりを失うより、マーカーを
                # 残して安全側に機能劣化させる。
                $action = 'Skipped-UnverifiedLegacy'
            }
            elseif ($process.StartTime.ToUniversalTime().Ticks -ne $recordedStartTicks) {
                $action = 'Skipped-Unverified'
            }
            else {
                $action = 'ConfirmedOwned'
            }
        }
        finally {
            # リースを解放してからでないとマーカーファイルを削除できない。
            $orphanLease.Dispose()
        }

        if ($action -eq 'Skipped-UnverifiedLegacy') {
            # マーカーは削除しない（.NOTES 参照）。
            $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = $action }
            continue
        }

        if ($action -ne 'ConfirmedOwned') {
            Remove-XlsOrphanMarker -MarkerPath $markerPath
            $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = $action }
            continue
        }

        if (-not $PSCmdlet.ShouldProcess("EXCEL.EXE (PID $markerPid)", 'Stop-Process')) {
            $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = 'WhatIf-WouldStop' }
            continue
        }

        # T-05 round 1 レビュー指摘対応（blocking、TOCTOU）: 照合に使った $process をそのまま停止に
        # 使う。PID の再解決（Stop-Process -Id）はしない。停止直前に同じオブジェクトを再検証し、
        # 一致しなくなっていたら（対象が終了し PID が再利用された等）停止せず安全側に倒す。
        try {
            $process.Refresh()
            if ($process.HasExited -or
                $process.ProcessName -ne 'EXCEL' -or
                $process.StartTime.ToUniversalTime().Ticks -ne $recordedStartTicks) {
                Remove-XlsOrphanMarker -MarkerPath $markerPath
                $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = 'Skipped-Unverified' }
                continue
            }

            Stop-Process -InputObject $process -Force -ErrorAction Stop
            if (-not $process.WaitForExit(5000)) {
                throw "Owned EXCEL.EXE PID $markerPid did not terminate within 5 seconds."
            }
        }
        catch {
            throw "Owned EXCEL.EXE PID $markerPid was not stopped safely: $($_.Exception.Message)"
        }

        Remove-XlsOrphanMarker -MarkerPath $markerPath
        $results += [PSCustomObject]@{ Pid = $markerPid; MarkerPath = $markerPath; Action = 'Stopped' }
    }

    return , $results
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
