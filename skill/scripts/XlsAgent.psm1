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

    # T-06 Get-XlsRange: Range.Value2 がエラーセルを返すときの生値。実機で確認済み（実装メモ参照）:
    # Value2 の .NET 型がそのままセル種別を表し、エラー値だけが [int]（数値・日付は必ず [double]）に
    # なるため、この型判定だけでエラー値を一意に検出できる。値そのものは CVErr の Long 表現。
    xlErrNull  = -2146826288
    xlErrDiv0  = -2146826281
    xlErrValue = -2146826273
    xlErrRef   = -2146826265
    xlErrName  = -2146826259
    xlErrNum   = -2146826252
    xlErrNA    = -2146826246
}

# T-06 Get-XlsRange: $script:Xl の xlErr* 値 -> 表示文字列。ConvertFrom-XlsValue2 が参照する
# 実務上のルックアップテーブル（$script:Xl 自体は「確認済みの名前付き定数」を集約する場所という
# 既存の役割を保つため、表示文字列マッピングはここに分離した）。
$script:XlValue2ErrorCodes = @{
    $script:Xl.xlErrNull  = '#NULL!'
    $script:Xl.xlErrDiv0  = '#DIV/0!'
    $script:Xl.xlErrValue = '#VALUE!'
    $script:Xl.xlErrRef   = '#REF!'
    $script:Xl.xlErrName  = '#NAME?'
    $script:Xl.xlErrNum   = '#NUM!'
    $script:Xl.xlErrNA    = '#N/A'
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

        **戻り値の契約（T-06 round 2 レビュー指摘対応で明文化）**: `ScriptBlock` の実行結果は
        「非列挙の単一の値」として扱い、そのまま Invoke-XlsSession 自身の戻り値にする。つまり
        `ScriptBlock` が配列（例えば要素数 1 のジャグ配列）を最後に返した場合、その配列は分解されずに
        1 個の値として呼び出し元に渡る。これは 01-design.md §3.1「戻り値: ScriptBlock の戻り値を
        そのまま返す」を配列にも一貫させるための意図的な契約であり、パイプラインのストリーム透過性
        （`ScriptBlock` が `1; 2` のように複数回出力した場合にそれぞれを別オブジェクトとして
        `Invoke-XlsSession | ForEach-Object` へ渡す、という意味の透過性）は保証しない
        （`ScriptBlock` の複数出力はすべて 1 個の `object[]` にまとめられて返る）。
        `tests/Invoke-XlsSession.Tests.ps1` にスカラー・要素数 1 の配列・要素数 2 以上の配列・
        複数回のパイプライン出力の 4 パターンを固定するテストがある。

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

        # 罠（T-06 で発見、実機確認。実装メモ参照）: `return & $ScriptBlock $app $wb` だと、
        # ScriptBlock の戻り値が「要素数 1 の配列」のとき、PowerShell の `return`（内部的には
        # Write-Output と同じ 1 段階だけ配列を展開する挙動）がその外側の配列を剥がしてしまい、
        # 呼び出し元には配列の中身（さらにそれ自体が配列なら、その要素）がそのまま渡ってしまう
        # （T-03 の受け入れ要点 (d)「戻り値が透過する」は、それまで文字列などスカラーでしか
        # 検証されておらず、この不具合は T-06 の Get-XlsRange が単一行・単一セルのジャグ配列を
        # 返すまで踏まれていなかった）。単項カンマで 1 段階だけ包み直すことで、`&` 呼び出しの結果を
        # スカラーならそのまま、配列ならその配列を要素数によらず 1 個の戻り値として透過させる
        # （このファイル内の Clear-XlsOrphans の `return , $results` と同じ、既存の確立済みパターン）。
        return , (& $ScriptBlock $app $wb)
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

function ConvertTo-XlsMarkdownTableCell {
    <#
    .SYNOPSIS
        1 セル分の表示文字列（Range.Text）を Markdown 表のセルとして安全に埋め込める形にエスケープする。
        COM は触らない純粋関数。Format-XlsSheetOverviewMarkdown の内部ヘルパー（T-07）。Export しない。
    .PARAMETER Text
        エスケープ対象の表示文字列（$null/空文字列も許容し、空セルとして扱う）。
    .EXAMPLE
        ConvertTo-XlsMarkdownTableCell -Text "a|b`nc"
    .NOTES
        実装判断（T-07、仕様に明記がないため記録）:
          1. バックスラッシュを先に `\\` へエスケープする（後続のエスケープが挿入する `\` 自身を
             二重エスケープしないよう、必ず最初に行う）。
          2. パイプ `|` を `\|` へエスケープする（Markdown 表の列区切りと衝突させないため）。
          3. 改行（CRLF/CR/LF いずれも）は表の行区切りと衝突する（Markdown の 1 行 1 レコード構造が
             壊れる）ため、リテラルの 2 文字 `\n`（バックスラッシュ + n）に置き換える。実際の改行文字
             ではなく可視化されたエスケープ表記にすることで、セル内改行があったという情報を失わずに
             1 行に収める。
        往復変換（エスケープを戻す）はサポート対象外（Get-XlsOverview は概観表示専用で「編集計画には
        使うな」という設計方針のため、逆変換ができる必要はない）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    $escaped = $Text -replace '\\', '\\'
    $escaped = $escaped -replace '\|', '\|'
    $escaped = $escaped -replace "`r`n", '\n'
    $escaped = $escaped -replace "`r", '\n'
    $escaped = $escaped -replace "`n", '\n'
    return $escaped
}

function Get-XlsRangeTextGrid {
    <#
    .SYNOPSIS
        指定した左上セルからの矩形範囲を、セル単位の Range.Text（表示文字列）でジャグ配列（行の配列の
        配列）として読み取る。Format-XlsSheetOverviewMarkdown の内部ヘルパー（T-07）。
        COM は Worksheet/Range オブジェクトのプロパティ読み取りのみ（Invoke-XlsSession の ScriptBlock
        内から呼ばれる前提。G-06）。Export しない。
    .PARAMETER Worksheet
        対象の Worksheet COM オブジェクト。
    .PARAMETER TopRow
        読み取る矩形の左上行（1-based）。
    .PARAMETER TopCol
        読み取る矩形の左上列（1-based）。
    .PARAMETER RowCount
        読み取る行数。
    .PARAMETER ColCount
        読み取る列数。
    .EXAMPLE
        Get-XlsRangeTextGrid -Worksheet $ws -TopRow 1 -TopCol 1 -RowCount 10 -ColCount 5
    .NOTES
        罠（T-07 実機確認、実装メモ参照）: `Range.Text` は複数セルの範囲に対しても例外を投げない。
        しかし `Range.NumberFormat`（T-06 の罠と同種）と同じく「範囲内のすべてのセルの表示文字列が
        完全に同一なら」その文字列 1 個をスカラーで返し、1 つでも異なれば `[DBNull]::Value` を返す
        （実機確認: 値の異なる 2x2 範囲で `System.DBNull` が返った）。つまり複数セルの表示文字列を
        一括取得する経路は存在せず、01-design.md §3.3 の「Range.Text（表示文字列）ベース」を満たすには
        セル単位で読むしかない（T-06 の Get-XlsRange が Value2 は一括、NumberFormat だけセル単位
        フォールバックするのとは違い、Get-XlsOverview は Text 自体が実質的に常にセル単位になる）。
        呼び出し元の Format-XlsSheetOverviewMarkdown が MaxRows/MaxCols でクリップした後の小さい範囲
        だけをこの関数に渡すことで、COM 呼び出し回数を既定 50×30=1500 回程度に抑えている（タスクカード
        T-07 の実装判断: 一括 Value2+NumberFormat での近似はせず、クリップ後のみセル単位 Text を許容）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$TopRow,

        [Parameter(Mandatory = $true)]
        [int]$TopCol,

        [Parameter(Mandatory = $true)]
        [int]$RowCount,

        [Parameter(Mandatory = $true)]
        [int]$ColCount
    )

    $rows = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt $RowCount; $r++) {
        $row = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt $ColCount; $c++) {
            $cell = $Worksheet.Cells.Item($TopRow + $r, $TopCol + $c)
            [void]$row.Add([string]$cell.Text)
        }
        [void]$rows.Add($row.ToArray())
    }

    return , $rows.ToArray()
}

function Format-XlsSheetOverviewMarkdown {
    <#
    .SYNOPSIS
        1 シート分の `## SheetName` 見出し + Markdown 表を組み立てる。座標（A1 等）は出さない。
        Format-XlsSheetOverviewMarkdown の呼び出し元は Get-XlsOverview（T-07）。
        COM は Worksheet.UsedRange と（Get-XlsRangeTextGrid 経由の）Cells.Item(...).Text の読み取りのみ
        （Invoke-XlsSession の ScriptBlock 内から呼ばれる前提。G-06）。Export しない。
    .PARAMETER Worksheet
        対象の Worksheet COM オブジェクト。
    .PARAMETER MaxRows
        表示する最大行数。
    .PARAMETER MaxCols
        表示する最大列数。
    .EXAMPLE
        Format-XlsSheetOverviewMarkdown -Worksheet $ws -MaxRows 50 -MaxCols 30
    .NOTES
        実装判断（T-07、仕様に明記がないため記録）:
          - 空シートの扱い: `UsedRange` が 1x1 かつその唯一のセルの `Text` が空文字列の場合だけ
            「空シート」とみなし `(empty)` を出す。真に空白のシートでは `UsedRange` が既定で `$A$1`
            （1x1、Value2=$null）になることを実機で確認済み（実装メモ参照）。単一セルにのみ実データが
            あるシート（例: A1 だけに "X"）はこの条件に当たらないため、通常どおり 1 行の表として出す
            （"X" というヘッダー行のみ、データ行なしの表になる）。数式が空文字列 "" を返す単一セルの
            シートは表示文字列が空になるため「空シート」に分類されるが、仕様に明記がない縮退ケースで
            あり実装判断として許容する。
          - ヘッダー行の扱い: クリップ後の範囲の 1 行目をそのまま Markdown 表の見出し行として使う
            （実際にヘッダーかどうかは問わない。座標由来の列名（A, B, C...）は「座標を出さない」という
            設計方針に反するため使わない。Markdown 表として構文的に妥当にするには見出し行が必要なので、
            シートの先頭行をそのまま流用する）。
          - `... (N rows × M cols total)` は `UsedRange` の合計行数・合計列数（クリップ前の値）を使い、
            行・列のどちらか一方でも上限を超えていれば出す。
          - `MaxRows`/`MaxCols` の下限（1 以上）は呼び出し元の Get-XlsOverview がパラメーター段階で
            検証する。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$MaxRows,

        [Parameter(Mandatory = $true)]
        [int]$MaxCols
    )

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("## $($Worksheet.Name)")
    [void]$lines.Add('')

    $usedRange = $Worksheet.UsedRange
    $totalRows = [int]$usedRange.Rows.Count
    $totalCols = [int]$usedRange.Columns.Count
    $topRow = [int]$usedRange.Row
    $topCol = [int]$usedRange.Column

    if ($totalRows -eq 1 -and $totalCols -eq 1 -and [string]::IsNullOrEmpty([string]$usedRange.Text)) {
        [void]$lines.Add('(empty)')
        [void]$lines.Add('')
        return ($lines.ToArray() -join "`n")
    }

    $clipRows = [Math]::Min($totalRows, $MaxRows)
    $clipCols = [Math]::Min($totalCols, $MaxCols)

    $grid = Get-XlsRangeTextGrid -Worksheet $Worksheet -TopRow $topRow -TopCol $topCol -RowCount $clipRows -ColCount $clipCols

    $headerCells = New-Object System.Collections.ArrayList
    foreach ($cellText in $grid[0]) {
        [void]$headerCells.Add((ConvertTo-XlsMarkdownTableCell -Text $cellText))
    }
    [void]$lines.Add('| ' + ($headerCells.ToArray() -join ' | ') + ' |')

    $sepCells = New-Object System.Collections.ArrayList
    for ($c = 0; $c -lt $clipCols; $c++) {
        [void]$sepCells.Add('---')
    }
    [void]$lines.Add('| ' + ($sepCells.ToArray() -join ' | ') + ' |')

    for ($r = 1; $r -lt $grid.Count; $r++) {
        $rowCells = New-Object System.Collections.ArrayList
        foreach ($cellText in $grid[$r]) {
            [void]$rowCells.Add((ConvertTo-XlsMarkdownTableCell -Text $cellText))
        }
        [void]$lines.Add('| ' + ($rowCells.ToArray() -join ' | ') + ' |')
    }

    if ($totalRows -gt $clipRows -or $totalCols -gt $clipCols) {
        # `u{00D7}` Unicode エスケープは PowerShell 6.2+ の構文のため使わない（G-11）。
        # [char]0x00D7（乗算記号 ×）をファイルの実バイト表現に依存せず明示的に組み立てる。
        $timesSign = [char]0x00D7
        [void]$lines.Add('')
        [void]$lines.Add("... ($totalRows rows $timesSign $totalCols cols total)")
    }

    [void]$lines.Add('')
    return ($lines.ToArray() -join "`n")
}

function Get-XlsOverview {
    <#
    .SYNOPSIS
        シートごとに Markdown 表で内容を概観する（座標なし、編集計画には使わない）。
    .PARAMETER Path
        対象ワークブックのパス。
    .PARAMETER Sheet
        対象シート名（省略時は全シート、ワークブックのタブ順）。指定した場合は指定した順序・
        指定したシートのみを出す。存在しないシート名があれば例外にする。
    .PARAMETER MaxRows
        シートごとに表示する最大行数（既定 50）。1 未満は次の一手が分かる例外にする。
    .PARAMETER MaxCols
        シートごとに表示する最大列数（既定 30）。1 未満は次の一手が分かる例外にする。
    .EXAMPLE
        Get-XlsOverview -Path 'C:\tmp\book.xlsx'
    .EXAMPLE
        Get-XlsOverview -Path 'C:\tmp\book.xlsx' -Sheet 'Summary', 'Data' -MaxRows 20 -MaxCols 10
    .NOTES
        内部で `Invoke-XlsSession -ReadOnly` を使う（呼び出し側は COM を見ない。G-06）。
        `Range.Text`（表示文字列）ベースで座標を出さない。SKILL.md では「編集計画には使うな」と
        書くこと（列幅が狭くて `#` などの表示になっているセルも、Excel が実際に表示している文字列を
        そのまま出す。Format-XlsSheetOverviewMarkdown の実装メモ・Get-XlsRangeTextGrid の実装メモ参照）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Sheet,

        [int]$MaxRows = 50,

        [int]$MaxCols = 30
    )

    if ($MaxRows -lt 1) {
        throw "MaxRows must be 1 or greater (got $MaxRows). Pass a positive row limit, e.g. -MaxRows 50."
    }
    if ($MaxCols -lt 1) {
        throw "MaxCols must be 1 or greater (got $MaxCols). Pass a positive column limit, e.g. -MaxCols 30."
    }

    return Invoke-XlsSession -Path $Path -ReadOnly -ScriptBlock {
        param($app, $wb)

        $availableNames = @($wb.Worksheets | ForEach-Object { $_.Name })

        $targetSheets = New-Object System.Collections.ArrayList
        if ($Sheet) {
            foreach ($name in $Sheet) {
                if ($availableNames -notcontains $name) {
                    throw "Sheet '$name' not found in '$Path'. Available sheets: $($availableNames -join ', ')."
                }
                [void]$targetSheets.Add($wb.Worksheets.Item($name))
            }
        }
        else {
            foreach ($ws in $wb.Worksheets) {
                [void]$targetSheets.Add($ws)
            }
        }

        $blocks = New-Object System.Collections.ArrayList
        foreach ($ws in $targetSheets) {
            [void]$blocks.Add((Format-XlsSheetOverviewMarkdown -Worksheet $ws -MaxRows $MaxRows -MaxCols $MaxCols))
        }

        return ($blocks.ToArray() -join "`n")
    }
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

function Split-XlsNumberFormatSections {
    <#
    .SYNOPSIS
        NumberFormat 文字列を、引用符リテラル・バックスラッシュエスケープの外側にある ';' で
        最大 4 セクション（正数;負数;ゼロ;テキスト）に分割する。COM は触らない純粋関数。
        Test-XlsIsDateNumberFormat の内部ヘルパー（T-06 round 2）。Export しない。
    .PARAMETER NumberFormat
        分割対象の書式文字列。
    .EXAMPLE
        Split-XlsNumberFormatSections -NumberFormat 'm/d/yyyy;@'
    .NOTES
        round 2 レビュー指摘（blocking）対応: Excel の書式コードは最大 4 セクション
        （正数;負数;ゼロ;テキスト）を持てる（Microsoft の書式仕様どおり）。セクション数が 1 なら
        すべての値に同じ書式、2 なら 正数&ゼロ;負数、3 なら 正数;負数;ゼロ、4 なら
        正数;負数;ゼロ;テキスト。区切りの ';' は、引用符で囲まれた文字リテラルの中や
        バックスラッシュエスケープの直後には現れないものとして扱う（その中の ';' はセクション
        区切りにしない）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$NumberFormat
    )

    $sections = New-Object System.Collections.ArrayList

    if ([string]::IsNullOrEmpty($NumberFormat)) {
        [void]$sections.Add('')
        return , $sections.ToArray()
    }

    $current = New-Object Text.StringBuilder
    $i = 0
    $len = $NumberFormat.Length
    while ($i -lt $len) {
        $ch = $NumberFormat[$i]
        if ($ch -eq '"') {
            [void]$current.Append($ch)
            $i++
            while ($i -lt $len -and $NumberFormat[$i] -ne '"') {
                [void]$current.Append($NumberFormat[$i])
                $i++
            }
            if ($i -lt $len) {
                [void]$current.Append($NumberFormat[$i])
                $i++
            }
        }
        elseif ($ch -eq '\' -and ($i + 1) -lt $len) {
            [void]$current.Append($ch)
            [void]$current.Append($NumberFormat[$i + 1])
            $i += 2
        }
        elseif ($ch -eq ';') {
            [void]$sections.Add($current.ToString())
            $current = New-Object Text.StringBuilder
            $i++
        }
        else {
            [void]$current.Append($ch)
            $i++
        }
    }
    [void]$sections.Add($current.ToString())

    return , $sections.ToArray()
}

function Get-XlsNumberFormatSectionCondition {
    <#
    .SYNOPSIS
        セクション文字列の先頭にある条件トークン（例 [<100]、[>=100]）を解析する。条件トークンが
        なければ $null を返す。COM は触らない純粋関数。Select-XlsNumberFormatSection の内部ヘルパー
        （T-06 round 3）。Export しない。
    .PARAMETER Section
        判定対象の 1 セクション分の書式文字列（Split-XlsNumberFormatSections の要素 1 つ）。
    .EXAMPLE
        Get-XlsNumberFormatSectionCondition -Section '[<100]0.00'
    .NOTES
        round 2 レビュー指摘（blocking）対応。Excel の条件付き書式（Microsoft の書式仕様どおり）は、
        セクション先頭の角括弧に比較演算子（<, >, <=, >=, =, <>）と数値を書くことで、既定の
        正数/負数/ゼロによるセクション振り分けを置き換える。セクション先頭から角括弧トークンを
        順に見ていき、比較演算子+数値のパターンに一致する最初のものを条件として採用する
        （色 `[Red]` やロケール/通貨コード `[$-411]` 等、条件でない角括弧トークンはスキップして
        次の角括弧を見る）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Section
    )

    if ([string]::IsNullOrEmpty($Section)) {
        return $null
    }

    $i = 0
    $len = $Section.Length
    while ($i -lt $len -and $Section[$i] -eq '[') {
        $close = $Section.IndexOf(']', $i)
        if ($close -lt 0) {
            return $null
        }
        $token = $Section.Substring($i + 1, $close - $i - 1)
        if ($token -match '^(?<op><=|>=|<>|<|>|=)\s*(?<num>-?\d+(\.\d+)?)$') {
            return [PSCustomObject]@{
                Operator = $Matches['op']
                Number   = [double]$Matches['num']
            }
        }
        $i = $close + 1
    }

    return $null
}

function Test-XlsNumberFormatConditionMatches {
    <#
    .SYNOPSIS
        Get-XlsNumberFormatSectionCondition が返した条件オブジェクトに、対象の値が一致するかを
        判定する。COM は触らない純粋関数。Select-XlsNumberFormatSection の内部ヘルパー
        （T-06 round 3）。Export しない。
    .PARAMETER Condition
        Get-XlsNumberFormatSectionCondition の戻り値（Operator/Number を持つオブジェクト）。
    .PARAMETER Value
        判定対象の数値。
    .EXAMPLE
        Test-XlsNumberFormatConditionMatches -Condition $cond -Value 45000
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Condition,

        [double]$Value = 0
    )

    switch ($Condition.Operator) {
        '<' { return [bool]($Value -lt $Condition.Number) }
        '>' { return [bool]($Value -gt $Condition.Number) }
        '<=' { return [bool]($Value -le $Condition.Number) }
        '>=' { return [bool]($Value -ge $Condition.Number) }
        '=' { return [bool]($Value -eq $Condition.Number) }
        '<>' { return [bool]($Value -ne $Condition.Number) }
    }

    return $false
}

function Select-XlsNumberFormatSection {
    <#
    .SYNOPSIS
        Split-XlsNumberFormatSections が返したセクション配列から、対象の値に適用されるセクションを
        1 つ選ぶ。COM は触らない純粋関数。Test-XlsIsDateNumberFormat の内部ヘルパー（T-06 round 2、
        条件付きセクション対応は round 3）。Export しない。
    .PARAMETER Sections
        Split-XlsNumberFormatSections の戻り値。
    .PARAMETER Value
        対象の数値（符号、または各セクションの条件でどのセクションを使うか決める）。
    .EXAMPLE
        Select-XlsNumberFormatSection -Sections $sections -Value -5.0
    .NOTES
        round 2 レビュー指摘（blocking）対応: Excel では `[<100]0.00;[>=100]yyyy-mm-dd` のように、
        セクション先頭の条件トークンが既定の符号による振り分けを置き換える（Microsoft の書式仕様
        どおり）。いずれかのセクションに条件トークンがあれば、まずそちらを優先する:
          1. 各セクションの条件を先頭から順に評価し、最初に値が一致したセクションを返す。
          2. どの条件にも一致しなければ、条件を持たない（無条件の）セクションを catch-all として
             返す（Excel の規則: 条件付きセクションが埋まらない残りの 1 セクションが無条件の
             フォールバックになる）。
          3. すべてのセクションに条件があり、かつどれにも一致しない場合は `$null` を返す
             （round 3 レビュー指摘 blocking 対応: 以前は「仕様上まず起こらない縮退ケース」として
             `$Sections[0]` にフォールバックしていたが、これは偽だった条件付きセクションへ戻って
             しまう誤り。例 `'[<0]yyyy-mm-dd;[>100]0.00'` と値 `50` は、どちらの条件も偽であり、
             Excel はどちらの数値表示も適用しない。第 1 条件（`<0`）へ戻って日付として解釈するのは
             誤変換であり、`$Sections[0]` を機械的に返す実装判断では正しさを保てないと判明した。
             呼び出し元の `Test-XlsFormatSectionHasDateToken` は `$null`/空文字列を安全に「日付でない」
             として扱う）。
        いずれのセクションにも条件トークンがない場合だけ、従来どおり符号による既定の振り分け
        （セクション数が 1 なら常にそのセクション。2 なら 正数&ゼロ=1 番目、負数=2 番目。3 以上
        （4 番目はテキスト用で今回対象の [double] の判定には使わない）なら 正数=1 番目、
        負数=2 番目、ゼロ=3 番目）を使う。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Sections,

        [double]$Value = 0
    )

    $conditions = New-Object System.Collections.ArrayList
    foreach ($section in $Sections) {
        [void]$conditions.Add((Get-XlsNumberFormatSectionCondition -Section $section))
    }

    $hasAnyCondition = $false
    foreach ($cond in $conditions) {
        if ($null -ne $cond) {
            $hasAnyCondition = $true
            break
        }
    }

    if ($hasAnyCondition) {
        for ($i = 0; $i -lt $Sections.Count; $i++) {
            $cond = $conditions[$i]
            if ($null -ne $cond -and (Test-XlsNumberFormatConditionMatches -Condition $cond -Value $Value)) {
                return $Sections[$i]
            }
        }
        for ($i = 0; $i -lt $Sections.Count; $i++) {
            if ($null -eq $conditions[$i]) {
                return $Sections[$i]
            }
        }
        # round 3 レビュー指摘（blocking）対応: すべて条件付きでどれにも一致しない場合、
        # 偽だった条件へ戻らず $null を返す（.NOTES 参照）。
        return $null
    }

    $count = $Sections.Count
    if ($count -le 1) {
        return $Sections[0]
    }
    if ($count -eq 2) {
        if ($Value -lt 0) { return $Sections[1] }
        return $Sections[0]
    }
    if ($Value -gt 0) { return $Sections[0] }
    if ($Value -lt 0) { return $Sections[1] }
    return $Sections[2]
}

function Test-XlsFormatSectionHasDateToken {
    <#
    .SYNOPSIS
        書式コード 1 セクション分の文字列に、日付/時刻トークン（d/m/y/h/s、または [h]/[m]/[s] 系の
        経過時間トークン）が「リテラルとして除外されずに」含まれるかを判定する。COM は触らない
        純粋関数。Test-XlsIsDateNumberFormat の内部ヘルパー（T-06 round 2）。Export しない。
    .PARAMETER Section
        判定対象の 1 セクション分の書式文字列（Split-XlsNumberFormatSections の要素 1 つ）。
    .EXAMPLE
        Test-XlsFormatSectionHasDateToken -Section 'm/d/yyyy'
    .NOTES
        round 2 レビュー指摘（blocking）対応: 引用符で囲まれた文字リテラル（例 "USD"）、
        バックスラッシュエスケープの次の 1 文字（例 `0\m` の m はリテラルの m であって月ではない）、
        アンダースコアの次の 1 文字（幅合わせのスペーサー）、アスタリスクの次の 1 文字（繰り返し
        埋め文字）は、日付トークンとして数えない。角括弧 `[...]` は、中身が `h`/`hh`/`m`/`mm`/`s`/
        `ss`（大文字小文字問わず）のときだけ経過時間トークンとして日付/時刻扱いにし、それ以外
        （色指定 `[Red]`、ロケール/通貨コード `[$-411]`、条件 `[>100]` 等）は読み飛ばして中身の
        文字を判定に使わない（`[Red]` の d、`[Yellow]` の y 等による誤検出を防ぐ）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Section
    )

    if ([string]::IsNullOrEmpty($Section)) {
        return $false
    }

    $i = 0
    $len = $Section.Length
    while ($i -lt $len) {
        $ch = $Section[$i]

        if ($ch -eq '"') {
            $i++
            while ($i -lt $len -and $Section[$i] -ne '"') { $i++ }
            $i++
        }
        elseif ($ch -eq '\') {
            $i += 2
        }
        elseif ($ch -eq '_') {
            $i += 2
        }
        elseif ($ch -eq '*') {
            $i += 2
        }
        elseif ($ch -eq '[') {
            $close = $Section.IndexOf(']', $i)
            if ($close -lt 0) {
                $i = $len
            }
            else {
                $token = $Section.Substring($i + 1, $close - $i - 1)
                if ($token -match '^[hHmMsS]+$') {
                    return $true
                }
                $i = $close + 1
            }
        }
        else {
            if ($ch -match '[dmyhsDMYHS]') {
                return $true
            }
            $i++
        }
    }

    return $false
}

function Test-XlsIsDateNumberFormat {
    <#
    .SYNOPSIS
        NumberFormat 文字列が（対象の値にとって）日付/時刻書式かどうかを判定する。COM は触らない
        純粋関数。ConvertFrom-XlsValue2（T-06）の内部ヘルパー。Export しない。
    .PARAMETER NumberFormat
        判定対象の Range.NumberFormat 文字列。
    .PARAMETER Value
        判定対象の数値。Excel の書式は最大 4 セクション（正数;負数;ゼロ;テキスト）を持てるため、
        どのセクションを見るかは値の符号で決まる（省略時は 0 として扱い、正数/ゼロ用のセクションを
        見る）。
    .EXAMPLE
        Test-XlsIsDateNumberFormat -NumberFormat 'm/d/yyyy;@' -Value 45000
    .NOTES
        round 2 レビュー指摘（blocking）対応: 旧実装は書式文字列のどこかに `@` があるだけで
        一律 `$false` にしていたため、`m/d/yyyy;@`（正数&ゼロ用は日付、負数用はテキスト）のような
        ごく一般的な複数セクション書式を非日付と誤判定していた。また角括弧を一律に除去していたため
        `[h]` 単独の経過時間書式を見落とし、逆にバックスラッシュエスケープ（例 `0\m`）を誤って
        日付と判定していた。値に対応するセクションだけを Select-XlsNumberFormatSection で選び、
        Test-XlsFormatSectionHasDateToken でリテラル除外込みのトークン判定をする方式に改めた。

        罠（実機確認）: 既定の「標準」書式は Range.NumberFormat（本来カルチャ非依存のはずの
        プロパティ）でも、日本語ロケールの Excel では英語の "General" ではなく "G/標準" として
        返ってくることを確認した。"General"/"G/標準" いずれも d/m/y/h/s トークンを含まないため、
        本関数はリテラル比較をせずトークン判定だけで両方とも正しく非日付と判定する。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$NumberFormat,

        [double]$Value = 0
    )

    if ([string]::IsNullOrEmpty($NumberFormat)) {
        return $false
    }

    $sections = Split-XlsNumberFormatSections -NumberFormat $NumberFormat
    $section = Select-XlsNumberFormatSection -Sections $sections -Value $Value

    return Test-XlsFormatSectionHasDateToken -Section $section
}

function ConvertTo-XlsIsoDateString {
    <#
    .SYNOPSIS
        Excel の日付シリアル値（Value2、1900 または 1904 日付体系）を ISO 8601 文字列に変換する。
        変換できない場合（OADate の有効範囲外）は $null を返す。COM は触らない純粋関数。
        ConvertFrom-XlsValue2 の内部ヘルパー（T-06 round 2）。Export しない。
    .PARAMETER OleAutomationValue
        Range.Value2 の生値（日付として解釈する Double）。
    .PARAMETER Date1904
        対象ワークブックが 1904 日付体系（Workbook.Date1904 = $true）かどうか。
    .EXAMPLE
        ConvertTo-XlsIsoDateString -OleAutomationValue 45000
    .NOTES
        round 2 レビュー指摘（blocking）対応。実機で確認した Excel シリアル値と
        [DateTime]::FromOADate() の差異（新規ワークブックのセルに直接シリアル値を入れ、
        Range.Text の表示と比較して確認。実装メモ参照）:
          - 1904 日付体系（Workbook.Date1904=$true）: Excel のシリアル値に 1462 を足すと
            .NET の OADate と一致する（1900 年始まりの OADate 基準と 1904 年始まりの Excel 基準の
            差。実機で serial 0/1/59/60/61/367 を確認し、すべて +1462 で Range.Text と一致した。
            1904 年は実際に閏年のため、下記の 1900 年体系のような架空日の補正は不要）。
          - 1900 日付体系（既定）: シリアル 1〜59（1900-01-01〜1900-02-28）は .NET の OADate より
            1 小さいため +1 の補正が要る。Excel は Lotus 1-2-3 互換のため 1900 年を閏年として扱う
            バグを意図的に継承しているが、.NET 側の OADate 基準はこのバグを持たないため、
            serial 61（1900-03-01）以降はこの補正をすると逆にずれる（補正なしで Range.Text と
            一致する）。実機で serial 1/2/59/60/61/367 を確認した。
          - シリアル 60 は Excel 上で "1900-02-29" と表示される架空の日付で、実在しないため
            .NET の DateTime では表現できない（FromOADate はこの日を作れない）。ここでは情報を
            失わないよう、Excel が表示する文字列をそのまま使う（整数部が 60 の間は "1900-02-29" を
            使い、端数（時刻）があれば ISO 形式で付加する。実装判断）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [double]$OleAutomationValue,

        [switch]$Date1904
    )

    $oaValue = $OleAutomationValue

    if ($Date1904) {
        $oaValue += 1462.0
    }
    elseif ($oaValue -ge 60.0 -and $oaValue -lt 61.0) {
        # 実在しない 1900-02-29（Excel が意図的に継承した架空の閏日）。.NOTES 参照。
        $fractionalDay = $oaValue - 60.0
        if ($fractionalDay -eq 0.0) {
            return '1900-02-29'
        }
        $timeOfDay = [TimeSpan]::FromDays($fractionalDay)
        return ('1900-02-29T{0}' -f $timeOfDay.ToString('hh\:mm\:ss'))
    }
    elseif ($oaValue -ge 1.0 -and $oaValue -lt 60.0) {
        $oaValue += 1.0
    }

    try {
        $dt = [DateTime]::FromOADate($oaValue)
    }
    catch {
        # 罠: 日付書式のセルだが OADate の有効範囲外（おおよそ西暦 100 年〜9999 年の外）の値。
        # 変換できないため呼び出し元に $null を返し、生の数値のまま使ってもらう（実装判断）。
        return $null
    }

    if ($dt.TimeOfDay -eq [TimeSpan]::Zero) {
        return $dt.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    }
    return $dt.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertFrom-XlsValue2 {
    <#
    .SYNOPSIS
        Range.Value2 の 1 セル分の生値を、日付・空セル・エラー値の変換規則に従って PowerShell の値に
        変換する。COM は一切触らない純粋関数。Get-XlsRange の内部関数（T-06）。T-07/T-08 も再利用する
        想定でこのシグネチャにしている（呼び出し側が Value2 と NumberFormat を渡すだけで済み、
        呼び出し側自身は COM オブジェクトを渡す必要がない）。Export しない。
    .PARAMETER Value2
        セル 1 つ分の Range.Value2（$null=空セル、[double]=数値/日付、[string]=文字列、[bool]=真偽値、
        [int]=エラー値のいずれか）。
    .PARAMETER NumberFormat
        そのセルの Range.NumberFormat（日付判定に使う）。意味を持たない場合は空文字列でよい。
    .PARAMETER Date1904
        対象ワークブックが 1904 日付体系（Workbook.Date1904 = $true）かどうか（round 2 レビュー
        指摘対応で追加。省略時は既定の 1900 日付体系として扱う）。
    .EXAMPLE
        ConvertFrom-XlsValue2 -Value2 45000 -NumberFormat 'm/d/yyyy'
    .NOTES
        変換規則（01-design.md §3.5、実機確認。実装メモに実測の詳細を記録）:
          - $null（空セル）はそのまま $null。
          - Value2 の .NET 型がそのままセルの種別を表す（実機確認、COM marshaling の性質）:
            数値・日付は必ず [double]（整数値でも [int] にはならない）、文字列は [string]、
            論理値は [bool]、エラー値だけが [int]（CVErr の Long 表現、例: #DIV/0! = -2146826281）。
            この型だけでエラー値かどうかを一意に判定でき、値の大小や別途のフラグ取得が要らない。
          - [double] かつ NumberFormat が日付/時刻書式（Test-XlsIsDateNumberFormat。値の符号で
            セクションを選ぶため `-Value $Value2` を渡す）なら ConvertTo-XlsIsoDateString で
            ISO 8601 文字列にする（Date1904・1900 年 1〜2 月の架空閏日の補正込み）。
          - [int] はエラー値。$script:XlValue2ErrorCodes にある既知 7 種（NULL/DIV0/VALUE/REF/NAME/
            NUM/N-A）は "#REF!" 等の表示文字列に変換する。未知のコード（新しい動的配列エラー
            #SPILL!/#CALC! 等、未確認）は情報を失わず "#ERR(<code>)" にフォールバックする
            （数式エラーの厳密な検証は T-10 の Test-XlsFormulas の責務であり、ここでは「読める」ことを
            優先する。実装判断）。
          - それ以外（文字列・真偽値・日付でない数値・OADate 範囲外の日付書式値）はそのまま返す。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value2,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$NumberFormat,

        [switch]$Date1904
    )

    if ($null -eq $Value2) {
        return $null
    }

    if ($Value2 -is [int]) {
        if ($script:XlValue2ErrorCodes.ContainsKey($Value2)) {
            return $script:XlValue2ErrorCodes[$Value2]
        }
        return "#ERR($Value2)"
    }

    if ($Value2 -is [double] -and (Test-XlsIsDateNumberFormat -NumberFormat $NumberFormat -Value $Value2)) {
        $iso = ConvertTo-XlsIsoDateString -OleAutomationValue $Value2 -Date1904:$Date1904
        if ($null -ne $iso) {
            return $iso
        }
        return $Value2
    }

    return $Value2
}

function ConvertFrom-XlsRangeToRows {
    <#
    .SYNOPSIS
        Range COM オブジェクトから Value2 を一括取得し、ConvertFrom-XlsValue2 を適用したジャグ配列
        （行の配列の配列）にして返す。Get-XlsRange（T-06）の内部関数。COM は Range オブジェクトの
        プロパティ読み取りのみを行う（Invoke-XlsSession の ScriptBlock 内から呼ばれる前提。G-06）。
        Export しない。
    .PARAMETER RangeObj
        対象の Range COM オブジェクト（Worksheet.Range(...) の戻り値）。
    .PARAMETER Date1904
        対象ワークブックが 1904 日付体系（Workbook.Date1904 = $true）かどうか（round 2 レビュー
        指摘対応で追加。呼び出し側が Worksheet.Parent.Date1904 を 1 回読んで渡す想定）。
    .EXAMPLE
        ConvertFrom-XlsRangeToRows -RangeObj $ws.Range('A1:C10')
    .NOTES
        罠（実機確認）:
          - 単一セル範囲は Value2 が 2 次元配列でなくスカラーで返る（01-design.md §5 既知の罠）。
            `$raw -is [Array]` で判定して分岐する。
          - Range.Value2 が返す 2 次元配列は .NET の既定である 0-based ではなく、Excel の
            Range.Cells(1,1) 起点に合わせて 1-based（GetLowerBound(0)/(1) がともに 1）で
            marshaling される。0 起点を仮定して `$raw[0,0]` のようにアクセスすると
            IndexOutOfRangeException になる。GetLowerBound/GetUpperBound で実際の境界を
            取得してループすること（Range.Value2 への**書き込み**側は逆に 0-based の
            [object[,]] で問題なく受け付ける非対称な挙動も実機で確認した。読みと書きで前提が
            違う点に注意）。
          - Range.NumberFormat は Value2 と違い、複数セルの範囲を一括で 2 次元配列にできない。
            範囲内の書式がすべて同一ならスカラー文字列 1 個、1 つでも異なれば [DBNull]::Value を
            返す（実機確認）。まず一括読みを試し、DBNull のときだけセル単位
            （`Range.Cells.Item(r,c).NumberFormat`。Value2 の 2 次元配列と同じ、範囲左上を (1,1)
            とする相対 1-based インデックスであることを実機で確認済み）にフォールバックする。
            書式が範囲全体で統一されている（よくあるケース）なら COM 呼び出し 1 回で済み、
            混在時のみセル数分の呼び出しになる（大きい範囲で書式が細かく混在していると遅くなりうる
            既知の制約。列単位の最適化は今回のスコープでは見送った。実装メモ参照）。
          - round 2 レビュー指摘（should-fix）対応: DBNull フォールバック時のセル単位
            NumberFormat 取得は、日付判定に使う `[double]` のセルだけに限定する（文字列・空セル・
            真偽値・エラー値は日付判定が不要で、NumberFormat を読んでも捨てるだけの無駄な COM
            呼び出しになるため）。ヘッダー行が既定書式・本文が日付書式という典型的な表でも、
            `[double]` セルだけに絞ることで無駄な呼び出しを避けられる。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $RangeObj,

        [switch]$Date1904
    )

    $raw = $RangeObj.Value2
    $rows = New-Object System.Collections.ArrayList

    if ($raw -isnot [Array]) {
        # 単一セル範囲（罠、.NOTES 参照）。NumberFormat は真の単一セルなら常にスカラー文字列。
        $nf = $RangeObj.NumberFormat
        $row = New-Object System.Collections.ArrayList
        [void]$row.Add((ConvertFrom-XlsValue2 -Value2 $raw -NumberFormat $nf -Date1904:$Date1904))
        [void]$rows.Add($row.ToArray())
        return , $rows.ToArray()
    }

    $rowLower = $raw.GetLowerBound(0)
    $rowUpper = $raw.GetUpperBound(0)
    $colLower = $raw.GetLowerBound(1)
    $colUpper = $raw.GetUpperBound(1)

    $bulkFormat = $RangeObj.NumberFormat
    $useUniformFormat = ($null -ne $bulkFormat -and $bulkFormat -isnot [DBNull])
    $uniformFormat = if ($useUniformFormat) { $bulkFormat } else { $null }

    for ($r = $rowLower; $r -le $rowUpper; $r++) {
        $row = New-Object System.Collections.ArrayList
        for ($c = $colLower; $c -le $colUpper; $c++) {
            $cellValue = $raw[$r, $c]

            # round 2 レビュー指摘（should-fix）: NumberFormat が日付判定に使われるのは [double] の
            # ときだけなので、それ以外の型ではセル単位の COM 呼び出しを発生させない（.NOTES 参照）。
            $nf = ''
            if ($cellValue -is [double]) {
                if ($useUniformFormat) {
                    $nf = $uniformFormat
                }
                else {
                    $nf = $RangeObj.Cells.Item($r, $c).NumberFormat
                }
            }

            [void]$row.Add((ConvertFrom-XlsValue2 -Value2 $cellValue -NumberFormat $nf -Date1904:$Date1904))
        }
        [void]$rows.Add($row.ToArray())
    }

    return , $rows.ToArray()
}

function ConvertTo-XlsHeaderObjects {
    <#
    .SYNOPSIS
        1 行目がヘッダーであるジャグ配列を PSCustomObject の配列に変換する。COM は触らない純粋関数。
        Get-XlsRange の -Header 用内部ヘルパー（T-06）。Export しない。
    .PARAMETER Rows
        1 行目がヘッダーであるジャグ配列（ConvertFrom-XlsRangeToRows の戻り値と同じ形）。
    .EXAMPLE
        ConvertTo-XlsHeaderObjects -Rows $rows
    .NOTES
        実装判断: ヘッダーセルが $null／空文字列の場合は "Column<列番号(1始まり)>" に置き換える。
        列名が重複する場合は後の列が [ordered] ハッシュテーブルのキーを上書きする（後勝ち）。
        いずれも仕様に明記がないための実装判断で、実装メモに記録する。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Rows
    )

    $objects = New-Object System.Collections.ArrayList

    if ($Rows.Count -eq 0) {
        return , $objects.ToArray()
    }

    $headerRow = $Rows[0]
    $keys = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $headerRow.Count; $i++) {
        $h = $headerRow[$i]
        $key = if ($null -eq $h -or [string]$h -eq '') { 'Column{0}' -f ($i + 1) } else { [string]$h }
        [void]$keys.Add($key)
    }

    for ($r = 1; $r -lt $Rows.Count; $r++) {
        $dataRow = $Rows[$r]
        $obj = [ordered]@{}
        for ($i = 0; $i -lt $keys.Count; $i++) {
            $obj[$keys[$i]] = if ($i -lt $dataRow.Count) { $dataRow[$i] } else { $null }
        }
        [void]$objects.Add([PSCustomObject]$obj)
    }

    return , $objects.ToArray()
}

function ConvertTo-XlsCsvField {
    <#
    .SYNOPSIS
        ConvertFrom-XlsValue2 済みの 1 値を CSV フィールド文字列にする（引用・エスケープ含む）。
        COM は触らない純粋関数。Get-XlsRange の -AsCsv 用内部ヘルパー（T-06）。Export しない。
    .PARAMETER Value
        変換済みの値（$null/[string]/[double]/[bool] のいずれか）。
    .EXAMPLE
        ConvertTo-XlsCsvField -Value 3.14
    .NOTES
        [double] は不変カルチャ・"G15"（Excel の倍精度浮動小数点が実用上持つ有効桁数相当）で
        文字列化する。実行環境のカルチャ（10 進区切りなど）に CSV の中身が左右されないようにする
        ため（S-01 のファイル I/O エンコーディング明示と同じ趣旨）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $text = if ($Value -is [double]) {
        $Value.ToString('G15', [Globalization.CultureInfo]::InvariantCulture)
    }
    elseif ($Value -is [bool]) {
        if ($Value) { 'True' } else { 'False' }
    }
    else {
        [string]$Value
    }

    if ($text -match '[",\r\n]') {
        return '"' + ($text -replace '"', '""') + '"'
    }

    return $text
}

function Export-XlsRowsToCsv {
    <#
    .SYNOPSIS
        ジャグ配列（行の配列の配列）を UTF-8（BOM 付き）CSV として書き出す。COM は触らない。
        Get-XlsRange の -AsCsv 用内部ヘルパー（T-06）。Export しない。
    .PARAMETER Rows
        書き出す行データ（ConvertFrom-XlsRangeToRows の戻り値と同じ形。既に変換済みの値であること）。
    .PARAMETER Path
        書き出し先パス。
    .EXAMPLE
        Export-XlsRowsToCsv -Rows $rows -Path 'C:\tmp\out.csv'
    .NOTES
        S-01: UTF-8 BOM 付きを明示する（PowerShell 5.1 の既定エンコーディング問題を避けるため、
        Invoke-XlsSession のマーカー書き込みと同じ GetPreamble() + GetBytes() 方式に揃えた）。
        改行は CRLF（RFC 4180、Excel 自身が生成する CSV と同じ）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($row in $Rows) {
        $fields = New-Object System.Collections.ArrayList
        foreach ($value in $row) {
            [void]$fields.Add((ConvertTo-XlsCsvField -Value $value))
        }
        [void]$lines.Add(($fields.ToArray() -join ','))
    }

    $text = ''
    if ($lines.Count -gt 0) {
        $text = ($lines.ToArray() -join "`r`n") + "`r`n"
    }

    $bytes = [Text.Encoding]::UTF8.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($text)
    try {
        [IO.File]::WriteAllBytes($Path, $bytes)
    }
    catch {
        throw "Failed to write CSV to '$Path'. Check that the parent directory exists and is writable. Underlying error: $($_.Exception.Message)"
    }
}

function Export-XlsJson {
    <#
    .SYNOPSIS
        任意のデータを UTF-8（BOM 付き）JSON として書き出す。COM は触らない。
        Get-XlsRange の -AsJson 用内部ヘルパー（T-06）。Export しない。
    .PARAMETER Data
        書き出すデータ（配列でもオブジェクトでも可）。
    .PARAMETER Path
        書き出し先パス。
    .EXAMPLE
        Export-XlsJson -Data $rows -Path 'C:\tmp\out.json'
    .NOTES
        罠: `ConvertTo-Json` はパイプライン経由でデータを渡すと、要素数 1 の配列を自動的に中身だけに
        アンラップしてしまう（PowerShell 全般の既知の挙動）。ここでは必ず `-InputObject`（パイプでは
        なく名前付き引数）で渡すことで、配列は要素数によらず配列のまま JSON 化されることを実機で
        確認した（実装メモ参照）。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Data,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $json = ConvertTo-Json -InputObject $Data -Depth 10

    $bytes = [Text.Encoding]::UTF8.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($json)
    try {
        [IO.File]::WriteAllBytes($Path, $bytes)
    }
    catch {
        throw "Failed to write JSON to '$Path'. Check that the parent directory exists and is writable. Underlying error: $($_.Exception.Message)"
    }
}

function Get-XlsRange {
    <#
    .SYNOPSIS
        Range.Value2 を一括取得し、日付・空セル・エラー値を変換したジャグ配列（既定）または
        ヘッダー付きオブジェクト配列（-Header）として返す。CSV/JSON への書き出しにも対応する
        （セル単位ループの代替。01-design.md §3.5）。
    .PARAMETER Worksheet
        対象の Worksheet COM オブジェクト（Invoke-XlsSession の ScriptBlock 内で得たもの。G-06）。
    .PARAMETER Range
        取得する範囲（A1 形式。例 'A1:C10'。単一セルも可）。
    .PARAMETER AsCsv
        指定した場合、結果を UTF-8（BOM 付き）CSV として書き出すパス。
    .PARAMETER AsJson
        指定した場合、結果を UTF-8（BOM 付き）JSON として書き出すパス。
    .PARAMETER Header
        1 行目をヘッダーとして扱う。戻り値が PSCustomObject の配列（ヘッダー名をプロパティ名とする）
        になり、-AsJson の出力もオブジェクト配列になる（実装判断、実装メモ参照）。
    .EXAMPLE
        Get-XlsRange -Worksheet $ws -Range 'A1:C10'
    .EXAMPLE
        Get-XlsRange -Worksheet $ws -Range 'A1:C10' -Header -AsJson 'C:\tmp\out.json'
    .NOTES
        実装判断（詳細は実装メモ）:
          - 既定の戻り値はジャグ配列（行の配列の配列）であり [object[,]] ではない。PowerShell
            ネイティブの配列にすることで、エージェントが `$rows[0][1]` のように素直に扱え、
            `ConvertTo-Json` への受け渡しも自然になる。
          - -Header は戻り値と -AsJson の出力に一貫して反映する（1 行目を読み飛ばしてヘッダー名を
            キーにしたオブジェクト配列にする）。-AsCsv の出力バイト列は -Header の有無で変わらない
            （CSV は元々「1 行目がヘッダーかどうか」を区別しないフォーマットで、ジャグ配列には
            ヘッダー行を含む全行がそのまま入っているため、-Header の有無にかかわらず同じ内容が
            書き出される。取り違えではなく意図した挙動）。
          - 空セルは $null、エラー値は "#REF!" 等の表示文字列、日付は ISO 8601 文字列に変換する
            （変換規則は ConvertFrom-XlsValue2 に集約。T-07/T-08 が再利用する想定）。
          - 単一セル範囲は Value2 がスカラーで返る既知の COM の罠を吸収し、1 行 1 列のジャグ配列にする。
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

    try {
        $rangeObj = $Worksheet.Range($Range)
    }
    catch {
        throw "Invalid range address '$Range' on worksheet '$($Worksheet.Name)'. Check the A1-style address (e.g. 'A1' or 'A1:C10'). Excel reported: $($_.Exception.Message)"
    }

    # round 2 レビュー指摘（blocking）対応: Workbook.Date1904 を 1 回だけ読み、日付変換に使う
    # （Worksheet.Parent が所属 Workbook を返すことを実機で確認済み。実装メモ参照）。
    $date1904 = [bool]$Worksheet.Parent.Date1904

    $rows = ConvertFrom-XlsRangeToRows -RangeObj $rangeObj -Date1904:$date1904

    # 罠（T-06 で発見、実機確認。実装メモ参照）: `$result = if (...) {...} else {...}`（if/else を
    # 式として使い代入する形）は、選ばれた枝の値が「要素数 1 の配列」で、かつその配列が既に 1 回
    # 関数の `return ,` 境界を通過済みの値（今回は $rows。ConvertFrom-XlsRangeToRows からの戻り値）
    # だった場合、もう 1 段階多く展開してしまい中身を剥がしてしまうことを実機で確認した
    # （同じ配列をインラインでその場で作った場合は再現しない。ローカル変数への代入元がどこかで
    # 関数境界をまたいだかどうかで挙動が変わる、PowerShell の出力ストリーム展開の紛らわしい罠）。
    # if/else を「式」ではなく「文」として使い、各枝の中で直接 `$result` に代入する形に変えることで
    # この追加展開を避ける。
    if ($Header) {
        $result = ConvertTo-XlsHeaderObjects -Rows $rows
    }
    else {
        $result = $rows
    }

    if ($AsCsv) {
        Export-XlsRowsToCsv -Rows $rows -Path $AsCsv
    }

    if ($AsJson) {
        Export-XlsJson -Data $result -Path $AsJson
    }

    return , $result
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
