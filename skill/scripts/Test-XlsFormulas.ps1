#Requires -Version 5.1
<#
.SYNOPSIS
    recalc.py 互換の CLI ラッパー。同じフォルダの XlsAgent.psm1 の Test-XlsFormulas を呼び出し、
    stdout に JSON を 1 個だけ出力する（01-design.md §3.6、G-04）。
.DESCRIPTION
    位置引数 <path> [timeoutSec=30]、スイッチ -Force を受け取る。
    検証を STA の別 Runspace 上で非同期実行する。Invoke-XlsSession が起動した Excel の所有権が
    実行ごとに一意なトークンで確立されたと確認できるまでは timeoutSec の対象外として無期限に待ち
    （打ち切ると Clear-XlsOrphans が絶対に追跡できない孤児 EXCEL.EXE を残すため。round 1/round 2
    レビュー blocking 対応、詳細は .NOTES）、所有権が確立してから timeoutSec 秒待っても完了しなければ
    { "error": "timeout" } を出力して非 0 で終了する。
.PARAMETER Path
    検証対象ワークブックのパス（位置引数 0）。
.PARAMETER TimeoutSec
    タイムアウト秒数（位置引数 1、既定 30）。Excel の起動〜所有権確立までは含まない（.NOTES 参照）。
.PARAMETER Force
    外部リンク refused 判定をスキップする（XlsAgent.psm1 の Test-XlsFormulas -Force と同じ意味）。
.EXAMPLE
    powershell -File Test-XlsFormulas.ps1 C:\tmp\book.xlsx
.EXAMPLE
    powershell -File Test-XlsFormulas.ps1 C:\tmp\book.xlsx 5 -Force
.NOTES
    exit code 契約（G-04、recalc.py 互換）: 出力 JSON に "error" キーがあるときだけ非 0
    （ファイル不在・Excel 起動失敗・タイムアウト）。"errors_found"/"refused" は正常系として exit 0。

    ラウンド 1・ラウンド 2 の Codex レビュー（harness/state/reviews/T-12.md,
    harness/state/reviews/T-12-r2.md）で計 4 件の blocking 指摘を受けて設計を変更した。対応内容は
    harness/state/tasks/T-12.md の実装メモに詳しく書いたが、要点のみここにも残す:

    **所有権ハンドシェイク（round 1 blocking #1・round 2 blocking 対応）**: round 1 では
    「`$env:TEMP\xlsagent` に新しいマーカーファイルが現れたら所有権確立」と判定していたが、
    round 2 レビューで 3 つの穴を指摘された:
      (a) 並行する別セッションが作った無関係な新しいマーカーでも「新しい」というだけでタイマーが
          開始されてしまう（このセッションのものだと相関確認していない）。
      (b) `Invoke-XlsSession` のマーカーは `FileMode.Create` で**可視化**した後に内容を書いて
          `Flush` するため、「存在する」時点ではまだ内容（PID・開始時刻）が空/不完全な競合窓がある。
          その状態で打ち切ると、開始時刻のない旧形式扱いになり `Clear-XlsOrphans` が
          `Skipped-UnverifiedLegacy`（停止しない）にしてしまい、実質回収不能になる。
      (c) PID 再利用で自分がこれから使う PID のマーカー名が「事前に存在した名前一覧」に既に
          含まれていると、「新しい」と判定できず今度は逆に timeoutSec が永久に始まらない。
    (a)(c) の根本原因はディレクトリのファイル名差分だけで相関を取っていたこと、(b) は
    「マーカー本体は `Clear-XlsOrphans` の排他判定のため `FileShare.None` で保持し続ける必要があり、
    CLI 側からは（進行中は必ず共有違反になるので）内容を読めない」という設計上の制約に起因する。

    修正（実行ごとに一意な runToken による ready ファイル方式）: CLI は起動のたびに
    `[guid]::NewGuid()` で `runToken` を生成し、`$env:XLSAGENT_RUN_TOKEN` としてプロセス環境変数に
    設定してから Runspace を起動する（同一プロセス内の別スレッドなので環境変数はそのまま見える）。
    `Invoke-XlsSession`（`skill/scripts/XlsAgent.psm1`）は、この環境変数が設定されているときだけ、
    マーカー本体の Write+Flush が完了した**直後**に、ロックしない別ファイル
    `$env:TEMP\xlsagent-ready\<runToken>.ready`（PID・開始時刻・runToken の 3 行）を書く
    （公開シグネチャ・既定動作は一切変更しない。環境変数未設定なら何もしない）。
    CLI 側はマーカーディレクトリではなく、**このファイル名（=自分の runToken）を厳密一致で**
    ポーリングする。存在するだけでなく内容（PID・開始時刻がパースできる・3 行目が自分の runToken と
    一致する）まで検証できて初めて所有権確立と判定する。これにより:
      (a) 対応: 無関係なセッションの ready ファイルは別の runToken（別ファイル名）になるため、
          絶対に自分のものと混同しない。
      (b) 対応: ready ファイルは Invoke-XlsSession 側で PID・開始時刻・runToken を組み立ててから
          1 回の `WriteAllText` で書く（マーカー本体の Write+Flush の**後**）。内容が不完全な
          ready ファイルを CLI 側で見つけても（3 行未満・PID/開始時刻がパース不能・runToken 不一致）
          「未確立」として扱いポーリングを続け、絶対に確立とはみなさない。
      (c) 対応: ファイル名がその実行専用の runToken（GUID）なので、PID 再利用や「事前一覧」との
          突き合わせという概念自体が不要になった。
    所有権確立前（＝ready ファイルがまだ現れない、または内容が検証を通らない間）は無期限に待ち、
    ここで強制終了しないことが安全性の核心（この開発機では Excel の COM アクティブ化自体に実測
    4〜61 秒の幅があることを確認している）。所有権確立が確認できてから初めて
    `timeoutSec` 秒の `AsyncWaitHandle.WaitOne` を開始する。

    **PID 再利用への耐性（round 3 blocking 対応）**: round 2 版は ready ファイルの StartTicks を
    `TryParse` するだけで、実際に生きているプロセスの開始時刻と突き合わせていなかった。これでは
    「たまたま同じ PID の何らかのプロセスが生きているだけ」でも所有権確立と誤認しうる（PID
    再利用への耐性がない）。修正: `Get-Process -Id $readyPid` で取得したプロセスを `Refresh()` した
    上で、`HasExited` が false かつ `StartTime.ToUniversalTime().Ticks` が ready ファイルに書かれた
    値と完全一致する場合にのみ所有権確立とみなす（`Clear-XlsOrphans` が孤児回収時に行っている
    PID+開始時刻の完全一致照合と同じ考え方）。

    **TerminateProcess は timeout 経路専用（round 1 should-fix・round 2 should-fix 対応）**:
    `PowerShell.EndInvoke()` が正常に返る（成功・エラーいずれも）経路や、Runspace 生成自体が
    失敗する最外周 catch では、パイプラインは既に完了しているか一度も開始していないため、通常の
    `exit $Code` で十分かつより安全（CLI を誤って既存のインタラクティブホスト内から dot-source 等で
    実行した場合でもホスト全体を巻き込まない）。round 1 では最外周 catch にも `-HardExit` を
    付けていたが、round 2 レビューで「この経路は内側の finally で Dispose/Close 済みなので不要」と
    指摘され、通常の `exit` に戻した。`TerminateProcess`（kernel32.dll、P/Invoke）は、所有権確立後に
    timeoutSec を超えて打ち切る場合（＝バックグラウンドの STA スレッドがまだ COM 呼び出しで
    ブロックしている可能性がある場合）にのみ使う。`[Environment]::Exit()` ではなく
    `TerminateProcess` を選んだ理由: 実機確認で `[Environment]::Exit()` はブロック中のスレッドが
    あるとすぐにはプロセスを終了させず（`timeoutSec=1` のテストで実際のプロセス終了に数十秒以上
    かかるケースを確認）、`TerminateProcess` は OS レベルで即座に終了させることを実機で確認した
    （`timeoutSec=2` で実測 2498ms など、指定値にほぼ一致する壁時計時間で終了する）。

    打ち切り後の EXCEL.EXE 回収（実装判断、G-10 順守）: CLI プロセス自身は `Clear-XlsOrphans` を
    呼ばない。呼んでも安全に何もできないため（`Invoke-XlsSession` のマーカーは `FileShare.None` の
    排他リースとして、まだ生きている可能性のある別スレッドが保持し続けているので、`Clear-XlsOrphans`
    は必ず `Skipped-ActiveSession` として何もせずスキップする — G-10 の「進行中セッションには触れない」
    という安全設計そのもの）。CLI プロセスが `TerminateProcess` で終了すると、OS がそのプロセスの
    全ハンドル（マーカーの排他リースを含む）を即座に解放するため、その後に**別プロセスから**
    `Clear-XlsOrphans` を呼べば、真の孤児として検出・回収できる（`Invoke-XlsSession` の finally が
    実行されないため、マーカーファイルもそのまま残る。上記の所有権ハンドシェイクにより、
    timeout で `TerminateProcess` が呼ばれる時点では必ずマーカーの内容が完全に書き終わっている
    ことが保証されている）。呼び出し元（テスト・エージェント）の責務とする。

    `$env:XLSAGENT_CLI_TEST_MODULE_PATH`（テスト専用の非公開フック）: 設定されていれば
    `XlsAgent.psm1` の代わりにそのパスのモジュールを読み込む。公開 CLI 引数（<path>/timeoutSec/
    -Force）の契約には含まれない。実 Excel の起動時間に左右されずタイムアウト機構そのものを
    決定的に検証するテスト（`tests/TestXlsFormulasCli.Tests.ps1`）でのみ使う。

    `$env:XLSAGENT_RUN_TOKEN`（本 CLI が自分で設定する内部フック）: `Invoke-XlsSession` 側の
    ready ファイル書き込みを有効にするためのもの。上記参照。CLI は実行ごとに新しい値を生成して
    設定し、終了前に元の値へ復元する（入れ子実行を壊さないため）。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Position = 1)]
    [int]$TimeoutSec = 30,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ProgressPreference = 'SilentlyContinue'

if (-not ('XlsAgentCli.NativeMethods' -as [type])) {
    # 実機確認（ファイル冒頭 .NOTES 参照）: 所有権確立後にタイムアウトで打ち切る場合、
    # [Environment]::Exit() はブロック中の Runspace スレッドがあるとすぐにはプロセスを終了させない
    # ことがあるため、OS レベルで即座に終了させる TerminateProcess を直接呼ぶ。XlsAgent.psm1 の
    # NativeMethods（GetWindowThreadProcessId）とは別名前空間にして、同一プロセス内で両方の
    # Add-Type が読み込まれても衝突しないようにする。
    Add-Type -Namespace XlsAgentCli -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetCurrentProcess();

[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
public static extern bool TerminateProcess(System.IntPtr hProcess, uint uExitCode);
'@
}

function Complete-XlsCli {
    <#
    .SYNOPSIS
        JSON を stdout に 1 個だけ書き、フラッシュしてからプロセスを終了する（内部ヘルパー）。
    .PARAMETER HardExit
        指定した場合のみ Win32 API TerminateProcess で即座に終了する。所有権確立後の timeout
        経路（バックグラウンドの STA スレッドがまだ COM 呼び出しでブロックしている可能性がある
        場合）専用（round 1 レビュー should-fix、round 2 レビュー should-fix で適用範囲をさらに
        限定）。それ以外の経路ではパイプラインが既に完了・後始末済み（または一度も開始していない）
        なので通常の `exit` で十分かつより安全。
    .NOTES
        PowerShell の既定の出力フォーマッタ（Out-Default 経由）を経由すると、非対話ホストでの
        コンソール幅推定によって長い文字列に意図しない改行が混入しうるため、必ず [Console]::Out に
        直接書く（「stdout に JSON を 1 個だけ」という契約を確実に守るため）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Code,

        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [switch]$HardExit
    )

    [Console]::Out.WriteLine($JsonText)
    try { [Console]::Out.Flush() } catch { }

    if ($HardExit) {
        try {
            $handle = [XlsAgentCli.NativeMethods]::GetCurrentProcess()
            [void][XlsAgentCli.NativeMethods]::TerminateProcess($handle, [uint32]$Code)
        }
        catch {
            # TerminateProcess の P/Invoke 自体が失敗するような極端なケースへのフォールバック。
        }
        # TerminateProcess が成功すればここには到達しない（プロセスは既に終了している）。
    }

    exit $Code
}

function ConvertTo-XlsCliErrorJson {
    <#
    .SYNOPSIS
        { "error": "<message>" } 形式の JSON 文字列を作る（recalc.py 互換、内部ヘルパー）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    return ConvertTo-Json -Compress -InputObject ([ordered]@{ error = $Message })
}

function Test-XlsCliOwnershipReady {
    <#
    .SYNOPSIS
        所有権ハンドシェイクの ready ファイル（`Invoke-XlsSession` が書く）を読み、内容が完全に
        検証できたときだけ PID・開始時刻を返す（内部ヘルパー、round 2 レビュー blocking 対応）。
        検証できなければ $null を返す（＝まだ所有権確立とはみなさない。呼び出し側はポーリングを
        続ける）。
    .PARAMETER ReadyPath
        `$env:TEMP\xlsagent-ready\<runToken>.ready` のフルパス。
    .PARAMETER ExpectedRunToken
        この CLI 実行が生成した runToken。ready ファイルの 3 行目と厳密一致する場合のみ有効。
    .NOTES
        ファイル名自体が runToken なので他プロセスのものと衝突しないが、念のため内容の 3 行目
        （runToken）も突き合わせる（多重の防御）。round 3 レビュー blocking 対応: PID が
        「生存していること」だけでは PID 再利用（別プロセスが偶然同じ PID を再利用している）に
        耐えられないため、記録された開始時刻（StartTicks）を `Get-Process` が返す実プロセスの
        `StartTime.ToUniversalTime().Ticks` と完全一致するまで確認する（`Clear-XlsOrphans` 自身が
        孤児回収時に行っているのと同じ照合パターン）。EXCEL.EXE であることまでは要求しない。
        理由: 実 Excel を使わない決定的なテスト（`tests/TestXlsFormulasCli.Tests.ps1` の
        フェイクモジュール）では、このフィールドにフェイクモジュール自身の（EXCEL ではない）PID を
        書くため。EXCEL.EXE かどうかの厳密な確認は `Clear-XlsOrphans` 自身が回収時に独立して行う。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReadyPath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedRunToken
    )

    if (-not (Test-Path -LiteralPath $ReadyPath -PathType Leaf)) {
        return $null
    }

    try {
        $text = Get-Content -LiteralPath $ReadyPath -Raw -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # 書き込みの最中でロックされている等、一時的に読めないだけかもしれない。次のポーリングで再試行。
        return $null
    }

    if ([string]::IsNullOrEmpty($text)) {
        return $null
    }

    $lines = @($text -split "`r`n|`n" | Where-Object { $_ -ne '' })
    if ($lines.Count -lt 3) {
        return $null
    }

    [int]$readyPid = 0
    if (-not [int]::TryParse($lines[0], [ref]$readyPid) -or $readyPid -le 0) {
        return $null
    }

    [long]$readyStartTicks = 0
    if (-not [long]::TryParse($lines[1], [ref]$readyStartTicks) -or
        $readyStartTicks -le 0) {
        return $null
    }

    if ($lines[2] -ne $ExpectedRunToken) {
        return $null
    }

    $proc = Get-Process -Id $readyPid -ErrorAction SilentlyContinue
    if (-not $proc) {
        return $null
    }

    # round 3 レビュー blocking 対応: PID の生存確認だけでは PID 再利用に耐えられない。
    # ready ファイルに記録された開始時刻と、いま生きているプロセスの実際の開始時刻を完全一致で
    # 照合して初めて「このプロセスが ready ファイルの書き手そのものである」と確認できる
    # （Clear-XlsOrphans の所有権確認と同じ考え方）。
    try {
        $proc.Refresh()
        if ($proc.HasExited -or
            $proc.StartTime.ToUniversalTime().Ticks -ne $readyStartTicks) {
            return $null
        }
    }
    catch {
        return $null
    }

    return [PSCustomObject]@{ Pid = $readyPid; StartTicks = $readyStartTicks }
}

try {
    try {
        # 実コンソールを持たないリダイレクト済み標準出力では例外になることがある（実機確認）。
        # 出力する文字列は現状すべて ASCII のため、失敗しても実害はないベストエフォート。
        [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    }
    catch { }

    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    # $env:XLSAGENT_CLI_TEST_MODULE_PATH: テスト専用の非公開フック（ファイル冒頭 .NOTES 参照）。
    # 公開引数の契約には含まれない。
    $modulePath = if ($env:XLSAGENT_CLI_TEST_MODULE_PATH) { $env:XLSAGENT_CLI_TEST_MODULE_PATH } else { Join-Path $scriptDir 'XlsAgent.psm1' }

    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        Complete-XlsCli -Code 1 -JsonText (ConvertTo-XlsCliErrorJson -Message "Module not found: '$modulePath'")
    }

    $waitMs = [Math]::Max(0, $TimeoutSec) * 1000
    $forceFlag = [bool]$Force

    # 所有権ハンドシェイク（round 2 レビュー blocking 対応、ファイル冒頭 .NOTES 参照）: 実行ごとに
    # 一意な runToken を生成し、Invoke-XlsSession（$env:XLSAGENT_RUN_TOKEN 経由）に ready ファイルを
    # 書かせる。CLI はこのファイル名（=runToken）を厳密一致でポーリングし、内容まで検証できて
    # 初めて所有権確立とみなす。
    $runToken = [guid]::NewGuid().ToString('N')
    $readyDir = Join-Path $env:TEMP 'xlsagent-ready'
    if (-not (Test-Path -LiteralPath $readyDir)) {
        try { New-Item -ItemType Directory -Path $readyDir -Force | Out-Null } catch { }
    }
    $readyPath = Join-Path $readyDir ("{0}.ready" -f $runToken)

    $previousRunToken = $env:XLSAGENT_RUN_TOKEN
    $env:XLSAGENT_RUN_TOKEN = $runToken

    $runspace = $null
    $ps = $null

    $timedOut = $false
    $errorMessage = $null
    $resultJson = $null

    try {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
            # 01-design.md §3.6 / タスクカード指定: Invoke-XlsSession は STA 必須。検証全体をこの
            # Runspace の上で動かすことで、timeoutSec による打ち切りが可能になる（メインスレッドで
            # 直接呼ぶと COM 呼び出しの完了を待つしかなく、外側から打ち切れない）。
            $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
            # ReuseThread: COM オブジェクト（Excel.Application）はアパートメント（スレッド）に
            # 紐付くため、同一 Runspace 内の呼び出しは常に同じ専用スレッドで実行される必要がある。
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $runspace.Open()

            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $runspace

            [void]$ps.AddScript({
                param($ModulePath, $WbPath, $ForceFlag)
                Import-Module -Name $ModulePath -Force -ErrorAction Stop
                Test-XlsFormulas -Path $WbPath -Force:$ForceFlag -AsJson
            }).AddArgument($modulePath).AddArgument($Path).AddArgument($forceFlag)

            $asyncResult = $ps.BeginInvoke()

            # フェーズ 1（所有権確立待ち、無期限、timeoutSec の対象外。round 2 blocking 対応）:
            # 自分の runToken に対応する ready ファイルの内容が検証できるようになるか、パイプラインが
            # （Excel 起動失敗などで ready ファイルを一度も書かずに）自然完了するまで待つ。この区間で
            # 強制終了しないことが今回の修正の核心。
            $ownership = $null
            while (-not $asyncResult.IsCompleted) {
                $ownership = Test-XlsCliOwnershipReady -ReadyPath $readyPath -ExpectedRunToken $runToken
                if ($null -ne $ownership) {
                    break
                }
                Start-Sleep -Milliseconds 100
            }
            if ($null -eq $ownership) {
                # ループを抜けた理由がパイプライン完了だった場合に備え、最後にもう一度だけ確認する
                # （ready ファイルの書き込みとパイプライン完了がほぼ同時に起きた場合の取りこぼし防止）。
                $ownership = Test-XlsCliOwnershipReady -ReadyPath $readyPath -ExpectedRunToken $runToken
            }

            # フェーズ 2（所有権確立後だけ timeoutSec を測る）:
            if ($null -ne $ownership -and -not $asyncResult.IsCompleted) {
                $completed = $asyncResult.AsyncWaitHandle.WaitOne($waitMs)
            }
            else {
                # 所有権が確立する前にパイプラインが完了した（例: Excel 起動失敗、モジュール読み込み
                # 失敗）。ready ファイルが一度も有効化されていないのでタイムアウトの対象外とし、
                # 完了扱いにする。
                $completed = $true
            }

            if (-not $completed) {
                $timedOut = $true
                # 打ち切り要求は投げっぱなし（完了を待たない）。COM のブロッキング呼び出し中は
                # BeginStop が即座には効かない可能性が高い（ファイル冒頭 .NOTES 参照）。この時点では
                # 所有権ハンドシェイクによりマーカーの内容が完全に書き終わっていることが保証されて
                # いるため、TerminateProcess で打ち切っても Clear-XlsOrphans が後から安全に回収できる。
                try { $ps.BeginStop($null, $null) | Out-Null } catch { }
            }
            else {
                $output = $null
                try {
                    $output = $ps.EndInvoke($asyncResult)
                }
                catch {
                    # PowerShell が .NET メソッド呼び出し（EndInvoke）の例外を
                    # MethodInvocationException でラップし、メッセージが
                    # 'Exception calling "EndInvoke" with "1" argument(s): "<本来のメッセージ>"' に
                    # なることを実機で確認した。次の一手が分かる本来のメッセージ（S-04）を保つため、
                    # InnerException があればそちらを使う。
                    if ($null -ne $_.Exception.InnerException) {
                        $errorMessage = $_.Exception.InnerException.Message
                    }
                    else {
                        $errorMessage = $_.Exception.Message
                    }
                }

                if (-not $errorMessage) {
                    if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) {
                        $errorMessage = $ps.Streams.Error[0].Exception.Message
                    }
                    elseif ($null -eq $output -or @($output).Count -eq 0) {
                        $errorMessage = 'Test-XlsFormulas produced no output.'
                    }
                    else {
                        $resultJson = [string](@($output) | Select-Object -Last 1)
                    }
                }
            }
        }
        finally {
            if ($null -ne $ps) {
                if (-not $timedOut) {
                    try { $ps.Dispose() } catch { }
                }
                # タイムアウト時は Dispose しない: パイプラインがまだ COM 呼び出しでブロックしている
                # 可能性があり、Dispose がその完了待ちでブロックしうるため（本関数の「即座に打ち切る」
                # という契約を壊す）。プロセスは直後に TerminateProcess で終了するので、ハンドルの
                # 後始末は OS に任せる。
            }
            if ($null -ne $runspace -and -not $timedOut) {
                try { $runspace.Close() } catch { }
                try { $runspace.Dispose() } catch { }
            }
            # ready ファイルは CLI 内部のハンドシェイク専用の使い捨てアーティファクトであり、
            # Clear-XlsOrphans はこれを一切参照しない（削除してもしなくても安全性・回収可能性には
            # 影響しない）。round 3 レビュー nit 対応: この finally は（タイムアウト時に呼ばれる
            # TerminateProcess を含む）Complete-XlsCli の呼び出しより**前**に必ず実行されるため、
            # $timedOut かどうかによらずベストエフォートで削除し、$env:TEMP への蓄積を防ぐ。
            Remove-Item -LiteralPath $readyPath -Force -ErrorAction SilentlyContinue
            if ($null -ne $previousRunToken) {
                $env:XLSAGENT_RUN_TOKEN = $previousRunToken
            }
            else {
                Remove-Item Env:\XLSAGENT_RUN_TOKEN -ErrorAction SilentlyContinue
            }
        }

    if ($timedOut) {
        # round 1 レビュー should-fix 対応: TerminateProcess は所有権確立後の timeout 経路専用。
        Complete-XlsCli -Code 1 -JsonText (ConvertTo-XlsCliErrorJson -Message 'timeout') -HardExit
    }

    if ($errorMessage) {
        Complete-XlsCli -Code 1 -JsonText (ConvertTo-XlsCliErrorJson -Message $errorMessage)
    }

    Complete-XlsCli -Code 0 -JsonText $resultJson
}
catch {
    # 上記のいずれの分岐にも当てはまらない、想定外の失敗（引数解析後の準備コード自体の失敗等）。
    # 次の一手が分かるメッセージのまま error キーへ格納する（S-04）。round 2 レビュー should-fix
    # 対応: この経路はパイプラインを一度も開始していないか、内側の finally で Dispose/Close 済み
    # の状態でしか到達しないため、通常の exit で十分（-HardExit は使わない）。
    try {
        Complete-XlsCli -Code 1 -JsonText (ConvertTo-XlsCliErrorJson -Message $_.Exception.Message)
    }
    catch {
        # ConvertTo-Json 自体が失敗するような極端なケースのための最後の砦。
        [Console]::Out.WriteLine('{"error":"unexpected CLI failure"}')
        exit 1
    }
}
