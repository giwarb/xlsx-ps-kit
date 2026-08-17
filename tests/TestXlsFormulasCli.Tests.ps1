#Requires -Version 5.1

<#
tests/TestXlsFormulasCli.Tests.ps1

T-12: skill/scripts/Test-XlsFormulas.ps1（recalc.py 互換の CLI ラッパー）の受け入れテスト。
harness/state/tasks/T-12.md の受け入れテスト要点、および harness/state/reviews/T-12.md
（round 1、判定 fix）・harness/state/reviews/T-12-r2.md（round 2、判定 fix）・
harness/state/reviews/T-12-r3.md（round 3、判定 fix）の指摘対応を反映している。

## round 1 レビュー対応の要点（詳細は harness/state/tasks/T-12.md の実装メモ）
- [blocking #1] CLI 側は「Invoke-XlsSession の所有権が確立するまでは timeoutSec の対象外」に
  設計変更した（skill/scripts/Test-XlsFormulas.ps1 参照。round 2 でさらに強化、下記参照）。
  このテストファイルはその変更を前提に、タイムアウト機構そのものを 2 系統に分けて検証する:
    1. `Context 'timeout mechanism (deterministic, no real Excel)'`: 実 Excel を使わない
       フェイクモジュールで、CLI のタイムアウト機構（所有権確立の検知 → timeoutSec の
       カウントダウン → 打ち切り）を実行環境の速度に依存せず決定的に検証する。
    2. `Context 'timeout with real Excel (post-ownership)'`: 実 Excel で「所有権確立後の区間
       （Workbooks.Open〜CalculateFullRebuild〜走査）」がタイムアウトする実地検証を 1 本だけ持つ。
  この分割は harness/state/tasks/T-12.md の「重いブックの作成が不安定なら、タイムアウト機構の
  単体検証で代替してよい」という判断と、round 1 レビュー should-fix「タイムアウトテストが実行環境の
  遅さに依存している」への対応。
- [blocking #2] タイムアウトテストの後始末は `Clear-XlsOrphans` のポーリングのみに限定した
  （直接 `Stop-Process -Id <ベースライン差分 PID>` は削除。G-10「PID の前後差分は所有権の証明に
  ならない」を再導入していたため）。未追跡の PID が残ればテストを失敗させ、CLI のライフサイクル
  不具合を隠さない。round 2 レビューでこの対応は確認済み（解消済みと認定）。
- [should-fix] `Start-Process -ArgumentList` は配列要素を空白で連結するだけで各要素の境界を
  保護しないため（Windows PowerShell 5.1 の既知の挙動）、CLI パスとスペースを含む引数を明示的に
  二重引用符で囲む。スペースを含む一時ディレクトリを使う受け入れテストも追加した。

## round 2 レビュー対応の要点（詳細は harness/state/tasks/T-12.md の実装メモ）
- [blocking] round 1 の「新しいマーカーファイルが現れたら所有権確立」という判定は、(a) 無関係な
  並行セッションのマーカーと自分のものを区別できない、(b) マーカーの可視化と内容の Write+Flush の
  間に競合窓がある、(c) PID 再利用で自分のマーカー名が事前一覧に含まれていると検出できない、という
  3 つの穴があった。修正: CLI が実行のたびに一意な runToken を生成して `$env:XLSAGENT_RUN_TOKEN`
  に設定し、`Invoke-XlsSession`（`skill/scripts/XlsAgent.psm1`、この環境変数が設定されている
  ときだけ動作する非公開フック）がマーカー本体の Write+Flush 完了後にロックしない別ファイル
  （PID・開始時刻・runToken の 3 行）を書く「ready ファイル」方式に変更した。CLI はこのファイル名
  （=自分の runToken）を厳密一致でポーリングし、内容まで検証できて初めて所有権確立と判定する。
  `New-XlsCliFakeReadyModule`（本ファイル）はこの ready ファイル書式を模倣し、(a) デコイの
  無関係な ready ファイル・マーカーが先に現れても惑わされない、(b) 本物の ready ファイルが
  書かれるまで（意図的に遅延させても）強制終了しない、の 2 点を決定的に固定するテストを追加した。
- [should-fix] 最外周 catch の `-HardExit` を外し、通常の `exit 1` に戻した（この経路は内側の
  finally で Dispose/Close 済み、またはパイプラインを一度も開始していない状態でしか到達しない
  ため）。テスト側の変更は不要（CLI 内部の実装のみ）。
- [should-fix] 実 Excel のタイムアウトテストを、Excel PID の有無にかかわらず `Clear-XlsOrphans`
  を最低 1 回呼び、戻り値の Action・新規マーカー・追加 EXCEL.EXE が最終的にゼロであることの
  両方を検証する形に強化した（実機確認: `BeginStop()` が競合に勝ってパイプラインの finally が
  先に完走し、マーカー・EXCEL.EXE とも自然に片付くケースと、`TerminateProcess` が勝って
  マーカー付きの孤児が残るケースの両方が実際に起こることを確認したため、フェイクモジュールの
  テストも含めてどちらのケースでも安全であることを固定する）。

## round 3 レビュー対応の要点（詳細は harness/state/tasks/T-12.md の実装メモ）
- [blocking] ready ファイルの StartTicks を `TryParse` するだけで、`Get-Process` が返す実際の
  `StartTime.ToUniversalTime().Ticks` と突き合わせていなかったため、PID 再利用（別プロセスが
  たまたま同じ PID で生きているだけ）に耐性がなかった。`skill/scripts/Test-XlsFormulas.ps1` の
  `Test-XlsCliOwnershipReady` を、`Get-Process` で取得したプロセスを `Refresh()` した上で
  `HasExited` と `StartTime` の完全一致を要求するよう修正した。`New-XlsCliFakeReadyModule` に
  `-WrongStartTicksFirst` を追加し、「最初は誤った StartTicks の ready、3 秒後に正しい値へ
  置き換える」ケースで、誤った ready が所有権確立と誤認されず、正しい ready まで待ってから
  timeoutSec が測られることを壁時計時間で決定的に固定した。
- [should-fix] `$clearResultsSeen` を収集するだけで一度も検証していなかった。
  `New-XlsCliFakeReadyModule` に `-WriteExclusiveMarker` を追加して本物と同じ排他マーカーを
  書かせ、CLI 強制終了後の `Clear-XlsOrphans` の結果に `Skipped-DeadPid` が含まれること、
  対応するマーカーファイルが実際に削除されることを明示的に検証するテストを追加した。実 Excel の
  タイムアウトテストにも、新規マーカーが実在したケースに限り、対応する結果に `Stopped` または
  安全側の `Skipped-*` が含まれることを確認する検証を追加した。
- [nit] `skill/scripts/Test-XlsFormulas.ps1` の内側 finally は、タイムアウト時でも ready
  ファイルを削除するように修正した（HardExit の前に必ず実行されるため安全にできる）。
  上記の `-WriteExclusiveMarker` テストで、タイムアウト後に新規の `.ready` が残っていないことも
  確認する。

CLI は「pwsh が無いこの開発機でも Windows PowerShell 5.1 の `powershell.exe -File` で動くこと」が
タスクカードの環境注記。このテスト自体もサブプロセスとして `powershell.exe -File
skill\scripts\Test-XlsFormulas.ps1 ...` を起動して確認する（`Start-Process` の
-RedirectStandardOutput/-RedirectStandardError を使い、`2>&1` によるストリーム混在を避ける）。

G-06: COM は Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中でのみ触る。
本テストファイル自身は CLI をサブプロセスとして起動するだけで、テストプロセス自身は COM を
一切触らない（フィクスチャ作成のみ Invoke-XlsSession 経由）。フェイクモジュールを使うタイムアウト
機構のテストは COM を一切使わない（Invoke-XlsSession のマーカー書式をファイル I/O だけで模倣する）。

G-10: `Clear-XlsOrphans` は自モジュールが書いたマーカーの PID しか対象にしない。本ファイルの
どのテストも「ベースラインとの PID 前後差分」だけを根拠に Excel プロセスを直接 Stop-Process しない
（round 1 レビュー blocking #2 対応）。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'
$cliPath = Join-Path $here '..\skill\scripts\Test-XlsFormulas.ps1'
$fakeModuleDir = Join-Path $env:TEMP 'xlsagent-tests-fake-modules'

function Invoke-XlsFormulasCliProcess {
    <#
    .SYNOPSIS
        skill/scripts/Test-XlsFormulas.ps1 を `powershell.exe -File` のサブプロセスとして起動し、
        stdout/stderr/終了コードをまとめて返す（COM は触らない、内部ヘルパー）。
    .PARAMETER Arguments
        CLI に渡す引数の配列（例: @($path, '1', '-Force')）。
    .PARAMETER ModulePathOverride
        指定した場合、`$env:XLSAGENT_CLI_TEST_MODULE_PATH` を設定してサブプロセスへ引き継ぐ
        （CLI 側の非公開テストフック。skill/scripts/Test-XlsFormulas.ps1 の .NOTES 参照）。
        実 Excel を起動せずタイムアウト機構だけを決定的に検証するために使う。
    .NOTES
        round 1 レビュー should-fix 対応: Windows PowerShell 5.1 の `Start-Process -ArgumentList`
        は配列要素を空白で連結して 1 本のコマンドラインを作るだけで、各要素の境界を保護しない
        （スペースを含むパスがあると引数が分割されてしまう）。CLI パスと、空白を含む引数だけを
        明示的に二重引用符で囲む（Windows のファイル名には `"` を含められないため、この方法で
        安全に区切れる）。
        `Start-Process -RedirectStandardOutput/-RedirectStandardError` を使う（`2>&1` による
        ストリーム混在を避けるため、標準出力と標準エラーを別ファイルへ確実に分離する）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Arguments,

        [string]$ModulePathOverride
    )

    $stdoutPath = Join-Path $env:TEMP ("xlsagent-tests-cli-out-{0}.txt" -f [guid]::NewGuid())
    $stderrPath = Join-Path $env:TEMP ("xlsagent-tests-cli-err-{0}.txt" -f [guid]::NewGuid())

    $quotedCliPath = '"{0}"' -f $cliPath
    $quotedArguments = @($Arguments | ForEach-Object {
        if ($_ -match '\s') { '"{0}"' -f $_ } else { $_ }
    })
    $allArgs = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $quotedCliPath) + $quotedArguments

    $previousOverride = $env:XLSAGENT_CLI_TEST_MODULE_PATH
    if ($ModulePathOverride) {
        $env:XLSAGENT_CLI_TEST_MODULE_PATH = $ModulePathOverride
    }
    else {
        Remove-Item Env:\XLSAGENT_CLI_TEST_MODULE_PATH -ErrorAction SilentlyContinue
    }

    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $allArgs -NoNewWindow -Wait -PassThru `
            -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath

        $stdout = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -Encoding UTF8 } else { '' }
        $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -Encoding UTF8 } else { '' }
        if ($null -eq $stdout) { $stdout = '' }
        if ($null -eq $stderr) { $stderr = '' }

        return [PSCustomObject]@{
            ExitCode = $proc.ExitCode
            Stdout   = $stdout
            Stderr   = $stderr
        }
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
        if ($null -ne $previousOverride) {
            $env:XLSAGENT_CLI_TEST_MODULE_PATH = $previousOverride
        }
        else {
            Remove-Item Env:\XLSAGENT_CLI_TEST_MODULE_PATH -ErrorAction SilentlyContinue
        }
    }
}

function New-TempXlsxPathWithSpace {
    <#
    .SYNOPSIS
        パスにスペースを含む一時 .xlsx パスを発行する（round 1 レビュー should-fix の受け入れ
        テスト用）。ファイル自体は作らない。COM は触らない。
    #>
    [CmdletBinding()]
    param()

    $dir = Join-Path $env:TEMP 'xlsagent tests space'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Join-Path $dir ("{0}.xlsx" -f [guid]::NewGuid())
}

function New-XlsCliFakeReadyModule {
    <#
    .SYNOPSIS
        実 Excel・COM を一切使わず、CLI の所有権ハンドシェイク（ready ファイル方式、round 2/
        round 3 レビュー blocking 対応）と timeoutSec のカウントダウンを決定的に検証するための
        フェイクモジュールを作る。
    .PARAMETER DecoyFirst
        指定した場合、まず無関係な別セッションを装ったデコイの ready ファイル・マーカーファイル
        （別の runToken・別の PID、CLI 自身の runToken とは一致しない）を書いてから
        PreReadyDelaySeconds だけ待つ。CLI がこのデコイに惑わされて所有権確立前にタイマーを
        開始しないことを確認するためのもの（round 2 レビュー blocking (a) の決定的検証）。
    .PARAMETER WrongStartTicksFirst
        指定した場合、まず正しい PID・runToken だが**誤った StartTicks**（生存プロセスの実際の
        開始時刻と一致しない固定値）で ready ファイルを書き、PreReadyDelaySeconds 秒後に正しい
        StartTicks で上書きする。CLI 側が StartTicks を実プロセスの StartTime と突き合わせずに
        受け入れてしまうと（round 3 レビュー blocking の指摘した欠陥）、誤った ready を最初の
        ポーリングで所有権確立と誤認し、timeoutSec 分だけで打ち切られる（＝
        PreReadyDelaySeconds を待たない）。正しく検証していれば PreReadyDelaySeconds 経過後の
        正しい ready まで待ってから timeoutSec を測るため、合計経過時間で判別できる
        （round 3 レビュー blocking の決定的検証）。
    .PARAMETER WriteExclusiveMarker
        指定した場合、`Invoke-XlsSession` の本物のマーカーと同じ書式・同じ排他リース方式
        （`FileShare.None` で開いたまま関数の終了までファイルストリームを保持する）で
        `$env:TEMP\xlsagent\<PID>.marker` を書く。CLI が TerminateProcess で強制終了された後に
        `Clear-XlsOrphans` を呼んだとき、対象 PID に対する具体的な回収結果（`Skipped-DeadPid`
        とマーカー削除）を検証できるようにする（round 3 レビュー should-fix の決定的検証）。
    .PARAMETER PreReadyDelaySeconds
        本物の（正しい）ready ファイルを書くまでの遅延秒数。所有権確立前（ready ファイルの内容が
        まだ検証を通らない間）に強制終了されないこと（round 2 レビュー blocking (b) / round 3
        レビュー blocking の決定的検証）を確認する。
    .PARAMETER PostReadySleepSeconds
        本物の ready ファイルを書いた後、さらに眠る秒数（CLI の timeoutSec より確実に長くする）。
    .NOTES
        `$env:XLSAGENT_RUN_TOKEN`（CLI が実行のたびに生成し、Runspace に引き継がれる環境変数）を
        読み、本物の `Invoke-XlsSession`（`skill/scripts/XlsAgent.psm1`、round 2 レビュー対応で
        追加した ready ファイル書き込みブロック）と同じ書式・書き込み先
        （`$env:TEMP\xlsagent-ready\<runToken>.ready`、PID・開始時刻・runToken の 3 行）を模倣する。
        COM は一切使わない。
    #>
    [CmdletBinding()]
    param(
        [switch]$DecoyFirst,

        [switch]$WrongStartTicksFirst,

        [switch]$WriteExclusiveMarker,

        [int]$PreReadyDelaySeconds = 0,

        [Parameter(Mandatory = $true)]
        [int]$PostReadySleepSeconds
    )

    if (-not (Test-Path -LiteralPath $fakeModuleDir)) {
        New-Item -ItemType Directory -Path $fakeModuleDir -Force | Out-Null
    }
    $path = Join-Path $fakeModuleDir ("FakeReady-{0}.psm1" -f [guid]::NewGuid())

    $template = @'
#Requires -Version 5.1
function Test-XlsFormulas {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$TimeoutSec = 30,
        [switch]$Force,
        [switch]$AsJson
    )

    $readyDir = Join-Path $env:TEMP 'xlsagent-ready'
    if (-not (Test-Path -LiteralPath $readyDir)) {
        New-Item -ItemType Directory -Path $readyDir -Force | Out-Null
    }

    if (__WRITE_EXCLUSIVE_MARKER__) {
        # 本物の Invoke-XlsSession のマーカーと同じ書式・同じ排他リース方式を模倣する
        # （round 3 レビュー should-fix の決定的検証用）。$markerLease は関数の終了まで
        # （＝CLI が TerminateProcess で強制終了するまで）開いたまま保持する。
        $markerDir = Join-Path $env:TEMP 'xlsagent'
        if (-not (Test-Path -LiteralPath $markerDir)) {
            New-Item -ItemType Directory -Path $markerDir -Force | Out-Null
        }
        $selfProcForMarker = Get-Process -Id $PID
        $exclusiveMarkerPath = Join-Path $markerDir ("{0}.marker" -f $PID)
        $markerLease = [IO.File]::Open($exclusiveMarkerPath, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $markerText = "{0}`r`n{1}`r`n" -f $PID, $selfProcForMarker.StartTime.ToUniversalTime().Ticks
        [byte[]]$markerBytes = [Text.Encoding]::UTF8.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($markerText)
        $markerLease.Write($markerBytes, 0, $markerBytes.Length)
        $markerLease.Flush()
    }

    if (__DECOY_FIRST__) {
        # 無関係な別セッションを装ったデコイ（別 runToken・別 PID）。CLI 自身の runToken とは
        # 一致しないため、これだけでは所有権確立と判定されないはず（round 2 blocking (a)）。
        $decoyToken = [guid]::NewGuid().ToString('N')
        $decoyReadyPath = Join-Path $readyDir ("{0}.ready" -f $decoyToken)
        $decoyText = "999999`r`n0`r`n{0}`r`n" -f $decoyToken
        [IO.File]::WriteAllText($decoyReadyPath, $decoyText, (New-Object Text.UTF8Encoding($true)))

        $markerDir2 = Join-Path $env:TEMP 'xlsagent'
        if (-not (Test-Path -LiteralPath $markerDir2)) {
            New-Item -ItemType Directory -Path $markerDir2 -Force | Out-Null
        }
        $decoyMarkerPath = Join-Path $markerDir2 '999998.marker'
        [IO.File]::WriteAllText($decoyMarkerPath, "999998`r`n0`r`n", (New-Object Text.UTF8Encoding($true)))
    }

    $realToken = $env:XLSAGENT_RUN_TOKEN
    $selfProc = Get-Process -Id $PID
    $readyPath = Join-Path $readyDir ("{0}.ready" -f $realToken)

    if (__WRONG_START_TICKS_FIRST__) {
        # 正しい PID・runToken だが誤った StartTicks（実プロセスの開始時刻と一致しない固定値）。
        # round 3 の修正が正しく機能していれば、これだけでは所有権確立と判定されないはず。
        $wrongText = "{0}`r`n123456789`r`n{1}`r`n" -f $PID, $realToken
        [IO.File]::WriteAllText($readyPath, $wrongText, (New-Object Text.UTF8Encoding($true)))
    }

    Start-Sleep -Seconds __PRE_READY_DELAY__

    $readyText = "{0}`r`n{1}`r`n{2}`r`n" -f $PID, $selfProc.StartTime.ToUniversalTime().Ticks, $realToken
    [IO.File]::WriteAllText($readyPath, $readyText, (New-Object Text.UTF8Encoding($true)))

    Start-Sleep -Seconds __POST_READY_SLEEP__
    return '{"status":"success","total_formulas":0,"total_errors":0,"error_summary":{}}'
}
Export-ModuleMember -Function Test-XlsFormulas
'@

    $decoyLiteral = if ($DecoyFirst) { '$true' } else { '$false' }
    $wrongTicksLiteral = if ($WrongStartTicksFirst) { '$true' } else { '$false' }
    $exclusiveMarkerLiteral = if ($WriteExclusiveMarker) { '$true' } else { '$false' }
    $content = $template -replace '__DECOY_FIRST__', $decoyLiteral -replace '__WRONG_START_TICKS_FIRST__', $wrongTicksLiteral -replace '__WRITE_EXCLUSIVE_MARKER__', $exclusiveMarkerLiteral -replace '__PRE_READY_DELAY__', [string]$PreReadyDelaySeconds -replace '__POST_READY_SLEEP__', [string]$PostReadySleepSeconds
    Set-Content -LiteralPath $path -Value $content -Encoding UTF8
    return $path
}

function New-XlsCliFakeFailingModule {
    <#
    .SYNOPSIS
        所有権（マーカー）が確立する前にパイプラインが失敗するケース（Excel 起動失敗の代替）を
        決定的に検証するためのフェイクモジュールを作る。マーカーは一切書かない。COM は使わない。
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $fakeModuleDir)) {
        New-Item -ItemType Directory -Path $fakeModuleDir -Force | Out-Null
    }
    $path = Join-Path $fakeModuleDir ("FakeFail-{0}.psm1" -f [guid]::NewGuid())

    $content = @'
#Requires -Version 5.1
function Test-XlsFormulas {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$TimeoutSec = 30,
        [switch]$Force,
        [switch]$AsJson
    )
    throw "Simulated Excel launch failure: class not registered"
}
Export-ModuleMember -Function Test-XlsFormulas
'@

    Set-Content -LiteralPath $path -Value $content -Encoding UTF8
    return $path
}

function New-XlsHeavyFormulaChainWorkbook {
    <#
    .SYNOPSIS
        「所有権確立後の区間（Workbooks.Open〜CalculateFullRebuild〜走査）」が短いタイムアウトを
        確実に超えるよう、数万セルの数式連鎖を持つブックを作る（実装メモ参照: 1 セルずつのループは
        しない。範囲一括の Formula 代入は相対参照が自動調整されるため、A2:A50000 へ '=A1+1' を
        1 回で代入するだけで A1 を先頭にした連鎖ができる）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Invoke-XlsSession -Path $Path -ScriptBlock {
        param($app, $wb)
        $ws = $wb.Worksheets.Item(1)
        $ws.Range('A1').Formula = '=1'
        $ws.Range('A2:A50000').Formula = '=A1+1'
        Save-XlsWorkbook -Workbook $wb -Path $Path
    }
}

Describe 'Test-XlsFormulas.ps1 CLI (T-12)' {

    BeforeEach {
        Import-Module $modulePath -Force
        $script:BaselinePids = Get-ExcelPids
    }

    AfterEach {
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }

    AfterAll {
        Remove-Module $moduleName -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $fakeModuleDir) {
            Remove-Item -LiteralPath $fakeModuleDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'normal verification results (exit 0, G-04)' {

        It 'exits 0 with status=errors_found for a workbook containing a formula error' {
            $path = New-TempXlsxPath
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = '=1/0'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path)

            $r.ExitCode | Should Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.status | Should Be 'errors_found'
            $parsed.total_formulas | Should Be 1
            $parsed.total_errors | Should Be 1
            $parsed.error_summary.'#DIV/0!'.count | Should Be 1
        }

        It 'exits 0 with status=success for a workbook without formula errors' {
            $path = New-TempXlsxPath
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = '=1+1'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path)

            $r.ExitCode | Should Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.status | Should Be 'success'
            $parsed.total_formulas | Should Be 1
            $parsed.total_errors | Should Be 0
        }

        It 'exits 0 with status=refused for a workbook with a missing external link, and -Force runs the normal flow instead' {
            $path = New-TempXlsxPath
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = "='C:\NoSuchFolderXYZ_T12\[Other.xlsx]Sheet1'!A1"
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $refused = Invoke-XlsFormulasCliProcess -Arguments @($path)
            $refused.ExitCode | Should Be 0
            $refusedParsed = $refused.Stdout.Trim() | ConvertFrom-Json
            $refusedParsed.status | Should Be 'refused'
            $refusedParsed.reason | Should Be 'external links present'
            (($refusedParsed.PSObject.Properties.Name | Sort-Object) -join ',') | Should Be 'links,reason,status'

            $forced = Invoke-XlsFormulasCliProcess -Arguments @($path, '30', '-Force')
            $forced.ExitCode | Should Be 0
            $forcedParsed = $forced.Stdout.Trim() | ConvertFrom-Json
            (@('errors_found', 'success') -contains $forcedParsed.status) | Should Be $true
        }

        It 'runs correctly when the workbook path contains a space (round 1 review should-fix: Start-Process argument quoting)' {
            $path = New-TempXlsxPathWithSpace
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = '=1+1'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path)

            $r.ExitCode | Should Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.status | Should Be 'success'
        }
    }

    Context 'error results (exit non-zero, { "error": ... }, G-04)' {

        It 'exits non-zero with a next-step error message when the file does not exist' {
            $path = New-TempXlsxPath  # わざと作らない（ファイル不在）

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path)

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Match 'File not found'
        }
    }

    Context 'timeout mechanism (deterministic, no real Excel; round 1/round 2 review should-fix)' {
        # 実 Excel の起動時間（実装メモ参照: この開発機で 4〜61 秒の幅を実機確認）に依存せず、
        # CLI のタイムアウト機構そのものを決定的に検証する。フェイクモジュールは COM を一切使わない。

        It 'fires timeout once the ownership ready handshake completes, and Clear-XlsOrphans processes the exclusive marker as Skipped-DeadPid (round 3 review should-fix: Action verification)' {
            # round 3 レビュー should-fix 対応: フェイクモジュールにも本物と同じ排他マーカーを
            # 持たせ、CLI 強制終了後の Clear-XlsOrphans の戻り値を具体的に検証できるようにする。
            $fakeModulePath = New-XlsCliFakeReadyModule -WriteExclusiveMarker -PostReadySleepSeconds 5

            $readyDirForTest = Join-Path $env:TEMP 'xlsagent-ready'
            $preExistingReadyNames = @()
            if (Test-Path -LiteralPath $readyDirForTest) {
                $preExistingReadyNames = @(Get-ChildItem -LiteralPath $readyDirForTest -Filter '*.ready' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            }

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-XlsFormulasCliProcess -Arguments @('dummy-path-unused', '1') -ModulePathOverride $fakeModulePath
            $sw.Stop()

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Be 'timeout'

            # フェイクモジュールはマーカー・ready ファイルを書いてから即座に眠るだけなので、
            # 所有権確立の検知はほぼ瞬時。timeoutSec=1 のカウントダウンにほぼ一致する壁時計時間で
            # 打ち切られることを確認する（実 Excel の起動待ちが混ざらないので決定的に短い上限を
            # 張れる）。
            $sw.Elapsed.TotalSeconds | Should BeLessThan 20

            # フェイクモジュールの PID（＝サブプロセス自身の PID、CLI は TerminateProcess で
            # 既に終了済み）に対する Clear-XlsOrphans の結果を具体的に検証する。対象プロセスは
            # 既に存在しないので Skipped-DeadPid になり、マーカーファイルは削除される
            # （round 3 レビュー should-fix: 「収集するだけで検証しない」への対応）。
            $markerDirForFake = Join-Path $env:TEMP 'xlsagent'
            $preExistingCount = @(Get-ChildItem -LiteralPath $markerDirForFake -Filter '*.marker' -File -ErrorAction SilentlyContinue).Count
            $preExistingCount | Should BeGreaterThan 0

            $results = @(Clear-XlsOrphans)
            $deadPidResults = @($results | Where-Object { $_.Action -eq 'Skipped-DeadPid' })
            $deadPidResults.Count | Should BeGreaterThan 0

            foreach ($deadPidResult in $deadPidResults) {
                Test-Path -LiteralPath $deadPidResult.MarkerPath | Should Be $false
            }

            # round 3 レビュー nit 対応: timeout 経路でも ready ファイルは finally で削除される
            # （Complete-XlsCli の TerminateProcess より前に実行されるため）。新規の ready が
            # 残っていないことを確認する。
            $currentReadyNames = @(Get-ChildItem -LiteralPath $readyDirForTest -Filter '*.ready' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            $extraReadyNames = @($currentReadyNames | Where-Object { $preExistingReadyNames -notcontains $_ })
            $extraReadyNames.Count | Should Be 0
        }

        It 'does not accept a ready file until its StartTicks matches the live process (round 3 review blocking: PID-reuse resistance)' {
            # 正しい PID・runToken だが誤った StartTicks の ready を最初に書き、3 秒後に正しい
            # StartTicks へ置き換える。round 3 の修正（StartTicks を Get-Process の実際の
            # StartTime と完全一致させる）が効いていれば、最初の誤った ready では所有権確立と
            # 判定されず、3 秒後の正しい ready まで待ってから timeoutSec=1 が測られるため、
            # 合計経過時間は 3 秒強になる。修正前のコード（TryParse するだけ）なら誤った ready を
            # 即座に受理してしまい、timeoutSec=1 秒前後で打ち切られる。
            $fakeModulePath = New-XlsCliFakeReadyModule -WrongStartTicksFirst -PreReadyDelaySeconds 3 -PostReadySleepSeconds 5

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-XlsFormulasCliProcess -Arguments @('dummy-path-unused', '1') -ModulePathOverride $fakeModulePath
            $sw.Stop()

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Be 'timeout'

            $sw.Elapsed.TotalSeconds | Should BeGreaterThan 3.4
            $sw.Elapsed.TotalSeconds | Should BeLessThan 20

            Clear-XlsOrphans | Out-Null
        }

        It 'does not start the timeoutSec countdown when an unrelated ready file/marker appears first (round 2 review blocking (a))' {
            # 無関係な別セッションを装ったデコイ（別 runToken）が先に現れても、CLI が惑わされずに
            # 自分の runToken の ready ファイルが出るまで（PreReadyDelaySeconds ぶん）待つことを
            # 固定する。惑わされていれば、デコイ出現とほぼ同時に timeoutSec=1 で打ち切られ、
            # 総経過時間が PreReadyDelaySeconds よりずっと短くなるはず。
            $fakeModulePath = New-XlsCliFakeReadyModule -DecoyFirst -PreReadyDelaySeconds 3 -PostReadySleepSeconds 5

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-XlsFormulasCliProcess -Arguments @('dummy-path-unused', '1') -ModulePathOverride $fakeModulePath
            $sw.Stop()

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Be 'timeout'

            # デコイに惑わされていれば ~1 秒程度で打ち切られる。正しく自分の ready を待てば
            # 3 秒（PreReadyDelaySeconds）+ 1 秒（timeoutSec）前後になるはず。
            $sw.Elapsed.TotalSeconds | Should BeGreaterThan 3.4
            $sw.Elapsed.TotalSeconds | Should BeLessThan 20

            Clear-XlsOrphans | Out-Null
        }

        It 'does not force-kill while the ownership handshake has not completed yet (round 2 review blocking (b): delayed marker content)' {
            # 本物の ready ファイルが書かれるまで PreReadyDelaySeconds だけ意図的に遅らせる。
            # 所有権確立前に強制終了していれば、総経過時間は timeoutSec=1 秒前後で終わるはず。
            $fakeModulePath = New-XlsCliFakeReadyModule -PreReadyDelaySeconds 3 -PostReadySleepSeconds 5

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-XlsFormulasCliProcess -Arguments @('dummy-path-unused', '1') -ModulePathOverride $fakeModulePath
            $sw.Stop()

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Be 'timeout'

            $sw.Elapsed.TotalSeconds | Should BeGreaterThan 3.4
            $sw.Elapsed.TotalSeconds | Should BeLessThan 20

            Clear-XlsOrphans | Out-Null
        }

        It 'converts a pipeline failure before ownership is established into an error JSON without waiting for timeoutSec (simulated Excel launch failure)' {
            $fakeModulePath = New-XlsCliFakeFailingModule

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-XlsFormulasCliProcess -Arguments @('dummy-path-unused', '30') -ModulePathOverride $fakeModulePath
            $sw.Stop()

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Match 'Simulated Excel launch failure'

            # ready ファイルが一度も存在しないので所有権未確立のまま完了する経路（フェーズ 1 で
            # asyncResult.IsCompleted になり timeoutSec=30 を待たない）。時間内に確実に完了する
            # ことを固定する。
            $sw.Elapsed.TotalSeconds | Should BeLessThan 20
        }
    }

    Context 'timeout with real Excel (post-ownership, round 1/round 2 review blocking fix verification)' {

        It 'exits non-zero with { "error": "timeout" } when timeoutSec elapses after ownership is established, and Clear-XlsOrphans reclaims any leftover Excel/marker (G-10)' {
            $path = New-TempXlsxPath
            # 所有権確立後の区間（Workbooks.Open〜CalculateFullRebuild〜走査）が 1 秒を確実に
            # 超えるよう、数万セルの数式連鎖を持つ重いブックを使う（New-XlsHeavyFormulaChainWorkbook
            # 参照。範囲一括の Formula 代入で作るので G-08 のセル単位ループ禁止に抵触しない）。
            New-XlsHeavyFormulaChainWorkbook -Path $path

            $markerDir = Join-Path $env:TEMP 'xlsagent'
            $baselineMarkerNames = @()
            if (Test-Path -LiteralPath $markerDir) {
                $baselineMarkerNames = @(Get-ChildItem -LiteralPath $markerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            }

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path, '1')

            $r.ExitCode | Should Not Be 0
            $parsed = $r.Stdout.Trim() | ConvertFrom-Json
            $parsed.error | Should Be 'timeout'

            # このタイムアウトが実際に真の孤児（マーカー付き）を残したかどうかを、最初の
            # Clear-XlsOrphans 呼び出しより前の時点で記録しておく（round 3 レビュー should-fix:
            # 新規マーカーに対応する結果が Stopped または安全側の Action であることの検証用）。
            $currentMarkerNamesAtStart = @(Get-ChildItem -LiteralPath $markerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            $sawNewMarkerAtStart = (@($currentMarkerNamesAtStart | Where-Object { $baselineMarkerNames -notcontains $_ })).Count -gt 0

            # round 1 レビュー blocking #2 対応: 直接 Stop-Process はしない。回収は Clear-XlsOrphans
            # のみ。round 2 レビュー should-fix 対応: 実機確認（実装メモ参照）で、CLI の
            # BeginStop() がタイミング次第でパイプラインの正常な finally（Close→Quit→Release→
            # マーカー削除）を先に完走させてしまい、EXCEL.EXE も自然消滅するケースと、
            # TerminateProcess が先に確定してマーカー付きの孤児が残るケースの両方が実機で起こる
            # ことを確認した。どちらのケースでも安全（前者は後始末不要、後者は Clear-XlsOrphans が
            # 回収可能）なので、Excel PID の有無にかかわらず Clear-XlsOrphans を最低 1 回呼び、
            # 戻り値の Action と、新規マーカー・追加 EXCEL.EXE が最終的にゼロであることの両方を
            # 検証する。
            $clearResultsSeen = New-Object System.Collections.ArrayList
            $clearCallCount = 0
            $deadline = (Get-Date).AddSeconds(150)
            do {
                $iterationResults = @(Clear-XlsOrphans)
                $clearCallCount++
                foreach ($item in $iterationResults) { [void]$clearResultsSeen.Add($item) }

                $currentMarkerNames = @(Get-ChildItem -LiteralPath $markerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
                $extraMarkers = @($currentMarkerNames | Where-Object { $baselineMarkerNames -notcontains $_ })

                $currentPids = Get-ExcelPids
                $extraPids = @($currentPids | Where-Object { $script:BaselinePids -notcontains $_ })

                if ($extraMarkers.Count -eq 0 -and $extraPids.Count -eq 0) {
                    break
                }
                Start-Sleep -Seconds 3
            } while ((Get-Date) -lt $deadline)

            # round 2 レビュー should-fix 対応: Excel PID の有無にかかわらず Clear-XlsOrphans を
            # 最低 1 回呼んだことを明示的に固定する（do-while なので構造的には保証されるが、
            # テストの意図を読み手にも明確にするため実測値でも確認する）。
            $clearCallCount | Should BeGreaterThan 0

            # round 3 レビュー should-fix 対応: 新規マーカーが実際に存在したケースでは、
            # Clear-XlsOrphans の戻り値に「対応が取れた」ことを示す Action（真の孤児として停止した
            # Stopped、または生存確認できず安全側にマーカーだけ処理した各種 Skipped-*）が
            # 最低 1 件含まれることを確認する。マーカーが最初から存在しなかった場合（BeginStop() が
            # 競合に勝ってグレースフルに片付いたケース、実装メモ参照）はこの検証をスキップする。
            if ($sawNewMarkerAtStart) {
                $safeActions = @('Stopped', 'Skipped-DeadPid', 'Skipped-UnverifiedLegacy', 'Skipped-Unverified', 'Skipped-NotExcel', 'Skipped-InvalidMarker')
                $matchingActions = @($clearResultsSeen | Where-Object { $safeActions -contains $_.Action })
                $matchingActions.Count | Should BeGreaterThan 0
            }

            $currentMarkerNames = @(Get-ChildItem -LiteralPath $markerDir -Filter '*.marker' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            $extraMarkers = @($currentMarkerNames | Where-Object { $baselineMarkerNames -notcontains $_ })
            $extraMarkers.Count | Should Be 0

            $currentPids = Get-ExcelPids
            $extraPids = @($currentPids | Where-Object { $script:BaselinePids -notcontains $_ })
            $extraPids.Count | Should Be 0
        }
    }

    Context 'stdout contract (G-04: JSON を 1 個だけ)' {

        It 'writes exactly one JSON value to stdout with no extra output' {
            $path = New-TempXlsxPath
            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Formula = '=1+1'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $r = Invoke-XlsFormulasCliProcess -Arguments @($path)
            $trimmed = $r.Stdout.Trim()

            $trimmed.StartsWith('{') | Should Be $true
            $trimmed.EndsWith('}') | Should Be $true
            { $trimmed | ConvertFrom-Json } | Should Not Throw
            $r.ExitCode | Should Be 0
        }
    }
}
