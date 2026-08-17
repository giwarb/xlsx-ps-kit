#Requires -Version 5.1

<#
tests/T01.Tests.ps1

T-01: XlsAgent.psm1 の骨組みが Import-Module でき、公開 7 関数（+ Clear-XlsOrphans = 計 8）が
Export されていること、Pester が 3.x で動いていることを確認する。COM には触らない。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

$script:ExpectedFunctions = @(
    'Invoke-XlsSession',
    'Save-XlsWorkbook',
    'Get-XlsOverview',
    'Get-XlsModel',
    'Get-XlsRange',
    'Set-XlsRange',
    'Test-XlsFormulas',
    'Clear-XlsOrphans'
)

Describe 'XlsAgent module skeleton (T-01)' {

    BeforeEach {
        $script:BaselinePids = Get-ExcelPids
    }

    AfterEach {
        Assert-NoOrphanExcel -Baseline $script:BaselinePids
    }

    It 'runs under Pester 3.x (not v5)' {
        $pesterModule = Get-Module Pester
        $pesterModule.Version.Major | Should Be 3
    }

    It 'imports XlsAgent.psm1 without error' {
        { Import-Module $modulePath -Force -ErrorAction Stop } | Should Not Throw
    }

    It 'exports exactly the 8 designed public functions' {
        Import-Module $modulePath -Force
        $exported = @(Get-Command -Module $moduleName | Select-Object -ExpandProperty Name)

        foreach ($fn in $script:ExpectedFunctions) {
            $exported -contains $fn | Should Be $true
        }

        $exported.Count | Should Be $script:ExpectedFunctions.Count
    }

    It 'defines the $script:Xl COM constants table container' {
        Import-Module $modulePath -Force
        # 内部変数なので & でモジュールスコープに入って確認する
        $xl = & (Get-Module $moduleName) { $script:Xl }
        $xl | Should Not Be $null
        $xl.GetType().Name | Should Be 'Hashtable'
    }

    It 'contains the verified constant required by Invoke-XlsSession' {
        # T-01 時点では空だったが、T-03 で Invoke-XlsSession が Application.Calculation に
        # xlCalculationManual を使うようになったため増分で入った。
        # 罠（T-03 round 1 レビュー指摘）: 総件数を固定する assertion（$xl.Count | Should Be N）は
        # 後続タスクが正当に定数を追加するたびに陳腐化し、同じ修正を繰り返すことになる。
        # G-12（実行して確認したものだけ載せる）の趣旨は「未確認の定数を先回りで詰め込まない」ことなので、
        # ここでは T-03 が使う既知のキーの有無と値だけを確認し、テーブル全体の件数は固定しない。
        Import-Module $modulePath -Force
        $xl = & (Get-Module $moduleName) { $script:Xl }

        $xl.ContainsKey('xlCalculationManual') | Should Be $true
        $xl.xlCalculationManual | Should Be -4135
    }

    # 罠（T-04 で判明、T-03 round 1 レビューの「T-01 assertion 陳腐化」指摘と同種）:
    # このループは「まだ未実装（スタブ）の公開関数は throw 文を含む」ことを一律に検証するものだったが、
    # 公開関数が実タスクで実装されると、その throw は「未実装です」という一様なものではなくなり、
    # 関数によっては本体に throw 文が一切なくなることもある（Save-XlsWorkbook は拡張子検証やパス比較を
    # 内部ヘルパーに委譲しており、自分自身の ScriptBlock には throw 文がない）。したがって、実装済みの
    # 関数はこのループの対象から外す。実装済みかどうかは各関数の担当タスクカード（T-03, T-04, ...）が
    # done になった時点でここから除外していく。
    $script:StillStubFunctions = $script:ExpectedFunctions | Where-Object {
        # T-03: Invoke-XlsSession 実装済み（ただし自身のエラー処理に throw が残っているため、この
        # ループに残しても偶然パスする。実装済みという事実を明示するため、ここでも除外しておく）。
        # T-04: Save-XlsWorkbook 実装済み。
        # T-05: Clear-XlsOrphans 実装済み（Stop-Process 失敗時などの throw は残っているため、この
        # ループに残しても偶然パスしてしまう。実装済みという事実を明示するため除外する）。
        # T-06: Get-XlsRange 実装済み（無効な範囲アドレスの throw は残っているため、このループに
        # 残しても偶然パスしてしまう。実装済みという事実を明示するため除外する）。
        $_ -notin @('Invoke-XlsSession', 'Save-XlsWorkbook', 'Clear-XlsOrphans', 'Get-XlsRange')
    }

    foreach ($fn in $script:StillStubFunctions) {
        It "stub function '$fn' contains an explicit throw" {
            Import-Module $modulePath -Force
            $cmd = Get-Command $fn -Module $moduleName
            $throwAst = $cmd.ScriptBlock.Ast.Find(
                {
                    param($node)
                    $node -is [System.Management.Automation.Language.ThrowStatementAst]
                },
                $true
            )
            $throwAst | Should Not Be $null
        }
    }

    AfterAll {
        Remove-Module $moduleName -ErrorAction SilentlyContinue
    }
}
