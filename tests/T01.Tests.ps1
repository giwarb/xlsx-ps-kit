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

    It 'does not have $script:Xl populated yet (skeleton only)' {
        Import-Module $modulePath -Force
        $xl = & (Get-Module $moduleName) { $script:Xl }
        $xl.Count | Should Be 0
    }

    foreach ($fn in $script:ExpectedFunctions) {
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
