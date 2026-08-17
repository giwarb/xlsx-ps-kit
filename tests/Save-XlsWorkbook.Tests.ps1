#Requires -Version 5.1

<#
tests/Save-XlsWorkbook.Tests.ps1

T-04: Save-XlsWorkbook（保存の唯一の経路。G-09）の受け入れテスト。
harness/state/tasks/T-04.md の受け入れテスト要点:
  (a) .xlsm を保存してマクロが残る（VBProject アクセス不可でも HasVBProject で確認）
  (b) 保存後に開き直して Calculation が Automatic
  追加: .xlsx 新規保存→開き直して内容一致、同一パスへの再保存（Save 経路）、
        拡張子から FileFormat が正しく選ばれる、CalculateFullRebuild により数式の値が
        保存前に最新化される

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中でのみ
触る。このテストファイルも例外ではない。ただし Get-XlsSaveFileFormat / Test-XlsSamePath は
COM を一切使わない純粋関数（Save-XlsWorkbook 内部ヘルパー）なので、モジュールスコープに
直接入って（T-01/T-02 と同じ `& $script:ModuleRef { ... }` パターンで）COM なしに単体テストする。

(b) の「保存後に開き直すと Automatic」の検証は、Invoke-XlsSession をもう一度使って COM で
開き直す方法は取らない。Invoke-XlsSession は Open/Add の直後に必ず Application.Calculation を
Manual に強制するため（G-06 上コレを避ける手段がない）、COM 経由で開き直すと常に Manual に
上書きされてしまい、ファイルに焼き込まれた状態を観測できない。代わりに、保存された .xlsx を
（Excel を経由せず）.NET の System.IO.Compression で ZIP として開き、xl/workbook.xml の
<calcPr calcMode="..."/> を直接読む。これは COM ではないので G-06 に抵触しない。
Save-XlsWorkbook の実装メモに、Automatic で保存すると calcMode 属性自体が省略される
（Excel の既定値が automatic のため）ことを実機確認した記録がある。

.xlsm の HasVBProject 検証はタスクカードの2段構え方針のうち、実行可能な1段目のみを行う:
「マクロなし .xlsm でも FileFormat 52 で保存され HasVBProject が false のまま拡張子と形式が
一致する」ことを検証する。2段目（マクロ付き既存 .xlsm を使った HasVBProject 維持の検証）は、
リポジトリにマクロ付き fixture が存在しないため実施しない。マクロ付き .xlsm を新規に作るには
Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定（Trust Center）が
有効である必要があり、これはテスト実行環境のセキュリティ設定を書き換える行為に当たるため
行わない（実装メモ参照）。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

function New-XlsSingleCellArray {
    <#
    .SYNOPSIS
        単一値を 1x1 の [object[,]] にラップする。COM は触らない。
    .PARAMETER Value
        Range.Value2 に代入したい単一の値。
    .EXAMPLE
        $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'hello')
    .NOTES
        レビュー指摘対応（T-04 round 1 blocking）: G-08「Range.Value2 への代入は [object[,]] のみ」に
        単一セルの例外はない。テストコードでも scalar を直接代入せず、この関数で必ず [object[,]] に
        包んでから Value2 へ渡す。
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    [object[,]]$cell = [Array]::CreateInstance([object], 1, 1)
    $cell[0, 0] = $Value
    return , $cell
}

function New-TempXlsPathWithExtension {
    <#
    .SYNOPSIS
        任意の拡張子で一時テストパスを発行する（New-TempXlsxPath の拡張子可変版）。
        ファイル自体は作らない。COM は触らない。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $dir = Join-Path $env:TEMP 'xlsagent-tests'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    Join-Path $dir ("{0}{1}" -f [guid]::NewGuid().ToString(), $Extension)
}

function Get-XlsxCalcModeFromDisk {
    <#
    .SYNOPSIS
        保存済み .xlsx を Excel を使わず ZIP として開き、xl/workbook.xml の
        <calcPr calcMode="..."/> を読む。COM は一切使わない（G-06 に抵触しない）。
    .EXAMPLE
        Get-XlsxCalcModeFromDisk -Path 'C:\tmp\book.xlsx'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('xl/workbook.xml')
        if (-not $entry) {
            throw "xl/workbook.xml not found inside '$Path'."
        }
        $stream = $entry.Open()
        # レビュー指摘対応（T-04 round 1 should-fix）: エンコーディングを明示する（S-01/P-3）。
        # xl/workbook.xml は BOM なし UTF-8 なので detectEncodingFromByteOrderMarks は $true のままでよい。
        $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList @($stream, [Text.Encoding]::UTF8, $true)
        try {
            $xml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        if ($xml -match '<calcPr[^>]*\bcalcMode="([^"]*)"') {
            return $matches[1]
        }
        elseif ($xml -match '<calcPr\b') {
            # calcMode 属性省略時の Excel の既定値は automatic。
            return 'auto-default'
        }
        else {
            return $null
        }
    }
    finally {
        $zip.Dispose()
    }
}

Describe 'Save-XlsWorkbook (T-04)' {

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

    Context 'Get-XlsSaveFileFormat (pure function, no COM)' {

        It 'maps each supported extension to the designed FileFormat constant' {
            $cases = @{
                '.xlsx' = 51
                '.xlsm' = 52
                '.xltx' = 54
                '.xltm' = 53
                '.csv'  = 6
            }

            foreach ($ext in $cases.Keys) {
                $result = & $script:ModuleRef { param($p) Get-XlsSaveFileFormat -Path $p } "C:\tmp\book$ext"
                $result | Should Be $cases[$ext]
            }
        }

        It 'is case-insensitive on the extension' {
            $result = & $script:ModuleRef { param($p) Get-XlsSaveFileFormat -Path $p } 'C:\tmp\BOOK.XLSM'
            $result | Should Be 52
        }

        It 'throws a next-step error for an unsupported extension' {
            $errorSeen = $null
            try {
                & $script:ModuleRef { param($p) Get-XlsSaveFileFormat -Path $p } 'C:\tmp\book.txt'
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match '\.txt'
            $errorSeen.Exception.Message | Should Match 'xlsx'
        }
    }

    Context 'Test-XlsSamePath (pure function, no COM)' {

        It 'treats identical paths as the same file' {
            $p = 'C:\tmp\xlsagent-tests\book.xlsx'
            $result = & $script:ModuleRef { param($a, $b) Test-XlsSamePath -PathA $a -PathB $b } $p $p
            $result | Should Be $true
        }

        It 'ignores case (Windows paths are case-insensitive)' {
            $result = & $script:ModuleRef { param($a, $b) Test-XlsSamePath -PathA $a -PathB $b } 'C:\Tmp\Book.xlsx' 'c:\tmp\book.xlsx'
            $result | Should Be $true
        }

        It 'treats different paths as different files' {
            $result = & $script:ModuleRef { param($a, $b) Test-XlsSamePath -PathA $a -PathB $b } 'C:\tmp\a.xlsx' 'C:\tmp\b.xlsx'
            $result | Should Be $false
        }

        It 'treats an empty PathA (unsaved workbook) as not the same file' {
            $result = & $script:ModuleRef { param($a, $b) Test-XlsSamePath -PathA $a -PathB $b } '' 'C:\tmp\a.xlsx'
            $result | Should Be $false
        }
    }

    Context '.xlsx save/reopen and FileFormat selection (integration, real Excel)' {

        It 'saves a new workbook via SaveAs, and reopening it shows the same content' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'
            $marker = 'xlsagent-{0}' -f ([guid]::NewGuid())

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value $marker)
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $seen = Invoke-XlsSession -Path $path -ReadOnly -ScriptBlock {
                param($app, $wb)
                $wb.Worksheets.Item(1).Range('A1').Value2
            }

            $seen | Should Be $marker
        }

        It 'saves to the same path a second time via Save() and picks up the new content' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'first')
                Save-XlsWorkbook -Workbook $wb -Path $path

                # 同一パスへの2回目の保存（Save() 経路）。SaveAs の「既に開いているファイルへの上書き」
                # 確認ダイアログは出ない（DisplayAlerts=$false）ことも、ここでハングしないことで確認する。
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 'second')
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $seen = Invoke-XlsSession -Path $path -ReadOnly -ScriptBlock {
                param($app, $wb)
                $wb.Worksheets.Item(1).Range('A1').Value2
            }

            $seen | Should Be 'second'
        }

        It 'selects the correct FileFormat for every supported extension (checked live via Workbook.FileFormat right after SaveAs)' {
            # 罠（実装メモ参照）: .xltx/.xltm は Workbooks.Open で開き直すと、テンプレート自体ではなく
            # それを元にした新規ブックが作られる。そのため FileFormat の確認は reopen ではなく
            # SaveAs 直後にその場で行う（1つの Invoke-XlsSession の中、G-06 を満たしたまま）。
            $cases = @{
                '.xlsx' = 51
                '.xlsm' = 52
                '.xltx' = 54
                '.xltm' = 53
                '.csv'  = 6
            }
            $paths = @{}
            foreach ($ext in $cases.Keys) {
                $paths[$ext] = New-TempXlsPathWithExtension -Extension $ext
            }

            $seenFormats = Invoke-XlsSession -Path (New-TempXlsPathWithExtension -Extension '.xlsx') -ScriptBlock {
                param($app, $wb)
                $results = @{}
                foreach ($ext in $paths.Keys) {
                    Save-XlsWorkbook -Workbook $wb -Path $paths[$ext]
                    $results[$ext] = $wb.FileFormat
                }
                $results
            }

            foreach ($ext in $cases.Keys) {
                $seenFormats[$ext] | Should Be $cases[$ext]
            }
        }
    }

    Context '(a) .xlsm keeps HasVBProject consistent with the FileFormat/extension' {

        It 'saves a macro-less workbook as .xlsm with FileFormat 52 and HasVBProject stays false' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsm'

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                Save-XlsWorkbook -Workbook $wb -Path $path
                [PSCustomObject]@{
                    FileFormat   = $wb.FileFormat
                    HasVBProject = $wb.HasVBProject
                }
            }

            $result.FileFormat | Should Be 52
            $result.HasVBProject | Should Be $false
        }
    }

    Context '(b) Calculation handling around save' {

        It 'recalculates stale formula results (CalculateFullRebuild) before saving' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'

            $values = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 10)
                $ws.Range('A2').Formula = '=A1*2'
                $before = $ws.Range('A2').Value2

                # Manual のまま先行セルを書き換える。手動計算のままなら A2 は古い値のままになる。
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 99)
                $stillStale = $ws.Range('A2').Value2

                Save-XlsWorkbook -Workbook $wb -Path $path

                [PSCustomObject]@{
                    Before  = $before
                    Stale   = $stillStale
                    AfterSave = $ws.Range('A2').Value2
                }
            }

            $values.Before | Should Be 20
            $values.Stale | Should Be 20
            $values.AfterSave | Should Be 198
        }

        It 'bakes an Automatic calculation mode into the saved .xlsx (verified without COM, via the ZIP contents)' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                # セッションは Manual で始まる（Invoke-XlsSession）。Save-XlsWorkbook が
                # 保存前に Automatic へ戻すことをこのテストで確認する。
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $calcMode = Get-XlsxCalcModeFromDisk -Path $path
            # Automatic で保存すると calcMode 属性自体が省略される（実機確認、実装メモ参照）。
            $calcMode | Should Be 'auto-default'
        }

        It 'restores Application.Calculation to Manual on the live session after saving' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'
            $expectedManual = & $script:ModuleRef { $script:Xl.xlCalculationManual }

            $calcAfterSave = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                Save-XlsWorkbook -Workbook $wb -Path $path
                $app.Calculation
            }

            $calcAfterSave | Should Be $expectedManual
        }
    }

    Context 'error handling when Save/SaveAs fails (T-04 round 1 should-fix)' {

        It 'wraps a SaveAs failure in an actionable message and still restores Calculation to Manual' {
            $path = New-TempXlsPathWithExtension -Extension '.xlsx'
            # 親ディレクトリが存在しないパス。SaveAs は生の COM エラーを投げる（実機確認）。
            $missingParentDir = Join-Path $env:TEMP ('xlsagent-tests-missing-{0}' -f ([guid]::NewGuid()))
            $badPath = Join-Path $missingParentDir 'book.xlsx'
            $expectedManual = & $script:ModuleRef { $script:Xl.xlCalculationManual }

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $errorSeen = $null
                try {
                    Save-XlsWorkbook -Workbook $wb -Path $badPath
                }
                catch {
                    $errorSeen = $_
                }

                [PSCustomObject]@{
                    ErrorSeen        = $errorSeen
                    CalculationAfter = $app.Calculation
                }
            }

            $result.ErrorSeen | Should Not Be $null
            # 生の HRESULT メッセージのままではなく、次の一手（パス・権限・ロック・拡張子の確認）が
            # 分かる文に包み直されていること（S-04、レビュー指摘対応）。
            $result.ErrorSeen.Exception.Message | Should Match ([regex]::Escape($badPath))
            $result.ErrorSeen.Exception.Message | Should Match 'parent directory'
            # catch で例外を包んだ後も、finally で確実に Manual に戻っていること。
            $result.CalculationAfter | Should Be $expectedManual
        }
    }
}
