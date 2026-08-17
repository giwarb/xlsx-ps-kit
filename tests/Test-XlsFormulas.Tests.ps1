#Requires -Version 5.1

<#
tests/Test-XlsFormulas.Tests.ps1

T-10: Test-XlsFormulas（全シートを再計算し、recalc.py 互換の JSON 契約でエラー有無を返す。
01-design.md §3.6。G-04: JSON 契約のキー名は変えない）の受け入れテスト。
harness/state/tasks/T-10.md の受け入れテスト要点:
  - 意図的に #REF! #DIV/0! #NAME? を含むブックで、種別ごとの count と locations が正しい。
  - 101 個以上のエラーで locations が 100 で切れ locations_truncated が正しい
    （範囲一括の Formula 代入で作る。1 セルずつのループはしない）。
  - 数式はあるがエラーなし → status="success"、total_formulas > 0、error_summary が空。
  - 数式ゼロ → success / 0 / 0。
  - Manual のまま古い値が残ったブックでも CalculateFullRebuild 後の値で判定される
    （T-04 罠 5 の機序。実装メモに詳細あり）。
  - 異常系: ファイル不在で次の一手が分かる例外。
  - locations の順序: シート順→セル順（決定的であること）。

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中
（Test-XlsFormulas 自体を含む）でのみ触る。フィクスチャ作成は Common.ps1 の New-XlsSingleCellArray
で作った [object[,]]（G-08）を ScriptBlock 内で Range.Value2 に代入する形と、Range.Formula への
文字列代入（COM プロパティ代入であり Value2 ではないため G-08 の対象外）で行う。

「Manual のまま古い値が残ったブック」の再現について: Save-XlsWorkbook は保存の度に必ず
Automatic + CalculateFullRebuild を行う（G-09/Set-XlsCalculationAutomatic）ため、この経路だけでは
「ファイルに保存された数式のキャッシュ値が実際の依存元セルと食い違う」状態を作れない。
Save-XlsWorkbook.Tests.ps1 の Get-XlsxCalcModeFromDisk（.xlsx を COM を使わず ZIP として直接読む）と
同じ精神で、ここでは New-XlsxWithStaleFormulaCache が ZIP を直接書き換えて依存元セルの値だけを
書き換え、依存先の数式セルのキャッシュ値を古いまま残す。COM を一切使わないので G-06 に抵触しない。

罠（実機確認、実装メモ参照）: 同じセッション内で複数セルへ Formula/Value2 を連続で書き込んだ直後に
Save-XlsWorkbook を呼ぶと、ごく稀に最後の 1 セルが保存後のファイルに反映されないことがある
（Range.Formula/Value2 の読み戻しで直後に確認しても未反映のまま見えるため、読み取り側のキャッシュや
CalculateFullRebuild 不足が原因ではない）。単純な Start-Sleep での「一呼吸置く」対処では再現を
安定して防げなかった（実装メモに検証記録あり）ため、Set-XlsFormulaVerified が「書き込み直後に
読み戻して一致することを確認し、一致しなければ再試行する」防御的な書き込みを行う。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

function Set-XlsFormulaVerified {
    <#
    .SYNOPSIS
        Range.Formula に書き込み、直後に読み戻して一致することを確認する。一致しなければ
        短い待機を挟んで再試行する（最大 10 回）。COM は Range オブジェクトの読み書きのみ
        （呼び出し元の Invoke-XlsSession ScriptBlock 内から呼ばれる前提。G-06）。
    .PARAMETER Range
        書き込み先の Range COM オブジェクト（単一セルを想定）。
    .PARAMETER Formula
        書き込む数式文字列（例 '=1/0'）。
    .EXAMPLE
        Set-XlsFormulaVerified -Range $ws.Range('A1') -Formula '=1/0'
    .NOTES
        罠（実機確認、実装メモ参照）: 同一セッションで複数セルへ Formula を連続代入した直後に
        Save-XlsWorkbook を呼ぶと、ごく稀に最後の 1 セルの書き込みが保存後のファイルに反映されない
        現象を実機で確認した。Range.Formula を代入直後に読み戻しても未反映のまま見えるため
        （＝読み取り側のキャッシュではなく、書き込み自体が未確定）、単純な待機（Start-Sleep）だけでは
        安定して防げなかった。読み戻して不一致なら代入をやり直す防御的な書き込みで確実に反映させる。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Range,

        [Parameter(Mandatory = $true)]
        [string]$Formula
    )

    # 複数セル範囲へ同一の数式文字列を代入すると、Range.Formula の読み戻しはセル数分だけ空白区切りで
    # 繰り返した文字列になる（実機確認）。範囲サイズによらず一貫して検証できるよう、先頭セル
    # （Cells.Item(1)、範囲内の相対 1-based インデックス）の Formula だけを読み戻して比較する。
    $firstCell = $Range.Cells.Item(1)
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $Range.Formula = $Formula
        if ([string]$firstCell.Formula -eq $Formula) {
            return
        }
        Start-Sleep -Milliseconds 200
    }

    throw "Failed to persist formula '$Formula' on '$($Range.Address(0, 0))' after 10 attempts (last read back: '$($firstCell.Formula)')."
}

function Set-XlsValue2Verified {
    <#
    .SYNOPSIS
        単一セル Range.Value2 に [object[,]] を書き込み、直後に読み戻して一致することを確認する。
        一致しなければ短い待機を挟んで再試行する（最大 10 回）。Set-XlsFormulaVerified の
        Value2 版（同じ罠への対処。.NOTES 参照）。COM は Range オブジェクトの読み書きのみ（G-06）。
    .PARAMETER Range
        書き込み先の Range COM オブジェクト（単一セルを想定）。
    .PARAMETER Value2
        書き込む値（Common.ps1 の New-XlsSingleCellArray で作った 1x1 の [object[,]]。G-08）。
    .EXAMPLE
        Set-XlsValue2Verified -Range $ws.Range('C1') -Value2 (New-XlsSingleCellArray -Value 'plain text')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Range,

        [Parameter(Mandatory = $true)]
        [object[,]]$Value2
    )

    $expected = $Value2[0, 0]
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $Range.Value2 = $Value2
        if ($Range.Value2 -eq $expected) {
            return
        }
        Start-Sleep -Milliseconds 200
    }

    throw "Failed to persist value '$expected' on '$($Range.Address(0, 0))' after 10 attempts (last read back: '$($Range.Value2)')."
}

function New-XlsxWithStaleFormulaCache {
    <#
    .SYNOPSIS
        保存済みの .xlsx を Excel を使わず ZIP として直接編集し、依存元セル A1 の値だけを書き換えて、
        依存先の数式セル（A2 を想定）のキャッシュ済み計算結果を古いまま残す。COM は一切使わない
        （G-06 に抵触しない。Save-XlsWorkbook.Tests.ps1 の Get-XlsxCalcModeFromDisk と対になる
        「直接書く」版）。
    .PARAMETER Path
        書き換え対象の .xlsx パス。あらかじめ Invoke-XlsSession + Save-XlsWorkbook で
        シート 1 枚目の A1=<OldPrecedentValue>（数値）、A2 に A1 を参照する数式を保存しておくこと。
    .PARAMETER OldPrecedentValue
        現在ファイルに保存されている A1 の値の文字列表現（置換元。ユニークにマッチさせるため）。
    .PARAMETER NewPrecedentValue
        A1 に書き換える新しい値の文字列表現。
    .EXAMPLE
        New-XlsxWithStaleFormulaCache -Path $path -OldPrecedentValue '5' -NewPrecedentValue '0'
    .NOTES
        実機確認済み（実装メモ参照）: xlsx の xl/worksheets/sheet1.xml を直接書き換えて
        `<c r="A1"><v>5</v></c>` を `<c r="A1"><v>0</v></c>` に変えても、A2 の
        `<c r="A2"><f>...</f><v>20</v></c>`（キャッシュ済み計算結果）はそのまま残る。さらに
        xl/calcChain.xml と、それを参照する [Content_Types].xml / xl/_rels/workbook.xml.rels の
        エントリを削除しておかないと Excel がファイルを壊れているとみなす（実機確認）。
        この状態を COM（Invoke-XlsSession、Manual で開始）で開いて Value2 を読むとキャッシュ済みの
        古い値（20）が返り、Set-XlsCalculationAutomatic 相当（Automatic + CalculateFullRebuild）を
        実行して初めて正しい値（この関数の用途では #DIV/0!）に更新されることを実機で確認した。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$OldPrecedentValue,

        [Parameter(Mandatory = $true)]
        [string]$NewPrecedentValue
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $extractDir = Join-Path $env:TEMP ("xlsagent-tests-extract-{0}" -f [guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $extractDir)

        $sheetPath = Join-Path $extractDir 'xl\worksheets\sheet1.xml'
        $sheetXml = Get-Content -LiteralPath $sheetPath -Raw -Encoding UTF8
        $oldCell = '<c r="A1"><v>{0}</v></c>' -f $OldPrecedentValue
        $newCell = '<c r="A1"><v>{0}</v></c>' -f $NewPrecedentValue
        if (-not $sheetXml.Contains($oldCell)) {
            throw "Expected to find '$oldCell' in '$sheetPath'. The fixture's A1 cell XML did not match the expected shape (New-XlsxWithStaleFormulaCache assumes an unstyled numeric A1 cell)."
        }
        $sheetXml = $sheetXml.Replace($oldCell, $newCell)
        [IO.File]::WriteAllText($sheetPath, $sheetXml, [Text.Encoding]::UTF8)

        # calcChain.xml が残っていると、参照元セルだけを直接書き換えたことと矛盾し Excel が
        # ファイル破損とみなす（実機確認）。存在すれば関連する 2 箇所の参照ごと削除する。
        $calcChainPath = Join-Path $extractDir 'xl\calcChain.xml'
        if (Test-Path -LiteralPath $calcChainPath) {
            Remove-Item -LiteralPath $calcChainPath -Force

            $ctPath = Join-Path $extractDir '[Content_Types].xml'
            $ct = Get-Content -LiteralPath $ctPath -Raw -Encoding UTF8
            $ct = $ct -replace '<Override PartName="/xl/calcChain.xml"[^>]*/>', ''
            [IO.File]::WriteAllText($ctPath, $ct, [Text.Encoding]::UTF8)

            $relsPath = Join-Path $extractDir 'xl\_rels\workbook.xml.rels'
            $rels = Get-Content -LiteralPath $relsPath -Raw -Encoding UTF8
            $rels = $rels -replace '<Relationship[^>]*Target="calcChain.xml"[^>]*/>', ''
            [IO.File]::WriteAllText($relsPath, $rels, [Text.Encoding]::UTF8)
        }

        Remove-Item -LiteralPath $Path -Force
        [System.IO.Compression.ZipFile]::CreateFromDirectory($extractDir, $Path)
    }
    finally {
        Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Test-XlsFormulas (T-10)' {

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

    Context 'error detection and the recalc.py JSON contract' {

        It 'reports per-type count and locations for #REF!/#DIV/0!/#NAME? and classifies status as errors_found' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsFormulaVerified -Range $ws.Range('A1') -Formula '=1/0'         # #DIV/0!
                Set-XlsFormulaVerified -Range $ws.Range('A2') -Formula '=NoSuchName'  # #NAME?
                Set-XlsFormulaVerified -Range $ws.Range('A3') -Formula '=#REF!'       # #REF!（T-06 com-constants.md と同じ手法）
                Set-XlsFormulaVerified -Range $ws.Range('B1') -Formula '=1+1'         # エラーなし数式
                Set-XlsValue2Verified -Range $ws.Range('C1') -Value2 (New-XlsSingleCellArray -Value 'plain text')
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $result = Test-XlsFormulas -Path $path

            $result.status | Should Be 'errors_found'
            $result.total_formulas | Should Be 4
            $result.total_errors | Should Be 3

            $result.error_summary.'#DIV/0!'.count | Should Be 1
            (@($result.error_summary.'#DIV/0!'.locations) -join ',') | Should Be 'Sheet1!A1'
            $result.error_summary.'#DIV/0!'.locations_truncated | Should Be 0

            $result.error_summary.'#NAME?'.count | Should Be 1
            (@($result.error_summary.'#NAME?'.locations) -join ',') | Should Be 'Sheet1!A2'

            $result.error_summary.'#REF!'.count | Should Be 1
            (@($result.error_summary.'#REF!'.locations) -join ',') | Should Be 'Sheet1!A3'
        }

        It 'orders locations deterministically by sheet order then row/column order (not creation order)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws1 = $wb.Worksheets.Item(1)
                $ws1.Name = 'Sheet1'
                $ws2 = $wb.Worksheets.Add([System.Type]::Missing, $ws1)
                $ws2.Name = 'Sheet2'

                # 行・列の昇順とは逆の順序でわざと作る（B2 を先、A1 を後）。
                Set-XlsFormulaVerified -Range $ws1.Range('B2') -Formula '=1/0'
                Set-XlsFormulaVerified -Range $ws1.Range('A1') -Formula '=1/0'
                Set-XlsFormulaVerified -Range $ws2.Range('C3') -Formula '=1/0'

                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $result = Test-XlsFormulas -Path $path

            $result.total_errors | Should Be 3
            $locations = @($result.error_summary.'#DIV/0!'.locations)
            $locations.Count | Should Be 3
            ($locations -join ',') | Should Be 'Sheet1!A1,Sheet1!B2,Sheet2!C3'
        }

        It 'truncates locations at 100 and reports locations_truncated for 101 errors of the same type (bulk range assignment, not a per-cell loop)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                # 範囲一括の Formula 代入。複数セル範囲への Range.Formula への scalar 代入は
                # 各セルへ同一の数式文字列を適用する一括操作として機能する（実機確認、実装メモ参照）。
                # Formula は Value2 ではないため G-08（Value2 代入は [object[,]] のみ）の対象外。
                Set-XlsFormulaVerified -Range $ws.Range('A1:A101') -Formula '=1/0'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $result = Test-XlsFormulas -Path $path

            $result.status | Should Be 'errors_found'
            $result.total_formulas | Should Be 101
            $result.total_errors | Should Be 101

            $entry = $result.error_summary.'#DIV/0!'
            $entry.count | Should Be 101
            $locations = @($entry.locations)
            $locations.Count | Should Be 100
            $locations[0] | Should Be 'Sheet1!A1'
            $locations[99] | Should Be 'Sheet1!A100'
            $entry.locations_truncated | Should Be 1
        }

        It 'returns status success with an empty error_summary when formulas exist but none error' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsValue2Verified -Range $ws.Range('A1') -Value2 (New-XlsSingleCellArray -Value 5.0)
                Set-XlsFormulaVerified -Range $ws.Range('A2') -Formula '=A1*2'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $result = Test-XlsFormulas -Path $path

            $result.status | Should Be 'success'
            $result.total_formulas | Should Be 1
            $result.total_errors | Should Be 0
            $result.error_summary.Count | Should Be 0
        }

        It 'returns success/0/0 for a workbook with no formulas at all' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsValue2Verified -Range $ws.Range('A1') -Value2 (New-XlsSingleCellArray -Value 'just text')
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $result = Test-XlsFormulas -Path $path

            $result.status | Should Be 'success'
            $result.total_formulas | Should Be 0
            $result.total_errors | Should Be 0
            $result.error_summary.Count | Should Be 0
        }

        It '-AsJson returns a recalc.py-compatible JSON string (G-04 key names)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsFormulaVerified -Range $ws.Range('A1') -Formula '=1/0'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $json = Test-XlsFormulas -Path $path -AsJson
            $parsed = $json | ConvertFrom-Json

            $parsed.status | Should Be 'errors_found'
            $parsed.total_formulas | Should Be 1
            $parsed.total_errors | Should Be 1

            # G-04: キー集合そのものを固定する (Codex T-10 レビュー should-fix 1)
            (($parsed.PSObject.Properties.Name | Sort-Object) -join ',') |
                Should Be 'error_summary,status,total_errors,total_formulas'

            $summary = $parsed.error_summary.'#DIV/0!'
            (($summary.PSObject.Properties.Name | Sort-Object) -join ',') |
                Should Be 'count,locations,locations_truncated'

            $summary.count | Should Be 1
            ($summary.locations -is [System.Array]) | Should Be $true
            $summary.locations.Count | Should Be 1
            $summary.locations[0] | Should Be 'Sheet1!A1'
            $summary.locations_truncated | Should Be 0
        }
    }

    Context 'recalculation and read-only behavior' {

        It 'computes the fully recalculated value rather than a stale cached value left over from Manual mode (T-04 trap 5)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsValue2Verified -Range $ws.Range('A1') -Value2 (New-XlsSingleCellArray -Value 5.0)
                Set-XlsFormulaVerified -Range $ws.Range('A2') -Formula '=100/A1'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            # ZIP を直接編集して A1 を 5→0 に書き換える。A2 のキャッシュ済み <v>（=20、エラーなし）は
            # そのまま残るので、ファイル上は「A2=20、エラーなし」という古い状態になる。
            New-XlsxWithStaleFormulaCache -Path $path -OldPrecedentValue '5' -NewPrecedentValue '0'

            $result = Test-XlsFormulas -Path $path

            # CalculateFullRebuild を経て初めて正しい #DIV/0! が検出される。
            $result.status | Should Be 'errors_found'
            $result.total_errors | Should Be 1
            $result.error_summary.'#DIV/0!'.count | Should Be 1
            (@($result.error_summary.'#DIV/0!'.locations) -join ',') | Should Be 'Sheet1!A2'
        }

        It 'does not modify the workbook on disk (ReadOnly session, verification only)' {
            $path = New-TempXlsxPath

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsFormulaVerified -Range $ws.Range('A1') -Formula '=1/0'
                Save-XlsWorkbook -Workbook $wb -Path $path
            }

            $hashBefore = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            Test-XlsFormulas -Path $path | Out-Null
            $hashAfter = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash

            $hashAfter | Should Be $hashBefore
        }
    }

    Context 'error handling' {

        It 'throws a next-step error and leaves no Excel process when the file does not exist' {
            $path = New-TempXlsxPath
            $errorSeen = $null

            try {
                Test-XlsFormulas -Path $path
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'File not found'
        }
    }
}
