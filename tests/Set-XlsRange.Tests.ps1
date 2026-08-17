#Requires -Version 5.1

<#
tests/Set-XlsRange.Tests.ps1

T-09: Set-XlsRange（一括書き込みの唯一の経路。01-design.md §3.5）と、内部関数
Test-XlsIsoDateString / ConvertFrom-XlsIsoDateStringToSerial / ConvertTo-XlsWriteCellValue /
ConvertTo-XlsValue2Grid / ConvertFrom-XlsObjectArray2DToRows / ConvertFrom-XlsHeaderObjectsToRows /
ConvertTo-XlsWriteRows / ConvertFrom-XlsCsvText / ConvertFrom-XlsCsvFieldValue / Import-XlsCsvRows /
Import-XlsJsonRows の受け入れテスト。
harness/state/tasks/T-09.md の受け入れテスト要点:
  - 往復テスト: Set → Get で 2 次元配列・CSV・JSON が同値。
  - 日付列が ISO 文字列で往復（Date1904 ブックも）。
  - 空セルが $null で往復。
  - 左上セル指定の自動拡張。
  - 異常系: 不正な範囲、寸法不一致（明示範囲とデータ寸法が矛盾）の扱い。

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中
（Set-XlsRange/Get-XlsRange 自体を含む）でのみ触る。フィクスチャ作成は Common.ps1 の
New-XlsObjectArray2D / New-XlsSingleCellArray で作った [object[,]]（G-08）を ScriptBlock 内で
Range.Value2 に代入する形で行う。

純粋関数は COM を一切使わないため、他の *.Tests.ps1 と同じ `& $script:ModuleRef { ... }` パターンで
モジュールスコープに入って直接テストする。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

function New-XlsAgentTempFilePath {
    param([Parameter(Mandatory = $true)][string]$Extension)

    $dir = Join-Path $env:TEMP 'xlsagent-tests'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Join-Path $dir ("{0}.{1}" -f ([guid]::NewGuid().ToString()), $Extension)
}

Describe 'Set-XlsRange (T-09)' {

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

    Context 'Test-XlsIsoDateString (pure function, no COM)' {

        It 'matches date-only and date-time ISO strings' {
            (& $script:ModuleRef { Test-XlsIsoDateString -Text '2024-01-01' }) | Should Be $true
            (& $script:ModuleRef { Test-XlsIsoDateString -Text '2024-01-01T12:34:56' }) | Should Be $true
        }

        It 'does not match non-ISO-shaped text' {
            $cases = @('2024-1-1', 'hello', '45000', '', '2024/01/01', '2024-01-01T12:34')
            foreach ($c in $cases) {
                (& $script:ModuleRef { param($t) Test-XlsIsoDateString -Text $t } $c) | Should Be $false
            }
        }
    }

    Context 'ConvertFrom-XlsIsoDateStringToSerial (pure function, no COM)' {

        # round-trip boundary の期待値は tests/Get-XlsRange.Tests.ps1 の
        # '1900/1904 date system boundaries' と対になる（ConvertTo-XlsIsoDateString の逆変換）。
        It 'converts 1900-01-01 back to serial 1 (the +1 correction boundary)' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-01-01' }) | Should Be 1.0
        }

        It 'converts 1900-02-28 back to serial 59' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-02-28' }) | Should Be 59.0
        }

        It 'converts the fictitious 1900-02-29 back to serial 60' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-02-29' }) | Should Be 60.0
        }

        It 'converts 1900-02-29T12:00:00 back to serial 60.5' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-02-29T12:00:00' }) | Should Be 60.5
        }

        It 'converts 1900-03-01 back to serial 61 (no correction needed)' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-03-01' }) | Should Be 61.0
        }

        It 'converts a Date1904-system ISO string back using the -1462 offset' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1904-01-01' -Date1904 }) | Should Be 0.0
        }

        It 'does not apply the 1900 fictitious-leap-day correction when Date1904 is set' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1904-02-29' -Date1904 }) | Should Be 59.0
        }

        It 'returns $null for a string that matches the ISO shape but is not a real date' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '2024-02-30' }) | Should Be $null
        }

        # round 1 レビュー should-fix 対応: 架空日 1900-02-29 の時刻部が「2 桁:2 桁:2 桁」という
        # 正規表現の形には一致するが実在しない時刻（例 99:99:99）の場合、以前は
        # [TimeSpan]::ParseExact の未捕捉の生例外が漏れていた。他の実在しない日付（2024-02-30 等）と
        # 同じく $null を返す（呼び出し元がそのまま元の文字列を書き込む）。
        It 'returns $null (does not throw) for the fictitious 1900-02-29 with an invalid time-of-day' {
            (& $script:ModuleRef { ConvertFrom-XlsIsoDateStringToSerial -Text '1900-02-29T99:99:99' }) | Should Be $null
        }

        # round 1 レビュー should-fix 対応: この関数は ConvertTo-XlsIsoDateString の厳密な逆関数では
        # ない。保証できるのは「Get-XlsRange が実際に出力しうる、秒単位・正準（canonical）な OADate」
        # についてのみ（関数 .NOTES 参照）。以下はその保証範囲内の代表値（Get 側が出力する形式に
        # サブ秒が含まれず、かつ正準な OADate であることが分かっている値）で、Get→Set が逆変換になる
        # ことを確認する回帰であり、任意の [double] に対する全単射を主張するものではない
        # （テスト名を「arbitrary」から変更）。
        It 'round-trips second-granularity canonical serials representable by Get through ConvertTo-XlsIsoDateString and back (1900 system)' {
            $serials = @(1.0, 2.0, 30.0, 59.0, 61.0, 100.0, 45000.0, 45000.5)
            foreach ($s in $serials) {
                $iso = & $script:ModuleRef { param($v) ConvertTo-XlsIsoDateString -OleAutomationValue $v } $s
                $back = & $script:ModuleRef { param($t) ConvertFrom-XlsIsoDateStringToSerial -Text $t } $iso
                $back | Should Be $s
            }
        }

        It 'round-trips second-granularity canonical serials representable by Get through ConvertTo-XlsIsoDateString and back (Date1904 system)' {
            $serials = @(0.0, 1.0, 59.0, 100.0, 45000.0)
            foreach ($s in $serials) {
                $iso = & $script:ModuleRef { param($v) ConvertTo-XlsIsoDateString -OleAutomationValue $v -Date1904 } $s
                $back = & $script:ModuleRef { param($t) ConvertFrom-XlsIsoDateStringToSerial -Text $t -Date1904 } $iso
                $back | Should Be $s
            }
        }
    }

    Context 'ConvertTo-XlsWriteCellValue (pure function, no COM)' {

        It 'passes $null through unchanged' {
            (& $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value $null }) | Should Be $null
        }

        It 'converts an ISO date string to its numeric serial' {
            (& $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value '1900-01-01' }) | Should Be 1.0
        }

        # round 1 レビュー should-fix 対応の回帰: 架空日の不正な時刻は例外にせず、そのまま元の文字列を
        # 返す（ConvertFrom-XlsIsoDateStringToSerial が $null を返す経路をここでも確認する）。
        It 'does not throw and leaves the original text unchanged for the fictitious 1900-02-29 with an invalid time-of-day' {
            $result = & $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value '1900-02-29T99:99:99' }
            $result | Should Be '1900-02-29T99:99:99'
        }

        It 'leaves a non-ISO string unchanged' {
            (& $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value 'hello world' }) | Should Be 'hello world'
        }

        It 'leaves booleans unchanged' {
            (& $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value $true }) | Should Be $true
            (& $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value $false }) | Should Be $false
        }

        It 'recasts integer and double values to [double]' {
            $r1 = & $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value ([int]42) }
            $r1 | Should Be 42.0
            ($r1 -is [double]) | Should Be $true

            $r2 = & $script:ModuleRef { ConvertTo-XlsWriteCellValue -Value ([double]3.5) }
            $r2 | Should Be 3.5
            ($r2 -is [double]) | Should Be $true
        }
    }

    Context 'ConvertTo-XlsValue2Grid (pure function, no COM)' {

        It 'builds an [object[,]] with the correct dimensions and converted values' {
            $rows = @(@(1, 'a', $null), @('1900-01-01', $true, 2.5))
            $grid = & $script:ModuleRef { param($r) ConvertTo-XlsValue2Grid -Rows $r } $rows

            $grid.GetLength(0) | Should Be 2
            $grid.GetLength(1) | Should Be 3
            $grid[0, 0] | Should Be 1.0
            $grid[0, 1] | Should Be 'a'
            $grid[0, 2] | Should Be $null
            $grid[1, 0] | Should Be 1.0
            $grid[1, 1] | Should Be $true
            $grid[1, 2] | Should Be 2.5
        }

        It 'throws a next-step error when row lengths are inconsistent' {
            $rows = @(@(1, 2, 3), @(4, 5))
            $errorSeen = $null
            try {
                & $script:ModuleRef { param($r) ConvertTo-XlsValue2Grid -Rows $r } $rows
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Row 2'
        }
    }

    Context 'ConvertFrom-XlsObjectArray2DToRows (pure function, no COM)' {

        It 'converts a 2D array into a jagged array of rows' {
            $arr = New-XlsObjectArray2D -Rows @(@(1.0, 2.0), @(3.0, 4.0))
            $rows = & $script:ModuleRef { param($a) ConvertFrom-XlsObjectArray2DToRows -Array $a } $arr

            $rows.Count | Should Be 2
            $rows[0][0] | Should Be 1.0
            $rows[0][1] | Should Be 2.0
            $rows[1][0] | Should Be 3.0
            $rows[1][1] | Should Be 4.0
        }
    }

    Context 'ConvertFrom-XlsHeaderObjectsToRows (pure function, no COM)' {

        It 'rebuilds a header row and data rows from an array of PSCustomObjects, preserving key order' {
            $objs = @(
                [PSCustomObject][ordered]@{ Name = 'Alice'; Age = 30.0 },
                [PSCustomObject][ordered]@{ Name = 'Bob'; Age = 41.0 }
            )
            $rows = & $script:ModuleRef { param($i) ConvertFrom-XlsHeaderObjectsToRows -Items $i } $objs

            $rows.Count | Should Be 3
            $rows[0][0] | Should Be 'Name'
            $rows[0][1] | Should Be 'Age'
            $rows[1][0] | Should Be 'Alice'
            $rows[1][1] | Should Be 30.0
            $rows[2][0] | Should Be 'Bob'
        }

        It 'accepts an array of ordered hashtables' {
            $objs = @([ordered]@{ X = 1.0; Y = 2.0 })
            $rows = & $script:ModuleRef { param($i) ConvertFrom-XlsHeaderObjectsToRows -Items $i } $objs

            $rows.Count | Should Be 2
            $rows[0][0] | Should Be 'X'
            $rows[1][0] | Should Be 1.0
        }

        It 'returns an empty jagged array for an empty item collection' {
            $rows = & $script:ModuleRef { ConvertFrom-XlsHeaderObjectsToRows -Items @() }
            , $rows | Should Not Be $null
            $rows.Count | Should Be 0
        }

        It 'throws a next-step error when the first element has no named properties' {
            $errorSeen = $null
            try {
                & $script:ModuleRef { ConvertFrom-XlsHeaderObjectsToRows -Items @(1, 2, 3) }
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match '-Header'
        }
    }

    Context 'ConvertTo-XlsWriteRows (pure function, no COM)' {

        It 'dispatches a 2D array to ConvertFrom-XlsObjectArray2DToRows' {
            # 罠: `@(@(1.0, 2.0))` は内側の配列をそのまま展開して 2 要素の配列に「平坦化」されて
            # しまう（tests/Get-XlsRange.Tests.ps1 の `@(, @(...))` パターンと同じ理由）。単一行の
            # ジャグ配列を作るには `,` で明示的に 1 段階包む。
            $arr = New-XlsObjectArray2D -Rows @(, @(1.0, 2.0))
            $rows = & $script:ModuleRef { param($d) ConvertTo-XlsWriteRows -Data $d } $arr
            $rows.Count | Should Be 1
            $rows[0][1] | Should Be 2.0
        }

        It 'passes a jagged array of rows through unchanged when -Header is not set' {
            $data = @(@(1, 2), @(3, 4))
            $rows = & $script:ModuleRef { param($d) ConvertTo-XlsWriteRows -Data $d } $data
            $rows.Count | Should Be 2
            $rows[1][0] | Should Be 3
        }

        It 'dispatches to header reconstruction when -Header is set' {
            $data = @([PSCustomObject][ordered]@{ A = 1.0; B = 2.0 })
            $rows = & $script:ModuleRef { param($d) ConvertTo-XlsWriteRows -Data $d -Header } $data
            $rows.Count | Should Be 2
            $rows[0][0] | Should Be 'A'
        }

        It 'throws when -Data is $null' {
            $errorSeen = $null
            try {
                & $script:ModuleRef { ConvertTo-XlsWriteRows -Data $null }
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
        }

        It 'throws a next-step error when a row element is a bare string' {
            $errorSeen = $null
            try {
                & $script:ModuleRef { ConvertTo-XlsWriteRows -Data @('not-a-row') }
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match '-Header'
        }
    }

    Context 'ConvertFrom-XlsCsvText (pure function, no COM)' {

        It 'parses simple comma-separated rows terminated by CRLF' {
            $rows = & $script:ModuleRef { ConvertFrom-XlsCsvText -Text "a,b`r`nc,d`r`n" }
            $rows.Count | Should Be 2
            $rows[0][0] | Should Be 'a'
            $rows[0][1] | Should Be 'b'
            $rows[1][1] | Should Be 'd'
        }

        It 'flushes a trailing row that has no terminating newline' {
            $rows = & $script:ModuleRef { ConvertFrom-XlsCsvText -Text 'a,b' }
            $rows.Count | Should Be 1
            $rows[0][1] | Should Be 'b'
        }

        It 'returns an empty jagged array for empty text' {
            $rows = & $script:ModuleRef { ConvertFrom-XlsCsvText -Text '' }
            , $rows | Should Not Be $null
            $rows.Count | Should Be 0
        }

        It 'round-trips a value containing a comma, a double quote, and a CRLF newline through ConvertTo-XlsCsvField' {
            # Get-XlsRange.Tests.ps1 の ConvertTo-XlsCsvField テストと同じ値。書き出し側の
            # エスケープ規則をそのまま読み戻せることを確認する。
            $tricky = "a,b`r`n`"c"
            $escaped = & $script:ModuleRef { param($v) ConvertTo-XlsCsvField -Value $v } $tricky
            $csvText = "plain,$escaped`r`n"

            $rows = & $script:ModuleRef { param($t) ConvertFrom-XlsCsvText -Text $t } $csvText
            $rows.Count | Should Be 1
            $rows[0][0] | Should Be 'plain'
            $rows[0][1] | Should Be $tricky
        }
    }

    Context 'ConvertFrom-XlsCsvFieldValue (pure function, no COM)' {

        It 'converts an empty field to $null' {
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field '' }) | Should Be $null
        }

        It 'converts exact True/False to booleans but leaves other casings as text' {
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field 'True' }) | Should Be $true
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field 'False' }) | Should Be $false
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field 'true' }) | Should Be 'true'
        }

        It 'converts an invariant-culture-parseable number to a double' {
            $result = & $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field '1234.5' }
            $result | Should Be 1234.5
            ($result -is [double]) | Should Be $true
        }

        It 'leaves an ISO date string as text (date conversion happens later)' {
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field '2024-01-01' }) | Should Be '2024-01-01'
        }

        It 'leaves ordinary text unchanged' {
            (& $script:ModuleRef { ConvertFrom-XlsCsvFieldValue -Field 'hello' }) | Should Be 'hello'
        }
    }

    Context 'Import-XlsCsvRows / Import-XlsJsonRows error handling (no COM)' {

        It 'throws a next-step error when the CSV file does not exist' {
            $missing = Join-Path $env:TEMP ("xlsagent-tests\{0}.csv" -f [guid]::NewGuid())
            $errorSeen = $null
            try {
                & $script:ModuleRef { param($p) Import-XlsCsvRows -Path $p } $missing
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'CSV file not found'
        }

        It 'throws a next-step error when the JSON file does not exist' {
            $missing = Join-Path $env:TEMP ("xlsagent-tests\{0}.json" -f [guid]::NewGuid())
            $errorSeen = $null
            try {
                & $script:ModuleRef { param($p) Import-XlsJsonRows -Path $p } $missing
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'JSON file not found'
        }

        It 'throws a next-step error when the JSON file is not valid JSON' {
            $badPath = New-XlsAgentTempFilePath -Extension 'json'
            [IO.File]::WriteAllText($badPath, 'not valid json {{{', [Text.Encoding]::UTF8)
            $errorSeen = $null
            try {
                & $script:ModuleRef { param($p) Import-XlsJsonRows -Path $p } $badPath
            }
            catch {
                $errorSeen = $_
            }
            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'parse JSON'
        }
    }

    Context 'Set-XlsRange integration (real Excel via Invoke-XlsSession)' {

        It 'writes a 2D [object[,]] array into an explicit matching range and round-trips through Get-XlsRange' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $data = New-XlsObjectArray2D -Rows @(
                    @(1.0, 'hello', $true),
                    @(2.0, $null, $false)
                )
                Set-XlsRange -Worksheet $ws -Range 'A1:C2' -Data $data
                Get-XlsRange -Worksheet $ws -Range 'A1:C2'
            }

            $result.Count | Should Be 2
            $result[0][0] | Should Be 1.0
            $result[0][1] | Should Be 'hello'
            $result[0][2] | Should Be $true
            $result[1][0] | Should Be 2.0
            $result[1][1] | Should Be $null
            $result[1][2] | Should Be $false
        }

        It 'writes a jagged array (Get-XlsRange''s own return shape) and round-trips' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(10.0, 'x'),
                    @(20.0, 'y')
                ))
                $src = Get-XlsRange -Worksheet $ws -Range 'A1:B2'

                Set-XlsRange -Worksheet $ws -Range 'D1:E2' -Data $src
                Get-XlsRange -Worksheet $ws -Range 'D1:E2'
            }

            $result[0][0] | Should Be 10.0
            $result[0][1] | Should Be 'x'
            $result[1][0] | Should Be 20.0
            $result[1][1] | Should Be 'y'
        }

        It 'writes a PSCustomObject array (Get-XlsRange -Header output) with -Header and round-trips' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Name', 'Score'),
                    @('Alice', 10.0),
                    @('Bob', 20.0)
                ))
                $srcObjs = Get-XlsRange -Worksheet $ws -Range 'A1:B3' -Header

                Set-XlsRange -Worksheet $ws -Range 'D1:E3' -Data $srcObjs -Header
                Get-XlsRange -Worksheet $ws -Range 'D1:E3' -Header
            }

            $result.Count | Should Be 2
            $result[0].Name | Should Be 'Alice'
            $result[0].Score | Should Be 10.0
            $result[1].Name | Should Be 'Bob'
        }

        It 'auto-resizes from a single top-left cell to the data dimensions' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $data = @(@(1.0, 2.0), @(3.0, 4.0), @(5.0, 6.0))

                Set-XlsRange -Worksheet $ws -Range 'B2' -Data $data
                Get-XlsRange -Worksheet $ws -Range 'B2:C4'
            }

            $result.Count | Should Be 3
            $result[0][0] | Should Be 1.0
            $result[0][1] | Should Be 2.0
            $result[2][0] | Should Be 5.0
            $result[2][1] | Should Be 6.0
        }

        It 'writes and round-trips a single-cell auto-resized write (1x1 data into a single top-left cell)' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsRange -Worksheet $ws -Range 'C3' -Data @(, @('solo'))
                Get-XlsRange -Worksheet $ws -Range 'C3'
            }

            $result[0][0] | Should Be 'solo'
        }

        It 'writes $null as empty cells and round-trips through Get-XlsRange' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Set-XlsRange -Worksheet $ws -Range 'A1:B2' -Data @(@(1.0, $null), @($null, 4.0))
                Get-XlsRange -Worksheet $ws -Range 'A1:B2'
            }

            $result[0][0] | Should Be 1.0
            $result[0][1] | Should Be $null
            $result[1][0] | Should Be $null
            $result[1][1] | Should Be 4.0
        }

        It 'writes the correct raw Value2 serial for ISO date strings at the 1900/1904 boundaries (pre-formatted target cell)' {
            $path = New-TempXlsxPath

            $rawValues = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)

                # NumberFormat は Set-XlsRange が触らない値のみの経路なので、往復検証のために
                # あらかじめ日付書式を設定しておく（タスクカード・実装メモの既知のトレードオフ）。
                $ws.Range('A1:A5').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
                Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(
                    @('1900-01-01'),
                    @('1900-02-28'),
                    @('1900-02-29'),
                    @('1900-02-29T12:00:00'),
                    @('1900-03-01')
                )

                @(
                    [double]$ws.Range('A1').Value2,
                    [double]$ws.Range('A2').Value2,
                    [double]$ws.Range('A3').Value2,
                    [double]$ws.Range('A4').Value2,
                    [double]$ws.Range('A5').Value2
                )
            }

            $rawValues[0] | Should Be 1.0
            $rawValues[1] | Should Be 59.0
            $rawValues[2] | Should Be 60.0
            $rawValues[3] | Should Be 60.5
            $rawValues[4] | Should Be 61.0
        }

        It 'writes the correct raw Value2 serial for ISO date strings in a Date1904 workbook' {
            $path = New-TempXlsxPath

            $rawValue = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $wb.Date1904 = $true
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').NumberFormat = 'yyyy-mm-dd'

                Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(, @('1904-01-01'))

                [double]$ws.Range('A1').Value2
            }

            $rawValue | Should Be 0.0
        }

        # round 1 レビュー should-fix 対応: 上のテストは生の Value2 だけを確認しており、タスクカードが
        # 求める「Date1904 ブックでの Set-XlsRange -> Get-XlsRange の往復」（公開関数経由）を検証して
        # いなかった。日付のみ・架空閏日・時刻付き日時の 3 パターンで公開関数経由の往復を確認する
        # （レビュー指摘のスニペットに準拠）。
        It 'round-trips date-only, fictitious-leap-day, and date-time ISO strings through the public Set-XlsRange -> Get-XlsRange path in a Date1904 workbook' {
            $path = New-TempXlsxPath

            $actual = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $wb.Date1904 = $true
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:A3').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
                $expected = @('1904-01-01', '1904-02-29', '2024-01-02T12:34:56')

                Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(
                    @($expected[0]),
                    @($expected[1]),
                    @($expected[2])
                )

                Get-XlsRange -Worksheet $ws -Range 'A1:A3'
            }

            $actual[0][0] | Should Be '1904-01-01'
            $actual[1][0] | Should Be '1904-02-29'
            $actual[2][0] | Should Be '2024-01-02T12:34:56'
        }

        It 'round-trips an ISO date column through Get-XlsRange when the target cell is pre-formatted as a date' {
            $path = New-TempXlsxPath

            $iso = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').NumberFormat = 'yyyy-mm-dd'
                Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(, @('2023-06-15'))
                Get-XlsRange -Worksheet $ws -Range 'A1'
            }

            $iso[0][0] | Should Be '2023-06-15'
        }

        It 'does not modify the NumberFormat of the target range' {
            $path = New-TempXlsxPath

            $formats = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').NumberFormat = '0.00%'
                $before = [string]$ws.Range('A1').NumberFormat

                Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(, @(5.0))

                $after = [string]$ws.Range('A1').NumberFormat
                @($before, $after)
            }

            $formats[0] | Should Be '0.00%'
            $formats[1] | Should Be '0.00%'
        }

        It 'round-trips numbers, comma-containing text, booleans, null, and a pre-formatted date column through -FromCsv' {
            $path = New-TempXlsxPath
            $csvPath = New-XlsAgentTempFilePath -Extension 'csv'

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)

                $ws.Range('A1:D3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Tokyo', 10.5, $true, 45000.0),
                    @('Osaka,Japan', 20.0, $false, 45010.0),
                    @($null, 0.0, $true, 45020.5)
                ))
                $ws.Range('D1:D3').NumberFormat = 'yyyy-mm-dd hh:mm:ss'

                $expected = Get-XlsRange -Worksheet $ws -Range 'A1:D3'
                Get-XlsRange -Worksheet $ws -Range 'A1:D3' -AsCsv $csvPath | Out-Null

                # 書き込み先はあらかじめ日付列だけ同じ書式にしておく（NumberFormat は Set-XlsRange が
                # 触らないため。実装メモ参照）。
                $ws.Range('I1:I3').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
                Set-XlsRange -Worksheet $ws -Range 'F1' -FromCsv $csvPath

                $actual = Get-XlsRange -Worksheet $ws -Range 'F1:I3'
                @($expected, $actual)
            }

            $expected = $result[0]
            $actual = $result[1]

            for ($r = 0; $r -lt 3; $r++) {
                for ($c = 0; $c -lt 4; $c++) {
                    $actual[$r][$c] | Should Be $expected[$r][$c]
                }
            }

            Test-Path -LiteralPath $csvPath | Should Be $true
        }

        It 'round-trips a jagged-array JSON export (no -Header) through -FromJson' {
            $path = New-TempXlsxPath
            $jsonPath = New-XlsAgentTempFilePath -Extension 'json'

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)

                $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(1.0, 'first'),
                    @($null, 'second')
                ))
                $expected = Get-XlsRange -Worksheet $ws -Range 'A1:B2'
                Get-XlsRange -Worksheet $ws -Range 'A1:B2' -AsJson $jsonPath | Out-Null

                Set-XlsRange -Worksheet $ws -Range 'D1' -FromJson $jsonPath
                $actual = Get-XlsRange -Worksheet $ws -Range 'D1:E2'
                @($expected, $actual)
            }

            $expected = $result[0]
            $actual = $result[1]
            $actual[0][0] | Should Be $expected[0][0]
            $actual[0][1] | Should Be $expected[0][1]
            $actual[1][0] | Should Be $expected[1][0]
            $actual[1][1] | Should Be $expected[1][1]
        }

        It 'round-trips an object-array JSON export (-Header) through -FromJson -Header' {
            $path = New-TempXlsxPath
            $jsonPath = New-XlsAgentTempFilePath -Extension 'json'

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)

                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Name', 'Score'),
                    @('Alice', 10.0),
                    @('Bob', 20.0)
                ))
                Get-XlsRange -Worksheet $ws -Range 'A1:B3' -Header -AsJson $jsonPath | Out-Null

                Set-XlsRange -Worksheet $ws -Range 'D1' -FromJson $jsonPath -Header
                Get-XlsRange -Worksheet $ws -Range 'D1:E3' -Header
            }

            $result.Count | Should Be 2
            $result[0].Name | Should Be 'Alice'
            $result[0].Score | Should Be 10.0
            $result[1].Name | Should Be 'Bob'
            $result[1].Score | Should Be 20.0
        }

        It 'throws a next-step error for an invalid range address' {
            $path = New-TempXlsxPath

            $errorSeen = $null
            try {
                Invoke-XlsSession -Path $path -ScriptBlock {
                    param($app, $wb)
                    $ws = $wb.Worksheets.Item(1)
                    Set-XlsRange -Worksheet $ws -Range 'NotAValidAddr$$' -Data @(, @(1.0))
                }
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Invalid range address'
        }

        It 'throws a next-step error when an explicit multi-cell range does not match the data dimensions' {
            $path = New-TempXlsxPath

            $errorSeen = $null
            try {
                Invoke-XlsSession -Path $path -ScriptBlock {
                    param($app, $wb)
                    $ws = $wb.Worksheets.Item(1)
                    Set-XlsRange -Worksheet $ws -Range 'A1:B2' -Data @(@(1.0, 2.0, 3.0), @(4.0, 5.0, 6.0), @(7.0, 8.0, 9.0))
                }
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match '2 x 2'
            $errorSeen.Exception.Message | Should Match '3 x 3'
        }

        It 'throws when row lengths in -Data are inconsistent (no partial write)' {
            $path = New-TempXlsxPath

            $errorSeen = $null
            try {
                Invoke-XlsSession -Path $path -ScriptBlock {
                    param($app, $wb)
                    $ws = $wb.Worksheets.Item(1)
                    Set-XlsRange -Worksheet $ws -Range 'A1' -Data @(@(1.0, 2.0), @(3.0))
                }
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Row 2'
        }
    }
}
