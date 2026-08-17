#Requires -Version 5.1

<#
tests/Get-XlsRange.Tests.ps1

T-06: Get-XlsRange（一括読み出しの唯一の経路。01-design.md §3.5）と、値変換規則を集約した内部関数
ConvertFrom-XlsValue2 / Test-XlsIsDateNumberFormat / ConvertTo-XlsHeaderObjects / ConvertTo-XlsCsvField /
Export-XlsRowsToCsv / Export-XlsJson の受け入れテスト。
harness/state/tasks/T-06.md の受け入れテスト要点:
  - 数値/文字/日付(ISO 文字列)/空セル($null)/エラー値が正しく読める。
  - CSV 出力: UTF-8 BOM 付き、日本語文字が壊れない、Excel の日付が ISO 文字列になっている。
  - JSON 出力: -Header 指定でオブジェクト配列。
  - 単一セル・単一行・単一列の範囲が正しく扱える。
  - 異常系: 存在しない範囲指定(無効なアドレス)で次の一手が分かる例外。

G-06: COM はすべて Invoke-XlsSession の ScriptBlock 内、または Invoke-XlsSession 自身の中でのみ触る。
このテストファイルも例外ではない。テストフィクスチャ（多セルへの書き込み）は T-09（Set-XlsRange）が
まだ実装されていないため、Common.ps1 の New-XlsObjectArray2D / New-XlsSingleCellArray で作った
[object[,]]（G-08）を ScriptBlock 内で直接 Range.Value2 に代入して用意する。

純粋関数（ConvertFrom-XlsValue2 等）は COM を一切使わないため、Save-XlsWorkbook.Tests.ps1 と同じ
`& $script:ModuleRef { ... }` パターンでモジュールスコープに入って直接テストする。
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Common.ps1')

$moduleName = 'XlsAgent'
$modulePath = Join-Path $here '..\skill\scripts\XlsAgent.psm1'

Describe 'Get-XlsRange (T-06)' {

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

    Context 'Test-XlsIsDateNumberFormat (pure function, no COM)' {

        It 'recognizes common US date/time format codes as date formats' {
            $cases = @('m/d/yyyy', 'mm/dd/yyyy', 'yyyy-mm-dd', 'yyyy-mm-dd hh:mm:ss', 'h:mm AM/PM', 'mmm yyyy', 'hh:mm:ss')
            foreach ($fmt in $cases) {
                $result = & $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f } $fmt
                $result | Should Be $true
            }
        }

        It 'does not treat the localized General format as a date' {
            # 罠（実機確認）: 日本語ロケールの Excel では既定書式が英語の 'General' ではなく
            # 'G/標準' として Range.NumberFormat から返ってくる。
            $cases = @('General', 'G/標準', '')
            foreach ($fmt in $cases) {
                $result = & $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f } $fmt
                $result | Should Be $false
            }
        }

        It 'does not treat numeric/currency/percent/scientific/text formats as date' {
            $cases = @('0.00', '$#,##0.00', '0%', '0.00E+00', '@', '#,##0')
            foreach ($fmt in $cases) {
                $result = & $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f } $fmt
                $result | Should Be $false
            }
        }

        It 'strips quoted literals and bracketed color/locale codes before checking for date tokens' {
            # [Red] や [Yellow] は色名の中に d/y を含むため、角括弧を除去しないと誤検出する。
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[Red]0.00' }) | Should Be $false
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[Yellow]General' }) | Should Be $false
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[Red]m/d/yyyy' }) | Should Be $true
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '"days" 0' }) | Should Be $false
        }

        # round 2 レビュー指摘（blocking）対応: 複数セクション書式・エスケープ・経過時間トークンの
        # 回帰テスト。
        It 'treats a multi-section format as a date when the section for the value''s sign is a date section' {
            # 'm/d/yyyy;@' は 2 セクション（正数&ゼロ=日付、負数=テキスト）。値を渡さない（既定 0=ゼロ）
            # 場合は 1 番目（日付）セクションが選ばれる。
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat 'm/d/yyyy;@' }) | Should Be $true
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat 'm/d/yyyy;@' -Value 45000 }) | Should Be $true
        }

        It 'does not treat a backslash-escaped literal character as a date token' {
            # '0\m' は「0 の後にリテラルの m を1文字表示する」書式であり、月を表す m ではない。
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '0\m' }) | Should Be $false
        }

        It 'treats a bracketed elapsed-time token alone as a date/time format' {
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[h]' }) | Should Be $true
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[hh]:mm:ss' }) | Should Be $true
            (& $script:ModuleRef { Test-XlsIsDateNumberFormat -NumberFormat '[mm]:ss' }) | Should Be $true
        }

        It 'selects the negative-number section based on the sign of -Value' {
            # 3 セクション: 正数=日付ではない、負数=日付、ゼロ=日付ではない。
            $fmt = '0.00;m/d/yyyy;0.00'
            (& $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f -Value 5 } $fmt) | Should Be $false
            (& $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f -Value -5 } $fmt) | Should Be $true
            (& $script:ModuleRef { param($f) Test-XlsIsDateNumberFormat -NumberFormat $f -Value 0 } $fmt) | Should Be $false
        }

        # round 2 レビュー指摘（blocking）対応: 条件付き書式（[<100] 等）はセクションの符号規則を
        # 置き換える。値の符号だけで選ぶと誤判定する回帰テスト。
        It 'selects a conditional section by its [<N]/[>=N] condition instead of the value''s sign' {
            (& $script:ModuleRef {
                Test-XlsIsDateNumberFormat -NumberFormat '[<100]0.00;[>=100]yyyy-mm-dd' -Value 45000
            }) | Should Be $true

            (& $script:ModuleRef {
                Test-XlsIsDateNumberFormat -NumberFormat '[<100]yyyy-mm-dd;[>=100]0.00' -Value 45000
            }) | Should Be $false
        }

        It 'falls back to the unconditional section when no condition matches the value' {
            # [<100] だけが条件、2 番目は無条件（catch-all）。45000 は条件に一致しないので 2 番目
            # （日付）が使われる。
            (& $script:ModuleRef {
                Test-XlsIsDateNumberFormat -NumberFormat '[<100]0.00;yyyy-mm-dd' -Value 45000
            }) | Should Be $true

            (& $script:ModuleRef {
                Test-XlsIsDateNumberFormat -NumberFormat '[<100]0.00;yyyy-mm-dd' -Value 50
            }) | Should Be $false
        }

        # round 3 レビュー指摘（blocking）対応: 全セクションが条件付きで、どの条件にも一致せず、
        # 無条件の catch-all セクションもない場合、偽だった条件付きセクションへ戻ってはいけない
        # （$Sections[0] への機械的なフォールバックは誤変換になる）。
        It 'is not a date when every section is conditional and none of the conditions match the value' {
            # '[<0]yyyy-mm-dd;[>100]0.00' + 50: 50 は 0 未満でも 100 超でもないため、
            # どちらのセクションも適用されない。第 1 セクション（日付）へ戻って誤変換してはいけない。
            (& $script:ModuleRef {
                Test-XlsIsDateNumberFormat -NumberFormat '[<0]yyyy-mm-dd;[>100]0.00' -Value 50
            }) | Should Be $false
        }
    }

    Context 'ConvertFrom-XlsValue2 (pure function, no COM)' {

        It 'returns $null for an empty cell (Value2 = $null)' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 $null -NumberFormat 'General' }
            $result | Should Be $null
        }

        It 'passes through non-date numbers, strings, and booleans unchanged' {
            (& $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 42.0 -NumberFormat 'General' }) | Should Be 42.0
            (& $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 'hello' -NumberFormat 'General' }) | Should Be 'hello'
            (& $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 $true -NumberFormat 'General' }) | Should Be $true
        }

        It 'converts all 7 known Value2 error codes to their display strings' {
            $expected = @{
                xlErrNull  = '#NULL!'
                xlErrDiv0  = '#DIV/0!'
                xlErrValue = '#VALUE!'
                xlErrRef   = '#REF!'
                xlErrName  = '#NAME?'
                xlErrNum   = '#NUM!'
                xlErrNA    = '#N/A'
            }
            foreach ($key in $expected.Keys) {
                $code = & $script:ModuleRef { param($k) $script:Xl[$k] } $key
                $result = & $script:ModuleRef { param($c) ConvertFrom-XlsValue2 -Value2 ([int]$c) -NumberFormat 'General' } $code
                $result | Should Be $expected[$key]
            }
        }

        It 'falls back to #ERR(code) for an unrecognized negative int error code' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 ([int]-1) -NumberFormat 'General' }
            $result | Should Be '#ERR(-1)'
        }

        It 'converts a date-only OADate value (midnight) to a plain yyyy-MM-dd string' {
            $expected = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 45000.0 -NumberFormat 'm/d/yyyy' }
            $result | Should Be $expected
        }

        It 'converts a datetime OADate value (with time) to a yyyy-MM-ddTHH:mm:ss string' {
            $expected = [DateTime]::FromOADate(45000.5).ToString('yyyy-MM-ddTHH:mm:ss')
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 45000.5 -NumberFormat 'yyyy-mm-dd hh:mm:ss' }
            $result | Should Be $expected
        }

        It 'does not treat a plain number stored with a General format as a date' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 45000.0 -NumberFormat 'General' }
            $result | Should Be 45000.0
        }
    }

    # round 2 レビュー指摘（blocking）対応: 1900 年 1〜2 月の Excel/OADate 差異と、Date1904 の
    # 境界を固定する回帰テスト（実機で Range.Text と比較して確認した値。実装メモ参照）。
    # 罠: Pester 3.4 は Context のネスト（Context の中に Context）でメタエラーになることを実機で
    # 確認したため、Describe 直下の兄弟 Context にした。
    Context '1900/1904 date system boundaries (round 2 blocking fix)' {

        It 'converts serial 1 (1900-01-01, the OADate/Excel 1-day-off boundary) correctly' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 1.0 -NumberFormat 'yyyy-mm-dd' }
            $result | Should Be '1900-01-01'
        }

        It 'converts serial 59 (1900-02-28, still inside the +1 correction range) correctly' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 59.0 -NumberFormat 'yyyy-mm-dd' }
            $result | Should Be '1900-02-28'
        }

        It 'converts serial 60 (the fictitious 1900-02-29) to the literal string Excel displays' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 60.0 -NumberFormat 'yyyy-mm-dd' }
            $result | Should Be '1900-02-29'
        }

        It 'converts serial 60.5 (fictitious 1900-02-29 with a time-of-day) with the time appended' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 60.5 -NumberFormat 'yyyy-mm-dd hh:mm:ss' }
            $result | Should Be '1900-02-29T12:00:00'
        }

        It 'converts serial 61 (1900-03-01, no correction needed) correctly' {
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 61.0 -NumberFormat 'yyyy-mm-dd' }
            $result | Should Be '1900-03-01'
        }

        It 'converts a Date1904-system serial using the +1462 offset' {
            # 実機確認: Date1904=$true のワークブックで serial 0 は 1904-01-01（実装メモ参照）。
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 0.0 -NumberFormat 'yyyy-mm-dd' -Date1904 }
            $result | Should Be '1904-01-01'
        }

        It 'does not apply the 1900 fictitious-leap-day correction when Date1904 is set' {
            # Date1904 のときは 1462 加算だけでよく、1900 年 1〜2 月特有の +1 補正は適用しない。
            $result = & $script:ModuleRef { ConvertFrom-XlsValue2 -Value2 59.0 -NumberFormat 'yyyy-mm-dd' -Date1904 }
            $result | Should Be '1904-02-29'
        }
    }

    Context 'ConvertTo-XlsHeaderObjects (pure function, no COM)' {

        It 'converts row 1 + data rows into PSCustomObjects keyed by the header row' {
            $rows = @(
                @('Name', 'Age'),
                @('Alice', 30.0),
                @('Bob', 41.0)
            )
            $result = & $script:ModuleRef { param($r) ConvertTo-XlsHeaderObjects -Rows $r } $rows

            $result.Count | Should Be 2
            $result[0].Name | Should Be 'Alice'
            $result[0].Age | Should Be 30.0
            $result[1].Name | Should Be 'Bob'
        }

        It 'falls back to ColumnN for a null or empty header cell' {
            $rows = @(
                @('Name', $null, ''),
                @('Alice', 1.0, 2.0)
            )
            $result = & $script:ModuleRef { param($r) ConvertTo-XlsHeaderObjects -Rows $r } $rows

            # 罠: Pester 3.4 の `Should Contain` は「コレクションに要素が含まれるか」ではなく
            # 「(パイプ元を)ファイルパスとみなしてその中身に文字列/正規表現がマッチするか」を見る
            # アサーションのため、ここでは使えない（実機確認、$props をファイルパスとして開こうとして
            # 例外になった）。配列の要素判定は -contains 演算子を使う。
            $props = $result[0].PSObject.Properties.Name
            ($props -contains 'Name') | Should Be $true
            ($props -contains 'Column2') | Should Be $true
            ($props -contains 'Column3') | Should Be $true
        }

        It 'returns an empty array when the range is only a header row' {
            $rows = , @('Name', 'Age')
            $result = & $script:ModuleRef { param($r) ConvertTo-XlsHeaderObjects -Rows $r } $rows

            , $result | Should Not Be $null
            $result.Count | Should Be 0
        }
    }

    Context 'ConvertTo-XlsCsvField (pure function, no COM)' {

        It 'renders $null as an empty field' {
            (& $script:ModuleRef { ConvertTo-XlsCsvField -Value $null }) | Should Be ''
        }

        It 'quotes and escapes a value containing a comma, a double quote, and a newline' {
            # nit（round 2 レビュー指摘対応）: テスト名どおり、実データに CRLF の改行を含める。
            $value = "a,b`r`n`"c"
            $result = & $script:ModuleRef { param($v) ConvertTo-XlsCsvField -Value $v } $value
            $result | Should Be "`"a,b`r`n`"`"c`""
        }

        It 'formats a double using invariant culture without locale-dependent separators' {
            $result = & $script:ModuleRef { ConvertTo-XlsCsvField -Value ([double]1234.5) }
            $result | Should Be '1234.5'
        }

        It 'renders booleans as True/False' {
            (& $script:ModuleRef { ConvertTo-XlsCsvField -Value $true }) | Should Be 'True'
            (& $script:ModuleRef { ConvertTo-XlsCsvField -Value $false }) | Should Be 'False'
        }
    }

    Context 'Get-XlsRange integration (real Excel via Invoke-XlsSession)' {

        It 'reads numbers, strings, dates (with and without time), empty cells, and error values' {
            $path = New-TempXlsxPath
            $expectedDateOnly = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')
            $expectedDateTime = [DateTime]::FromOADate(45000.5).ToString('yyyy-MM-ddTHH:mm:ss')

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)

                $ws.Range('A1:D2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(42.0, 'hello world', 45000.0, $null),
                    @($null, $true, 45000.5, 0.0)
                ))
                $ws.Range('C1').NumberFormat = 'm/d/yyyy'
                $ws.Range('C2').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
                $ws.Range('D1').Formula = '=1/0'
                $ws.Range('D2').Formula = '=NA()'

                Get-XlsRange -Worksheet $ws -Range 'A1:D2'
            }

            $rows.Count | Should Be 2
            $rows[0][0] | Should Be 42.0
            $rows[0][1] | Should Be 'hello world'
            $rows[0][2] | Should Be $expectedDateOnly
            $rows[0][3] | Should Be '#DIV/0!'
            $rows[1][0] | Should Be $null
            $rows[1][1] | Should Be $true
            $rows[1][2] | Should Be $expectedDateTime
            $rows[1][3] | Should Be '#N/A'
        }

        It 'handles a single-cell range without leaking the raw COM scalar shape' {
            $path = New-TempXlsxPath

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('B2').Value2 = (New-XlsSingleCellArray -Value 'solo')
                Get-XlsRange -Worksheet $ws -Range 'B2'
            }

            , $rows | Should Not Be $null
            $rows.Count | Should Be 1
            $rows[0].Count | Should Be 1
            $rows[0][0] | Should Be 'solo'
        }

        It 'handles an empty single-cell range as a single $null value, not an empty jagged array' {
            $path = New-TempXlsxPath

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                Get-XlsRange -Worksheet $ws -Range 'Z99'
            }

            $rows.Count | Should Be 1
            $rows[0].Count | Should Be 1
            $rows[0][0] | Should Be $null
        }

        It 'handles a single-row multi-column range' {
            $path = New-TempXlsxPath

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:C1').Value2 = (New-XlsObjectArray2D -Rows @(, @(1.0, 2.0, 3.0)))
                Get-XlsRange -Worksheet $ws -Range 'A1:C1'
            }

            $rows.Count | Should Be 1
            $rows[0].Count | Should Be 3
            $rows[0][0] | Should Be 1.0
            $rows[0][2] | Should Be 3.0
        }

        It 'handles a single-column multi-row range' {
            $path = New-TempXlsxPath

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:A3').Value2 = (New-XlsObjectArray2D -Rows @(@(1.0), @(2.0), @(3.0)))
                Get-XlsRange -Worksheet $ws -Range 'A1:A3'
            }

            $rows.Count | Should Be 3
            $rows[0].Count | Should Be 1
            $rows[0][0] | Should Be 1.0
            $rows[2][0] | Should Be 3.0
        }

        It 'falls back to per-cell NumberFormat reads when formats are mixed within the range' {
            # A1:B2 の書式をわざと混在させ、一括 NumberFormat が DBNull になるケースを踏ませる
            # （実装メモ参照）。それでも各セルの日付判定が正しく行われることを確認する。
            $path = New-TempXlsxPath
            $expectedDate = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(45000.0, 7.0),
                    @(7.0, 45000.0)
                ))
                $ws.Range('A1').NumberFormat = 'm/d/yyyy'
                $ws.Range('B1').NumberFormat = '0.00'
                $ws.Range('A2').NumberFormat = '0.00'
                $ws.Range('B2').NumberFormat = 'm/d/yyyy'

                Get-XlsRange -Worksheet $ws -Range 'A1:B2'
            }

            $rows[0][0] | Should Be $expectedDate
            $rows[0][1] | Should Be 7.0
            $rows[1][0] | Should Be 7.0
            $rows[1][1] | Should Be $expectedDate
        }

        It 'handles non-double cells (text/bool/error) correctly in a mixed-format range (round 2 should-fix)' {
            # round 2 レビュー指摘（should-fix）対応: DBNull フォールバック時、NumberFormat のセル単位
            # 取得は [double] のセルだけに限定した。文字列・真偽値・エラー値が混ざる範囲でも
            # 正しく読めることを確認する（[double] 以外のセルで NumberFormat 取得をスキップしても
            # 結果が変わらないことの回帰）。
            $path = New-TempXlsxPath
            $expectedDate = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:C2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @(45000.0, 'text cell', $true),
                    @(7.0, $null, 7.0)
                ))
                $ws.Range('A1').NumberFormat = 'm/d/yyyy'
                $ws.Range('B1').NumberFormat = '0.00'
                $ws.Range('C1').NumberFormat = '0.00'
                $ws.Range('A2').Formula = '=1/0'
                $ws.Range('C2').NumberFormat = 'm/d/yyyy'

                Get-XlsRange -Worksheet $ws -Range 'A1:C2'
            }

            $rows[0][0] | Should Be $expectedDate
            $rows[0][1] | Should Be 'text cell'
            $rows[0][2] | Should Be $true
            $rows[1][0] | Should Be '#DIV/0!'
            $rows[1][1] | Should Be $null
        }

        It 'converts dates using the Date1904 date system when the workbook uses it (round 2 blocking fix)' {
            $path = New-TempXlsxPath

            $rows = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $wb.Date1904 = $true
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1').Value2 = (New-XlsSingleCellArray -Value 0.0)
                $ws.Range('A1').NumberFormat = 'yyyy-mm-dd'

                Get-XlsRange -Worksheet $ws -Range 'A1'
            }

            # 実機確認: Date1904=$true のワークブックで serial 0 は 1904-01-01（実装メモ参照）。
            $rows[0][0] | Should Be '1904-01-01'
        }

        It '-Header returns a PSCustomObject array keyed by the header row, excluding the header row itself' {
            $path = New-TempXlsxPath

            $result = Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Name', 'Score'),
                    @('Alice', 10.0),
                    @('Bob', 20.0)
                ))
                Get-XlsRange -Worksheet $ws -Range 'A1:B3' -Header
            }

            $result.Count | Should Be 2
            $result[0].Name | Should Be 'Alice'
            $result[0].Score | Should Be 10.0
            $result[1].Name | Should Be 'Bob'
        }

        It '-AsCsv writes a UTF-8 BOM CSV file with Japanese text and ISO date strings intact' {
            $path = New-TempXlsxPath
            $csvPath = Join-Path $env:TEMP ("xlsagent-tests\{0}.csv" -f ([guid]::NewGuid()))
            $expectedDate = [DateTime]::FromOADate(45000).ToString('yyyy-MM-dd')

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B2').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('日本語テキスト,カンマ入り', 45000.0),
                    @('two', 3.0)
                ))
                $ws.Range('B1').NumberFormat = 'm/d/yyyy'

                Get-XlsRange -Worksheet $ws -Range 'A1:B2' -AsCsv $csvPath | Out-Null
            }

            Test-Path -LiteralPath $csvPath | Should Be $true
            $bytes = [IO.File]::ReadAllBytes($csvPath)
            # UTF-8 BOM: EF BB BF
            $bytes[0] | Should Be 0xEF
            $bytes[1] | Should Be 0xBB
            $bytes[2] | Should Be 0xBF

            $text = [IO.File]::ReadAllText($csvPath, [Text.Encoding]::UTF8)
            $text | Should Match ([regex]::Escape('"日本語テキスト,カンマ入り"'))
            $text | Should Match ([regex]::Escape($expectedDate))
            $text | Should Not Match '45000'
        }

        It '-AsJson with -Header writes a JSON array of objects' {
            $path = New-TempXlsxPath
            $jsonPath = Join-Path $env:TEMP ("xlsagent-tests\{0}.json" -f ([guid]::NewGuid()))

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:B3').Value2 = (New-XlsObjectArray2D -Rows @(
                    @('Name', 'Score'),
                    @('Alice', 10.0),
                    @('Bob', 20.0)
                ))
                Get-XlsRange -Worksheet $ws -Range 'A1:B3' -Header -AsJson $jsonPath | Out-Null
            }

            Test-Path -LiteralPath $jsonPath | Should Be $true
            $bytes = [IO.File]::ReadAllBytes($jsonPath)
            $bytes[0] | Should Be 0xEF
            $bytes[1] | Should Be 0xBB
            $bytes[2] | Should Be 0xBF

            # 罠: `@(Get-Content ... | ConvertFrom-Json)` は要素数がおかしくなる（実機確認）。
            # ConvertFrom-Json は JSON 配列をパイプラインへ「展開せず 1 個のオブジェクトとして」書き出す
            # ため、`@( コマンド呼び出し )` がそのパイプライン出力をさらに 1 段階配列で包んでしまい、
            # 要素数 2 のはずが 1（中身は元の配列 1 個）になる。`@(変数)`（既に受け取った変数を包む）は
            # この問題を起こさないため、いったん変数で受けてから必要なら @() で包む。
            $rawJsonText = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8
            $parsed = ConvertFrom-Json $rawJsonText
            $parsed.Count | Should Be 2
            $parsed[0].Name | Should Be 'Alice'
            $parsed[1].Score | Should Be 20
        }

        It '-AsJson without -Header writes a JSON array of arrays, even for a single-row range' {
            $path = New-TempXlsxPath
            $jsonPath = Join-Path $env:TEMP ("xlsagent-tests\{0}.json" -f ([guid]::NewGuid()))

            Invoke-XlsSession -Path $path -ScriptBlock {
                param($app, $wb)
                $ws = $wb.Worksheets.Item(1)
                $ws.Range('A1:C1').Value2 = (New-XlsObjectArray2D -Rows @(, @(1.0, 2.0, 3.0)))
                Get-XlsRange -Worksheet $ws -Range 'A1:C1' -AsJson $jsonPath | Out-Null
            }

            $rawJson = [IO.File]::ReadAllText($jsonPath, [Text.Encoding]::UTF8)
            $parsedOuter = ConvertFrom-Json $rawJson
            # 罠（実装メモ参照）: 要素数1の配列を素朴に扱うと PowerShell がアンラップしてしまうため、
            # ここではさらに @() で確実に配列として扱う。
            $parsedRows = @($parsedOuter)
            $parsedRows.Count | Should Be 1
            @($parsedRows[0]).Count | Should Be 3
        }

        It 'throws a next-step error for an invalid range address' {
            $path = New-TempXlsxPath

            $errorSeen = $null
            try {
                Invoke-XlsSession -Path $path -ScriptBlock {
                    param($app, $wb)
                    $ws = $wb.Worksheets.Item(1)
                    Get-XlsRange -Worksheet $ws -Range 'NotAValidAddr$$'
                }
            }
            catch {
                $errorSeen = $_
            }

            $errorSeen | Should Not Be $null
            $errorSeen.Exception.Message | Should Match 'Invalid range address'
            $errorSeen.Exception.Message | Should Match ([regex]::Escape('NotAValidAddr$$'))
        }
    }
}
