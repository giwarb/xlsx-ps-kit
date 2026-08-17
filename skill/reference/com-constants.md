# COM constants (Excel)

<!-- T-13 で棚卸し・実測充填を完了。「確認」列は実行して確認した日付・タスク・確認方法を必ず入れる。
     未確認の行は残さない（G-12）。書式: 日付（T-XX、実機。<何を設定し何を読み戻して確認したか>）。 -->

`$script:Xl`（`XlsAgent.psm1`）に載っているのは「モジュール本体の公開関数が実際に使う定数」のみ。
書式・罫線・テーブル・グラフ系の定数はモジュール本体が高水準ラッパーを持たない方針（G-02）のため
`$script:Xl` には追加せず、このファイルにのみ実測結果を記録する（reference 専用）。

## 計算モード・保存形式

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlCalculationManual | -4135 | Application.Calculation | 2026-08-17（T-03、実機。Invoke-XlsSession が Workbook を開いた後に `Application.Calculation = -4135` を設定し、ScriptBlock 内で同プロパティを読み戻すと -4135 のままであることを `Invoke-XlsSession.Tests.ps1` で確認） |
| xlCalculationAutomatic | -4105 | Application.Calculation | 2026-08-17（T-04 Save-XlsWorkbook、実機。Automatic のまま保存すると xl/workbook.xml の calcPr から calcMode 属性が省略されることまで確認） |
| xlOpenXMLWorkbook | 51 | SaveAs .xlsx | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認） |
| xlOpenXMLWorkbookMacroEnabled | 52 | SaveAs .xlsm | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後・再オープン後とも FileFormat=52 を確認） |
| xlOpenXMLTemplateMacroEnabled | 53 | SaveAs .xltm | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認） |
| xlOpenXMLTemplate | 54 | SaveAs .xltx | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat で確認。Workbooks.Open で開き直すとテンプレートから新規ブックが生成されるため、reopen では確認できない罠あり） |
| xlCSV | 6 | SaveAs .csv | 2026-08-17（T-04 Save-XlsWorkbook、実機。SaveAs 直後の Workbook.FileFormat とファイル内容で確認） |
| xlExcelLinks | 1 | Workbook.LinkSources の Type 引数 | 2026-08-17（T-08 Get-XlsModel、実機。`Workbook.LinkSources(1)` が未解決の外部リンク先パスを 1-based string[] で返す（リンクなしは $null）ことを確認。T-11 Test-XlsFormulas で refused 判定にも再利用: 実機プローブで、リンク数が 1 件でも常に `string[]`（スカラー化しない）で返ること、要素は `[シート名]` を含まないプレーンなファイルパス（例 `C:\NoSuch\Other.xlsx`）で `Test-Path -LiteralPath` でそのまま存在確認できること、解決済み・未解決を問わず両方が一覧に含まれることを確認） |

## SpecialCells / エラー判定

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlCellTypeFormulas | -4123 | Range.SpecialCells の Type 引数 | 2026-08-17（T-08 Get-XlsModel、実機。`Range.SpecialCells(-4123)` で数式セルのみの Range を取得、ヒットなしは HRESULT 0x800A03EC 例外になることを確認） |
| xlErrors | 16 | Range.SpecialCells の第 2 引数（Value フィルター） | 2026-08-17（T-10 Test-XlsFormulas、実機。`Range.SpecialCells(-4123, 16)` で「エラー値を返す数式セルだけ」の Range を取得できることを確認。ヒットなしは xlCellTypeFormulas 単独と同じ HRESULT 0x800A03EC 例外） |

### `$script:Xl.xlErr*` について（レビュー指摘 T-13 round 1 blocking 対応）

`$script:Xl.xlErr*`（`xlErrNull`/`xlErrDiv0`/`xlErrValue`/`xlErrRef`/`xlErrName`/`xlErrNum`/`xlErrNA`）は、
**Excel VBA の `XlCVError` 列挙値ではない**。これらは `Range.Value2` を PowerShell COM 相互運用で
読んだときにエラーセルが取る .NET 側の生の `Int32` 値（HRESULT 形式）に付けた、このモジュール
ローカルの別名（Value2 コードエイリアス）であり、Get-XlsRange（`ConvertFrom-XlsValue2`）が
表示文字列（`#VALUE!` 等）へ変換するためだけに使う。Excel の型ライブラリが定義する `XlCVError`
列挙値（VBA の `CVErr()` などで使う値）とは名前は似ているが値がまったく異なる別物であり、
スキル利用エージェントがこの値を `XlCVError` として他の COM API（例: `CVErr`）に渡すと壊れる。

下表は「Excel 公式の `XlCVError` 値（未実測・文献値）」と「`Range.Value2` を実機で読んだときの
実測 `Int32` 値（＝`$script:Xl.xlErr*` の値）」を明示的に分離して記録する。

| Excel 名 | XlCVError 値（公式、未実測。出典: [Microsoft Learn: XlCVError enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcverror)） | `Range.Value2` の実測 Int32（`$script:Xl.xlErr*` の値） | 用途 | 確認 |
|---|---:|---:|---|---|
| xlErrNull | 2000 | -2146826288 | Range.Value2（エラーセルの生値、#NULL!） | 2026-08-17（T-06 Get-XlsRange、実機。`=SUM(1:1 2:2)` で確認。XlCVError 値 2000 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrDiv0 | 2007 | -2146826281 | Range.Value2（エラーセルの生値、#DIV/0!） | 2026-08-17（T-06 Get-XlsRange、実機。`=1/0` で確認。XlCVError 値 2007 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrValue | 2015 | -2146826273 | Range.Value2（エラーセルの生値、#VALUE!） | 2026-08-17（T-06 Get-XlsRange、実機。`=1/""` で確認。XlCVError 値 2015 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrRef | 2023 | -2146826265 | Range.Value2（エラーセルの生値、#REF!） | 2026-08-17（T-06 Get-XlsRange、実機。`=#REF!` で確認。XlCVError 値 2023 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrName | 2029 | -2146826259 | Range.Value2（エラーセルの生値、#NAME?） | 2026-08-17（T-06 Get-XlsRange、実機。`=NoSuchName` で確認。XlCVError 値 2029 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrNum | 2036 | -2146826252 | Range.Value2（エラーセルの生値、#NUM!） | 2026-08-17（T-06 Get-XlsRange、実機。`=SQRT(-1)` で確認。XlCVError 値 2036 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |
| xlErrNA | 2042 | -2146826246 | Range.Value2（エラーセルの生値、#N/A） | 2026-08-17（T-06 Get-XlsRange、実機。`=NA()` で確認。XlCVError 値 2042 は Microsoft Learn の公式列挙定義を転記したのみで実測はしていない） |

## 配置・罫線（reference 専用、$script:Xl には追加しない）

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlCenter | -4108 | Range.HorizontalAlignment | 2026-08-18（T-13、実機。セルに値を入れて `HorizontalAlignment = -4108` を代入した直後に同じプロパティを読み戻し、-4108 のまま保持されることを確認） |
| xlLeft | -4131 | Range.HorizontalAlignment | 2026-08-18（T-13、実機。xlCenter と同じ手順で `HorizontalAlignment = -4131` の代入・読み戻しを確認） |
| xlRight | -4152 | Range.HorizontalAlignment | 2026-08-18（T-13、実機。xlCenter と同じ手順で `HorizontalAlignment = -4152` の代入・読み戻しを確認） |
| xlEdgeBottom | 9 | Range.Borders.Item(index) | 2026-08-18（T-13、実機。`Range.Borders.Item(9)` で下辺の Border オブジェクトを取得できること、`.LineStyle`/`.Weight` への代入がそのオブジェクトに反映されることを確認） |
| xlInsideHorizontal | 12 | Range.Borders.Item(index)（複数行範囲の内側水平線） | 2026-08-18（T-13、実機。2 行×2 列の範囲に対し `Range.Borders.Item(12)` を取得し、`.LineStyle`/`.Weight` への代入・読み戻しを確認） |
| xlContinuous | 1 | Border.LineStyle | 2026-08-18（T-13、実機。xlEdgeBottom/xlInsideHorizontal で取得した Border オブジェクトに `LineStyle = 1` を代入し、読み戻し値が 1 であることを確認） |
| xlThin | 2 | Border.Weight | 2026-08-18（T-13、実機。同じ Border オブジェクトに `Weight = 2` を代入し、読み戻し値が 2 であることを確認） |

## 条件付き書式・データ検証（reference 専用）

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlValidateList | 3 | Range.Validation.Add の Type 引数 | 2026-08-18（T-13、実機。`Validation.Add(3, 1, 3, 'A,B,C')`（Type=3, AlertStyle=xlValidAlertStop, Operator は list では無視される）を実行後、`Range.Validation.Type` が 3 であることを確認） |
| xlValidAlertStop | 1 | Range.Validation.Add の AlertStyle 引数（第 2 引数） | 2026-08-18（T-14 round 2、実機。`Validation.Add(3, 1, [Type]::Missing, 'A,B,C')` 実行後、`Range.Validation.AlertStyle` が 1 であることを確認。`patterns.md` のデータ検証パターンで、list 検証では使われない `Operator` にマジックナンバーを渡さず `[Type]::Missing` にした際に追加実測した） |
| xlCellValue | 1 | FormatConditions.Add の Type 引数 | 2026-08-18（T-13、実機。`FormatConditions.Add(1, 5, '=3')` を実行後、生成された FormatCondition の `.Type` が 1 であることを確認） |
| xlGreater | 5 | FormatConditions.Add の Operator 引数 | 2026-08-18（T-13、実機。上と同じ呼び出しで、生成された FormatCondition の `.Operator` が 5 であることを確認） |

## グラフ（reference 専用）

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlColumnClustered | 51 | Chart.ChartType | 2026-08-18（T-13、実機。`Worksheet.ChartObjects().Add(...).Chart` で作ったグラフに `ChartType = 51` を代入し、読み戻し値が 51 であることを確認。使用後は `ChartObject.Delete()` で後始末） |
| xlLine | 4 | Chart.ChartType | 2026-08-18（T-13、実機。同じ Chart オブジェクトで `ChartType = 4` に切り替え、読み戻し値が 4 であることを確認） |

## セル挿入（reference 専用）

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlShiftDown | -4121 | Range.Insert の Shift 引数 | 2026-08-18（T-13、実機。1 セルに値を入れて `Range.Insert(-4121)` した後、元セルが空になり 1 行下のセルへ元の値が移動していることを確認） |
| xlShiftToRight | -4161 | Range.Insert の Shift 引数 | 2026-08-18（T-13、実機。1 セルに値を入れて `Range.Insert(-4161)` した後、元セルが空になり 1 列右のセルへ元の値が移動していることを確認。レビュー指摘対応（T-13 round 1 should-fix）: `XlInsertShiftDirection` 列挙の公式名は `xlShiftToRight`（`xlToRight` は誤記、値・実測内容は変更なし）） |

## テーブル化（ListObjects、reference 専用）

| 名前 | 値 | 用途 | 確認 |
|---|---|---|---|
| xlSrcRange | 1 | ListObjects.Add の SourceType 引数 | 2026-08-18（T-13、実機。`ListObjects.Add(1, Range, [Type]::Missing, 1, [Type]::Missing, [Type]::Missing)` でテーブル化し、生成された ListObject の `.SourceType` が 1 であることを確認） |
| xlYes | 1 | ListObjects.Add の HasHeaders 引数（第 4 引数） | 2026-08-18（T-13、実機。上と同じ呼び出しで HasHeaders 引数に 1 を渡し、生成された ListObject の `.HeaderRowRange.Value2` が指定したヘッダー文字列と一致することを確認） |

## 財務モデルの色規約（Font.Color への RGB 直接代入）

Excel の COM には VBA の `RGB()` 関数に相当する組み込み関数がないため、PowerShell 側で
`R + G * 256 + B * 65536`（Windows の BGR 順 Long 値）を自分で計算し、`Range.Font.Color` に
直接代入する。以下は財務モデルでよく使われる色規約（青=ハードコード入力／黒=同一シート内数式／
緑=他シート参照／赤=他ブック参照・警告）の Long 値と、計算式そのものの実機確認。

| 用途 | R,G,B | Font.Color に代入する Long 値 | 確認 |
|---|---|---|---|
| ハードコード入力（青） | 0, 0, 255 | 16711680 | 2026-08-18（T-13、実機。`Range.Font.Color = 16711680` を代入し、読み戻し値が 16711680 のままであることを確認） |
| 同一シート内の数式（黒） | 0, 0, 0 | 0 | 2026-08-18（T-13、実機。数式セル（`=A1*2`）に `Font.Color = 0` を代入し、読み戻し値が 0 であることを確認） |
| 他シート参照の数式（緑） | 0, 128, 0 | 32768 | 2026-08-18（T-13、実機。`Range.Font.Color = 32768` を代入し、読み戻し値が 32768 のままであることを確認） |
| 他ブック参照・警告（赤） | 255, 0, 0 | 255 | 2026-08-18（T-13、実機。`Range.Font.Color = 255` を代入し、読み戻し値が 255 のままであることを確認） |
| （検証用）任意の RGB(12,34,56) | 12, 34, 56 | 3678732 | 2026-08-18（T-13、実機。`R + G*256 + B*65536` = `12 + 34*256 + 56*65536` = 3678732 を計算し、`Range.Font.Color` に代入した上で読み戻し値が 3678732 と一致することを確認。計算式自体の正しさを確認する目的の行） |

`Font.Color` は、**今回確認した単一セル、および全セルが同一色の範囲**では、書き込み・読み戻しとも
BGR 順 Long（実測では .NET 型は `System.Double`）のスカラーで、Windows 標準の `System.Drawing.Color`
の `R/G/B` から上記の式でそのまま変換できることを実機で確認した（`ColorTranslator` 等の外部依存は
不要）。「常に」という一般化はしない（レビュー指摘対応、T-13 round 1 blocking）。

**混在色の範囲での挙動（追加実測、T-13 round 2）**: 2 セルの範囲で各セルの `Font.Color` を異なる値
（255 と 16711680）に設定してから、その範囲 `Range.Font.Color`（複数セルへの読み取りアクセス）を
まとめて読むと、値ではなく `System.DBNull`（`$null` ではない）が返ることを実機で確認した。これは
`Get-XlsRange`/`Get-XlsOverview` の `NumberFormat`/`Text` が範囲内で値が揃っていないときに
`DBNull` を返すのと同じ挙動パターンである（同一色の範囲では `System.Double` のスカラーが返ることも
同じ実機確認で対照区として確かめた）。エージェントが複数セル範囲へ直接 `Font.Color` を読みに行く
場合は、`DBNull`（色が揃っていない）を考慮すること。
