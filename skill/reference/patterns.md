# Tested COM snippets

<!-- T-14 で充填。各スニペットは Invoke-XlsSession の ScriptBlock 内で動作確認済みのものだけ載せる（G-12）。
     形式: 見出し / 3 行以内の説明 / 5〜10 行のコード / 確認日 -->

実行環境（全パターン共通）: 2026-08-18、実機、Excel `$app.Version`=`16.0` / `$app.Build`=`20228`
（`Invoke-XlsSession` の `ScriptBlock` 内で `$app.Version`/`$app.Build` を読んで確認）。
定数の値・出典は `skill/reference/com-constants.md` を参照。ここでは名前とマジックナンバーを
コメントで併記するだけに留め、値の一次資料はそちらに置く。

## Header row formatting

ヘッダー行に太字・塗りつぶし・中央揃えを設定する。`Value2` への代入は 1 行分の `[object[,]]`
（G-08）。`xlCenter` は `com-constants.md` で実測済み（T-13）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$hdr = $ws.Range('A1:C1')
$data = New-Object 'object[,]' 1,3
$data[0,0] = 'Name'; $data[0,1] = 'Qty'; $data[0,2] = 'Price'
$hdr.Value2 = $data
$hdr.Font.Bold = $true
$hdr.Interior.Color = 15773696   # 見た目用の任意 BGR Long（財務モデル色規約の値ではない）
$hdr.HorizontalAlignment = -4108 # xlCenter（com-constants.md で実測済み）
```

確認: 2026-08-18（T-14、実機。`Invoke-XlsSession` 内で上記を実行した直後に同じ `ScriptBlock` の中で
`$hdr.Font.Bold`（`True`）・`$hdr.Interior.Color`（`15773696`）・`$hdr.HorizontalAlignment`（`-4108`）・
`$hdr.Value2[1,1]`（`'Name'`。複数セル範囲の `Value2` 読み取りは 1-based であることに注意）を
読み戻して確認）。

## Convert range to table (ListObjects.Add)

セル範囲を `ListObjects`（Excel テーブル）に変換する。`SourceType`/`HasHeaders` の値は
`com-constants.md` で実測済み（T-13）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$data = New-Object 'object[,]' 3,2
$data[0,0]='Item'; $data[0,1]='Amount'
$data[1,0]='A'; $data[1,1]=10
$data[2,0]='B'; $data[2,1]=20
$ws.Range('A1:B3').Value2 = $data
$lo = $ws.ListObjects.Add(1, $ws.Range('A1:B3'), [Type]::Missing, 1, [Type]::Missing, [Type]::Missing)
# 引数: SourceType=xlSrcRange(1), Source=Range, [省略], HasHeaders=xlYes(1)
$lo.Name = 'TblSample'
```

確認: 2026-08-18（T-14、実機。実行後に `$lo.SourceType`（`1`）・`$lo.HeaderRowRange.Value2[1,1]`
（`'Item'`）・`$lo.HeaderRowRange.Value2[1,2]`（`'Amount'`）・`$lo.Range.Address($false,$false)`
（`'A1:B3'`）を読み戻して確認）。

## Auto-fit columns

列幅をセル内容に合わせて自動調整する。`Worksheet.Columns`/`.Item(n)` が返すのは行・列全体を表す
`Range` であり、`AutoFit()` は行または列を表す `Range` に対するメソッド（任意の `Range` に使える
汎用メソッドではない）。`Value2` への代入は単一セルでも `[object[,]]`（G-08。レビュー指摘 T-14
round 1 blocking 対応: 以前はスカラー代入だった）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$data = New-Object 'object[,]' 1,1
$data[0,0] = 'This is a fairly long piece of text for column width test'
$ws.Range('A1').Value2 = $data
$before = $ws.Columns.Item(1).ColumnWidth
[void]$ws.Columns.Item(1).AutoFit()
$after = $ws.Columns.Item(1).ColumnWidth
```

確認: 2026-08-18（T-14 round 2、実機。既定の列幅から長い文字列を `[object[,]]` で入れて `AutoFit()`
を呼んだ結果、`ColumnWidth` が `8.38` → `50.5` に変化し、`$ws.Range('A1').Value2` が代入した文字列と
一致することを確認。数値は既定テーマ・既定フォントでの実測値であり、フォント・DPI・テーマが変われば
変わりうる一般的な傾向を示すものではない点に注意）。

## Add a comment (legacy `AddComment` and threaded `AddCommentThreaded`)

セルにコメントを付ける。従来型（レガシー）の `Range.AddComment` と、スレッド形式コメント
`Range.AddCommentThreaded` の両方を試す。**この環境（Excel Version `16.0` / Build `20228`）では
両方とも動作した**。他の Excel ビルドでは `AddCommentThreaded` が存在しない／失敗する可能性がある
（タスクカード想定どおり。動かない環境に当たった場合はビルド番号とエラー内容をここに追記すること）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$c1 = $ws.Range('A1')
$cmt = $c1.AddComment('Legacy comment text')

$c2 = $ws.Range('B1')
$tc = $c2.AddCommentThreaded('Threaded comment text')
```

確認: 2026-08-18（T-14 round 2、実機、Excel Version `16.0` Build `20228`。コメント対象セルは空のまま
（レビュー指摘 T-14 round 1 blocking 対応: 以前は `Value2 = 'legacy'` のようなスカラー代入で
セル値を入れていたが G-08 違反のため削除。コメント自体の動作確認にセル値は不要）。
`$c1.AddComment(...)` の戻り値 `Comment` の `.Text()`（引数なし呼び出しで現在のテキストを返す）が
`'Legacy comment text'`、`$c2.AddCommentThreaded(...)` の戻り値 `CommentThreaded` の `.Text()` が
`'Threaded comment text'` であることを確認。どちらも例外なく実行できた（両方とも `try/catch` で
実行し、この環境では `catch` に入らなかった）。

## Conditional formatting

セルの値が条件を満たすと書式を変える。`FormatConditions.Add` の `Type`/`Operator` は
`com-constants.md` で実測済み（T-13）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$rng = $ws.Range('A1:A5')
$data = New-Object 'object[,]' 5,1
for ($i = 0; $i -lt 5; $i++) { $data[$i,0] = $i + 1 }
$rng.Value2 = $data
$fc = $rng.FormatConditions.Add(1, 5, '=3')  # Type=xlCellValue(1), Operator=xlGreater(5)
$fc.Interior.Color = 65535                   # 黄色（BGR Long）
```

確認: 2026-08-18（T-14、実機。実行後に `$fc.Type`（`1`）・`$fc.Operator`（`5`）・`$fc.Formula1`
（`'=3'`）・`$fc.Interior.Color`（`65535`）・`$rng.FormatConditions.Count`（`1`）を読み戻して確認）。

## Data validation (list)

セルに入力できる値をリストで制限する。`Validation.Add` の `Type`（xlValidateList=3）と `AlertStyle`
（xlValidAlertStop=1）はどちらも `com-constants.md` に実測記録がある（xlValidateList は T-13、
xlValidAlertStop は T-14 round 2 で追加）。`Operator` はリスト検証では使われないため、値を持たない
マジックナンバーを渡さず `[Type]::Missing` にする（レビュー指摘 T-14 round 1 should-fix 対応）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$cell = $ws.Range('A1')
$cell.Validation.Add(3, 1, [Type]::Missing, 'A,B,C')
# Type=xlValidateList(3), AlertStyle=xlValidAlertStop(1)
```

確認: 2026-08-18（T-14 round 2、実機。`Operator` を `[Type]::Missing` にした呼び出しに変更した上で
再実行し、`$cell.Validation.Type`（`3`）・`$cell.Validation.AlertStyle`（`1`）・
`$cell.Validation.Formula1`（`'A,B,C'`）を読み戻して確認）。

## Insert a chart

ワークシートに埋め込みグラフを挿入する。このパターンは成果物としてグラフを残すためのものなので、
`com-constants.md` の実測作業（T-13）で使った `ChartObject.Delete()` による後始末はここでは行わない
（レビュー指摘 T-14 round 1 should-fix 対応: 後始末目的の Delete を成果物編集用パターンに持ち込んで
いた）。`ChartType` の値は `com-constants.md` で実測済み（T-13、xlColumnClustered=51）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$data = New-Object 'object[,]' 3,2
$data[0,0]='Q1'; $data[0,1]=10
$data[1,0]='Q2'; $data[1,1]=20
$data[2,0]='Q3'; $data[2,1]=15
$ws.Range('A1:B3').Value2 = $data
$chart = $ws.ChartObjects().Add(100, 10, 300, 200).Chart
$chart.SetSourceData($ws.Range('A1:B3'))
$chart.ChartType = 51  # xlColumnClustered
```

確認: 2026-08-18（T-14 round 2、実機。上記コードをそのまま実行し、`$ws.ChartObjects().Count` が
実行前の `0` から実行後 `1`（+1）に増えたこと、`$chart.ChartType` の読み戻しが `51` であることを
確認。グラフは削除せず残したままセッションを終える）。

## Named ranges

セルに名前を付けて数式や参照から使えるようにする。シート名を含む `RefersTo` はシート名を単引用符で
囲み、シート名自体に単引用符が含まれる場合はそれを 2 つ重ねてエスケープする（`docs/01-design.md`
§5 の「スペース入りシート名は数式内でクォート」と同じ理由。単引用符の二重化は round 1 の実測が
`Sheet1`（クォート不要なシート名）でしか行われておらず、シート名に単引用符が含まれる場合の挙動を
確認していなかったため round 2 で追加した。レビュー指摘 T-14 round 1 should-fix 対応）。
`Value2` への代入は単一セルでも `[object[,]]`（G-08。レビュー指摘 T-14 round 1 blocking 対応:
以前はスカラー代入だった）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Sales O'Brien Q1"
$data = New-Object 'object[,]' 1,1
$data[0,0] = 42
$ws.Range('A1').Value2 = $data
$safeName = $ws.Name.Replace("'", "''")
$refersTo = "='{0}'!`$A`$1" -f $safeName
$wb.Names.Add('MyNamedCell', $refersTo)
```

確認: 2026-08-18（T-14 round 2、実機。シート名をスペースと単引用符の両方を含む `"Sales O'Brien Q1"`
に変更し、`.Replace("'", "''")` で二重化してから組み立てた
`$refersTo`（`"='Sales O''Brien Q1'!$A$1"`）を渡したところ、`$wb.Names.Item('MyNamedCell').RefersTo`
の読み戻し値は渡した文字列と完全に一致した。round 1 で確認した「クォート不要な単純シート名
（`Sheet1`）だと Excel が単引用符を自動的に取り除く」という正規化は、このケースでは起きなかった。
シート名にスペースまたは単引用符が含まれ、クォートが実際に必要な場合は渡した文字列がそのまま
保持されると見られる。`$ws.Range('MyNamedCell').Value2` が `42` であることも確認）。

## Sheet add / copy / delete

シートを追加・複製・削除する。`Worksheets.Delete()` は確認ダイアログを出しうるが、
`Invoke-XlsSession` がセッション開始時に `DisplayAlerts=$false` を設定済みのため、ここで改めて
設定し直す必要はない。`Worksheets.Add()` は**アクティブシートの直前**に挿入される(タブの末尾に
追加されるわけではない。T-07 実装メモで確認)。末尾に置きたい場合は追加後に `.Move` で並べ替えるか、
末尾のシートをアクティブにしてから `Add()` する。

```powershell
param($app, $wb)
$countStart = $wb.Worksheets.Count
$newWs = $wb.Worksheets.Add()
$newWs.Name = 'AddedSheet'
$newWs.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
$copiedWs = $wb.Worksheets.Item($wb.Worksheets.Count)
$copiedWs.Name = 'CopiedSheet'
$copiedWs.Delete()
```

確認: 2026-08-18（T-14、実機。新規ブック（シート 1 枚）から開始し、`$wb.Worksheets.Count` が
`Add()` 後に `1`→`2`、`Copy()` 後に `2`→`3`、`Delete()` 後に `3`→`2` と推移することを確認。
`$newWs.Name` が `'AddedSheet'` のまま残っていることも確認）。

## Misc PowerShell/COM gotchas

SKILL.md 本体には載せていない、より細かい・頻度の低い罠。T-15 round 3 のレビュー指摘
（COM gotchas 章を 15 項目程度へ圧縮する）で、SKILL.md から詳細をここへ移送した。

**`[int]` キャストは切り捨てではなく四捨五入する。** 切り捨てが必要な場面（例: 列文字変換の
26 進数風の桁上がり計算）で `[int]` を使うと、26 の倍数の境界で 1 桁多い誤った結果になる
（`ConvertTo-XlsColumnLetter` で実際に踏んだ罠。T-08 実装メモ参照）。`[Math]::Floor()` を使う。

```powershell
# (26 - 1) / 26 = 0.9615... を [int] にキャストすると切り捨てでなく四捨五入で 1 になってしまう
[int]((26 - 1) / 26)          # => 1（誤り。26=Z を AZ と誤変換する原因になった）
[int][Math]::Floor((26 - 1) / 26)  # => 0（正しい）
```

**2 次元配列を手で組み立てるときの 2 つの構文の罠**（`Set-XlsRange`/`Get-XlsRange` を使えば
どちらも回避できる。自前で `[object[,]]` やジャグ配列を組む場合だけの注意）:

```powershell
# 罠1: 多次元配列の添字をメソッド引数へ直書きすると、PowerShell のパーサーが
# 添字のカンマを引数区切りと誤認して構文エラーになる（T-09 実装メモ）。
# NG: $list.Add($grid[$r, $c])
$cell = $grid[$r, $c]
$list.Add($cell)          # OK: いったんローカル変数に取り出す

# 罠2: @(@(1.0, 2.0)) は 1 段階フラット化され、1 行 2 列ではなく 2 行 1 列になる（T-09 実装メモ）。
@(@(1.0, 2.0)).Count        # => 2（誤り: 2 行になってしまう）
@(, @(1.0, 2.0)).Count      # => 1（正しい: 1 行 2 列のジャグ配列を維持）
```

**型指定のないパラメーター経由で受け取った `[double]` を、関数・`ScriptBlock` の境界を越えて
そのまま `Range.Value2 =` へ代入すると、まれに `COMException (0x800A03EC)` になることがある**
（T-06、`ConvertTo-XlsWriteCellValue` 実装メモ。根本原因は未特定だが、代入の直前に明示的に
`[double]` へ再キャストすると再現しなかった)。

```powershell
function Set-MyCell {
    param($Worksheet, $Range, $Value)   # $Value は型指定なし。呼び出し元境界を越えてきた [double]
    $grid = New-Object 'object[,]' 1,1
    $grid[0,0] = [double]$Value          # 代入直前に明示的に再キャストする（罠を回避）
    $Worksheet.Range($Range).Value2 = $grid
}
```

**既定（General）書式のセルに `Range.Value2` で数値に見える文字列（`"2024"` など）を書き込むと、
Excel が自動的に数値（`[double]`）へ変換し、文字列のまま保持されない**（T-16、fixtures 生成で
実機確認。SKILL.md 財務モデル節の「年は文字列として（`"2024"`、`2,024` ではない）」を満たすのに
文字列を渡すだけでは不十分）。`Set-XlsRange` 自身は `NumberFormat` を一切変更しない仕様のため、
これだけでは防げない。書き込む**前**に対象セルへ `NumberFormat = '@'`（テキスト書式）を設定して
おくと自動変換を回避できる（先頭にアポストロフィを付ける方法でも防げる。アポストロフィは値から
取り除かれて文字列として保持される（実測）。書式意図を明示できるため、ここでは書き込み前の
`NumberFormat='@'` を推奨する）。`NumberFormat` を後から `'@'` に変えても、既に
数値化された値は文字列に戻らない（順序が重要）。

```powershell
param($app, $wb)
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').NumberFormat = '@'   # 書き込み前に設定する（後からでは手遅れ）
$data = New-Object 'object[,]' 1,1
$data[0,0] = '2024'
$ws.Range('A1').Value2 = $data
```

確認: 2026-08-18（T-16、実機。既定書式のまま書き込むと `Range.Value2` の読み戻し型が
`System.Double`（2024）になり、`NumberFormat='@'` を先に設定してから同じ文字列を書き込むと
`System.String`（"2024"）のまま保持されることを比較確認した）。
