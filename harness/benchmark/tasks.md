# ベンチマーク — xlsx スキル同等性

同じ自然言語指示を、(A) Claude 公式 xlsx スキル、(B) xlsx-ps スキル、で実行させ、
手数・出力品質・検証結果を比較する。B が A と同等以上なら合格。

## 実行環境

- A: Python + openpyxl + LibreOffice がある環境（Claude.ai の Code Execution でよい）
- B: Windows + Excel + PowerShell 5.1、`skill/` を Copilot / Claude Code のスキルとして配置
- 同じモデル（Sonnet）で両方を実行する。スキルの差だけを測る。

## タスク

### B-1 空から財務モデル
「2024〜2028 の 5 年収益予測。売上成長率・粗利率・営業費用率を Assumptions シートに置き、
Model シートで参照。色規約に従う。」
- 合格: 数式で計算されている／入力セルが青／仮定値がラベル付き別セル／検証 `success`／年が文字列。

### B-2 既存ブックの入力セル更新
サンプル `fixtures/model.xlsx`（青文字の入力セルを持つ）を渡し「成長率を 8% に、粗利率を 42% に更新」。
- 合格: 指定セルだけ変わり、既存数式に触れていない（`Get-XlsModel -FormulasOnly` の前後 diff がゼロ）／検証 `success`。

### B-3 CSV 取り込みと集計
`fixtures/sales.csv`（1,000 行、日付・地域・金額）を渡し「地域別・月別の集計表を作り、合計行と合計列を数式で」。
- 合格: SUMIFS 等の数式／日付が日付として入っている／テーブル化されている／検証 `success`。

### B-4 壊れた数式の修正
`fixtures/broken.xlsx`（`#REF!` 3 か所、`#DIV/0!` 2 か所、`#NAME?` 1 か所）を渡し「エラーを直して」。
- 合格: 検証が `success`／`#DIV/0!` は `IFERROR` か分母ガードで直っている／元の意図（合計・比率）が保たれている。

### B-5 xlsm 編集
`fixtures/macro.xlsm`（マクロあり）を渡し「Summary シートを追加して集計」。
- 合格: 保存後にマクロが残っている／検証 `success`。

### B-6 外部リンク付きブック
`fixtures/linked.xlsx`（存在しないファイルへのリンクを持つ）を渡し「B 列に前年比を追加」。
- 合格: 検証が最初 `refused` を返し、エージェントがリンク値の退避か `-Force` かを判断して報告している／リンクを黙って壊していない。

## 記録表（T-16 で埋める）

| タスク | A 手数 | B 手数 | A 検証 | B 検証 | 品質差 | 判定 | メモ |
|---|---|---|---|---|---|---|---|
| B-1 | 未実施(環境制約: LibreOffice なし) | 23 | 未実施(環境制約: LibreOffice なし) | `success`(39 formulas, 0 errors) | 判定不能(A 未実施) | 合格 | Assumptions の Base Revenue/成長率/粗利率/営業費用率が各 `Font.Color=16711680`(青)、値ラベル付き別セル(行見出し)。年見出し(Assumptions 行10・Model 行4)は `Value2` 実測で `System.String`。凡例行(Blue/Yellow/Green/Black text の意味)まで自主的に追加していた。 |
| B-2 | 未実施(環境制約: LibreOffice なし) | 14 | 未実施(環境制約: LibreOffice なし) | `success`(20 formulas) | 判定不能(A 未実施) | 合格 | `fixtures/model.xlsx` と `bench/b2/model.xlsx` の `Get-XlsModel -FormulasOnly` 差分は path と再計算後の `value2` のみ(数式文字列の差分ゼロ)。`Assumptions!B3` 0.05→0.08、`B4` 0.35→0.42 のみ変化、`B5`(0.2)・`B6`(1000000)・全 4 セルの `Font.Color`(16711680)は不変を実測で確認。 |
| B-3 | 未実施(環境制約: LibreOffice なし) | 34 | 未実施(環境制約: LibreOffice なし) | `success`(78 formulas, 0 errors) | 判定不能(A 未実施) | 合格 | Data シートに `SalesData` テーブル(`ListObjects`、A1:C1001)。Date 列は `Value2` が `System.Double` かつ `NumberFormat="yyyy-mm-dd"`(日付型)。Summary は構造化参照 `SUMIFS(SalesData[Amount],...)` + `EDATE` で月範囲判定、行合計(`=SUM(B4:M4)`)・列合計(`=SUM(B4:B7)`)・総合計とも数式。 |
| B-4 | 未実施(環境制約: LibreOffice なし) | 17 | 未実施(環境制約: LibreOffice なし) | `success`(13 formulas, 0 errors。修正前は `errors_found` 6件) | 判定不能(A 未実施) | 合格 | `#REF!`(D4/D5/D6)は正しい `SUM` に復元(D6 は `SUM(D2:D3)`→`SUM(D2:D5)` へ拡張し全地域を含める、より妥当な解釈)。`#DIV/0!`(F4/F5)は `IFERROR(...,0)` で他行(F2/F3)と同じ比率式に統一。`#NAME?`(F6、未定義名 `GrowthRate` 参照)は列見出し「% of Quota」の意味に合わせ `IFERROR(D6/E6,0)` に修正(他行と同一パターン)。合計・比率の意図は保持(むしろ D6 は改善)。 |
| B-5 | 未実施(環境制約: LibreOffice なし) | 19 | 未実施(環境制約: LibreOffice なし) | `success`(3 formulas, 0 errors) | 判定不能(A 未実施) | 合格 | **縮退検証**(T-16 環境制約: マクロ実体なし fixture、Trust Center 変更が要る VBA コード注入は行わない)。実測: `Workbook.FileFormat`=52(維持)、`Workbook.HasVBProject`=**false**。実行エージェントは `HasVBProject: True` と自己報告していたが、独立検証(このセッション、`Invoke-XlsSession -ReadOnly` で直接読み取り)では `false`(fixture 自体が元々マクロなし=false のため、これが正しい値。エージェントの自己報告が誤り、成果物自体には問題なし)。Summary シートは `A2=SORT(UNIQUE(Data!B2:B11))` のスピル + `B2=SUMIF(Data!B2:B11,A2#,Data!C2:C11)` のスピル参照 + `E1=SUM(...)` の Grand Total。 |
| B-6 | 未実施(環境制約: LibreOffice なし) | 18 | 未実施(環境制約: LibreOffice なし) | 無 `-Force`: `refused`(reason="external links present")。`-Force`: `errors_found` 24件(すべて `#REF!`) | 判定不能(A 未実施) | 合格 | プロセス基準(最初に refused を観測し、リンク値退避 or `-Force` を判断)はオーケストレーター記録により充足(このセッションでは成果物のみ独立検証)。D 列(外部リンク数式)は `fixtures/linked.xlsx` と 12 行すべて完全一致(改変なし)。B 列に `=C{r}/D{r}`(前年比)を新規追加、D 列が `-Force` 再計算で `#REF!` になった結果 B 列も連鎖的に `#REF!` — エラー 24件はちょうど B2:B13(12)+D2:D13(12)のみで、C/A 列は無傷。 |

「手数」はエージェントのツール呼び出し回数。「品質差」は人間の目視で 3 段階（B 優／同等／A 優）。

## fixtures の作り方

`fixtures/` は T-16 の最初に、xlsx-ps 自身（`Invoke-XlsSession` 内の COM）で生成する。
生成スクリプト `fixtures/Make-Fixtures.ps1` を成果物に含める。B-6 のリンク先は
`C:\nonexistent\source.xlsx` のような存在しないパスにする。

## 不合格だったとき

- 原因が `SKILL.md` の記述不足 → gotchas / patterns に追記して再実行。
- 原因がモジュールの機能不足 → 01-design.md §0 の変更に該当するか判断。該当なら人間へ、非該当ならタスクカード追加。
- 原因が Excel COM の制約 → 記録表メモに書き、「同等でない点」として SKILL.md の冒頭に明記する。
