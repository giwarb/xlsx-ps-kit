# 01 方式設計

## 0. 確定事項（変更には人間の承認が要る）

| 項目 | 決定 |
|---|---|
| 実行基盤 | Excel デスクトップ版の COM オートメーション。OOXML 直接操作・ImportExcel・Open XML SDK は使わない |
| 言語 | PowerShell 5.1 以上（7 でも動くこと）。7 専用構文は禁止 |
| 利用者 | コーディングエージェント（Copilot / Claude Code）。人間の対話利用は副次 |
| 設計思想 | xlsx スキルと同型：**汎用ライブラリ相当 ＋ 検証スクリプト ＋ ルール文書**。JSON コマンド語彙で能力を絞らない |
| 依存 | Excel（M365 Apps）、PowerShell 5.1+。外部モジュール・OSS なし |
| 完了条件 | `harness/benchmark/tasks.md` の 6 タスク通過 |

## 1. 部品対応

| xlsx スキル | xlsx-ps | 備考 |
|---|---|---|
| openpyxl（作成・編集） | Excel COM 直接記述 ＋ `XlsAgent.psm1` のセッション補助 | 自由度は openpyxl 以上 |
| pandas（一括入出力） | `Get-XlsRange` / `Set-XlsRange` | `Range.Value2` 2 次元配列 ⇄ CSV/JSON。セル単位ループを避ける唯一の経路 |
| markitdown（概観） | `Get-XlsOverview` | Markdown 表ダンプ、座標なし、編集計画には使わない |
| 2 回ロード（数式＋値） | `Get-XlsModel` | COM なら 1 回で `Formula` と `Value2` が取れる |
| recalc.py | `Test-XlsFormulas.ps1` | JSON 契約を完全一致させる |
| LibreOffice 互換の関数制限 | なし | Excel 本体で計算 |
| openpyxl gotchas | COM gotchas | 新規執筆の中心 |
| 出力必須要件・財務モデル規約 | そのまま転記 | 基盤非依存 |

## 2. 成果物

```
skill/
  SKILL.md
  scripts/
    XlsAgent.psm1              公開 6 関数
    Test-XlsFormulas.ps1       recalc.py 互換の検証
  reference/
    com-constants.md
    patterns.md
tests/
  *.Tests.ps1                  Pester（v3 互換で書く。5.1 同梱の Pester 3.4 で動くこと）
```

## 3. `XlsAgent.psm1` 公開関数

エージェントが毎回書くと事故る部分だけを吸収する。書式・罫線・結合・列幅・コメント・テーブル・
条件付き書式・データ検証・グラフは吸収しない（素の COM を書かせる）。

### 3.1 `Invoke-XlsSession`

```
Invoke-XlsSession -Path <string> [-ReadOnly] [-Visible] -ScriptBlock { param($app, $wb) ... }
```

- 新規 `Excel.Application` を起動（`GetActiveObject` 禁止）。
- 初期状態: `Visible=$false, DisplayAlerts=$false, ScreenUpdating=$false, EnableEvents=$false, AskToUpdateLinks=$false, Calculation=xlCalculationManual`。
- `Workbooks.Open($Path, UpdateLinks:=0, ReadOnly:=$ReadOnly)`。ファイル不在で `-ReadOnly` でなければ `Workbooks.Add()`。
- 開いた直後に `$wb.ReadOnly` を確認し、`-ReadOnly` 指定なしで読み取り専用になった場合は例外（他プロセスが開いている）。
- STA でなければ例外（`[Threading.Thread]::CurrentThread.ApartmentState`）。
- 起動した PID を `$env:TEMP\xlsagent\<pid>.marker` に記録。
- `finally` で `Workbook.Close(SaveChanges:=$false)`（保存は ScriptBlock 内で `Save-XlsWorkbook` を明示）、`Application.Quit`、`ReleaseComObject` を逆順、`[GC]::Collect(); [GC]::WaitForPendingFinalizers()` を 2 回、マーカー削除。
- 戻り値: ScriptBlock の戻り値をそのまま返す。

### 3.2 `Save-XlsWorkbook`

```
Save-XlsWorkbook -Workbook <obj> -Path <string>
```

- 拡張子から `FileFormat` を決定: `.xlsx`→51, `.xlsm`→52, `.xltx`→54, `.xltm`→53, `.csv`→6。
- 保存前に `Calculation=xlCalculationAutomatic` に戻し `CalculateFullRebuild` を実行（手動のまま保存すると開いた人の環境で古い値が見える）。
- 上書き時は `DisplayAlerts=$false` 前提で `SaveAs`。既存パスと同一なら `Save`。

### 3.3 `Get-XlsOverview`

```
Get-XlsOverview -Path <string> [-Sheet <string[]>] [-MaxRows 50] [-MaxCols 30]
```

- シートごとに `## SheetName` 見出し ＋ Markdown 表。`Range.Text`（表示文字列）ベース。
- 座標を出さない。SKILL.md で「編集計画には使うな」と書く。
- `MaxRows/MaxCols` 超過は末尾に `... (N rows × M cols total)` を付ける。

### 3.4 `Get-XlsModel`

```
Get-XlsModel -Path <string> [-Sheet <string[]>] [-Range <string>] [-FormulasOnly] [-AsJson]
```

- 返却（オブジェクト or JSON）:
  ```
  {
    workbook: { path, sheets: [name], names: [{name, refersTo}], links: [string], version, build },
    sheets: [{
      name, usedRange, tables: [{name, range}], protected,
      cells: [{ address, formula, value2, numberFormat }]   // FormulasOnly なら数式セルのみ
    }]
  }
  ```
- `-Range` 指定時は該当範囲のみ。既定はシート全体の UsedRange。
- SKILL.md で「まず `-FormulasOnly` か `-Range` で絞れ、全部取ると溢れる」と書く。

### 3.5 `Get-XlsRange` / `Set-XlsRange`

```
Get-XlsRange -Worksheet <obj> -Range <string> [-AsCsv <path>] [-AsJson <path>] [-Header]
Set-XlsRange -Worksheet <obj> -Range <string> (-Data <object[,]> | -FromCsv <path> | -FromJson <path>) [-Header]
```

- `Value2` の 2 次元配列一括転送のみ。内部で `[object[,]]` を必ず作る。
- CSV は UTF-8（BOM 付き）で読み書き。5.1 の既定エンコーディング問題を吸収。
- `-Range` は左上セルだけ指定でもよい（データ寸法で自動拡張）。
- 日付: `Get` は OLE シリアル → ISO 8601 文字列、`Set` は ISO 文字列を検出したら `[DateTime]::ToOADate()`。

### 3.6 `Test-XlsFormulas`（関数）＋ `Test-XlsFormulas.ps1`（CLI）

CLI:
```
pwsh -File Test-XlsFormulas.ps1 <path> [timeoutSec=30] [-Force]
```

- `Invoke-XlsSession -ReadOnly` で開き `CalculateFullRebuild`（`Calculation` は Automatic に一時変更）。
- 全シートで `SpecialCells(xlCellTypeFormulas, xlErrors)`。`SpecialCells` はヒットなしで例外を投げるので握りつぶす。
- 出力 JSON（recalc.py 完全互換）:
  ```
  { "status": "success" | "errors_found",
    "total_formulas": int,
    "total_errors": int,
    "error_summary": { "#REF!": { "count": int, "locations": ["Sheet1!B5", ...], "locations_truncated": int }, ... } }
  ```
- 外部リンク（`$wb.LinkSources(xlExcelLinks)` が非 null）があり、リンク元ファイルが存在しない場合は
  `{ "status": "refused", "reason": "external links present", "links": [...] }` を返して**再計算しない**。`-Force` で続行。
- 終了コード: `error` キー（Excel 起動失敗・ファイル不在・タイムアウト）のときのみ非 0。`errors_found` / `refused` は 0。
- タイムアウト: `CalculateFullRebuild` を Runspace で走らせて `timeoutSec` で打ち切り、`{ "error": "timeout" }`。

## 4. `SKILL.md` 章立て（xlsx スキルと 1 対 1）

| 章 | 扱い |
|---|---|
| 冒頭表（Task / Approach） | 部品対応表の右列で書き換え |
| Requirements for every output（7 項目） | **一字も変えず**転記 |
| Recalculate | COM 版に書き換え。「`Test-XlsFormulas` を通さず納品するな」「JSON を読め、exit code を信じるな」は維持 |
| Choosing formulas that survive verification | 大幅短縮。`Formula`（英語名・US 区切り）を使う／`FormulaLocal` 禁止／動的配列は `Formula2`。互換性制限は消えたと明記 |
| openpyxl gotchas | COM gotchas に全面置換（§5） |
| Financial models | そのまま転記 |
| Dependencies | Excel デスクトップ版 ＋ PowerShell 5.1+ |
| 「緑の検証は正しさの証明ではない」警告 | **一字も変えず**転記 |

## 5. COM gotchas（SKILL.md に載せる候補）

`Invoke-XlsSession` で潰せるもの（新プロセス強制、ReadOnly 検出、STA 確認、後始末）は載せない。
エージェントが ScriptBlock の中で踏むものだけ載せる。

- `Value2` を使う。`Value` は日付を `DateTime` で返し丸め問題、`Text` は表示文字列。
- インデックスは 1-based。PowerShell 配列は 0-based。
- 結合セルは `MergeArea` の左上にだけ書く。
- `Range.Value2` への一括代入は `[object[,]]`。ジャグ配列だと 1 行しか入らない（`Set-XlsRange` を使えば回避）。
- スペース入りシート名は数式内でクォート（openpyxl と共通）。
- 保護シートは `Unprotect` してから書き、終了時に戻す。
- `Formula` に動的配列関数を書くと `@` が付く。`Formula2` を使う。
- `.xlsm` は `Save-XlsWorkbook` を通す（`FileFormat` 52）。
- 対象ファイルを Excel で開いたまま実行しない（`Invoke-XlsSession` が例外を出す）。
- `SpecialCells` はヒットなしで例外。

## 6. reference の中身

- `com-constants.md`: `xlCellTypeFormulas(-4123) / xlErrors(16) / xlOpenXMLWorkbook(51) / xlOpenXMLWorkbookMacroEnabled(52) / xlCalculationManual(-4135) / xlCalculationAutomatic(-4105) / xlCenter(-4108) / xlContinuous(1) / xlThin(2) / xlExcelLinks(1) / xlSrcRange(1) / xlYes(1)` など 30 個程度。実装時に実測して充填。
- `patterns.md`: ヘッダー行書式／テーブル化（`ListObjects.Add`）／列幅自動調整／コメント追加（`AddComment` と `AddCommentThreaded` の両方）／条件付き書式／データ検証／グラフ挿入／名前定義／シート追加・複製・削除。各 5〜10 行。

## 7. 決めていないこと（実装中に判断してよい）

- `Get-XlsOverview` の既定 `MaxRows/MaxCols`。
- `Get-XlsModel` の JSON で `value2` が `null`（空セル）と `""`（空文字数式結果）を区別するか。recalc.py 側は区別しないので、しなくてよい。
- Pester のバージョン。5.1 同梱の 3.4 で動くことを優先し、v5 の構文は使わない。
