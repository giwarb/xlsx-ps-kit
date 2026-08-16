# 02 実装計画

各タスクは「1 タスク = 1 実装セッション（Sonnet）＋ 1 レビューセッション（Codex）」で回す。
タスクカードは `harness/state/tasks/T-XX.md` にコピーして使う（雛形は §3）。

## 1. フェーズ

| Phase | 内容 | 完了条件 |
|---|---|---|
| P0 | 環境確認・骨組み | `tests/P0.Tests.ps1` が通る（Excel 起動・終了・プロセス残骸なし） |
| P1 | セッション管理 | `Invoke-XlsSession`, `Save-XlsWorkbook` |
| P2 | 読み取り系 | `Get-XlsOverview`, `Get-XlsModel`, `Get-XlsRange` |
| P3 | 書き込み系 | `Set-XlsRange` |
| P4 | 検証 | `Test-XlsFormulas`（関数 + CLI） |
| P5 | 文書 | `SKILL.md`, `reference/*.md` |
| P6 | ベンチマーク | `harness/benchmark/tasks.md` 6 本を通す |

P1〜P4 は直列。P5 は P4 完了後だが、`reference/com-constants.md` は P1 から随時追記してよい。

## 2. タスク一覧

| ID | Phase | タスク | 依存 | 見積（Sonnet ターン） |
|---|---|---|---|---|
| T-01 | P0 | `XlsAgent.psm1` 骨組み・`Export-ModuleMember`・Pester 3.4 実行確認 | - | 1 |
| T-02 | P0 | Excel 起動→終了のスモークテスト。EXCEL.EXE 残骸ゼロを Pester で検証 | T-01 | 1 |
| T-03 | P1 | `Invoke-XlsSession`（起動・初期状態・STA 検査・ReadOnly 検査・PID マーカー・finally 後始末） | T-02 | 2 |
| T-04 | P1 | `Save-XlsWorkbook`（拡張子→FileFormat、保存前 CalculateFullRebuild、Automatic 復帰） | T-03 | 1 |
| T-05 | P1 | 孤児プロセス回収 `Clear-XlsOrphans`（自モジュール起動 PID のみ） | T-03 | 1 |
| T-06 | P2 | `Get-XlsRange`（Value2 → 2 次元配列 → CSV/JSON、日付変換、Header） | T-03 | 2 |
| T-07 | P2 | `Get-XlsOverview`（Markdown 表、上限、省略表記） | T-06 | 1 |
| T-08 | P2 | `Get-XlsModel`（Formula+Value2 同時取得、names/links/tables/protected、Range/FormulasOnly 絞り込み） | T-06 | 2 |
| T-09 | P3 | `Set-XlsRange`（`[object[,]]` 生成、CSV/JSON 入力、左上指定の自動拡張、ISO 日付→OADate） | T-06 | 2 |
| T-10 | P4 | `Test-XlsFormulas` 関数（CalculateFullRebuild、SpecialCells 走査、error_summary、truncation） | T-04 | 2 |
| T-11 | P4 | 外部リンク検出と `refused` / `-Force` | T-10 | 1 |
| T-12 | P4 | CLI `Test-XlsFormulas.ps1`（引数、タイムアウト Runspace、exit code 契約） | T-11 | 1 |
| T-13 | P5 | `reference/com-constants.md` 実測充填 | T-12 | 1 |
| T-14 | P5 | `reference/patterns.md`（9 パターン、各実行確認済み） | T-12 | 2 |
| T-15 | P5 | `SKILL.md` 確定（xlsx スキル章立てに 1 対 1） | T-13, T-14 | 1 |
| T-16 | P6 | ベンチマーク B-1〜B-6 実行、差分記録、SKILL.md 反映 | T-15 | 3 |

## 3. タスクカード雛形

```markdown
# T-XX <タイトル>

## 目的
（01-design.md の該当節を引用）

## 入力
- 参照ファイル: docs/01-design.md §x.y, docs/03-guardrails.md
- 前提タスク: T-YY（完了済み）

## 成果物
- skill/scripts/XlsAgent.psm1 の <関数名>
- tests/<関数名>.Tests.ps1

## 完了条件（Definition of Done）
- [ ] Pester が緑
- [ ] テスト後に EXCEL.EXE が残っていない（tests/Common.ps1 の Assert-NoOrphanExcel）
- [ ] guardrails.md の禁止事項に抵触なし（自己チェック欄に○）
- [ ] Codex レビューで blocking 指摘ゼロ

## 実装メモ（Sonnet が記入）

## レビュー結果（Codex が記入 → harness/state/reviews/T-XX.md）
```

## 4. 各タスクの受け入れテスト要点

- **T-03**: (a) 存在しないパスで `Workbooks.Add` になる (b) 別プロセスで開いたファイルを渡すと ReadOnly 例外 (c) ScriptBlock 内で `throw` しても Quit されプロセスが残らない (d) 戻り値が透過する。
- **T-04**: (a) `.xlsm` を保存してマクロが残る（VBProject アクセス不可でも `HasVBProject` で確認）(b) 保存後に開き直して `Calculation` が Automatic。
- **T-06/T-09**: 往復テスト。`Set` → `Get` で 2 次元配列・CSV・JSON が同値。日付列が ISO 文字列で往復。空セルが `$null` で往復。
- **T-08**: 数式セルで `formula` が `=` 始まり、`value2` が計算結果。名前定義・外部リンクが取れる。
- **T-10**: 意図的に `#REF!` `#DIV/0!` `#NAME?` を含むブックで、種別ごとの count と locations が正しい。101 個以上のエラーで `locations` が 100 で切れ `locations_truncated` が正しい。
- **T-11**: 存在しない `C:\nonexistent.xlsx` へのリンクを持つブックで `refused`。`-Force` で `errors_found`。
- **T-12**: `errors_found` で exit 0、ファイル不在で exit 1 と `{"error": ...}`。
- **T-16**: `harness/benchmark/tasks.md` の記録表が埋まっている。

## 5. 順序に関する注意

- T-06 を T-07/T-08 より先にするのは、`Value2` の日付・空セル・エラー値の変換規則を 1 か所（`Get-XlsRange` の内部関数 `ConvertFrom-XlsValue2`）に集約し、Overview/Model がそれを再利用するため。
- T-10 は T-04 の後。`CalculateFullRebuild` 前後の `Calculation` モード制御を `Save-XlsWorkbook` と共通化する。
- `reference/*.md` は「実行して動いたスニペット」だけを載せる。書いてから動作確認するのではなく、Pester か手動で通した後に転記する（guardrails G-12）。
