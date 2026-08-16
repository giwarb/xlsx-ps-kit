# レビューチェックリスト

各項目に ○ / × / -（該当なし）。× は指摘欄に必ず対応する記述を書く。

## 契約・スコープ

| # | 項目 |
|---|---|
| C-1 | 関数シグネチャが 01-design.md §3 と一致（引数名・型・既定値） |
| C-2 | 戻り値の形（オブジェクト／JSON キー名）が設計と一致 |
| C-3 | `Test-XlsFormulas` の JSON が recalc.py 契約と一致（`status`/`total_formulas`/`total_errors`/`error_summary`/`locations_truncated`） |
| C-4 | exit code 契約（`error` キーのときだけ非 0） |
| C-5 | 確定事項（§0）に触れる変更がない |
| C-6 | タスクカードの範囲外を触っていない |

## COM 安全性

| # | 項目 |
|---|---|
| S-1 | `GetActiveObject` を使っていない |
| S-2 | COM が `Invoke-XlsSession` 外に漏れていない（テスト含む） |
| S-3 | `finally` で Close→Quit→Release（逆順）→GC×2 |
| S-4 | ScriptBlock 内で `throw` してもプロセスが残らない（テストで確認） |
| S-5 | `Value2` を使い、`Value`/`Text` を数値・日付取得に使っていない |
| S-6 | 2 次元代入が `[object[,]]` |
| S-7 | `SpecialCells` のヒットなし例外を処理 |
| S-8 | `DisplayAlerts=$false` 中に `Close(SaveChanges:=$true)` していない |
| S-9 | 全殺し（`Get-Process EXCEL \| Stop-Process`）がない |
| S-10 | ReadOnly 検出・STA 検査がある（T-03 のみ） |

## PowerShell 互換

| # | 項目 |
|---|---|
| P-1 | `#Requires -Version 5.1` |
| P-2 | 三項演算子 / `??` / `?.` / `-Parallel` / `$IsWindows` なし |
| P-3 | ファイル I/O に `-Encoding UTF8` |
| P-4 | Pester 3.4 で動く構文（`Should Be` 形式、`-Because` なし、`BeforeAll` 内の変数スコープに注意） |

## テスト

| # | 項目 |
|---|---|
| T-1 | Pester が緑（自分で実行） |
| T-2 | `AfterEach` に `Assert-NoOrphanExcel` |
| T-3 | 実 Excel を起動している（モックでない） |
| T-4 | 02-implementation-plan.md §4 の受け入れ要点が全部テストにある |
| T-5 | 異常系（ファイル不在・保護シート・他プロセスが開いている）が最低 1 つある |

## 品質

| # | 項目 |
|---|---|
| Q-1 | コメントベースヘルプがある |
| Q-2 | エラーメッセージが次の一手を示す |
| Q-3 | COM 定数が `$script:Xl` に集約され、`com-constants.md` に追記されている |
| Q-4 | 実装メモに COM の罠が書かれている（空欄なら should-fix） |
| Q-5 | 自己チェック欄の○が事実と一致 |
