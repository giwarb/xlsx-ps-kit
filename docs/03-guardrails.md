# 03 ガードレール

実装者（Sonnet）とレビュアー（Codex）の両方に適用する。
G-xx は「守られていなければ blocking」、S-xx は「守られていなければ should-fix」。

## A. スコープ

- **G-01** 01-design.md §0「確定事項」を変更しない。変更が必要だと思ったら実装せず、タスクカードの実装メモに理由を書いて止める。
- **G-02** JSON コマンドインタプリタ方式に戻さない。公開関数は §3 の 6 つ（＋ `Clear-XlsOrphans`）。書式・罫線・グラフ等の高水準ラッパーを追加しない。
- **G-03** OSS・外部モジュールを導入しない（Pester は 5.1 同梱の 3.4 を使う。v5 をインストールしない）。
- **G-04** `recalc.py` の JSON 契約（キー名・`errors_found` で exit 0・`error` キーで exit 非 0）を変えない。

## B. COM 安全性

- **G-05** `New-Object -ComObject Excel.Application` 以外で Excel を取得しない。`GetActiveObject` 禁止。
- **G-06** COM を触るコードは必ず `Invoke-XlsSession` の ScriptBlock 内、または `Invoke-XlsSession` 自身の中に書く。テストコードも同様。
- **G-07** すべてのテストは `tests/Common.ps1` の `Assert-NoOrphanExcel` を `AfterEach` で呼ぶ。テスト開始前の EXCEL.EXE PID 集合と終了後の差分がゼロでなければ失敗。
- **G-08** `Range.Value2` への代入は `[object[,]]` のみ。`foreach` で `Cells(r,c)` に代入するコードを書かない（`Set-XlsRange` の内部実装を含む）。
- **G-09** `DisplayAlerts=$false` の状態で `Workbook.Close(SaveChanges:=$true)` を呼ばない。保存は `Save-XlsWorkbook` のみ。
- **G-10** ユーザーの Excel を殺さない。`Clear-XlsOrphans` は自モジュールが書いたマーカーの PID しか対象にしない。`Get-Process EXCEL | Stop-Process` のような全殺しを書かない（テストコード含む）。

## C. PowerShell 互換

- **G-11** `#Requires -Version 5.1`。三項演算子、`??`、`?.`、`ForEach-Object -Parallel`、`$IsWindows` 依存を使わない。
- **S-01** ファイル I/O は `-Encoding UTF8` を明示。CSV は BOM 付き UTF-8。
- **S-02** 承認動詞（`Get-Verb`）のみ。関数名は `Verb-XlsNoun`。

## D. 品質

- **G-12** `reference/patterns.md` と `com-constants.md` に載せる内容は、実行して確認したものだけ。未確認のものは載せない。
- **G-13** 各公開関数に Pester テストがある。テストは実 Excel を起動する統合テストでよい（モックしない）。
- **S-03** 公開関数にはコメントベースヘルプ（`.SYNOPSIS` `.PARAMETER` `.EXAMPLE`）を付ける。
- **S-04** エラーは COM の HRESULT をそのまま投げず、`throw "Sheet '$name' not found"` のように次の一手が分かる文にする。

## E. 実装者（Sonnet）の振る舞い

- **G-14** 1 タスクカードで 1 関数（＋そのテスト）。隣の関数に手を出さない。
- **G-15** テストが赤のまま「完了」と報告しない。通せないなら実装メモに現状と仮説を書いて止める。
- **G-16** レビュー指摘への対応は、指摘ごとに「対応した／対応しない（理由）」を書く。黙って直さない。
- **S-05** 実装メモに、踏んだ COM の罠を書く。それが `SKILL.md` gotchas 章の原料になる。

## F. レビュアー（Codex）の振る舞い

- **G-17** レビューは `harness/checklists/review-checklist.md` の項目に沿って行い、項目ごとに ○/×/該当なし を付ける。印象論だけのレビューは無効。
- **G-18** 指摘は `blocking` / `should-fix` / `nit` の 3 段階で必ずラベルを付ける。`blocking` は G-xx 違反かテスト不通過か契約違反のみ。
- **G-19** レビュアーはコードを直接修正しない。修正案はレビュー結果に diff またはスニペットとして書く。
- **G-20** レビュー結果は `harness/state/reviews/T-XX.md` に書く。ターミナル出力だけで終わらせない。

## G. 自己チェック欄（タスクカード末尾にコピー）

```
| # | 項目 | ○/× |
|---|---|---|
| G-05 | GetActiveObject を使っていない | |
| G-06 | COM は Invoke-XlsSession 内のみ | |
| G-07 | Assert-NoOrphanExcel が AfterEach にある | |
| G-08 | Value2 代入は [object[,]] のみ | |
| G-11 | 7 専用構文なし | |
| G-13 | Pester テストがある | |
| G-15 | テストは緑 | |
```
