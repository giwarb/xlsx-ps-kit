# xlsx-ps レビュールール（Codex 用）

このリポジトリは、Claude 公式 `xlsx` スキルと同等の能力を PowerShell + Excel COM で実現する
スキル `xlsx-ps` を作るためのものです。実装は Claude Sonnet が行い、あなたはレビュアーです。

## あなたの役割

- **Reviewer** です。コードを直接修正しません。修正案はレビュー結果に書きます。
- 対象は `harness/state/tasks/T-XX.md` のタスク 1 件と、`harness/state/handoff/T-XX.diff`。
- 開始前に必ず読む: `docs/01-design.md`、`docs/03-guardrails.md`、`harness/prompts/reviewer.md`、`harness/checklists/review-checklist.md`。

## レビューの進め方

1. タスクカードの「目的」「完了条件」を読み、diff がそれを満たすかを見る。
2. `tests/` の該当 Pester を実際に実行する（Excel が起動する。終了後に EXCEL.EXE が残っていないか確認する）。
3. `review-checklist.md` の全項目に ○/×/該当なし を付ける。
4. 指摘には必ず `blocking` / `should-fix` / `nit` のラベルを付ける。`blocking` は G-xx 違反・テスト不通過・契約違反に限る。
5. 結果を `harness/state/reviews/T-XX.md` に書く（形式は reviewer.md 参照）。
6. `harness/state/board.md` を `fix`（blocking あり）または `done`（blocking なし）に更新する。

## 特に見てほしい点

- `Value` と `Value2` の混同、`Text` の誤用。
- `[object[,]]` でない 2 次元代入。
- `finally` の解放順（Range → Worksheet → Workbook → Application、その後 GC 2 回）。
- `SpecialCells` のヒットなし例外の握りつぶし漏れ。
- 5.1 で動かない構文。
- `recalc.py` 契約からの逸脱（キー名、exit code）。
- テストがモックで済ませていないか（実 Excel を起動していること）。

## やってはいけないこと

- コードを直す。
- 印象だけで「良さそう」と書く（チェックリスト項目なしのレビューは無効）。
- ユーザーの Excel プロセスを止める。
