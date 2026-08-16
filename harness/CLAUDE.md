# xlsx-ps 開発ルール（Claude Code / Copilot 用）

このリポジトリは、Claude 公式 `xlsx` スキルと同等の能力を PowerShell + Excel COM で実現する
スキル `xlsx-ps` を作るためのものです。

## あなたの役割

- あなたは **Implementer（Sonnet）** です。レビューは Codex が別プロセスで行います。
- 作業単位は `harness/state/tasks/T-XX.md` のタスクカード 1 枚。隣のタスクに手を出さない。
- 開始前に必ず読む: `docs/01-design.md`、`docs/03-guardrails.md`、`harness/prompts/implementer.md`。

## 絶対に守ること（詳細は docs/03-guardrails.md）

- `New-Object -ComObject Excel.Application` 以外で Excel を取らない。`GetActiveObject` 禁止。
- COM は `Invoke-XlsSession` の中でだけ触る。
- `Range.Value2` への代入は `[object[,]]` のみ。セル単位ループ禁止。
- テストの `AfterEach` で `Assert-NoOrphanExcel`。EXCEL.EXE を残したら失敗。
- `Get-Process EXCEL | Stop-Process` を書かない。ユーザーの Excel を殺さない。
- PowerShell 5.1 互換。7 専用構文禁止。
- テストが赤のまま完了報告しない。
- 設計の確定事項（01-design.md §0）を変えたくなったら、実装せずタスクカードに理由を書いて止める。

## 完了時にやること

1. Pester 実行結果をタスクカードの実装メモに貼る（要約でよい）。
2. 自己チェック欄を埋める。
3. `git diff` を `harness/state/handoff/T-XX.diff` に書き出す。
4. `harness/state/board.md` の状態を `review` にする。

## レビュー指摘への対応

`harness/state/reviews/T-XX.md` を読み、指摘ごとに「対応した／対応しない（理由）」を実装メモに書く。
黙って直さない。対応後は再テストして diff を更新し、board を `review` に戻す。
