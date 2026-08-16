# 04 ハーネス — Sonnet 実装 / Codex レビュー

## 1. 役割

| 役割 | 担当 | 入口 | 書く場所 |
|---|---|---|---|
| Orchestrator | 人間、または上位エージェント（Claude Code / Copilot の親エージェント） | `harness/state/board.md` | タスクカード発行、判定 |
| Implementer | Claude Sonnet | `harness/prompts/implementer.md` ＋ タスクカード | `skill/scripts/`, `tests/`, タスクカード実装メモ |
| Reviewer | Codex（Codex CLI） | `harness/prompts/reviewer.md` ＋ タスクカード ＋ diff | `harness/state/reviews/T-XX.md` |

Implementer と Reviewer は互いのターミナル出力を見ない。受け渡しは **すべてファイル**（`harness/state/`）で行う。
これは以前の Claude Code ⇄ Codex CLI 連携と同じ流儀で、どちらのツールに載せても動く。

## 2. 状態ディレクトリ

```
harness/state/
  board.md                 タスク一覧と状態（todo / impl / review / fix / done）
  tasks/T-XX.md            タスクカード（02-implementation-plan.md §3 の雛形から）
  reviews/T-XX.md          レビュー結果（1 タスクに複数ラウンドあれば T-XX-r2.md）
  handoff/T-XX.diff        実装者が生成した diff（`git diff` の出力）
```

`board.md` の行例:
```
| T-03 | Invoke-XlsSession | review | r1 | 2026-08-18 |
```

## 3. 1 タスクのループ

```
Orchestrator: tasks/T-XX.md を作成 → board を impl に
   ↓
Sonnet:  implementer.md を読む → T-XX.md を読む → 実装 → Pester 緑 → 自己チェック記入
         → git diff > handoff/T-XX.diff → board を review に
   ↓
Codex:   reviewer.md を読む → T-XX.md と handoff/T-XX.diff とテスト結果を読む
         → checklist に沿ってレビュー → reviews/T-XX.md を書く → board を fix or done に
   ↓
Sonnet（fix の場合）: reviews/T-XX.md を読む → 指摘ごとに対応/不対応を実装メモに書く
         → 再テスト → handoff/T-XX.diff 更新 → board を review に（r2）
   ↓
Orchestrator: done なら次のタスクカードを発行。3 ラウンド超えたら人間に上げる
```

## 4. Codex にレビューさせる理由と設計上の含み

- Sonnet が書いたコードを同系統モデルにレビューさせると盲点が重なる。Codex は PowerShell / COM の知識分布が異なるので、`Value` vs `Value2`、`[object[,]]`、finally の順序といった典型的な罠を別角度で拾える。
- Codex には**コードを直させない**（G-19）。直させると Sonnet 側の学習（実装メモ→gotchas）が積み上がらない。
- Codex には Excel を起動させてよい（テストを実行して確認させる）。ただし G-07 のプロセス残骸チェックは Codex 側でも走る。

## 5. Claude Code で回す場合

- ルートに `harness/CLAUDE.md` を置く（Claude Code はカレントの `CLAUDE.md` を読む）。
- Sonnet への委譲: `/task T-03` のようなスラッシュコマンドを `.claude/commands/task.md` として用意し、中身は `implementer.md` を読んで `state/tasks/$ARGUMENTS.md` を実行、とする。モデルは `--model sonnet` 相当のサブエージェント設定。
- Codex への委譲: `codex exec "$(cat harness/prompts/reviewer.md) タスク: T-03"` をシェルから叩く。Codex は `harness/AGENTS.md` を読む。

## 6. Copilot で回す場合

- `harness/CLAUDE.md` の内容を `.github/copilot-instructions.md` に転記。
- Implementer サブエージェントの frontmatter で model を Sonnet に固定、`prompts/implementer.md` を本文にする。
- Reviewer は Copilot 内で Codex 系モデルを指定するか、外部で Codex CLI を叩く。後者のほうが「互いの出力を見ない」を守りやすい。

## 7. エスカレーション条件（人間が介入する）

- 同一タスクで review ラウンドが 3 を超えた。
- G-01（確定事項の変更）に該当する提案が実装メモに書かれた。
- `Assert-NoOrphanExcel` が失敗する原因が 2 ラウンドで特定できない。
- ベンチマーク B-x が「同等でない」と判定され、原因が設計側にある。

## 8. ハーネス自体の完了条件

T-01〜T-16 が done で、`harness/benchmark/tasks.md` の記録表が埋まり、`SKILL.md` に実装メモ由来の gotchas が反映されていること。
