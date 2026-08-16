# xlsx-ps — PowerShell + Excel COM によるスプレッドシート操作スキル 開発キット

Claude 公式 `xlsx` スキル（openpyxl + pandas + markitdown + LibreOffice recalc）と
**同等の能力**を、Python が使えず PowerShell と Excel デスクトップ版だけが使える環境で
実現するためのスキル `xlsx-ps` を、コーディングエージェントに実装させるための初期セット。

- 実装担当: Claude Sonnet（Copilot / Claude Code のサブエージェント）
- レビュー担当: Codex（Codex CLI）
- オーケストレーション: 人間 or 上位エージェント（`harness/04-harness.md` 参照）

## ディレクトリ

```
xlsx-ps-kit/
  README.md                       このファイル
  docs/
    01-design.md                  方式設計（確定事項・部品対応・I/F 契約）
    02-implementation-plan.md     フェーズ分割・タスクカード・完了条件
    03-guardrails.md              実装・レビュー双方に課す禁止/必須事項
    04-harness.md                 Sonnet 実装 / Codex レビューの回し方
  skill/
    SKILL.md                      スキル本文ドラフト（実装完了後に確定）
    reference/com-constants.md    COM 定数表（骨子、実装時に充填）
    reference/patterns.md         COM スニペット集（骨子、実装時に充填）
    scripts/                      実装物の置き場（初期は空）
  harness/
    AGENTS.md                     Codex 用ルート指示
    CLAUDE.md                     Claude Code 用ルート指示（Copilot なら copilot-instructions.md に転記）
    prompts/implementer.md        Sonnet への委譲プロンプト雛形
    prompts/reviewer.md           Codex への委譲プロンプト雛形
    checklists/review-checklist.md
    benchmark/tasks.md            xlsx スキル同等性を測る 6 本のタスク
    state/                        タスクカード・レビュー結果の受け渡し場所
  tests/
    test-plan.md                  Pester テスト計画
```

## 最初に読む順序

1. `docs/01-design.md` — 何を作るか
2. `docs/03-guardrails.md` — 何をしてはいけないか
3. `docs/04-harness.md` — どう回すか
4. `docs/02-implementation-plan.md` — どの順で作るか

## 「同等」の定義

`harness/benchmark/tasks.md` の 6 タスクを、xlsx スキルと同じ手数・同じ品質で完走できること。
機能の網羅ではなく、このベンチマークの通過を Definition of Done とする。
