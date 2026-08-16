# Pester テスト計画

Pester 3.4（PowerShell 5.1 同梱）で動く構文で書く。実 Excel を起動する統合テスト。

## 共通（tests/Common.ps1）

- `Get-ExcelPids`: `Get-Process EXCEL -ErrorAction SilentlyContinue | % Id`
- `Assert-NoOrphanExcel -Baseline $pids`: 現在の PID 集合と差分がゼロでなければ `throw`
- `New-TempXlsxPath`: `$env:TEMP\xlsagent-tests\<guid>.xlsx`
- 各 Tests.ps1 の `BeforeEach` で baseline 取得、`AfterEach` で Assert、`AfterAll` で temp 掃除

## ファイル対応

| Tests.ps1 | 対象 | 主なケース |
|---|---|---|
| T01 | モジュール | ロード、Export 7 関数、Pester 3.x |
| T02 | スモーク | 起動→Quit→PID 差分ゼロ |
| Invoke-XlsSession | T-03 | 新規/既存/ReadOnly 例外/throw 時後始末/戻り値透過/STA |
| Save-XlsWorkbook | T-04 | xlsx/xlsm(HasVBProject)/保存後 Automatic |
| Clear-XlsOrphans | T-05 | 自 PID のみ回収、他 PID は残す |
| Get-XlsRange | T-06 | 数値/文字/日付 ISO/空 $null/エラー値、CSV/JSON 出力 |
| Get-XlsOverview | T-07 | 見出し、上限、省略表記 |
| Get-XlsModel | T-08 | formula+value2、names、links、tables、protected、-Range、-FormulasOnly |
| Set-XlsRange | T-09 | Get との往復同値、左上指定拡張、CSV/JSON 入力、日付 |
| Test-XlsFormulas | T-10〜12 | 種別カウント、101 件 truncation、refused/-Force、CLI exit code、timeout |

## 実行

```powershell
Invoke-Pester -Path tests -OutputFile tests/results.xml -OutputFormat NUnitXml
```
