# teams-daily-report - GitHub Copilot 指示ファイル

## スラッシュコマンド

Excel 日報への書き込みには Prompt File（カスタムスラッシュコマンド）を使用してください。

| コマンド | 用途 | 定義ファイル |
|---|---|---|
| `/daily-plan` | 計画を Excel の C列に書き込む | `.github/prompts/daily-plan.prompt.md` |
| `/daily-result` | 実績を Excel の F列に書き込む | `.github/prompts/daily-result.prompt.md` |

> 詳細な手順はすべて各 `prompt.md` に記述されています。
> `copilot-instructions.md` 側に手順を重複して書かないこと。

---

## 重要な設計ルール

### ファイルの役割分担

| ファイル | 役割 |
|---|---|
| `src/teams_chat_from_outlook.py` | Outlook メール経由の従来モード（引数なしで実行） |
| `src/daily_report_writer.py` | CLI 直接書き込みモード（スラッシュコマンドから呼び出す） |

### CLI 引数仕様（daily_report_writer.py）

```
--mode    plan または result（必須）
--summary 書き込む文字列（必須、50文字以内に自動切り詰め）
--date    MM/DD 形式（省略時は今日の日付を自動使用）
```

### Excel 書き込み先

| モード | 列 | 内容 |
|---|---|---|
| `plan` | C列 | 日報計画 |
| `result` | F列 | 日報実績 |

### 禁止事項

- `win32com.client`（Outlook COM）を `daily_report_writer.py` に追加しないこと
- スラッシュコマンドのプロンプトに Excel 書き込みロジックを直接記述しないこと（スクリプトに委譲する）
