# teams-daily-report

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)

![teams-daily-report](Gemini_Generated_Image_ko1xckko1xckko1x.png)

Teams のグループチャットに投稿した内容を **Outlook 共有** を起点に取得し、  
ローカル PC 上の **Excel 日報** に自動追記するツールです。

本ツールは以下の制約下でも動作する設計になっています。

- Microsoft Graph の **アプリ登録が不可**
- Power Automate / Power Apps を使わない
- 管理者権限を必要としない
- Teams の通常運用（チャット投稿）を変えない

---

## 全体アーキテクチャ

```
Teams チャット
  └─ メッセージ … → Outlook で共有
        ↓
Outlook（Windows ローカル）
        ↓
Python（pywin32）
  - 日報タグ抽出
  - 日付情報の解析
  - 要約（50文字）取得
  - 重複防止
        ↓
Excel（日報テンプレート）
```

---

## チャット投稿ルール（重要）

日報に反映したいメッセージには、以下の形式でタグと要約を記載します。

### 日付を指定する場合（推奨）

過去のチャットを後からまとめて共有する場合や、正確な日付を記録したい場合に使用します。

```
#日報 12/20 会議資料作成
#日報計画 12/25 プロジェクト計画書作成
#日報結果 12/24 テスト完了
```

### 日付を指定しない場合（後方互換）

メール受信日（Outlookで共有した日）に記録されます。

```
#日報 ログ保存先不具合対応
#日報計画 要約:設計レビュー準備
#日報結果 要約:顧客対応完了
```

### タグの種類

- `#日報計画` - 計画列（C列）に記録
- `#日報結果` - 実績列（F列）に記録
- `#日報` - 実績列（F列）に記録（結果として扱う）

### ルール

- 要約は **50文字以内**（自動で切り詰められます）
- `要約:` プレフィックスは省略可能
- 1つのメッセージに複数のタグを含めることが可能
- 投稿者は **自分／他人どちらでも可**
  - 日報に残したいメッセージだけを Outlook 共有する

---

## セットアップ

### 1. Python 環境

Python 3.9 以上（Windows）

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

※ pywin32 の COM 登録が必要な場合：

```powershell
python -m pywin32_postinstall -install
```

---

### 2. 環境設定ファイル（.env）の作成

`.env.example` をコピーして `.env` ファイルを作成し、実際の環境に合わせて設定します。

```powershell
Copy-Item .env.example .env
notepad .env
```

`.env` ファイルで設定する主な項目：

```bash
# Outlook設定
OUTLOOK_FOLDER=Teams日報

# Excelパス（方法1: 直接指定 - 推奨）
EXCEL_PATH=C:\Users\yamada\OneDrive - 株式会社ABC\営業部\業務管理\業務内容報告書\{fiscal_year}年度\{month_folder}\営業1課\（{month_name}_山田太郎）業務内容報告書.xlsm

# または（方法2: テンプレートから生成）
EXCEL_COMPANY=株式会社ABC
EXCEL_DEPARTMENT=営業部-01.庶務
EXCEL_TEAM=営業1課
EXCEL_USER_NAME=山田太郎
```

**ポイント**:
- `.env` ファイルは個人情報を含むため、Git にコミットされません
- `{fiscal_year}`, `{month_folder}`, `{month_name}` は実行時に自動置換されます
- **月が替わっても `.env` ファイルの更新は不要です**（テンプレート変数が自動的に置換されます）
- 詳細は `.env.example` を参照してください

---

### 3. Outlook 側の準備

1. Outlook（Windowsアプリ）を起動
2. **受信トレイ配下に**フォルダを作成  
   例: `Teams日報`（`.env` の `OUTLOOK_FOLDER` と同じ名前にする）
3. Teams チャットで日報にしたいメッセージを  
   **「… → Outlook で共有」**

（必要であれば Outlook ルールで自動振り分け可能）

---

### 4. Excel 日報の構造

本ツールは特定のExcel構造を前提としています。

#### ファイル配置

Excel ファイルのパスは `.env` ファイルで設定します（上記参照）。

自動生成されるパスの例:
```text
{OneDrive}\業務内容報告書\{年度}年度\{月フォルダ}\{グループ名}\{氏名}.xlsm
```

具体例:
```text
C:\Users\{username}\OneDrive - 株式会社サンプル\部署名-01.庶務事項\
業務管理\業務内容報告書\2025年度\09_12月\チームA\
（12月_山田太郎）業務内容報告書.xlsm
```

**注意**: 年度・月フォルダーは実行日から自動計算されます（4月始まりの年度）。

#### テンプレート変数の自動置換

Excel パスには以下のテンプレート変数が使用できます：

| 変数 | 説明 | 例（2025年12月実行時） | 例（2026年1月実行時） |
|------|------|------------------------|----------------------|
| `{fiscal_year}` | 年度（4月始まり） | 2025 | 2025 |
| `{month_folder}` | 月フォルダー形式 | 09_12月 | 10_1月 |
| `{month_name}` | 月名のみ | 12月 | 1月 |

**メリット**: 一度設定すれば、月が替わっても自動的に正しいパスが生成されます。

```bash
# .env ファイルの設定例
EXCEL_PATH=C:\Users\yamada\OneDrive\業務内容報告書\{fiscal_year}年度\{month_folder}\（{month_name}_山田太郎）.xlsm

# 2025年12月に実行 → C:\Users\yamada\OneDrive\業務内容報告書\2025年度\09_12月\（12月_山田太郎）.xlsm
# 2026年1月に実行 → C:\Users\yamada\OneDrive\業務内容報告書\2025年度\10_1月\（1月_山田太郎）.xlsm
# 2026年4月に実行 → C:\Users\yamada\OneDrive\業務内容報告書\2026年度\01_4月\（4月_山田太郎）.xlsm
```

#### シート構造

- シート名: `実施申請・報告書（{年}{月}）` のパターンを含むシート
- 日付列: **B列** （datetime型、表示形式はmm/dd）
- 計画列: **C列** （入力エリア全体）
- 実績列: **F列** （入力エリア全体）
- 行間隔: 1日あたり **6行** + 2行の余白 = 8行間隔

**重要**: 日付セルの3行上から入力エリアが始まる構造に対応しています。

#### 日付ブロックの例

各日付は6行の入力ブロックで構成され、その後2行の余白があります。  
日付表示は入力ブロックの4行目（B列）にあります。

| 行 | B列 | C列(計画) | F列(実績) |
|----|-----|-----------|-----------|
| 205 | | 計画1 | 実績1 |
| 206 | | 計画2 | 実績2 |
| 207 | | | 実績3 |
| 208 | 12/24 | | 実績4 |
| 209 | (水) | | 実績5 |
| 210 | | | 実績6 |
| 211 | | （余白） | （余白） |
| 212 | | （余白） | （余白） |
| 213 | | 次の日付の入力開始 | |

スクリプトは自動的に各列の空き行を探して書き込みます。

---

## requirements.txt

```txt
requests
openpyxl
pywin32
python-dotenv
```

**依存パッケージの説明**:
- `pywin32`: Outlook COM 自動化
- `openpyxl`: Excel ファイル読み書き
- `python-dotenv`: 環境変数ファイル（`.env`）の読み込み
- `requests`: HTTP リクエスト（将来の拡張用）

---

## 実行

```powershell
python src\teams_chat_from_outlook.py
```

実行内容：

- Outlook フォルダから未処理メールを取得
- `#日報計画 / #日報結果 / #日報` を抽出
- 日付指定があればその日付、なければメール受信日に記録
- Excel の指定日付行に追記
- 処理済みメールは EntryID で管理（`processed_mail_ids.json` に保存、重複防止）

### 処理の流れ

1. **メール取得**: Outlook の `Teams日報` フォルダから未処理メールを取得
2. **日報抽出**: メール本文から `#日報` タグを検索
3. **日付判定**: 
   - タグに日付指定（例: `#日報 12/20`）があればその日付を使用
   - 日付指定がなければメール受信日を使用
4. **Excel書き込み**:
   - 計画（`#日報計画`）→ C列の1行目
   - 実績（`#日報結果`, `#日報`）→ F列の2〜6行目（空き行を自動検索）
5. **重複防止**: 処理済みメールIDを保存

### 再実行する場合

処理済みメールIDをクリアすれば再処理できます：

```powershell
Remove-Item processed_mail_ids.json
python src\teams_chat_from_outlook.py
```

---

## 運用方法

### 基本的な流れ

本ツールは **「Outlookで共有」と「スクリプト実行」の2ステップ** で動作します。

#### ステップ1: Teamsからメールを転送

1. Teamsチャットで日報を投稿
   ```
   #日報計画 12/24 要約: 設計書作成
   ```

2. メッセージの **「… → Outlookで共有」** をクリック
   - Outlookの受信トレイに届く
   - 手動で「Teams日報」フォルダに移動
   - または、Outlookルールで自動振り分け設定

#### ステップ2: スクリプト実行

```powershell
python src\teams_chat_from_outlook.py
```

### 推奨運用パターン

#### パターンA: まとめて処理（推奨）✨

**「Outlookで共有」は投稿の都度、スクリプトは1日1回**

```
【朝】
Teams投稿: #日報計画 12/24 要約: 設計書作成
→ すぐに「Outlookで共有」をクリック

【昼】
Teams投稿: #日報結果 12/24 要約: 設計書50%完成
→ すぐに「Outlookで共有」をクリック

【夕方】
Teams投稿: #日報結果 12/24 要約: 設計書完成、レビュー依頼
→ すぐに「Outlookで共有」をクリック

【帰宅前（17:00-18:00）】
python src\teams_chat_from_outlook.py
→ 朝・昼・夕の3件がまとめてExcelに反映
```

**メリット**:
- スクリプト実行は1日1回でOK
- 投稿時は「Outlookで共有」ボタンを押すだけ（習慣化しやすい）
- 処理済み管理により、何度実行しても重複しない

#### パターンB: リアルタイム処理

投稿→共有→スクリプト実行を即座に行う方法。

**メリット**: Excelに即座に反映、確認しやすい  
**デメリット**: 投稿のたびにスクリプト実行が必要

### 複数の日報を一度に投稿

1つのメッセージに複数の日報を含めることも可能です：

```
#日報計画 12/24 要約: 設計書作成
#日報結果 12/24 要約: 要件定義完了
#日報結果 12/23 要約: 昨日のレビュー対応完了
```

→ スクリプト実行で3件すべてが処理されます

### デスクトップショートカット作成（オプション）

毎回コマンドを入力するのが面倒な場合、バッチファイルを作成：

```powershell
@"
@echo off
cd /d C:\Users\{あなたのユーザー名}\Docs\GitHub\keides2\teams-daily-report
python src\teams_chat_from_outlook.py
pause
"@ | Out-File -FilePath "$env:USERPROFILE\Desktop\日報更新.bat" -Encoding ascii
```

デスクトップの「日報更新.bat」をダブルクリックするだけで実行できます。

---

## 設計上のポイント

- **Graph API 不要** - pywin32 による COM 自動化のみ使用
- **管理者依存なし** - ローカル環境のみで完結
- **UI 作成なし** - コマンドライン実行のみ
- **Teams / Outlook の公式機能のみ使用** - 非公式APIやハックなし
- **壊れにくい運用** - シンプルな構成で保守性が高い
- **日付の柔軟性** - タグに日付を明示することで過去分も正確に記録可能
- **重複防止** - EntryID による処理済み管理で同じメールを二重処理しない
- **環境変数で設定管理** - `.env` ファイルで個人情報を管理、ソースコード修正不要

### 環境変数化のメリット

✅ **個人情報の保護**
- `.env` ファイルは `.gitignore` に含まれ、Git にコミットされない
- 氏名、会社名などの個人情報をリポジトリに含めない
- パブリックリポジトリとして安全に公開可能

✅ **簡単な環境設定**
- `.env.example` をコピーして編集するだけ
- ソースコード修正不要
- チームメンバーへの展開が容易

✅ **複数環境対応**
- 開発環境・本番環境で異なる設定を簡単に使い分け
- テスト用の設定ファイルも作成可能

---

## ファイル構成

```text
teams-daily-report/
├── README.md                      # このファイル
├── LICENSE                        # MITライセンス
├── requirements.txt               # Python依存パッケージ
├── .env.example                   # 環境変数設定例（コピーして使用）
├── .env                          # 環境変数設定（各自作成、Gitにコミットしない）
├── .gitignore                    # Git除外設定
├── processed_mail_ids.json        # 処理済みメールID（自動生成）
├── src/
│   └── teams_chat_from_outlook.py # メイン実行スクリプト
├── tests/
│   ├── README.md                  # テストスクリプトの説明書
│   ├── debug_mail_content.py      # メール内容デバッグ
│   ├── debug_mapi_properties.py   # MAPIプロパティ調査
│   ├── test_extract.py            # タグ抽出ロジックテスト
│   ├── test_excel_read.py         # Excel読み込みテスト
│   ├── check_excel_dates.py       # Excel日付確認
│   ├── debug_excel_dates.py       # Excel日付デバッグ
│   └── list_outlook_folders.py    # Outlookフォルダー一覧表示
└── evac/
    ├── README.md                  # 旧仕様ファイルの説明
    ├── fetch_chat_ids.py          # [使用不可] Graph API版（旧仕様）
    ├── .env                      # [使用不可] Graph API設定
    └── .env.example              # [使用不可] Graph API設定例
```

**重要**: `.env` ファイルは各自が作成する個人設定ファイルです。`.gitignore` に含まれているため、Git にコミットされません。

---

## トラブルシューティング

### Excelファイルが開かれているエラー

```
エラー: Excelファイルが開かれています
```

**解決方法**:
1. Excelファイルを閉じてください
2. もう一度スクリプトを実行してください

スクリプトは実行中にExcelファイルを読み書きするため、Excelで開いていると競合が発生します。

### Excel ファイルが見つからない

```
エラー: Excelファイルが見つかりません
```

**確認項目**:
- `.env` ファイルの `EXCEL_PATH` 設定を確認
- テンプレート変数（`{fiscal_year}`, `{month_folder}`, `{month_name}`）が正しく置換されているか確認
- 年度・月の計算ロジック（4月始まりの年度計算）を確認

**デバッグ**: 
```powershell
python tests\check_current_excel_dates.py
```

### 日付行が見つからない / セル位置がずれる

```
✗ スキップ: 2025-12-24 の日付行が見つかりません
```

**確認項目**:
- Excel のB列に日付が datetime 型で入力されているか確認
- シート名が月のパターン（例: "12月"）を含んでいるか確認
- `.env` ファイルの `DATE_COL`, `PLAN_COL`, `RESULT_COL` 設定を確認

**Excel構造の前提**:
- 日付セル（B列）の **3行上** から入力エリアが始まる
- 1日あたり6行の入力エリア + 2行の余白 = 合計8行

**デバッグ**: 
```powershell
python tests\check_current_excel_dates.py  # 日付行の位置を確認
```

### メールが処理されない

```
✗ スキップ: (件名)
  理由: 日報タグまたは要約が見つかりません
```

**確認項目**:
- Outlook フォルダ名が `.env` の `OUTLOOK_FOLDER` 設定と一致しているか確認
- メール本文に `#日報計画`, `#日報結果`, または `#日報` タグがあるか確認
- タグの後に要約テキストがあるか確認（空白のみは無効）

**対応パターン**:
```
#日報計画 12/24 要約: XXX      ← 日付指定あり（推奨）
#日報結果 12/24 要約: XXX      ← 日付指定あり（推奨）
#日報 12/24 XXX                ← 結果として処理
#日報計画 要約: XXX            ← 日付なし（受信日を使用）
```

**デバッグ**: 
```powershell
python tests\debug_mail_content.py  # メール内容とタグ検出を確認
```

### 処理済みメールを再処理したい

```powershell
Remove-Item processed_mail_ids.json
python src\teams_chat_from_outlook.py
```

`processed_mail_ids.json` を削除すると、すべてのメールが再処理されます。

---

## デバッグツール

問題が発生した場合は、`tests` フォルダー内のデバッグスクリプトを使用してください。
詳細は [tests/README.md](tests/README.md) を参照してください。

| スクリプト | 用途 |
|-----------|------|
| `check_current_excel_dates.py` | 現在の.env設定でExcelの日付行を確認 |
| `debug_mail_content.py` | Outlookメールの内容とタグ検出を確認 |
| `list_outlook_folders.py` | Outlookフォルダー一覧を表示 |
| `test_extract.py` | タグ抽出ロジックをテスト |

---

### 同じ内容が複数回書き込まれる

- `processed_mail_ids.json` が正しく更新されているか確認
- メール本文のプレビュー部分は自動的に除外されます（"Microsoft Teams" 以降のみ処理）
- **デバッグ**: `python tests\test_extract.py` でタグ抽出ロジックを確認

---

## 処理済みメール管理 (`processed_mail_ids.json`)

### 📝 ファイルの役割

このツールは、`processed_mail_ids.json` ファイルで処理済みメールを管理します。

**保存される情報:**
```json
[
  "00000000C431A3DBFEA37F47AD3334F526F6C83F07004E194AD5...",
  "00000000C431A3DBFEA37F47AD3334F526F6C83F07004E194AD5...",
  ...
]
```

各文字列は **Outlook メールの EntryID**（一意の識別子）です。

### 🎯 重複処理の防止

**動作の流れ:**

1. **起動時**: `processed_mail_ids.json` から処理済みIDを読み込み
2. **メール処理**: 各メールの EntryID を確認
   - 既に処理済み → スキップ
   - 未処理 → Excelに書き込み、IDをリストに追加
3. **終了時**: 更新されたIDリストを `processed_mail_ids.json` に保存

### 📊 Excel行番号の管理方法

**Q: Excelの何行目に入力したかという情報は記録していますか？**

**A: いいえ、記録していません。**

代わりに、**毎回Excelファイルを読み取って空セルを探します**。

```python
# 実行のたびに、Excelファイルの実際の状態を確認
for offset in range(0, ROWS_PER_DAY):
    cell = ws[f"{PLAN_COL}{target_row}"]
    if not cell.value:  # 空セルを発見
        cell.value = summary  # ここに書き込み
        break
```

**動作例:**

```
【1日目実行後】
C205: 会議の準備
C206: 資料作成
C207: (空)

【2日目実行時】
→ C205を確認: 値あり → スキップ
→ C206を確認: 値あり → スキップ
→ C207を確認: 空セル → ✅ ここに書き込み
```

**設計の利点:**
- ✅ 状態管理が不要（Excelファイルが唯一の真実の情報源）
- ✅ 手動編集に対応（ユーザーがExcelで行を削除/追加しても問題ない）
- ✅ シンプルな実装

### 🔄 処理済みメールを再処理したい場合

1. `processed_mail_ids.json` を削除（または該当IDを削除）
2. スクリプトを再実行

⚠️ **注意**: 既にExcelに書き込まれた内容は自動削除されません。手動で削除するか、空セルに追記されます。

### 🚫 なぜ `.gitignore` に含まれるのか？

**3つの理由:**

1. **個人情報**: あなたのOutlookメールボックス内のメール履歴情報
2. **環境固有**: EntryIDは各ユーザーのOutlook環境ごとに異なる
3. **運用データ**: ソースコードではなく、実行時に自動生成されるログファイルの性質

**類似する除外ファイル:**
```gitignore
.env                      # 個人の環境設定
processed_mail_ids.json   # 個人の処理履歴
*.xlsm                    # 個人のExcelファイル
*.log                     # ログファイル
```

---

## 想定利用シーン

- **個人日報の自動化** - チャットで仕事をしながら日報も自動記録
- **業務記録の証跡保存** - 「いつ何をしたか」を正確に記録
- **過去分の一括登録** - 日付を明示して過去のチャットをまとめて登録
- **チーム内の情報共有** - 他メンバーの重要な発言も日報に記録可能
- **依頼対応の記録** - 「誰の依頼で何をしたか」を残す用途

---

## 今後の拡張案（任意）

- Azure OpenAI / Azure AI による
  - チャット内容の自動要約
  - 日報の書式チェック
  - 週報・月報の自動生成
- 発言者名の Excel 自動記録
- 処理ログの詳細保存（CSVやログファイル）
- 複数人の日報を一括処理
- Outlook ルールの自動設定スクリプト

---

## ライセンス

MIT License

## 免責事項

このツールは個人の業務効率化を目的として作成されたものです。使用する場合は、各自の責任のもとで以下の点に注意してください：

- 実際の環境に合わせてパスや設定を変更してください
- Excelファイルのバックアップを必ず取得してください
- 会社のセキュリティポリシーを遵守してください
- Microsoft Teams、Outlook、Excel の利用規約を確認してください

作者は本ソフトウェアの使用によって生じたいかなる損害についても責任を負いません。
