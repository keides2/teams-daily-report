# GitHub Copilot スラッシュコマンドの使い方

このプロジェクトでは、Excel日報への書き込みを効率化するための GitHub Copilot カスタムスラッシュコマンドを提供しています。

## 利用可能なコマンド

| コマンド | 用途 | Excel 書き込み先 |
|---|---|---|
| `/daily-plan` | 本日の計画を登録 | C列（計画） |
| `/daily-result` | 本日の実績を登録 | F列（実績） |

## このプロジェクトでの使い方

1. GitHub Copilot Chat を開く（`Ctrl+Alt+I`）
2. `/daily-plan` または `/daily-result` と入力
3. プロンプトに従って計画または実績を入力
4. 自動的に Excel ファイルに書き込まれます

> **注意**: これらのコマンドはこのプロジェクト（リポジトリ）内でのみ使用できます。

---

## 他のプロジェクトでも使えるようにする方法

### 方法1: グローバルプロンプトディレクトリを設定（推奨）

すべてのプロジェクトで共通のスラッシュコマンドを使えるようにします。

#### 手順

1. **共通ディレクトリを作成**
   ```powershell
   mkdir C:\Users\<ユーザー名>\.copilot\prompts
   ```

2. **プロンプトファイルをコピー**
   ```powershell
   copy .github\prompts\daily-plan.prompt.md C:\Users\<ユーザー名>\.copilot\prompts\
   copy .github\prompts\daily-result.prompt.md C:\Users\<ユーザー名>\.copilot\prompts\
   ```

3. **VS Code の設定を開く**（`Ctrl+,`）

4. **設定を追加**
   
   `settings.json` を開いて以下を追加：
   ```json
   {
     "github.copilot.chat.prompts.directory": "C:\\Users\\<ユーザー名>\\.copilot\\prompts"
   }
   ```

   または、UI から直接設定：
   - 設定で `github.copilot.chat.prompts.directory` を検索
   - パスを入力: `C:\Users\<ユーザー名>\.copilot\prompts`

5. **VS Code を再起動**

これで、どのプロジェクトでも `/daily-plan` と `/daily-result` が使えるようになります。

#### メリット
- ✅ すべてのプロジェクトで利用可能
- ✅ 一元管理で更新が簡単
- ✅ プロジェクトごとにファイルをコピー不要

---

### 方法2: 各プロジェクトにコピー

プロジェクト固有のコマンドとして設定します。

#### 手順

他のプロジェクトで以下を実行：

```powershell
# .github/prompts ディレクトリを作成
mkdir .github\prompts

# プロンプトファイルをコピー
copy <このプロジェクトのパス>\.github\prompts\*.prompt.md .github\prompts\

# copilot-instructions.md もコピー（オプション）
copy <このプロジェクトのパス>\.github\copilot-instructions.md .github\
```

#### メリット
- ✅ プロジェクトごとにカスタマイズ可能
- ✅ プロジェクトと一緒にバージョン管理
- ✅ チームメンバーと共有しやすい

#### デメリット
- ❌ プロジェクトごとに管理が必要
- ❌ 更新時に全プロジェクトを修正

---

### 方法3: シンボリックリンク（上級者向け）

共通ディレクトリへのシンボリックリンクを作成します。

#### 手順（管理者権限が必要）

```powershell
# 管理者権限で PowerShell を起動
# 他のプロジェクトで実行
mkdir .github
mklink /D .github\prompts C:\Users\<ユーザー名>\.copilot\prompts
```

#### メリット
- ✅ 一元管理
- ✅ プロジェクトごとに設定の見た目は保たれる

#### デメリット
- ❌ 管理者権限が必要
- ❌ リポジトリにコミットできない
- ❌ 環境に依存

---

## カスタマイズ方法

### Excel ファイルのパスを変更する場合

Excel ファイルのパスは、プロンプトファイルではなく **環境変数**（または `.env`）で設定します。

`src/daily_report_writer.py` は起動時に `EXCEL_PATH` を参照し、`{fiscal_year}` / `{month_folder}` / `{month_name}` を自動置換して当月ファイルを特定します。

#### 例: Windows のユーザー環境変数に設定（推奨）

PowerShell で 1 回だけ実行（以後は永続）:

```powershell
[Environment]::SetEnvironmentVariable(
   "EXCEL_PATH",
   'C:\path\to\your\\{fiscal_year}年度\\{month_folder}\\（{month_name}_氏名）業務内容報告書.xlsm',
   "User"
)
```

反映確認（新しいターミナルで）:

```powershell
$env:EXCEL_PATH
```

#### 例: `.env` に設定（このリポジトリを開いて使う場合）

`.env.example` を `.env` にコピーして `EXCEL_PATH=` を設定します（`.env` は Git 管理対象にしないでください）。

### 書き込み先の列を変更する場合

`daily_report_writer.py` の以下の部分を修正：

```python
# 列の定義
PLAN_COLUMN = 3   # C列 → 他の列に変更可能
RESULT_COLUMN = 6  # F列 → 他の列に変更可能
```

---

## トラブルシューティング

### スラッシュコマンドが表示されない

1. VS Code を再起動
2. `.github/prompts/` ディレクトリが存在するか確認
3. ファイル名が `*.prompt.md` で終わっているか確認
4. グローバル設定の場合、パスが正しいか確認

### エラーが発生する

1. Python 環境が正しく設定されているか確認
2. Excel ファイルが閉じられているか確認（開いていると書き込めません）
3. `requirements.txt` の依存関係がインストールされているか確認

---

## 参考情報

- [GitHub Copilot カスタムプロンプト公式ドキュメント](https://docs.github.com/en/copilot/customizing-copilot/creating-custom-instructions-for-github-copilot)
- プロンプトファイルの場所:
  - プロジェクト固有: `.github/prompts/*.prompt.md`
  - グローバル: VS Code 設定で指定したディレクトリ
