<#
.SYNOPSIS
    複数の .xlsm ファイルの ThisWorkbook モジュールを一括更新する

.DESCRIPTION
    指定フォルダー配下の月次フォルダーにある .xlsm を対象に、
    vba\ThisWorkbook.cls の内容を ThisWorkbook モジュールとして上書きインポートする。

.PARAMETER RootFolder
    月次フォルダーが並ぶ親フォルダー（例: 2026年度）

.PARAMETER ClsFile
    インポートする ThisWorkbook.cls の絶対パス

.EXAMPLE
    cd "C:\Users\shimatani\Docs\GitHub\keides2\teams-daily-report\vba"
    .\Update-ThisWorkbook.ps1 `
        -RootFolder "C:\Users\shimatani\OneDrive - DAIKIN\ITソリューション開発グループ-01.庶務事項\コロナ対応\在宅勤務関連\提出BOX）業務内容報告書\2026年度" `
        -ClsFile    "C:\Users\shimatani\Docs\GitHub\keides2\teams-daily-report\vba\ThisWorkbook.cls"
    # 実行結果例:
    #   対象ファイル数: 12
    #   処理中: （4月_嶋谷圭介）業務内容報告書.xlsm
    #     ThisWorkbook を更新しました。
    #   ...
    #   === 完了: 成功 12 件 / エラー 0 件 ===
#>
param(
    [Parameter(Mandatory)][string]$RootFolder,
    [Parameter(Mandatory)][string]$ClsFile
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- 前提確認 ---
if (-not (Test-Path $RootFolder)) { Write-Error "RootFolder が見つかりません: $RootFolder"; exit 1 }
if (-not (Test-Path $ClsFile))    { Write-Error "ClsFile が見つかりません: $ClsFile";    exit 1 }

$xlsmFiles = Get-ChildItem -Path $RootFolder -Recurse -Filter "*.xlsm" |
             Where-Object { -not $_.Name.StartsWith("~") }   # 開きかけファイル除外

if ($xlsmFiles.Count -eq 0) { Write-Warning "対象 .xlsm が見つかりません。"; exit 0 }

Write-Host "対象ファイル数: $($xlsmFiles.Count)" -ForegroundColor Cyan
$xlsmFiles | ForEach-Object { Write-Host "  $_" }

# --- VBOM アクセスを有効化（未設定の場合） ---
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$origValue = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
if ($origValue -ne 1) {
    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1
    Write-Host "VBOM アクセスを有効化しました。" -ForegroundColor Yellow
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$ok  = 0
$err = 0

foreach ($file in $xlsmFiles) {
    try {
        Write-Host "`n処理中: $($file.Name)" -ForegroundColor Cyan
        $wb = $excel.Workbooks.Open($file.FullName)

        # ThisWorkbook コンポーネントを削除して再インポート
        $proj = $wb.VBProject
        foreach ($comp in @($proj.VBComponents)) {
            if ($comp.Name -eq "ThisWorkbook") {
                # ThisWorkbook は削除不可のため CodeModule をクリアして上書き
                $lines = $comp.CodeModule.CountOfLines
                if ($lines -gt 0) { $comp.CodeModule.DeleteLines(1, $lines) }

                # cls ファイルを読み込み（Shift-JIS のまま）
                $code = [System.IO.File]::ReadAllText($ClsFile, [System.Text.Encoding]::GetEncoding(932))

                # Attribute 行（ヘッダー部）を除いてコードのみ挿入
                $codeLines = $code -split "`r`n|`n"
                $startLine = ($codeLines | Select-String "^Option Explicit" | Select-Object -First 1).LineNumber
                if (-not $startLine) { $startLine = 10 }   # フォールバック
                $pureCode = ($codeLines | Select-Object -Skip ($startLine - 1)) -join "`r`n"

                $comp.CodeModule.InsertLines(1, $pureCode)
                Write-Host "  ThisWorkbook を更新しました。" -ForegroundColor Green
                break
            }
        }

        $wb.Save()
        $wb.Close($false)
        $ok++
    }
    catch {
        Write-Warning "  エラー: $($_.Exception.Message)"
        try { $wb.Close($false) } catch {}
        $err++
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# --- VBOM を元に戻す ---
if ($origValue -ne 1) {
    if ($null -eq $origValue) {
        Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $origValue
    }
    Write-Host "`nVBOM アクセスを元に戻しました。" -ForegroundColor Yellow
}

Write-Host "`n=== 完了: 成功 $ok 件 / エラー $err 件 ===" -ForegroundColor Cyan
