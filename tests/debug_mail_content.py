"""
メールの詳細内容を確認するスクリプト
"""

import win32com.client
import datetime as dt

OUTLOOK_FOLDER = "Teams日報"

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)

target_folder = None
for f in inbox.Folders:
    if f.Name == OUTLOOK_FOLDER:
        target_folder = f
        break

if not target_folder:
    print(f"フォルダが見つかりません: {OUTLOOK_FOLDER}")
    exit(1)

print(f"メール件数: {target_folder.Items.Count}")
print("=" * 80)

for idx, mail in enumerate(target_folder.Items, 1):
    print(f"\n--- メール {idx} ---")
    print(f"件名: {mail.Subject}")
    print(f"差出人: {mail.SenderName}")
    print(f"受信日時: {mail.ReceivedTime}")
    print(f"送信日時: {mail.SentOn}")
    print(f"作成日時: {mail.CreationTime}")
    print()
    body = mail.Body or ""
    print(f"本文全体（{len(body)}文字）:")
    print("-" * 80)
    print(body)
    print("-" * 80)
    print()
    
    # #日報が何回出現するか
    count = body.count("#日報")
    print(f"#日報の出現回数: {count}回")
    print()
