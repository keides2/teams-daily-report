"""
Outlookの受信トレイ配下のフォルダ一覧を表示するスクリプト（詳細版）
"""

import win32com.client

def list_folders(parent, indent=0):
    """フォルダを再帰的に表示"""
    for folder in parent.Folders:
        print("  " * indent + f"- {folder.Name} (Count: {folder.Items.Count})")
        if folder.Folders.Count > 0:
            list_folders(folder, indent + 1)

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

print(f"受信トレイ: {inbox.Name}")
print("=" * 60)

list_folders(inbox)

if inbox.Folders.Count == 0:
    print("  (フォルダが見つかりません)")

print("=" * 60)
