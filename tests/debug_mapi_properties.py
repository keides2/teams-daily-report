"""
Teamsチャットメールの詳細なMAPIプロパティを調査
"""
import win32com.client
import datetime
import re

def debug_mapi_properties():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Teams日報フォルダを取得
    inbox = outlook.GetDefaultFolder(6)  # 6 = olFolderInbox
    teams_folder = inbox.Folders("Teams日報")
    
    messages = teams_folder.Items
    messages.Sort("[ReceivedTime]", True)
    
    print(f"メール数: {messages.Count}")
    print("=" * 80)
    
    if messages.Count > 0:
        mail = messages.GetFirst()
        
        print(f"Subject: {mail.Subject}")
        print(f"ReceivedTime: {mail.ReceivedTime}")
        print(f"SentOn: {mail.SentOn}")
        print(f"CreationTime: {mail.CreationTime}")
        print()
        
        # メール本文を詳しく調べる
        body = mail.Body
        html_body = mail.HTMLBody
        
        print("=" * 80)
        print("Body全文:")
        print(body)
        print()
        
        print("=" * 80)
        print("HTMLBody全文:")
        print(html_body)
        print()
        
        # タイムスタンプのパターンを検索
        # 例: "2025年12月20日 10:30" や "12/20 10:30" など
        timestamp_patterns = [
            r'(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})',
            r'(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})',
            r'(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})',
            r'Posted:\s*(.+?)(?:\n|<)',
            r'Sent:\s*(.+?)(?:\n|<)',
            r'datetime="(.+?)"',
            r'data-time="(.+?)"',
        ]
        
        print("=" * 80)
        print("タイムスタンプ候補:")
        for pattern in timestamp_patterns:
            matches = re.findall(pattern, body, re.IGNORECASE)
            if matches:
                print(f"Pattern '{pattern}': {matches}")
            
            html_matches = re.findall(pattern, html_body, re.IGNORECASE)
            if html_matches:
                print(f"HTML Pattern '{pattern}': {html_matches}")
        
        print()
        
        # MAPIプロパティを調べる
        print("=" * 80)
        print("主要なMAPIプロパティ:")
        
        # よく使われるMAPIプロパティのタグ
        mapi_properties = {
            "PR_CLIENT_SUBMIT_TIME": "http://schemas.microsoft.com/mapi/proptag/0x0039",
            "PR_MESSAGE_DELIVERY_TIME": "http://schemas.microsoft.com/mapi/proptag/0x0E06",
            "PR_CREATION_TIME": "http://schemas.microsoft.com/mapi/proptag/0x3007",
            "PR_LAST_MODIFICATION_TIME": "http://schemas.microsoft.com/mapi/proptag/0x3008",
            "PR_TRANSPORT_MESSAGE_HEADERS": "http://schemas.microsoft.com/mapi/proptag/0x007D",
        }
        
        try:
            property_accessor = mail.PropertyAccessor
            for name, tag in mapi_properties.items():
                try:
                    value = property_accessor.GetProperty(tag)
                    print(f"{name}: {value}")
                except Exception as e:
                    print(f"{name}: エラー - {e}")
        except Exception as e:
            print(f"PropertyAccessorエラー: {e}")
        
        print()
        
        # メッセージヘッダーがあれば表示
        try:
            headers = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
            print("=" * 80)
            print("Message Headers:")
            print(headers)
        except:
            print("Message Headers: 取得できませんでした")

if __name__ == "__main__":
    debug_mapi_properties()
