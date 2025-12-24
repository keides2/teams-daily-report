"""
extract_daily_reports関数のテスト
"""
import re
import datetime as dt

def extract_daily_reports(text: str):
    """テスト用の抽出関数"""
    reports = []
    current_year = dt.date.today().year
    
    print("=" * 80)
    print("入力テキスト:")
    print(repr(text))
    print()
    
    # #日報計画を全て抽出（日付付き）
    pattern1 = r"#日報計画\s+(\d{1,2})/(\d{1,2})\s*(?:要約[:：]\s*)?(.+?)(?:\n|$)"
    matches = list(re.finditer(pattern1, text))
    print(f"#日報計画 (日付付き): {len(matches)}件")
    for m in matches:
        print(f"  - {m.group(0)}")
    
    # #日報結果を全て抽出（日付付き）
    pattern2 = r"#日報結果\s+(\d{1,2})/(\d{1,2})\s*(?:要約[:：]\s*)?(.+?)(?:\n|$)"
    matches = list(re.finditer(pattern2, text))
    print(f"#日報結果 (日付付き): {len(matches)}件")
    for m in matches:
        print(f"  - {m.group(0)}")
    
    # #日報（単独）を全て抽出（日付付き）
    pattern3 = r"#日報\s+(\d{1,2})/(\d{1,2})\s+(.+?)(?:\n|$)"
    matches = list(re.finditer(pattern3, text))
    print(f"#日報 (日付付き): {len(matches)}件")
    for m in matches:
        print(f"  - {m.group(0)}")
    
    # 日付なしのパターン
    pattern4 = r"#日報\s+(.+?)(?:\n|$)"
    matches = list(re.finditer(pattern4, text))
    print(f"#日報 (全て): {len(matches)}件")
    for m in matches:
        summary = m.group(1).strip()[:50]
        is_date = re.match(r"\d{1,2}/\d{1,2}", summary)
        print(f"  - マッチ: {repr(m.group(0))}")
        print(f"    要約: {repr(summary)}")
        print(f"    日付形式: {is_date is not None}")
        if summary and "計画" not in summary and "結果" not in summary and not is_date:
            reports.append(("result", summary, None))
            print(f"    → 採用")
        else:
            print(f"    → スキップ")
    
    return reports


# テスト
test_text = """
SUZUKI Ichiro 鈴木 一郎 ログ保存先不具合の件、テストまで完了しました   #日報 ログ保存先不具合対応 ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌                                  ‌
 ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌                                                                                               ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌                                                                                                                                           Microsoft Teams
"""

reports = extract_daily_reports(test_text)
print()
print("=" * 80)
print(f"抽出結果: {len(reports)}件")
for kind, summary, date in reports:
    print(f"  - {kind}: {summary} (日付: {date})")
