import os
import json
import re
import gspread
from google.oauth2.service_account import Credentials

def clean_html_advanced(html_text):
    """HTMLタグを除去してテキストのみにする（style, scriptは中身ごと削除）"""
    if not html_text: return ""
    # すでにCSSが混じっているテキストからも、特定のパターンを除去
    text = re.sub(r'<style.*?>.*?</style>', '', html_text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script.*?>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
    # CSSの断片（中身だけ残ってしまった場合）を除去
    text = re.sub(r'\{[^\}]+\}', '', text) # { ... } を削除
    text = re.sub(r'[a-z0-9\-\.#\s,]+:[^;]+;', '', text) # property: value; を削除
    
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    
    # 認証（ブラウザ認証）
    print("認証を開始します...")
    client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
    gc = gspread.oauth(
        credentials_filename=client_secret_path,
        authorized_user_filename="authorized_user.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet("Newsletters")
    rows = ws.get_all_values()
    
    if len(rows) <= 1:
        print("データがありません。")
        return

    header = rows[0]
    data = rows[1:]
    
    # 本文の列（index 3）
    updated_count = 0
    for i, row in enumerate(data):
        original_body = row[3]
        # CSSが含まれているかチェック ( Theory や .isPc など)
        if "{" in original_body or "font-family" in original_body or "@media" in original_body:
            new_body = clean_html_advanced(original_body)
            if new_body != original_body:
                row[3] = new_body
                ws.update_cell(i + 2, 4, new_body) # 4列目（本文）を更新
                updated_count += 1
                print(f"[{updated_count}] 行目 {i+2} を清掃しました")

    print(f"完了！合計 {updated_count} 件のデータを清掃しました。")

if __name__ == "__main__":
    main()
