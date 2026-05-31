import os
import json
import re
import gspread
import imaplib
import email
from email.header import decode_header

def clean_html_new(html_text):
    """HTMLタグとstyle/scriptの中身を完全に除去"""
    if not html_text: return ""
    text = re.sub(r'<style.*?>.*?</style>', '', html_text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script.*?>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    gmail_user = os.environ.get("GMAIL_ADDRESS")
    gmail_pass = os.environ.get("GMAIL_APP_PASSWORD")

    if not gmail_user or not gmail_pass:
        print("エラー: GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません。")
        return

    # スプレッドシート認証
    print("スプレッドシート認証中...")
    client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
    gc = gspread.oauth(
        credentials_filename=client_secret_path,
        authorized_user_filename="authorized_user.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet("Newsletters")
    rows = ws.get_all_values()

    # IMAPでGmailに接続
    print(f"Gmail (IMAP) に接続中: {gmail_user}")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_pass)
        mail.select("inbox")

        # Theory関連のメールを検索（エラー回避のためドメイン名のみで検索）
        search_query = '(FROM "theory.co.jp")'
        status, messages = mail.search(None, search_query)
        
        if status != "OK":
            print("メールの検索に失敗しました。")
            return

        msg_ids = messages[0].split()
        print(f"該当するメールを {len(msg_ids)} 件見つけました。再取得を開始します...")

        for m_id in msg_ids:
            res, msg_data = mail.fetch(m_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    message_id = msg.get("Message-ID", "").strip("<>")
                    
                    # 本文抽出
                    html_body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/html":
                                html_body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                                break
                    else:
                        if msg.get_content_type() == "text/html":
                            html_body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

                    if html_body:
                        clean_text = clean_html_new(html_body)
                        
                        # スプレッドシート内の該当する行を探して更新
                        for i, row in enumerate(rows):
                            if len(row) > 5 and row[5] == message_id:
                                ws.update_cell(i + 1, 4, clean_text[:2000])
                                print(f"Message-ID: {message_id} の本文を更新しました。")
                                break

        mail.logout()
        print("完了しました！")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
