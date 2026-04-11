import imaplib
import email
from email.header import decode_header
import os
import json
import sys
from datetime import datetime, timedelta, timezone

import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ────────────────────────────────────────────────────────────────────

# 監視対象のドメイン（これらが含まれる送信元からのメールを収集）
TARGET_DOMAINS = [
    "baycrews.co.jp", "baycrews.jp",
    "tomorrowland.co.jp",
    "jun.co.jp", "junonline.jp",
    "world.co.jp",
    "mash-holdings.com",
    "daytona-park.com",
    "tsi-holdings.jp",
    "stripe-intl.com",
    "united-arrows.co.jp"
]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HEADERS = ["取得日時", "送信者", "件名", "送信日時", "Message-ID"]

# ── 認証 ────────────────────────────────────────────────────────────────────

def get_spreadsheet(spreadsheet_id, sheet_name):
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise EnvironmentError("GOOGLE_CREDENTIALS_JSON が設定されていません。")
    
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
        worksheet.append_row(HEADERS)
        print(f"シート '{sheet_name}' を新規作成しました。")
    return worksheet

# ── Gmail 取得 ───────────────────────────────────────────────────────────────

def fetch_emails():
    user = os.environ.get("GMAIL_ADDRESS")
    password = os.environ.get("GMAIL_APP_PASSWORD")
    
    if not user or not password:
        raise EnvironmentError("GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません。")

    # Gmail IMAP 接続
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)
    mail.select("inbox")

    # 過去3日間のメールを検索（余裕を持って取得）
    date_since = (datetime.now() - timedelta(days=3)).strftime("%d-%b-%Y")
    
    found_emails = []
    
    # ドメインごとに検索
    for domain in TARGET_DOMAINS:
        search_query = f'FROM "{domain}" SINCE {date_since}'
        status, response = mail.search(None, search_query)
        
        if status != "OK":
            continue
            
        for msg_id in response[0].split():
            status, data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue
                
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # 件名のデコード
            subject, encoding = decode_header(msg.get("Subject", ""))[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")
                
            # 送信者のデコード
            from_, encoding = decode_header(msg.get("From", ""))[0]
            if isinstance(from_, bytes):
                from_ = from_.decode(encoding or "utf-8")
                
            # 日時
            date_str = msg.get("Date")
            
            found_emails.append({
                "sender": from_,
                "subject": subject,
                "date": date_str,
                "message_id": msg.get("Message-ID", "")
            })
            
    mail.logout()
    return found_emails

# ── メイン ───────────────────────────────────────────────────────────────────

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        print("エラー: SPREADSHEET_ID が必要です。")
        sys.exit(1)
        
    sheet_name = "Newsletters"
    
    print("Google スプレッドシートに接続中...")
    worksheet = get_spreadsheet(spreadsheet_id, sheet_name)
    
    print("既存のMessage-IDを取得中...")
    all_values = worksheet.get_all_values()
    existing_ids = set()
    for row in all_values[1:]:
        if len(row) >= 5:
            existing_ids.add(row[4])
            
    print("Gmailからメルマガを検索中...")
    emails = fetch_emails()
    print(f"  検索結果: {len(emails)} 件")
    
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []
    
    for em in emails:
        if em["message_id"] in existing_ids:
            continue
            
        new_rows.append([
            now,
            em["sender"],
            em["subject"],
            em["date"],
            em["message_id"]
        ])
        existing_ids.add(em["message_id"])
        
    if new_rows:
        print(f"\n{len(new_rows)} 件の新しいメルマガを追記します...")
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        print("完了しました。")
    else:
        print("\n新しいメルマガはありませんでした。")

if __name__ == "__main__":
    main()
