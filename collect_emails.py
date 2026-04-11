import imaplib
import email
from email.header import decode_header
import os
import json
import sys
import re
from datetime import datetime, timedelta, timezone

import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ────────────────────────────────────────────────────────────────────

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
# 新しいヘッダー構成（本文を追加）
HEADERS = ["取得日時", "送信者", "件名", "本文（抜粋）", "送信日時", "Message-ID"]

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

# ── 本文抽出 ────────────────────────────────────────────────────────────────

def clean_html(html):
    """HTMLタグを除去してテキストのみにする"""
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_body(msg, limit=500):
    """メール本文を抽出して指定文字数で切り出す"""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if content_type == "text/plain" and "attachment" not in content_disposition:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="ignore")
                break
            elif content_type == "text/html" and not body:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body = clean_html(payload.decode(charset, errors="ignore"))
    else:
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset() or "utf-8"
        body = payload.decode(charset, errors="ignore")
        if msg.get_content_type() == "text/html":
            body = clean_html(body)

    return (body[:limit] + "...") if len(body) > limit else body

# ── Gmail 取得 ───────────────────────────────────────────────────────────────

def fetch_emails():
    user = os.environ.get("GMAIL_ADDRESS")
    password = os.environ.get("GMAIL_APP_PASSWORD")
    if not user or not password:
        raise EnvironmentError("GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません。")
    
    password = password.replace(" ", "") # スペースが入っていた場合の対策

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)
    mail.select("inbox")

    # 過去3日間のメールを検索
    date_since = (datetime.now() - timedelta(days=3)).strftime("%d-%b-%Y")
    found_emails = []
    
    for domain in TARGET_DOMAINS:
        search_query = f'FROM "{domain}" SINCE {date_since}'
        status, response = mail.search(None, search_query)
        if status != "OK": continue
            
        for msg_id in response[0].split():
            status, data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK": continue
            msg = email.message_from_bytes(data[0][1])
            
            subject, encoding = decode_header(msg.get("Subject", ""))[0]
            if isinstance(subject, bytes): subject = subject.decode(encoding or "utf-8")
                
            from_, encoding = decode_header(msg.get("From", ""))[0]
            if isinstance(from_, bytes): from_ = from_.decode(encoding or "utf-8")
            
            found_emails.append({
                "sender": from_,
                "subject": subject,
                "body": extract_body(msg, limit=500),
                "date": msg.get("Date"),
                "message_id": msg.get("Message-ID", "")
            })
            
    mail.logout()
    return found_emails

# ── メイン ───────────────────────────────────────────────────────────────────

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id: sys.exit(1)
        
    worksheet = get_spreadsheet(spreadsheet_id, "Newsletters")
    all_rows = worksheet.get_all_values()
    
    # ヘッダーの確認と更新（旧形式からの移行）
    if not all_rows or all_rows[0] != HEADERS:
        print("スプレッドシートのヘッダーを更新します...")
        if not all_rows:
            all_rows = [HEADERS]
        else:
            # 既存のデータを新形式にマップし直す（簡易的なマイグレーション）
            new_data = [HEADERS]
            for row in all_rows[1:]:
                # [取得日時, 送信者, 件名, 送信日時, Message-ID] -> [取得日時, 送信者, 件名, 本文, 送信日時, Message-ID]
                if len(row) == 5:
                    new_data.append([row[0], row[1], row[2], "", row[3], row[4]])
                else:
                    new_data.append(row)
            all_rows = new_data
        worksheet.clear()
        worksheet.update("A1", all_rows)
        print("ヘッダーとデータの構造を更新しました。")

    # 既存データの Message-ID インデックスを作成
    id_map = {row[5]: i for i, row in enumerate(all_rows) if len(row) > 5}
    
    print("Gmailからメルマガを取得中...")
    emails = fetch_emails()
    
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    update_count = 0
    new_count = 0
    
    for em in emails:
        msg_id = em["message_id"]
        if msg_id in id_map:
            idx = id_map[msg_id]
            # 本文がまだ入っていない、または短い場合は更新（過去分への反映用）
            if len(all_rows[idx][3]) < 10:
                all_rows[idx][3] = em["body"]
                update_count += 1
        else:
            all_rows.append([now, em["sender"], em["subject"], em["body"], em["date"], msg_id])
            id_map[msg_id] = len(all_rows) - 1
            new_count += 1
            
    if update_count > 0 or new_count > 0:
        print(f"更新: {update_count} 件, 新規: {new_count} 件")
        worksheet.update("A1", all_rows)
        print("スプレッドシートの内容を最新化しました。")
    else:
        print("新しいメルマガはありませんでした。")

if __name__ == "__main__":
    main()
