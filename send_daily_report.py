import os
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
import re
import html

def linkify(text):
    """テキスト内のURLを<a>タグに変換する"""
    url_pattern = r'(https?://[\w/:%#\$&\?\(\)~\.=\+\-]+)'
    return re.sub(url_pattern, r'<a href="\1">\1</a>', text)

def clean_body(text):
    """メルマガの定型ヘッダーなどを除外して本文を抽出する"""
    if not text:
        return ""
    lines = text.split('\n')
    cleaned_lines = []
    
    # 無視したいキーワードやパターン
    ignore_keywords = ["━━━━━━━━━━━", "───────────", "−−−−−−−−−−−", "BAYCREW'S STORE", "ベイクルーズストア", "▼"]
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 定型文を含まない行だけを抽出
        if not any(k in line for k in ignore_keywords):
            cleaned_lines.append(line)
            
    # 結合して返す
    full_text = "\n".join(cleaned_lines)
    
    # 200文字程度で切り取るが、URLの途中なら最後まで含める
    limit = 200
    if len(full_text) > limit:
        # 制限位置の後にスペースや改行、URLの終わりが来るまで伸ばす
        extended = full_text[limit:]
        match = re.search(r'[\s\n]', extended)
        if match:
            return full_text[:limit + match.start()]
        return full_text
    return full_text

def send_daily_report():
    # ── 設定 ────────────────────────────────────────────────────────────────────
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    email_sender = os.environ.get("EMAIL_SENDER")
    email_password = os.environ.get("EMAIL_PASSWORD")
    if email_password:
        # スペースや特殊な空白文字を除去
        email_password = email_password.replace(" ", "").replace("\xa0", "")
    email_receiver = os.environ.get("EMAIL_RECEIVER", email_sender) # デフォルトは自分宛
    
    if not email_sender or not email_password:
        print("エラー: EMAIL_SENDER または EMAIL_PASSWORD が設定されていません。")
        return

    # 過去24時間のデータを抽出 (UTCベース)
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    threshold_time = now_utc - datetime.timedelta(hours=24)
    target_label = "過去24時間"
    print(f"{target_label} のデータを抽出中...")

    # ── スプレッドシートからデータ取得 ──────────────────────────────────────────
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    json_creds = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if json_creds:
        import json
        creds_dict = json.loads(json_creds)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
    else:
        # ローカル実行用（ブラウザ認証）
        print("GOOGLE_CREDENTIALS_JSON が見つからないため、ブラウザ認証を開始します...")
        client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
        gc = gspread.oauth(
            credentials_filename=client_secret_path,
            authorized_user_filename="authorized_user.json",
            scopes=scopes
        )
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet("Master_Timeline")
    
    all_rows = worksheet.get_all_values()
    if len(all_rows) <= 1:
        print("データがありません。")
        return

    header = all_rows[0]
    data_rows = all_rows[1:]

    # 列番号の特定（ヘッダー名から判断）
    try:
        col_date = header.index("取得日時")
        col_source = header.index("出所")
        col_company = header.index("対象ブランド/メディア")
        col_title = header.index("タイトル")
        col_url = header.index("URL")
        col_summary = header.index("AI要約")
        col_body = header.index("本文")
    except ValueError as e:
        print(f"エラー: 必要な列が見つかりません: {e}")
        return

    # 過去24時間のデータを会社ごとにグループ化
    report_data = defaultdict(list)
    count = 0
    
    for row in data_rows:
        if len(row) <= max(col_date, col_source, col_company, col_title, col_url, col_summary, col_body):
            continue
            
        row_date_str = row[col_date]
        try:
            # 取得日時は YYYY-MM-DD HH:MM:SS の形式 (UTC) を想定
            row_dt = datetime.datetime.strptime(row_date_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=datetime.timezone.utc)
            is_target = row_dt >= threshold_time
        except ValueError:
            # 時間がパースできない場合(古いデータなど)のフォールバック
            yesterday_str = (now_utc - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
            today_str = now_utc.strftime("%Y-%m-%d")
            is_target = row_date_str.startswith(yesterday_str) or row_date_str.startswith(today_str)

        if is_target:
            company = row[col_company] or "その他"
            report_data[company].append({
                "source": row[col_source],
                "title": row[col_title],
                "url": row[col_url],
                "summary": row[col_summary],
                "body": row[col_body]
            })
            count += 1

    if count == 0:
        print(f"{target_label} の新着記事はありませんでした。")
        return

    # ── メール本文の作成 ────────────────────────────────────────────────────────
    body = f"<h2>【{now_utc.strftime('%Y-%m-%d')}】業界動向デイリーレポート</h2>"
    body += f"<p>合計 {count} 件の新着記事が見つかりました。</p><hr>"

    for company, articles in report_data.items():
        body += f"<h3 style='background-color: #f0f0f0; padding: 5px;'>■ {company}</h3>"
        for art in articles:
            body += f"<div style='margin-bottom: 25px; border-left: 4px solid #ddd; padding-left: 10px;'>"
            body += f"<strong>【{art['source']}】{art['title']}</strong><br>"
            
            # AI要約があれば表示
            if art['summary']:
                body += f"<div style='margin: 5px 0; color: #333; background: #fffde7; padding: 5px;'>要約: {art['summary']}</div>"
            
            # メルマガ（または要約がない場合）は本文の一部を表示
            if art['source'] == "NEWSLETTER" and art['body']:
                display_body = clean_body(art['body'])
                # HTMLエスケープしてから改行を<br>に変換し、URLをリンク化
                safe_body = html.escape(display_body).replace('\n', '<br>')
                snippet = linkify(safe_body) + ("..." if len(art['body']) > len(display_body) else "")
                body += f"<div style='margin: 5px 0; color: #666; font-size: 0.85em; font-style: italic;'>本文抜粋: {snippet}</div>"
            
            if art['url']:
                body += f"<a href='{art['url']}' style='font-size: 0.8em;'>{art['url']}</a>"
            body += f"</div>"
        body += "<br>"

    # ── メールの送信 ──────────────────────────────────────────────────────────
    # 複数宛先（カンマ区切り）に対応し、お互いのアドレスが見えないように1件ずつ個別に送信する
    receivers = [r.strip() for r in email_receiver.split(',') if r.strip()]
    
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(email_sender, email_password)
            
            for receiver in receivers:
                msg = MIMEMultipart()
                msg['From'] = email_sender
                msg['To'] = receiver
                msg['Subject'] = f"【Fashion Data】デイリーレポート ({now_utc.strftime('%Y-%m-%d')})"
                msg.attach(MIMEText(body, 'html'))
                
                server.send_message(msg)
                print(f"メール送信完了: {count}件の記事を {receiver} へ送りました。")
    except Exception as e:
        print(f"メール送信エラー: {e}")

if __name__ == "__main__":
    send_daily_report()
