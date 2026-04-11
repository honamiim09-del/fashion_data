import requests
from bs4 import BeautifulSoup
import os
import json
import sys
from datetime import datetime, timezone
import time

import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ────────────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HEADERS = ["取得日時", "メディア", "タイトル", "URL", "公開日"]

SOURCES = [
    {
        "name": "WWDJAPAN",
        "url": "https://www.wwdjapan.com/category/fashion/news-fashion",
        "container": "article.entry-item",
        "title_selector": "h2.entry-title a",
        "date_selector": "div.time",
        "base_url": "https://www.wwdjapan.com"
    },
    {
        "name": "繊研新聞",
        "url": "https://senken.co.jp/tags/breaking-news",
        "container": "div.post-large, div.post-small",
        "title_selector": "h2 a, h3 a",
        "date_selector": "p.m-b-h, small",
        "base_url": "https://senken.co.jp"
    }
]

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

# ── スクレイピング ───────────────────────────────────────────────────────────

def scrape_news():
    all_news = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36"
    }

    for source in SOURCES:
        print(f"{source['name']} を取得中...")
        try:
            response = requests.get(source["url"], headers=headers, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            
            items = soup.select(source["container"])
            print(f"  {len(items)} 件の要素が見つかりました。")

            for item in items[:15]: # 各サイト上位15件程度
                title_elem = item.select_one(source["title_selector"])
                if not title_elem: continue
                
                title = title_elem.get_text(strip=True)
                link = title_elem.get("href")
                if not link.startswith("http"):
                    link = source["base_url"] + link
                
                date_elem = item.select_one(source["date_selector"])
                date_str = date_elem.get_text(strip=True) if date_elem else ""

                all_news.append({
                    "media": source["name"],
                    "title": title,
                    "url": link,
                    "date": date_str
                })
            
            # サーバー負荷軽減のため少し待つ
            time.sleep(2)
            
        except Exception as e:
            print(f"エラー ({source['name']}): {e}")

    return all_news

# ── メイン ───────────────────────────────────────────────────────────────────

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id: sys.exit(1)
        
    sheet_name = "Industry_News"
    worksheet = get_spreadsheet(spreadsheet_id, sheet_name)
    
    print("既存のURLを取得中...")
    all_values = worksheet.get_all_values()
    existing_urls = set()
    for row in all_values[1:]:
        if len(row) >= 4:
            existing_urls.add(row[3])
            
    print("最新ニュースをスクレイピング中...")
    news_list = scrape_news()
    
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []
    
    for news in news_list:
        if news["url"] in existing_urls:
            continue
            
        new_rows.append([
            now,
            news["media"],
            news["title"],
            news["url"],
            news["date"]
        ])
        existing_urls.add(news["url"])
        
    if new_rows:
        print(f"\n{len(new_rows)} 件の新しい業界ニュースを追記します...")
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        print("完了しました。")
    else:
        print("\n新しいニュースはありませんでした。")

if __name__ == "__main__":
    main()
