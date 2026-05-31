import requests
from bs4 import BeautifulSoup
import os
import json
import sys
from datetime import datetime, timezone
import time
import gspread
from google.oauth2.service_account import Credentials
import feedparser # RSSの読み取り用
import urllib.parse

# ── 設定 ────────────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HEADERS = ["取得日時", "メディア", "タイトル", "URL", "公開日"]

# 重点的に追いたいキーワード（ファッション・アパレルを組み合わせて精度向上）
KEYWORDS = [
    "ベイクルーズ ファッション", "ユナイテッドアローズ", "アダストリア アパレル", 
    "Theory ファッション", "セオリー アパレル",
    "ワールド アパレル", "マッシュホールディングス", "TSIホールディングス", 
    "ストライプインターナショナル", "トゥモローランド ファッション",
    "ファッション 決算", "アパレル 店舗 出店", "百貨店 ファッション ニュース"
]

# 除外キーワード（これらがタイトルに含まれる場合はスキップ）
EXCLUDE_KEYWORDS = [
    "ゲーム", "アニメ", "スポーツ", "ワールドカップ", "世界情勢", "事件", "事故", 
    "理論", "相対性理論", "学説", "政治", "芸能", "映画", "ドラマ"
]

# RSSソース
RSS_SOURCES = [
    {"name": "FASHIONSNAP", "url": "https://www.fashionsnap.com/feed/"},
    {"name": "流通ニュース", "url": "https://www.ryutsuu.biz/feed"},
    {"name": "繊研新聞", "url": "https://senken.co.jp/posts.rss"},
    {"name": "Fashion Business", "url": "https://www.fashion-business.co.jp/feed"}
]

# ── 認証 ────────────────────────────────────────────────────────────────────

def get_spreadsheet(spreadsheet_id, sheet_name):
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        client = gspread.authorize(creds)
    else:
        print("GOOGLE_CREDENTIALS_JSON が見つからないため、ブラウザ認証を開始します...")
        client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
        client = gspread.oauth(
            credentials_filename=client_secret_path,
            authorized_user_filename="authorized_user.json",
            scopes=SCOPES
        )
    return client.open_by_key(spreadsheet_id).worksheet(sheet_name)

# ── 収集ロジック ─────────────────────────────────────────────────────────────

def fetch_rss_news():
    news_list = []
    for src in RSS_SOURCES:
        print(f"RSS取得中: {src['name']}...")
        try:
            d = feedparser.parse(src["url"])
            for entry in d.entries[:20]:
                news_list.append({
                    "media": src["name"],
                    "title": entry.title,
                    "url": entry.link,
                    "date": entry.get("published", "")
                })
        except Exception as e:
            print(f"エラー ({src['name']}): {e}")
    return news_list

def fetch_google_news():
    """Googleニュースからキーワード検索で記事を取得"""
    news_list = []
    for kw in KEYWORDS:
        print(f"Googleニュース検索中: {kw}...")
        # 検索期間を「過去7日間（when:7d）」に限定して古い記事を除外
        encoded_kw = urllib.parse.quote(kw + " when:7d")
        url = f"https://news.google.com/rss/search?q={encoded_kw}&hl=ja&gl=JP&ceid=JP:ja"
        try:
            d = feedparser.parse(url)
            for entry in d.entries[:10]: # 各キーワード上位10件
                news_list.append({
                    "media": f"Googleニュース({kw})",
                    "title": entry.title,
                    "url": entry.link,
                    "date": entry.published
                })
            time.sleep(1) # 負荷軽減
        except Exception as e:
            print(f"エラー (Google News {kw}): {e}")
    return news_list

# ── メイン ───────────────────────────────────────────────────────────────────

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    sheet_name = "Industry_News"
    
    worksheet = get_spreadsheet(spreadsheet_id, sheet_name)
    
    # 既存のURLを取得して重複を避ける
    existing_urls = set(worksheet.col_values(4))
    
    # ニュース収集
    all_raw_news = fetch_rss_news() + fetch_google_news()
    
    new_rows = []
    now = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M")
    
    # ブランド名や重要語が含まれるものだけに絞り込む（精度向上）
    # ※Googleニュースはすでにキーワードで検索済みだが、RSSの方はここでフィルタリング
    filter_keywords = [
        "ベイクルーズ", "ユナイテッドアローズ", "アダストリア", "Theory", "セオリー",
        "ワールド", "マッシュ", "TSI", "ストライプ", "トゥモローランド", "決算", "店舗", "出店"
    ]

    for news in all_raw_news:
        if news["url"] in existing_urls:
            continue
            
        # タイトルに関連ワードが含まれているか確認
        is_relevant = any(k.lower() in news["title"].lower() for k in filter_keywords)
        
        # 除外キーワードが含まれているか確認
        is_excluded = any(ex.lower() in news["title"].lower() for ex in EXCLUDE_KEYWORDS)
        
        if (not is_relevant and "Googleニュース" not in news["media"]) or is_excluded:
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
        print(f"\n{len(new_rows)} 件の新しいニュースを追加します...")
        # 重複を排除して最新順になるように（実際にはappendなので下に追加）
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        print("完了しました。")
    else:
        print("\n新しいニュースはありませんでした。")

if __name__ == "__main__":
    main()
