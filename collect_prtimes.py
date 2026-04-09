"""
PR TIMES RSS から特定企業のニュースを取得し、Google スプレッドシートに保存するスクリプト

必要な環境変数:
  GOOGLE_CREDENTIALS_JSON  : サービスアカウントの認証情報 JSON 文字列
  SPREADSHEET_ID           : 保存先スプレッドシートの ID
  SHEET_NAME (任意)        : シート名（デフォルト: "PR_TIMES"）
"""

import json
import os
import sys
from datetime import datetime, timezone

import feedparser
import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ────────────────────────────────────────────────────────────────────

KEYWORDS = [
    "ベイクルーズ",
    "トゥモローランド",
    "株式会社ジュン",
    "JUN GROUP",
]

# PR TIMES 全体RSS
PRTIMES_RSS_URL = "https://prtimes.jp/index.rdf"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

# スプレッドシートのヘッダー定義
HEADERS = ["取得日時", "企業名（キーワード）", "タイトル", "URL", "発表日時", "概要"]

# ── 認証 ────────────────────────────────────────────────────────────────────


def get_spreadsheet(spreadsheet_id: str, sheet_name: str) -> gspread.Worksheet:
    """環境変数から認証情報を読み込み、ワークシートを返す。"""
    credentials_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not credentials_json:
        raise EnvironmentError(
            "環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。"
        )

    credentials_dict = json.loads(credentials_json)
    creds = Credentials.from_service_account_info(credentials_dict, scopes=SCOPES)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(spreadsheet_id)

    # シートが存在しなければ作成してヘッダーを挿入
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
        worksheet.append_row(HEADERS)
        print(f"シート '{sheet_name}' を新規作成し、ヘッダーを追加しました。")

    return worksheet


# ── RSS 取得 ─────────────────────────────────────────────────────────────────


def fetch_all_articles(url: str) -> list[dict]:
    """PR TIMESの全体RSSを取得し、パースされた記事一覧を返す。"""
    feed = feedparser.parse(url)
    if feed.bozo:
        print(f"  [警告] RSS の解析に問題がありました: {feed.bozo_exception}")

    articles = []
    for entry in feed.entries:
        published = ""
        if hasattr(entry, "published_parsed") and entry.published_parsed:
            published = datetime(*entry.published_parsed[:6], tzinfo=timezone.utc).strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        summary = getattr(entry, "summary", "")
        # HTML タグを簡易除去
        import re
        summary = re.sub(r"<[^>]+>", "", summary).strip()

        articles.append(
            {
                "title": getattr(entry, "title", ""),
                "url": getattr(entry, "link", ""),
                "published": published,
                "summary": summary,
            }
        )

    return articles


# ── 重複チェック ──────────────────────────────────────────────────────────────


def get_existing_urls(worksheet: gspread.Worksheet) -> set[str]:
    """シートに保存済みの URL を集合で返す（URL 列 = 4 列目）。"""
    all_values = worksheet.get_all_values()
    urls = set()
    for row in all_values[1:]:  # 1 行目はヘッダーなのでスキップ
        if len(row) >= 4 and row[3]:
            urls.add(row[3])
    return urls


# ── メイン ───────────────────────────────────────────────────────────────────


def main() -> None:
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        print("エラー: 環境変数 SPREADSHEET_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    sheet_name = os.environ.get("SHEET_NAME", "PR_TIMES")

    print("Google スプレッドシートに接続中...")
    worksheet = get_spreadsheet(spreadsheet_id, sheet_name)

    print("保存済み URL を取得中...")
    existing_urls = get_existing_urls(worksheet)
    print(f"  既存レコード数: {len(existing_urls)} 件")

    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    new_rows: list[list] = []

    print("\nPR TIMES全体RSS 取得中...")
    all_articles = fetch_all_articles(PRTIMES_RSS_URL)
    print(f"  全体の取得件数: {len(all_articles)} 件")

    for keyword in KEYWORDS:
        print(f"\nキーワード「{keyword}」を検索中...")
        
        matched_articles = []
        for article in all_articles:
            # 大文字・小文字を区別せずに検索
            search_text = (article["title"] + " " + article["summary"]).lower()
            if keyword.lower() in search_text:
                matched_articles.append(article)
                
        print(f"  キーワード一致件数: {len(matched_articles)} 件")

        skipped = 0
        for article in matched_articles:
            if article["url"] in existing_urls:
                skipped += 1
                continue

            new_rows.append(
                [
                    now,
                    keyword,
                    article["title"],
                    article["url"],
                    article["published"],
                    article["summary"],
                ]
            )
            existing_urls.add(article["url"])

        print(f"  スキップ（重複）: {skipped} 件 / 追加予定: {len(matched_articles) - skipped} 件")

    if new_rows:
        print(f"\n{len(new_rows)} 件をスプレッドシートに追記中...")
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        print("完了しました。")
    else:
        print("\n新規追加するデータはありませんでした。")


if __name__ == "__main__":
    main()
