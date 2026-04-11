import os
import json
import sys
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ────────────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MASTER_HEADERS = ["取得日時", "出所", "タイトル", "URL", "発信日", "抜粋", "対象ブランド/メディア"]
MASTER_SHEET_NAME = "Master_Timeline"

# ── 認証 ────────────────────────────────────────────────────────────────────

def get_client():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise EnvironmentError("GOOGLE_CREDENTIALS_JSON が設定されていません。")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        print("SPREADSHEET_ID が設定されていません。")
        sys.exit(1)

    client = get_client()
    spreadsheet = client.open_by_key(spreadsheet_id)

    print("各シートからデータを取得中...")
    
    all_data = []

    # 1. PR TIMES
    try:
        ws_pr = spreadsheet.worksheet("PR_TIMES")
        rows = ws_pr.get_all_values()
        if len(rows) > 1:
            for r in rows[1:]:
                # [取得日時, 企業名, タイトル, URL, 発表日時, 概要]
                if len(r) >= 5:
                    all_data.append([r[0], "PR TIMES", r[2], r[3], r[4], r[5] if len(r) > 5 else "", r[1]])
    except Exception as e:
        print(f"PR_TIMES 読み込みエラー: {e}")

    # 2. Newsletters
    try:
        ws_news = spreadsheet.worksheet("Newsletters")
        rows = ws_news.get_all_values()
        if len(rows) > 1:
            for r in rows[1:]:
                # [取得日時, 送信者, 件名, 本文, 送信日時, Message-ID]
                if len(r) >= 5:
                    all_data.append([r[0], "NEWSLETTER", r[2], "", r[4], r[3], r[1]])
    except Exception as e:
        print(f"Newsletters 読み込みエラー: {e}")

    # 3. Industry_News
    try:
        ws_ind = spreadsheet.worksheet("Industry_News")
        rows = ws_ind.get_all_values()
        if len(rows) > 1:
            for r in rows[1:]:
                # [取得日時, メディア, タイトル, URL, 公開日]
                if len(r) >= 5:
                    all_data.append([r[0], "INDUSTRY NEWS", r[2], r[3], r[4], "", r[1]])
    except Exception as e:
        print(f"Industry_News 読み込みエラー: {e}")

    if not all_data:
        print("統合するデータが見つかりませんでした。")
        return

    # 取得日時の降順（最新順）にソート
    # 取得日時が r[0] に入っている想定。文字列比較でも YYYY-MM-DD 形式なら概ね正しくソートされる
    all_data.sort(key=lambda x: x[0], reverse=True)

    # 上位500件に制限（スプレッドシートの肥大化・Looker Studioの重さを防ぐ）
    all_data = all_data[:500]

    # マスターシートの準備
    try:
        worksheet = spreadsheet.worksheet(MASTER_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=MASTER_SHEET_NAME, rows=1000, cols=10)
        print(f"シート '{MASTER_SHEET_NAME}' を新規作成しました。")

    # 書き込み
    print(f"マスターシート {MASTER_SHEET_NAME} に {len(all_data)} 件を書き込み中...")
    worksheet.clear()
    worksheet.update("A1", [MASTER_HEADERS] + all_data)
    print("統合完了！")

if __name__ == "__main__":
    main()
