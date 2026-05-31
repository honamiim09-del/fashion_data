import os
import time
import gspread
from collect_data import fetch_article_body, summarize_text

def backfill():
    print("認証中...")
    try:
        client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
        gc = gspread.oauth(
            credentials_filename=client_secret_path,
            authorized_user_filename="authorized_user.json",
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
    except Exception as e:
        print(f"認証エラー: {e}")
        return

    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    sheet_name = "PR_TIMES"

    print(f"スプレッドシートを開いています...")
    try:
        sh = gc.open_by_key(spreadsheet_id)
        worksheet = sh.worksheet(sheet_name)
    except Exception as e:
        print(f"シートの取得に失敗しました: {e}")
        return
    
    rows = worksheet.get_all_values()
    if len(rows) <= 1:
        print("処理するデータがありません。")
        return

    header = rows[0]
    
    # H列（8列目）に「AI要約」列を確保
    if len(header) < 8:
        print("ヘッダーに 'AI要約' を追加します。")
        worksheet.update_cell(1, 8, "AI要約")
    
    for i, row in enumerate(rows[1:], start=2):
        url = row[3] if len(row) > 3 else ""
        body_text = row[6] if len(row) > 6 else ""
        ai_summary = row[7] if len(row) > 7 else ""
        
        # 本文はあるが、AI要約がない場合のみ実行
        if body_text and not ai_summary:
            print(f"[{i-1}/{len(rows)-1}] 要約を生成中: {url[:50]}...")
            summary = summarize_text(body_text)
            
            if summary:
                try:
                    worksheet.update_cell(i, 8, summary)
                    print(f"  -> 保存完了")
                    # API制限（1分間に5回まで）を避けるため、15秒待機
                    time.sleep(15)
                except Exception as e:
                    print(f"  -> 保存エラー: {e}")

    print("\nすべての過去データの要約が完了しました！")

if __name__ == "__main__":
    backfill()
