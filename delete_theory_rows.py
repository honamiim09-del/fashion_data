import os
import gspread

def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "1kj2nX1v6LU9SUzxUB2Gc8A4ZkWWmQ2fnfojfIc6nmyY")
    # ローカルの認証ファイルを使用
    client_secret_path = "/Users/honami/Downloads/client_secret_54835506179-o2nulvaqu9qoggb1cba1tnbftcnusjvl.apps.googleusercontent.com.json"
    
    print("スプレッドシートに接続中...")
    gc = gspread.oauth(
        credentials_filename=client_secret_path,
        authorized_user_filename="authorized_user.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet("Newsletters")
    
    rows = ws.get_all_values()
    print(f"現在の全件数: {len(rows)} 件")
    
    deleted_count = 0
    # 2行目から逆順にチェックして、Theory関連の行を削除（インデックスがずれないように逆順）
    for i in range(len(rows), 1, -1):
        row = rows[i-1]
        sender = row[1]
        subject = row[2]
        # 送信者または件名に Theory が含まれる場合
        if "Theory" in sender or "theory" in sender or "Theory" in subject:
            ws.delete_rows(i)
            deleted_count += 1
            print(f"{i}行目を削除しました: {subject[:20]}...")

    print(f"削除完了！合計 {deleted_count} 件の Theory データを削除しました。")

if __name__ == "__main__":
    main()
