import os
import requests
from bs4 import BeautifulSoup
import re
import google.generativeai as genai

def fetch_article_body(url: str) -> str:
    """
    PR TIMES の記事詳細ページから本文テキストを抽出する。
    
    Args:
        url (str): 記事の URL
        
    Returns:
        str: 抽出された本文テキスト（エラー時は空文字）
    """
    try:
        # ページを取得
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        # 文字コードを自動判定して設定
        response.encoding = response.apparent_encoding
        
        # HTML をパース
        soup = BeautifulSoup(response.text, "html.parser")
        
        # 本文が入っている div.maintext を探す
        main_text_div = soup.find("div", class_="maintext")
        
        if not main_text_div:
            # 万が一 maintext が見つからない場合は、代替案として通常の記事本文エリアを探す
            # PR TIMES の別パターンの可能性も考慮
            main_text_div = soup.find("div", class_="content-main") or soup.find("article")
            
        if main_text_div:
            # テキストを抽出（改行を維持するために separator を指定）
            text = main_text_div.get_text(separator="\n")
            
            # 余計な空白や改行を整える
            lines = []
            for line in text.splitlines():
                stripped = line.strip()
                if stripped:
                    lines.append(stripped)
            
            # 連続する改行を1つにまとめる
            cleaned_text = "\n".join(lines)
            return cleaned_text
            
        return ""
        
    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL {url}: {e}")
        return ""
    except Exception as e:
        print(f"An unexpected error occurred while processing {url}: {e}")
        return ""

def summarize_text(text: str) -> str:
    """
    Gemini API を使用してテキストを要約する。
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("警告: GEMINI_API_KEY が設定されていません。")
        return ""
    
    if not text or len(text) < 10:
        return ""

    try:
        genai.configure(api_key=api_key)
        # 最も制限が緩いはずの標準モデル名を使用
        model = genai.GenerativeModel("gemini-flash-latest")
        
        prompt = (
            "あなたは優秀なニュース編集者です。以下の記事本文を、要点がわかるように日本語で3行程度に要約してください。"
            "箇条書きではなく、自然な文章にしてください。"
            "\n\n"
            f"--- 本文 ---\n{text}"
        )
        
        # 429エラー時に最大5回までリトライ
        for attempt in range(5):
            try:
                response = model.generate_content(prompt)
                return response.text.strip()
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg:
                    wait_time = 60 + (attempt * 30) # 60秒, 90秒, 120秒...と待機時間を増やす
                    print(f"    API制限(429)のため、{wait_time}秒待機してリトライします... (試行 {attempt + 1}/5)")
                    import time
                    time.sleep(wait_time)
                    continue
                else:
                    raise e
        return ""
        
    except Exception as e:
        print(f"要約エラー: {e}")
        return ""

if __name__ == "__main__":
    # テスト用のコード
    test_url = "https://prtimes.jp/main/html/rd/p/000001000.000000000.html" # 適当な例
    # 実際には存在するURLでテストする必要があるが、ここでは関数の定義のみを行う。
    print("fetch_article_body function defined.")
