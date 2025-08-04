import os
import time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
from bs4 import BeautifulSoup
import json

# ヘルパー関数: 列番号をアルファベットに変換
def col_to_letter(col_num):
    """Convert 1-based column number to a letter (e.g. 1 -> 'A', 27 -> 'AA')."""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string

# Google Sheets認証
try:
    with open('credentials.json', 'r') as f:
        credentials_info = json.load(f)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
    gc = gspread.authorize(credentials)
except Exception as e:
    print(f"Error loading credentials: {e}")
    exit()

# Google Sheets設定
INPUT_SPREADSHEET_ID = '19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok'
OUTPUT_SPREADSHEET_ID = '1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE'
DATE_STR = datetime.now().strftime('%y%m%d')

# Selenium設定
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
browser = webdriver.Chrome(options=chrome_options)

# 入力スプレッドシートからURLを取得
print(f"--- Getting URLs from sheet '{DATE_STR}' ---")
sh_input = gc.open_by_key(INPUT_SPREADSHEET_ID)
try:
    input_ws = sh_input.worksheet(DATE_STR)
    input_urls = [url for url in input_ws.col_values(3)[1:] if url]
    print(f"Found {len(input_urls)} URLs to process.")
except gspread.WorksheetNotFound:
    print(f"Worksheet '{DATE_STR}' not found in input spreadsheet. Exiting.")
    browser.quit()
    exit()

# 出力スプレッドシートを設定
sh_output = gc.open_by_key(OUTPUT_SPREADSHEET_ID)
print(f"--- Checking output sheet for '{DATE_STR}' ---")

if DATE_STR in [ws.title for ws in sh_output.worksheets()]:
    date_ws = sh_output.worksheet(DATE_STR)
    sh_output.del_worksheet(date_ws)
    print(f"Existing sheet '{DATE_STR}' deleted.")

new_ws = sh_output.add_worksheet(title=DATE_STR, rows="1000", cols=len(input_urls) + 1)
print(f"Created new sheet: {new_ws.title}")

# ヘッダーを1列目(A列)に書き込む
headers = ['No.', 'タイトル', 'URL', '発行日時', '本文', '', '', '', '', '', '', 'コメント数', 'コメント']
new_ws.update('A1:A13', [[h] for h in headers])

# ニュース記事の処理
print("--- Starting URL processing ---")
if not input_urls:
    print("No URLs to process. Exiting.")
    browser.quit()
    exit()

for idx, base_url in enumerate(input_urls, start=1):
    try:
        print(f"  - Processing URL {idx}/{len(input_urls)}: {base_url}")
        
        headers_req = {'User-Agent': 'Mozilla/5.0'}
        
        # 記事本文の取得（複数ページ対応）
        article_bodies = []
        page = 1
        print("    - Processing article body...")
        title = '取得不可'
        article_date = '取得不可'
        while True:
            url = base_url if page == 1 else f"{base_url}?page={page}"
            res = requests.get(url, headers=headers_req)
            soup = BeautifulSoup(res.text, 'html.parser')

            # 記事タイトルと投稿日は1ページ目のみから取得
            if page == 1:
                title_tag = soup.find('h1', class_='sc-1f7c32y-2')
                title = title_tag.get_text(strip=True) if title_tag else '取得不可'
                date_tag = soup.find('time')
                article_date = date_tag.get_text(strip=True) if date_tag else '取得不可'
            
            # 本文部分を取得
            body_elements = soup.find_all('p', class_='sc-1f7c32y-14')
            body_text = '\n'.join([p.get_text(strip=True) for p in body_elements])
            
            # 本文が見つからない場合、またはページが重複している場合は終了
            if not body_text or body_text in article_bodies:
                break
            
            article_bodies.append(body_text)
            page += 1
            if page > 10:  # 無限ループ防止のための上限設定
                break
                
        print(f"    - Article Title: {title}")
        print(f"    - Article Date: {article_date}")
        print(f"    - Found {len(article_bodies)} body pages.")

        # コメント取得（複数ページ対応）
        comments = []
        comment_page = 1
        print("    - Scraping comments with Selenium...")
        while True:
            comment_url = f"{base_url}/comments?page={comment_page}"
            browser.get(comment_url)
            time.sleep(2) # 読み込みを待つ
            
            soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
            comment_elements = soup_comments.find_all('p', class_='sc-169yn8p-10')
            page_comments = [p.get_text(strip=True) for p in comment_elements]
            
            # ページにコメントがないか、すでに取得済みのコメントページの場合は終了
            if not page_comments or page_comments[0] in comments:
                break
            
            comments.extend(page_comments)
            comment_page += 1
            if comment_page > 10: # 無限ループ防止のための上限設定
                break

        print(f"    - Found {len(comments)} comments.")

        # 出力シートに書き込み
        current_column_idx = idx + 1 # B列から開始
        
        # データをリストにまとめる
        data_to_write = [
            [idx], # 1行目
            [title], # 2行目
            [base_url], # 3行目
            [article_date] # 4行目
        ]
        
        # 本文を5行目以降に追加
        body_start_row = 5
        for body in article_bodies:
            data_to_write.append([body])

        # コメント数の行まで空行で埋める
        comment_count_row = 16
        current_row_count = len(data_to_write)
        if current_row_count < comment_count_row - 1:
            data_to_write.extend([['']] * (comment_count_row - 1 - current_row_count))
        
        # コメント数を16行目に追加
        data_to_write.append([len(comments)])

        # コメントを17行目以降に追加
        comment_start_row = 17
        for comment in comments:
            data_to_write.append([comment])
        
        start_cell = f'{col_to_letter(current_column_idx)}1'
        new_ws.update(start_cell, data_to_write, value_input_option='USER_ENTERED')
        
        print(f"  - Successfully wrote data for URL {idx} to column {col_to_letter(current_column_idx)}")

    except Exception as e:
        print(f"  - Error processing URL {idx}: {e}")
        print("  - An error occurred. Continuing to the next URL.")

browser.quit()
print("--- Scraping job finished ---")
