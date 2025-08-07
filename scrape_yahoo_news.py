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

# 新しいシートを作成し、ヘッダーを1行目に設定
new_ws = sh_output.add_worksheet(title=DATE_STR, rows="1000", cols="30")
header = ['No.', 'タイトル', 'URL', '発行日時']
body_headers = [f'本文({i}ページ)' for i in range(1, 11)] # 本文ヘッダーを10列分作成
comment_headers = ['コメント数', 'コメント']

# A1からQ1までを定義
full_header = header + body_headers + comment_headers
# Q列以降のコメントヘッダーは必要に応じて追加。ここでは動的に扱います。
new_ws.update('A1', [full_header])
print(f"Created new sheet: {new_ws.title}")

# ニュース記事の処理
print("--- Starting URL processing ---")
if not input_urls:
    print("No URLs to process. Exiting.")
    browser.quit()
    exit()

# 全データを一時的に保持するリスト
all_data_to_write = []

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
        while page <= 10: # 最大10ページまで取得
            url = base_url if page == 1 else f"{base_url}?page={page}"
            res = requests.get(url, headers=headers_req)
            soup = BeautifulSoup(res.text, 'html.parser')

            if page == 1:
                title_tag = soup.find('title')
                title = title_tag.get_text(strip=True).replace(' - Yahoo!ニュース', '') if title_tag else '取得不可'
                date_tag = soup.find('time')
                article_date = date_tag.get_text(strip=True) if date_tag else '取得不可'
            
            article_body_container = soup.find('article')
            if article_body_container:
                body_elements = article_body_container.find_all('p')
                body_text = '\n'.join([p.get_text(strip=True) for p in body_elements])
            else:
                body_text = ''
            
            if not body_text or (len(article_bodies) > 0 and body_text == article_bodies[-1]):
                break
            
            article_bodies.append(body_text)
            page += 1
                
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
            time.sleep(2)
            
            soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
            comment_elements = soup_comments.find_all('p', class_='sc-169yn8p-10')
            page_comments = [p.get_text(strip=True) for p in comment_elements]
            
            if not page_comments or (len(comments) > 0 and page_comments[0] == comments[-1]):
                break
            
            comments.extend(page_comments)
            comment_page += 1
            if comment_page > 10:
                break

        print(f"    - Found {len(comments)} comments.")

        # データを1行にまとめる
        row_data = [idx, title, base_url, article_date]
        
        # 本文データを追加
        row_data.extend(article_bodies)
        # 10ページに満たない場合は空欄で埋める
        row_data.extend([''] * (10 - len(article_bodies)))
        
        # コメント数とコメントを追加
        row_data.append(len(comments))
        row_data.extend([''] * 1) # P列を空欄にする
        row_data.extend(comments) # Q列以降にコメントを配置

        all_data_to_write.append(row_data)
        
        print(f"  - Successfully processed data for URL {idx}. Storing for batch update.")

    except Exception as e:
        print(f"  - Error processing URL {idx}: {e}")
        print("  - An error occurred. Continuing to the next URL.")

# 全URLの処理が完了したら、一括でシートに書き込み
if all_data_to_write:
    # 複数行にまたがるデータ（コメントが多い場合など）に対応するため、一括更新を調整
    max_cols = 0
    for row in all_data_to_write:
        if len(row) > max_cols:
            max_cols = len(row)
    
    # 全ての行の列数を揃える
    padded_data = [row + [''] * (max_cols - len(row)) for row in all_data_to_write]
    
    start_row = 2 # 2行目から開始
    new_ws.update(f'A{start_row}', padded_data)
    print(f"--- All processed data has been written to the sheet, starting from A{start_row} ---")
else:
    print("No data to write. The sheet will remain empty except for the header.")

browser.quit()
print("--- Scraping job finished ---")
