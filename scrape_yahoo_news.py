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
new_ws = sh_output.add_worksheet(title=DATE_STR, rows="1000", cols="30") # colsの数を適宜調整
header = ['No.', 'タイトル', 'URL', '発行日時', '本文']
comment_cols = ['コメント数', 'コメント']
header_row = header + [''] * 9 + comment_cols # 10列目(J列)まで空欄、11列目(K列)にコメント数、12列目(L列)以降にコメント

# 空のヘッダーを作成
full_header = [''] * 15 # O列まで空欄
full_header[0:5] = ['No.', 'タイトル', 'URL', '発行日時', '本文']
full_header[14] = 'コメント数'
full_header[15:] = ['コメント'] * (len(full_header) - 15)

new_ws.update('A1:P1', [full_header])
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
        while True:
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
            
            if not body_text or body_text in article_bodies:
                break
            
            article_bodies.append(body_text)
            page += 1
            if page > 10:
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
            time.sleep(2)
            
            soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
            comment_elements = soup_comments.find_all('p', class_='sc-169yn8p-10')
            page_comments = [p.get_text(strip=True) for p in comment_elements]
            
            if not page_comments or page_comments[0] in comments:
                break
            
            comments.extend(page_comments)
            comment_page += 1
            if comment_page > 10:
                break

        print(f"    - Found {len(comments)} comments.")

        # データを複数行にまとめる
        # 1行目のデータ
        row_data = [idx, title, base_url, article_date, article_bodies[0]]
        row_data.extend([''] * 9) # O列まで空欄
        row_data.append(len(comments))
        row_data.extend(comments) # コメントはP列以降に記載

        all_data_to_write.append(row_data)
        
        # 2ページ目以降の本文を次の行に追加
        for i in range(1, len(article_bodies)):
            all_data_to_write.append([''] * 4 + [article_bodies[i]] + [''] * 20) # 適切に空欄を追加

        print(f"  - Successfully processed data for URL {idx}. Storing for batch update.")

    except Exception as e:
        print(f"  - Error processing URL {idx}: {e}")
        print("  - An error occurred. Continuing to the next URL.")

# 全URLの処理が完了したら、一括でシートに書き込み
if all_data_to_write:
    start_row = 2 # 2行目から開始
    start_cell = f'A{start_row}'
    new_ws.update(range_name=start_cell, values=all_data_to_write)
    print(f"--- All processed data has been written to the sheet, starting from {start_cell} ---")
else:
    print("No data to write. The sheet will remain empty except for the header.")

browser.quit()
print("--- Scraping job finished ---")
