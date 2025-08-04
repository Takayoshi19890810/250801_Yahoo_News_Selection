import os
import time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
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
# INPUTのスプレッドシートID: https://docs.google.com/spreadsheets/d/19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok/
INPUT_SPREADSHEET_ID = '19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok'
# OUTPUTのスプレッドシートID: https://docs.google.com/spreadsheets/d/1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE/
OUTPUT_SPREADSHEET_ID = '1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE'
DATE_STR = datetime.now().strftime('%y%m%d')
BASE_SHEET = 'Base'

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
    # C列の2行目以降から読み込む
    input_urls = [url for url in input_ws.col_values(3)[1:] if url]
    print(f"Found {len(input_urls)} URLs to process.")
except gspread.WorksheetNotFound:
    print(f"Worksheet '{DATE_STR}' not found in input spreadsheet. Exiting.")
    browser.quit()
    exit()

# 出力スプレッドシートを設定
sh_output = gc.open_by_key(OUTPUT_SPREADSHEET_ID)
print(f"--- Checking output sheet for '{DATE_STR}' ---")

# 既存のシートがあれば削除し、新しいシートを作成
if DATE_STR in [ws.title for ws in sh_output.worksheets()]:
    date_ws = sh_output.worksheet(DATE_STR)
    sh_output.del_worksheet(date_ws)
    print(f"Existing sheet '{DATE_STR}' deleted.")

# 新しいシートを作成
new_ws = sh_output.add_worksheet(title=DATE_STR, rows="100", cols=len(input_urls) + 1)
print(f"Created new sheet: {new_ws.title}")

# ヘッダーを1列目(A列)に書き込む
headers = ['タイトル', '投稿日', 'URL', '本文1', '本文2', '本文3', '本文4', '本文5', '本文6', '本文7', '本文8', '本文9', '本文10', '本文11', '本文12', '本文13', '本文14', '本文15', '本文16', '本文17', 'コメント']
new_ws.update('A1', [[h] for h in headers])

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
        
        # 記事本文、タイトル、投稿日の取得
        article_bodies = []
        res = requests.get(base_url, headers=headers_req)
        soup = BeautifulSoup(res.text, 'html.parser')
        
        # タイトルの取得
        title = soup.find('h1').get_text(strip=True) if soup.find('h1') else '取得不可'
        
        # 投稿日の取得
        date_tag = soup.find('time')
        article_date = date_tag.get_text(strip=True) if date_tag else '取得不可'

        # 本文の取得
        body_elements = soup.find_all('p', class_='sc-1f7c32y-14 dYlVjE')
        article_bodies = [p.get_text(strip=True) for p in body_elements]

        print(f"    - Article Title: {title}")
        print(f"    - Article Date: {article_date}")
        print(f"    - Found {len(article_bodies)} body paragraphs.")

        # コメント取得（Selenium使用）
        comments = []
        comment_url = f"{base_url}/comments"
        browser.get(comment_url)
        time.sleep(3) # 読み込みを待つ
        soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
        comment_elements = soup_comments.find_all('p', class_='sc-fD-bZ kYJzEZ')
        comments = [p.get_text(strip=True) for p in comment_elements]
        print(f"    - Found {len(comments)} comments.")

        # 出力シートに書き込み
        current_column_idx = idx + 1 # B列から開始
        
        data_to_write = [
            [title],
            [article_date],
            [base_url],
        ]
        
        # 本文を追加
        for body in article_bodies:
            data_to_write.append([body])

        # コメントを20行目以降に追加
        # 現在のデータの行数に応じて空行を挿入
        current_rows = len(data_to_write)
        if current_rows < 20:
            data_to_write.extend([['']] * (20 - current_rows))

        # コメントを追加
        for comment in comments:
            data_to_write.append([comment])

        # データをまとめて書き込み
        start_cell = f'{col_to_letter(current_column_idx)}1'
        new_ws.update(start_cell, data_to_write)
        
        print(f"  - Successfully wrote data for URL {idx} to column {col_to_letter(current_column_idx)}")

    except Exception as e:
        print(f"  - Error writing to Google Sheets for URL {idx}: {e}")
        print("  - An error occurred. Continuing to the next URL.")

browser.quit()
print("--- Scraping job finished ---")
