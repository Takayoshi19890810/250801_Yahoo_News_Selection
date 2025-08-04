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
    # GitHub Actionsのシークレットからcredentials.jsonを読み込む
    with open('credentials.json', 'r') as f:
        credentials_info = json.load(f)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
    gc = gspread.authorize(credentials)
except Exception as e:
    print(f"Error loading credentials: {e}")
    exit()

# Google Sheets設定
# INPUTのスプレッドシートID (https://docs.google.com/spreadsheets/d/19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok)
INPUT_SPREADSHEET_ID = '19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok'
# OUTPUTのスプレッドシートID (https://docs.google.com/spreadsheets/d/1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE)
OUTPUT_SPREADSHEET_ID = '1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE'
DATE_STR = datetime.now().strftime('%y%m%d')

# Selenium設定
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
browser = webdriver.Chrome(options=chrome_options)

# ヘルパー関数: 列番号をアルファベットに変換
def col_to_letter(col_num):
    """Convert 1-based column number to a letter (e.g. 1 -> 'A', 27 -> 'AA')."""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string

# 入力スプレッドシートからURLを取得
print(f"--- Getting URLs from sheet '{DATE_STR}' ---")
sh_input = gc.open_by_key(INPUT_SPREADSHEET_ID)
try:
    input_ws = sh_input.worksheet(DATE_STR)
    # C列の2行目から読み込む
    input_urls = [url for url in input_ws.col_values(3)[1:] if url]
    print(f"Found {len(input_urls)} URLs to process.")
except gspread.WorksheetNotFound:
    print(f"Worksheet '{DATE_STR}' not found. Exiting.")
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

# ニュース記事の処理
print("--- Starting URL processing ---")
if not input_urls:
    print("No URLs to process. Exiting.")
    browser.quit()
    exit()

# ヘッダーを1列目(A列)に書き込む
headers = ['タイトル', '投稿日', 'URL', '本文1', '本文2', '本文3', '本文4', '本文5', '本文6', '本文7', '本文8', '本文9', '本文10', '本文11', '本文12', '本文13', '本文14', '本文15', '本文16', '本文17', 'コメント']
new_ws.update('A1', [[h] for h in headers])

for idx, base_url in enumerate(input_urls, start=1):
    try:
        print(f"  - Processing URL {idx}/{len(input_urls)}: {base_url}")
        
        headers_req = {'User-Agent': 'Mozilla/5.0'}
        
        # 記事本文、タイトル、投稿日の取得
        article_bodies = []
        page = 1
        print(f"    - Processing article body...")
        while True:
            url = base_url if page == 1 else f"{base_url}?page={page}"
            res = requests.get(url, headers=headers_req)
            soup = BeautifulSoup(res.text, 'html.parser')
            if '指定されたURLは存在しませんでした' in res.text:
                break
            body_tag = soup.find('article')
            body_text = body_tag.get_text(separator='\n').strip() if body_tag else ''
            if not body_text or body_text in article_bodies:
                break
            article_bodies.append(body_text)
            page += 1
        print(f"    - Found {len(article_bodies)} body pages.")

        res_main = requests.get(base_url, headers=headers_req)
        soup_main = BeautifulSoup(res_main.text, 'html.parser')
        page_title = soup_main.title.string if soup_main.title else '取得不可'
        title = page_title.replace(' - Yahoo!ニュース', '').strip() if page_title else '取得不可'
        date_tag = soup_main.find('time')
        article_date = date_tag.text.strip() if date_tag else '取得不可'
        print(f"    - Article Title: {title}")
        print(f"    - Article Date: {article_date}")

        # コメント取得（Selenium使用）
        comments = []
        comment_page = 1
        print("    - Scraping comments with Selenium...")
        while True:
            comment_url = f"{base_url}/comments?page={comment_page}"
            browser.get(comment_url)
            time.sleep(2)
            if '指定されたURLは存在しませんでした' in browser.page_source:
                break
            soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
            comment_paragraphs = soup_comments.find_all('p', class_='sc-169yn8p-10 hYFULX')
            page_comments = [p.get_text(strip=True) for p in comment_paragraphs if p.get_text(strip=True)]
            if not page_comments:
                break
            comments.extend(page_comments)
            comment_page += 1
        print(f"    - Found {len(comments)} comments.")

        # 出力シートに書き込み
        current_column_idx = idx + 1 # B列から開始
        print(f"    - Writing data to column {col_to_letter(current_column_idx)}...")
        
        data_to_write = [
            [title],
            [article_date],
            [base_url],
        ]
        
        # 本文を追加
        for body in article_bodies:
            data_to_write.append([body])

        # コメントを20行目以降に追加
        empty_rows_count = 20 - len(data_to_write)
        if empty_rows_count > 0:
            data_to_write.extend([['']] * empty_rows_count)

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
