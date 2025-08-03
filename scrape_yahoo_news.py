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

# Google Sheets認証
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

# Google Sheets設定
INPUT_SPREADSHEET_ID = '1yjHpQMHfJt7shjqZ6SYQNNlHougbrw0ZCgWpFUgv3Sc'
OUTPUT_SPREADSHEET_ID = '1ff9j8Dr2G6UO2GjsLNpgC8bW0KJmX994iJruw4X_qVM'
DATE_STR = datetime.now().strftime('%y%m%d')
BASE_SHEET = 'Base'

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
print("--- Getting URLs from spreadsheet ---")
sh_input = gc.open_by_key(INPUT_SPREADSHEET_ID)
input_ws = sh_input.worksheet("URLS")
input_urls = [url for url in input_ws.col_values(3)[1:] if url]
print(f"Found {len(input_urls)} URLs to process.")

# 出力スプレッドシートを設定
sh_output = gc.open_by_key(OUTPUT_SPREADSHEET_ID)
print("--- Checking output sheet ---")

if DATE_STR in [ws.title for ws in sh_output.worksheets()]:
    date_ws = sh_output.worksheet(DATE_STR)
    print(f"Using existing output sheet: {date_ws.title}")

    existing_urls = []
    try:
        all_urls_in_sheet = date_ws.row_values(3)
        existing_urls = [url for url in all_urls_in_sheet[1:] if url and url.startswith(('http://', 'https://'))]
    except gspread.exceptions.APIError as e:
        print(f"Could not retrieve existing URLs from sheet. Error: {e}")
    
    urls_to_add = [url for url in input_urls if url not in existing_urls]
    print(f"Found {len(existing_urls)} existing URLs. Adding {len(urls_to_add)} new URLs.")

else:
    sh_output.duplicate_sheet(sh_output.worksheet(BASE_SHEET).id, new_sheet_name=DATE_STR)
    date_ws = sh_output.worksheet(DATE_STR)
    print(f"Created new sheet: {date_ws.title}")
    urls_to_add = [url for url in input_urls if url]
    print(f"Found {len(urls_to_add)} new URLs to add.")

if not urls_to_add:
    print("No new URLs to add. Exiting.")
    browser.quit()
    exit()

# ニュース記事の処理
print("--- Starting URL processing for new articles ---")
output_column = len(existing_urls) + 2 if 'existing_urls' in locals() else 2

for idx, base_url in enumerate(urls_to_add, start=1):
    try:
        print(f"  - Processing URL: {base_url}")
        
        headers = {'User-Agent': 'Mozilla/5.0'}
        
        # 記事本文、タイトル、投稿日の取得
        article_bodies = []
        page = 1
        print(f"    - Processing article body...")
        while True:
            url = base_url if page == 1 else f"{base_url}?page={page}"
            res = requests.get(url, headers=headers)
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

        res_main = requests.get(base_url, headers=headers)
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
        current_column_idx = output_column + (idx - 1)
        print(f"    - Writing data to column {current_column_idx}...")
        
        # データをリストにまとめる
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
        date_ws.update(start_cell, data_to_write)
        
        print(f"  - Successfully wrote data for URL {idx} to column {current_column_idx}")

    except Exception as e:
        print(f"  - Error writing to Google Sheets for URL {idx}: {e}")
        print("  - An error occurred. Exiting the program to prevent further issues.")
        browser.quit()
        exit()

browser.quit()
print("--- Scraping job finished ---")
