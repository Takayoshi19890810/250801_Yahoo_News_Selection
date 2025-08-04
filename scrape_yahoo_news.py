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

new_ws = sh_output.add_worksheet(title=DATE_STR, rows="100", cols=len(input_urls) + 1)
print(f"Created new sheet: {new_ws.title}")

# ヘッダーを1列目(A列)に書き込む
headers = ['タイトル', '投稿日', 'URL', '本文', 'コメント']
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
        res = requests.get(base_url, headers=headers_req)
        soup = BeautifulSoup(res.text, 'html.parser')
        
        # タイトルの取得 (セレクタ修正)
        title_tag = soup.find('h1', class_='sc-1f7c32y-2')
        title = title_tag.get_text(strip=True) if title_tag else '取得不可'
        
        # 投稿日の取得
        date_tag = soup.find('time')
        article_date = date_tag.get_text(strip=True) if date_tag else '取得不可'

        # 本文の取得 (セレクタ修正)
        body_elements = soup.find_all('p', class_='sc-1f7c32y-14')
        article_body_text = '\n'.join([p.get_text(strip=True) for p in body_elements])
        
        print(f"    - Article Title: {title}")
        print(f"    - Article Date: {article_date}")
        print(f"    - Found {len(body_elements)} body paragraphs.")

        # コメント取得（Selenium使用）
        comments = []
        comment_url = f"{base_url}/comments"
        browser.get(comment_url)
        
        try:
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'p.sc-169yn8p-10'))
            )
            soup_comments = BeautifulSoup(browser.page_source, 'html.parser')
            comment_elements = soup_comments.find_all('p', class_='sc-169yn8p-10')
            comments = [p.get_text(strip=True) for p in comment_elements]
        except Exception:
            print("    - No comments found or timed out waiting for comments.")

        print(f"    - Found {len(comments)} comments.")

        # 出力シートに書き込み
        current_column_idx = idx + 1 # B列から開始
        
        # データをリストにまとめる
        data_to_write = [
            [title],
            [article_date],
            [base_url],
            [article_body_text],
        ]
        
        # コメントを別々の行に追加
        if comments:
            data_to_write.append(['----- Comments -----'])
            data_to_write.extend([[c] for c in comments])
        
        start_cell = f'{col_to_letter(current_column_idx)}1'
        new_ws.update(start_cell, data_to_write)
        
        print(f"  - Successfully wrote data for URL {idx} to column {col_to_letter(current_column_idx)}")

    except Exception as e:
        print(f"  - Error processing URL {idx}: {e}")
        print("  - An error occurred. Continuing to the next URL.")

browser.quit()
print("--- Scraping job finished ---")
