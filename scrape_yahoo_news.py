import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import json

# GitHub Actionsのシークレットからサービスアカウントのキーを読み込む
# ローカルでテストする場合は、service_account.jsonファイルを作成してください
try:
    with open('service_account.json', 'r') as f:
        credentials_info = json.load(f)
except FileNotFoundError:
    print("service_account.json not found, trying to get from environment variable")
    import os
    gcp_sa_key = os.environ.get('GCP_SA_KEY')
    if not gcp_sa_key:
        raise ValueError("GCP_SA_KEY environment variable is not set")
    credentials_info = json.loads(gcp_sa_key)

# Google Sheets APIの認証
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)

# スプレッドシートのURLと日付のフォーマット
INPUT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/19c6yIGr5BiI7XwstYhUPptFGksPPXE4N1bEq5iFoPok/edit#gid=0'
OUTPUT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1n7gXdU2Z3ykL7ys1LXFVRiGHI0VEcHsRZeK-Gr1ECVE/edit#gid=0'
TODAY_DATE = datetime.now().strftime('%y%m%d')

def get_urls_from_sheet():
    """入力スプレッドシートからURLのリストを取得する"""
    try:
        input_sheet = gc.open_by_url(INPUT_SHEET_URL)
        worksheet = input_sheet.worksheet(TODAY_DATE)
        urls = worksheet.col_values(3)[1:]  # C列の2行目から読み込む
        return [url for url in urls if url]  # 空のURLを除外
    except gspread.WorksheetNotFound:
        print(f"Worksheet for {TODAY_DATE} not found in input sheet.")
        return []

def scrape_yahoo_news(url):
    """Yahooニュースの記事からタイトルと本文をスクレイピングする"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # タイトルと本文のセレクタは変更される可能性があるため、適宜調整してください
        title = soup.find('h1', class_='sc-fD-bZ kYJzEZ').text.strip() if soup.find('h1', class_='sc-fD-bZ kYJzEZ') else 'タイトルが見つかりません'
        body_elements = soup.find('p', class_='sc-eQWqj kHwWGY').text.strip() if soup.find('p', class_='sc-eQWqj kHwWGY') else '本文が見つかりません'
        
        return title, body_elements

    except requests.exceptions.RequestException as e:
        print(f"Error scraping {url}: {e}")
        return None, None
    except AttributeError:
        print(f"Error parsing HTML for {url}. Selectors may have changed.")
        return 'タイトルが見つかりません', '本文が見つかりません'

def write_to_output_sheet(data):
    """結果を日付ごとのシートに書き込む"""
    try:
        output_sheet = gc.open_by_url(OUTPUT_SHEET_URL)
        try:
            worksheet = output_sheet.worksheet(TODAY_DATE)
        except gspread.WorksheetNotFound:
            print(f"Worksheet for {TODAY_DATE} not found. Creating new sheet.")
            worksheet = output_sheet.add_worksheet(title=TODAY_DATE, rows="100", cols="10")
        
        # ヘッダーを書き込む
        worksheet.clear()
        worksheet.append_row(['URL', 'タイトル', '本文'])
        
        # データを書き込む
        worksheet.append_rows(data)
        print(f"Successfully wrote {len(data)} rows to sheet {TODAY_DATE}.")

    except Exception as e:
        print(f"Error writing to output sheet: {e}")

def main():
    urls = get_urls_from_sheet()
    if not urls:
        print("No URLs to scrape.")
        return
    
    scraped_data = []
    for url in urls:
        title, content = scrape_yahoo_news(url)
        if title and content:
            scraped_data.append([url, title, content])
    
    if scraped_data:
        write_to_output_sheet(scraped_data)
    else:
        print("No data was scraped.")

if __name__ == '__main__':
    main()
