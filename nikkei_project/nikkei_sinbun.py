from selenium.webdriver import Chrome, ChromeOptions
import pandas as pd
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import os
import time
import datetime
import sys
import logging
import json

# loggingの設定
logging.basicConfig(format='%(asctime)s -  %(levelname)s - %(message)s', level=logging.INFO)


# 環境変数読み込み
dotenv_path = os.path.join(os.getcwd(), '.env')
load_dotenv(dotenv_path)
NIKKEI_URL = 'https://r.nikkei.com/'
NIKKEY_KEY = 'togtai1992@gmail.com'
NIKKEY_PASSWORD = 'ghg78551'
splead_sheets_key = '1YPx82yEIvVl7WDNAhb_Y9Io3s5iF0kSF_aXEMgV99VE'
jsonf = 'Obtained-Nikkei-newspaper-0eaefcf15f0b.json'
today = datetime.datetime.now().strftime('%Y年%m月%d日')

### main処理
def main():
    """
    日経新聞の保存記事を取得する
    """
    
    
    # driverを起動
    driver=set_driver(False)
    
    # webサイトの自動操作及びデータの取得
    browsing_articles = Automatic_operation_of_website(driver, NIKKEI_URL)
    
    #スプレッドシートに接続しワークシートの取得
    ws = connect_gspread(jsonf,splead_sheets_key,today)


    # csvデータとして出力する
    splead_sheets_output(browsing_articles, ws)






### Chromeを起動する関数
def set_driver(headless_flg):
    # Chromeドライバーの読み込み
    options = ChromeOptions()

    # ChromeのWebDriverオブジェクトを作成する。環境変数内にpathが存在しているため、executable_pathの指定はいらない
    return Chrome(options=options) 


# webサイトを自動操作及びデータを取得する関数
def Automatic_operation_of_website(driver, url: str):

        # nikkei新聞に接続
        logging.info('Connecting to Nikkei newspaper・・・')
        driver.get(url)
        logging.info('Successful connection')
        
        time.sleep(2)
        
        # ログインボタンクリック
        driver.find_element_by_class_name("k-button--primary").click()

        # 認証key及びパスワードの送信
        driver.find_element_by_id('LA7010Form01:LA7010Email').send_keys(NIKKEY_KEY)
        driver.find_element_by_id('LA7010Form01:LA7010Password').send_keys(NIKKEY_PASSWORD)
        driver.find_element_by_class_name('btnM1').click()
        
        logging.info('Authentication was successful')

        # マイニュースコンテンツに移動
        count = 2
        items = driver.find_elements_by_class_name("k-header-tool-nav__item-body")
        items[count].click()
        
        logging.info('Moved to My News Content')
                
             
        
        # 保存記事のクリック
        driver.find_element_by_link_text('保存記事').click()

        
        time.sleep(5)

        # 記事の取得
        dates = []
        titles = []
        hrefs = []
        number = 1

        logging.info('保存記事を取得します')

        for item in driver.find_elements_by_class_name("nui-feed__item"):
            i = item.find_element_by_tag_name('a')
            title = i.text
            if title:
                logging.info(f'{number}つ目の記事を取得中・・・')
                dates.append(datetime.datetime.now().strftime('%Y年%m月%d日'))
                href = i.get_attribute('href')
                titles.append(title)
                hrefs.append(href)
                logging.info(f'{number}つ目の記事の取得完了・・・')
                number += 1
                time.sleep(3)
            


        browsing_articles = []
        for date, title, href in zip(dates, titles, hrefs):
            browsing_articles.append([date, title, href])
            
        
        

        return browsing_articles


        
    
#csv出力する関数
def splead_sheets_output(browsing_articles, ws):
        for row_data in browsing_articles:
            ws.append_row(row_data)
        # df = pd.DataFrame(browsing_articles)
        # logging.info(f'NIKKEI{today}.csvとして/Users/toguchitaichi/Desktop/NIKKEI記事一覧/に保存します')
        # df.to_csv(f'/Users/toguchitaichi/Desktop/NIKKEI記事一覧/NIKKEI{today}.csv', encoding='utf-8')
        # logging.info('保存完了')



def connect_gspread(jsonf, key, today):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonf, scope)
    gc = gspread.authorize(credentials)
    SPREADSHEET_KEY = key
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    worksheet_list = workbook.worksheets()
    worksheet = workbook.get_worksheet(len(worksheet_list)-1)
    worksheet.update_title(today)
    workbook.add_worksheet(title='最後尾', rows=100, cols=26)
    return worksheet


### 直接起動された場合はmain()を起動(モジュールとして呼び出された場合は起動しないようにするため)
if __name__ == "__main__":
    main()
