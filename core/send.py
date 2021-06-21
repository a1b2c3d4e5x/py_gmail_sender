
import sys
sys.path.append("..")
from const.color import output_error
from concurrent.futures import ThreadPoolExecutor
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from datetime import datetime

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import traceback
import requests
import random
import time
import sys
import os

import smtplib
import ssl

def write_to_file(log_file: str, text: str):
    today = datetime.today()
    time = today.strftime('%Y-%m-%d-%H:%M:%S')
    log_df = pd.DataFrame([(text, time)], columns=['LOG', 'TIME'])
    with open(log_file, mode = 'a') as f:
        log_df.to_csv(f, header = f.tell() == 0, index = False)
    output_error(text)

def to_log(text: str):
    today = datetime.today()
    date = today.strftime('%Y-%m-%d')
    log_file = os.getcwd() + '/result/log-' + date + '.csv'
    write_to_file(log_file, text)

def to_error(text: str):
    today = datetime.today()
    date = today.strftime('%Y-%m-%d')
    log_file = os.getcwd() + '/result/error-' + date + '.csv'
    write_to_file(log_file, text)

def some_exception(e):
    error_class = e.__class__.__name__           # 取得錯誤類型
    detail = e.args[0]                           # 取得詳細內容
    cl, exc, tb = sys.exc_info()                 # 取得Call Stack
    lastCallStack = traceback.extract_tb(tb)[-1] # 取得Call Stack的最後一筆資料
    fileName = lastCallStack[0]                  # 取得發生的檔案名稱
    lineNum = lastCallStack[1]                   # 取得發生的行號
    funcName = lastCallStack[2]                  # 取得發生的函數名稱
    errMsg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, error_class, detail)
    to_error(errMsg)

class SendMail(object):

    def from_csv(sender_csv: str, mail_format: str, receiver_csv: str):
        # 建立資料夾
        result_folder = os.getcwd() + '/result'
        if False == os.path.isdir(result_folder):
            os.mkdir(result_folder)
            
        try:
            sender_df = pd.read_csv(sender_csv)
            receiver_df = pd.read_csv(receiver_csv)
        except Exception as e:
            print('無法轉換為 pandas')
            some_exception(e)
        
        mail_index: int = 0 # 用於記錄目前使用於寄出 email 的 index
        for index, row in receiver_df.iterrows():
            sender_email: str = sender_df['EMAIL'][sender_df.index[mail_index]]
            sender_password: str = sender_df['APPLICATION_PASSWORD'][sender_df.index[mail_index]]
 
            receiver_name = row["NAME"]
            receiver_email = row["EMAIL"]

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(host='smtp.gmail.com', port='465') as smtp:  
                try:
                    print('sender_email: ' + sender_email + ' -> ' + receiver_email)

                    content = MIMEMultipart()
                    content["subject"] = 'Hi: ' + receiver_name
                    content["from"] = sender_email
                    content["to"] = receiver_email
                    content.attach(MIMEText("Demo python send email"))
                    
                    smtp.ehlo()                                     # 驗證 SMTP 伺服器
                    smtp.login(sender_email, sender_password)       # 登入寄件者gmail

                    print('1')
                    smtp.send_message(content)
                    print('2')

                    to_log('[Done] ' + receiver_email)

                    time.sleep(5)
                except Exception as e:
                    print('無法送出 Email')
                    some_exception(e)
                    #traceback.print_exc()
                    time.sleep(5)

            mail_index += 1
            if mail_index >= len(sender_df):
                mail_index = 0

            

            

        return
        
        



        

        return 
        # Google SMTP Server
        
            

    def from_excel(sender_csv: str, mail_format: str, receiver_excel: str):
        print(sender_csv, mail_format, receiver_csv)

        return

        if False == isinstance(area_id, int):
            area_id = 0

        #fixed_url: str = 'https://www.iyp.com.tw/showroom.php?cate_name_eng_lv1=agriculture&cate_name_eng_lv3=agriculture-equip&a_id=4'
        fixed_url: str = 'https://www.iyp.com.tw/showroom.php?cate_name_eng_lv1=' + main_category + '&cate_name_eng_lv3=' + sub_category + '&a_id=' + str(area_id)
        page: int = 0
        total_count: int = 0
        total_email: int = 0

        """
        # test parse content
        # https://www.iyp.com.tw/082322152
        content = Spider_ipy.spider_content('https://www.iyp.com.tw/082322152')
        print(content)
        return
        """

        try:
            # 建立資料夾
            result_folder = os.getcwd() + '/result'
            if False == os.path.isdir(result_folder):
                os.mkdir(result_folder)
            main_cate_folder = result_folder + '/' + main_category
            if False == os.path.isdir(main_cate_folder):
                os.mkdir(main_cate_folder)
            sub_cate_file = result_folder + '/' + main_category + '/' + sub_category + '.csv'
            sub_cate_log = result_folder + '/' + main_category + '/__log__.csv'

        except:
            to_log('無法建立資料夾: ' + main_category + ', ' + sub_category)
            traceback.print_exc()
            return

        while True:
            try:
                # 組合出要爬得 url
                target_url = fixed_url + '&p=' + str(page)
                to_log('[TARGET] ' + target_url)
  
                # 假資料模擬瀏覽器
                headers = {'user-agent': UserAgent().random}

                # 取得網頁資料
                pageRequest = requests.get(target_url, headers = headers)
                pageRequest.encoding = pageRequest.apparent_encoding

            except:
                to_log('無法 request 列表: ' + target_url)
                traceback.print_exc()

                time.sleep((random.random() * 10) + 120)
                continue;

            try:
                soup = BeautifulSoup(pageRequest.text, 'html.parser')
            except:
                to_log('無法轉成 html: ' + target_url)
                to_log('細節: \n\t' + str(pageRequest).replace('\n', '\n\t'))
                traceback.print_exc()
                break;
 
            store_data_array = []
            try:
                # 爬到文章 List 區塊
                res_block_list = soup.find(id = 'search-res')

                # VIP 店家, 優質店家,  一般店家
                store_block_list = res_block_list.find_all('ol', class_ = ['recommend', 'diamond', 'general'], recursive = False)
                if 0 == len(store_block_list):
                    break

                for list_ol in store_block_list:
                    store_list = list_ol.find_all('li', recursive = False)
                    if 0 == len(store_list):
                        break

                    for list_li in store_list:
                        item_a = list_li.select_one('h3 a')

                        # 統一網站的 url 格式
                        store_name = item_a.text
                        store_name_url = item_a['href']
                        if '//ww' == store_name_url[:4]:
                            store_name_url = 'https:' + store_name_url
                        elif 'www.' == store_name_url[:4]:
                            store_name_url = 'https://' + store_name_url
   
                        store_data_array.append({'name': store_name, 'url': store_name_url})

            except TypeError:
                to_log('Parsing list url 失敗: ' + target_url)
                traceback.print_exc()
                break

            total_count += len(store_data_array)
            if 0 == len(store_data_array):
                to_log('找不到任何資料: ' + target_url)
                break
            
            # 爬店家 ipy 內容頁，取回有 Email 的資訊
            iyp_result = []

            """
            with ThreadPoolExecutor(max_workers = 5) as executor:
                time.sleep((random.random() + 0.5) * 2)
                results = executor.map(Spider_ipy.spider_content, data['url'])
            """


            for data in store_data_array:
                if 'https://www.iyp.com.tw/' in data['url']:
                    print(' [FETCH] ' + data['name'], data['url'])

                    time.sleep((random.random() + 0.5) * 2)
                    store_content = Spider_ipy.spider_content(data['url'])
                    if None != store_content:
                        print('\033[36m [EMAIL] ' + store_content[0], '\033[0m')
                        iyp_result.append((data['name'], data['url'], store_content[0], store_content[1]))
                else:
                    print('  [SKIP] ' + data['name'], data['url'])

            # 儲存有找到 Email 的那些資料
            if 0 != len(iyp_result):
                total_email += len(iyp_result)
                df = pd.DataFrame(iyp_result, columns=['NAME', 'URL', 'EMAIL', 'WEBSITE'])
                with open(sub_cate_file, mode = 'a') as f:
                    df.to_csv(f, header = f.tell() == 0, index = False)

            page += 1
            time.sleep(3)

        to_log('  [DONE] cate: ' + main_category + ', sub-cate: ' + sub_category + ', total: ' + str(total_count) + ', email: ' + str(total_email))

        # 記錄這一次的 Log
        fetch_log = [(str(total_count), str(total_email), sub_category, fixed_url)]
        

        df_log = pd.DataFrame(fetch_log, columns=['TOTAL', 'EMAIL', 'SUB_CATEGORY', 'URL'])
        with open(sub_cate_log, mode = 'a') as f:
            df_log.to_csv(f, header = f.tell() == 0, index = False)

