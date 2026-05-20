from os.path import join, dirname, realpath
import os
import datetime as dt
import time
import pandas as pd
from bs4 import BeautifulSoup
import requests
import json
import google.auth
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
from google.auth import credentials
from google.oauth2 import service_account

class NoNewsFound(Exception):
    pass


class vsdc:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException
        )

    def tinTCPH(self) -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        (tin từ tổ chức phát hành).
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """
        
        target_time = dt.datetime.strptime("19/05/2026 00:00:00","%d/%m/%Y %H:%M:%S")                    
        current_time = dt.datetime.now()
              
        url = 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        driver = webdriver.Chrome()
        #executable_path=self.PATH
        driver.get(url)
        driver.maximize_window()
        keywords = [ "Chi trả cổ tức",
                    "Chi bổ sung cổ tức","Trả cổ tức",'Cấp Giấy chứng nhận đăng ký chứng khoán thay đổi',
                     'Điều chỉnh thông tin số lượng', "Đăng ký, lưu ký"]
            
            
                    
        frames = []
        while current_time>= target_time:
            news_time = []
            news_headlines = []
            news_urls = []
            news_noidung = []

            time.sleep(1)
            driver_wait = WebDriverWait(driver, 5, ignored_exceptions=self.ignored_exceptions)
            list_news_elem = driver_wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'list-news')))
            tags = list_news_elem.find_elements(By.TAG_NAME, 'li')
            for tag_ in tags:
                tag_wait = WebDriverWait(tag_, 5, ignored_exceptions=self.ignored_exceptions)
                h3_tag = tag_wait.until(EC.presence_of_element_located((By.TAG_NAME, 'h3')))
                txt = h3_tag.text
                check = [word in txt for word in keywords]
                if any(check):
                    news_headlines += [txt]
                    tag_wait = WebDriverWait(h3_tag, 5, ignored_exceptions=self.ignored_exceptions)
                    sub_url = tag_wait.until(EC.presence_of_element_located((By.TAG_NAME, 'a'))).get_attribute('href')
                    news_urls += [f'=HYPERLINK("{sub_url}","Link")']
                    news_time_ = tag_.find_elements(By.TAG_NAME, 'div')[0]
                    news_time_= news_time_.text
                    news_time_ = dt.datetime.strptime(news_time_[-21:],'%d/%m/%Y - %H:%M:%S')
                    news_time += [news_time_]
                    sub_driver = webdriver.Chrome() #executable_path=self.PATH
                    sub_driver.get(sub_url)
                    sub_driver_wait = WebDriverWait(sub_driver, 5, ignored_exceptions=self.ignored_exceptions)
                    noidung_elem = sub_driver_wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'col-md-12')))
                    noidung = noidung_elem.text
                    news_noidung += [noidung]
                    sub_driver.quit()

            frame = pd.DataFrame({
                'Thời gian': news_time,
                'Tiêu đề': news_headlines,
                'Thông tin chi tiết': news_noidung,
                'Đường dẫn': news_urls})
            frames.append(frame)
            # Turn Page
            time.sleep(1)
            button_row = driver_wait.until(EC.presence_of_element_located((By.ID, 'd_number_of_page')))
            button_elems = button_row.find_elements(By.XPATH,".//*")
            nextpage_button = button_elems[-3]
            nextpage_button.click()
            current_time = news_time_
            time.sleep(1)
            output_table = pd.concat(frames, ignore_index=True)
        driver.quit()

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong giờ vừa qua')

        def f(frame):
            series = frame['Tiêu đề'].str.split(': ')
            table = pd.DataFrame(series.tolist(), columns=['Mã cổ phiếu', 'Nội dung'])
            return table
        output_table[['Mã cổ phiếu', 'Nội dung']] = output_table.transform(f)
        output_table.drop(['Tiêu đề'], axis=1, inplace=True)
        output_table = output_table[[
            'Thời gian',
            'Mã cổ phiếu',
            'Nội dung',
            'Thông tin chi tiết',
            'Đường dẫn']]
        return output_table
obj = vsdc()
result = obj.tinTCPH()
result = result.astype(str)
base_dir = os.path.dirname(__file__)
credPathRoot = os.path.join(base_dir, "credentials.json")
scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_file(credPathRoot, scopes=scope)
client = gspread.authorize(creds)

spreadsheet_id = '1vGICehEFWMUVCKReu9iNyBUu3cbOx7LnntBM5wXagd0'
sheet = client.open_by_key(spreadsheet_id).sheet1
headers = ["NGÀY", "MÃ CỔ PHIẾU", "NỘI DUNG", "THÔNG TIN CHI TIÊT","ĐƯỜNG DẪN"]
sheet.insert_row(headers, 1)

if len(sheet.get_all_values()) == 0:
    sheet.append_rows([result.columns.tolist()],value_input_option='USER_ENTERED')              
sheet.append_rows(result.values.tolist(),value_input_option='USER_ENTERED')
    
    


