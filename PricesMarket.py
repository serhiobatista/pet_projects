#!/usr/bin/env python
# coding: utf-8

# # Скрипт для выгрузки биржевых цен не нефтепродукты

# In[ ]:


import os
import glob
from os import path
from datetime import datetime, timedelta
import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import time
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

import shutil
import pandas as pd
from selenium import webdriver



# URL and paths

web_page = 'https://spimex.com/markets/oil_products/indexes/regional/'
download_dir = r'C:\Users\skolesov\Downloads'
dist_path = r'\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Общая\Реализация НП\План-факт\Ежедневный отчет по АЗС\Архив\Биржа'
path = r"C:\Users\skolesov\Скрипты\Выгрузки\Биржа.xlsx"

driver_service = Service(executable_path=r"C:\Users\skolesov\Desktop\chromedriver.exe")

driver = webdriver.Chrome(service=driver_service)
driver.maximize_window()
driver.set_window_size(1400, 1100)
driver.get(web_page)




#Elements of page
bookmark = '/html/body/main/section/div/div[2]/div/div[3]/div[1]/span[2]'
prices_csv = '/html/body/main/section/div/div[2]/div/div[3]/div[2]/div[2]/div/div[1]/div/a'
alert = '/html/body/main/section/div/div[2]/div/div[6]/div[2]/div/div[2]/div[2]/div[3]/form/input[4]'




check_alert = 'Термины и определения'

if check_alert in driver.page_source:
    driver.find_element(By.XPATH, alert).click()
    time.sleep(3)

    
driver.find_element(By.XPATH, bookmark).click()
time.sleep(3)
    
driver.find_element(By.XPATH, prices_csv).click()
time.sleep(5)
driver.quit()



path_to_file = sorted(Path(download_dir).iterdir(), key = os.path.getmtime)[-1]


df = pd.read_csv(path_to_file, sep=';', parse_dates = ['date'])

df = df[['date','product_code', 'center_code', 'index_value_per_liter']]
df = df.rename(columns={'date':'Дата', 'product_code':'Товар','center_code':'Регион','index_value_per_liter':'Биржевая цена'})

drop_goods = ['TRD','MZT']
df = df.query('Дата >= "2023-01-01" and Товар not in @drop_goods')

df['Дата'] = pd.to_datetime(df['Дата']).dt.date

goods = {'DTL':'ДТ', 'DTM':'ДТ', 'DTZ':'ДТ', 'PRM':'95', 'REG':'92', 'SUG':'СУГ'}
df['Товар'] = df['Товар'].replace(goods)

df['Биржевая цена'] = df['Биржевая цена'].apply(lambda x: x.replace(',','.') if type(x) != float  else x)
df['Биржевая цена'] = df['Биржевая цена'].astype(np.float64)




files = glob.glob(f'{dist_path}/*')
for f in files:
    os.remove(f)
    

df.to_excel(f'{dist_path}\Биржа.xlsx',index=False)

