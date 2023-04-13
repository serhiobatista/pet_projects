#!/usr/bin/env python
# coding: utf-8

# ##  Выгрузка данных из НУС. Общий отчет по заключенным  коммерческим договорам договорам с 01.01.2022 по текущую дату.

# In[1]:


import time
import pandas as pd
import win32com.client as win32
import os
import shutil

from dotenv import load_dotenv
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver

from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from datetime import datetime, date, timedelta
from pathlib import Path

date_strt= '01.01.2022'
date_endt = datetime.now().strftime('%d.%m.%Y') #datetime


driver = webdriver.Chrome(executable_path=r"C:\Users\skolesov\Desktop\chromedriver.exe")

driver.set_window_size(1400, 1100)
driver.get('http://10.20.14.57/les_reports/nus.php')


time.sleep(3)

type_report = driver.find_element(By.XPATH,"/html/body/div[2]/form/fieldset/input[3]")
type_report.click()

date_start = driver.find_element(By.XPATH, '/html/body/div[2]/form/fieldset/div[1]/input')
date_start.clear()
date_start.send_keys(date_strt)

time.sleep(1)

random_click = driver.find_element(By.XPATH,'/html/body/div[1]/div')
random_click.click()

time.sleep(1)

date_end = driver.find_element(By.XPATH, '/html/body/div[2]/form/fieldset/div[2]/input')
date_end.clear()
date_end.send_keys(date_endt)

time.sleep(1)

random_click = driver.find_element(By.XPATH,'/html/body/div[1]/div')
random_click.click()

driver.set_page_load_timeout(3000)
download = driver.find_element(By.XPATH, '/html/body/div[2]/form/input')
download.click()
        
driver.switch_to.alert.accept() #скипаем алерт


download_2 = driver.find_element(By.XPATH, '/html/body/center/h4/a')
download_2.click()

time.sleep(10)

dirpath = "C:/Users/skolesov/Downloads"
dstpath = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Общая\Реализация НП\рейтинги отчет\архив"

paths = sorted(Path(dirpath).iterdir(), key=os.path.getmtime)

shutil.move(f'{dirpath}/{paths[-1].name}',f'{dstpath}/period.csv')


# In[5]:


#обновляем модель на основе обновленной выгрузки

time.sleep(20)
xlapp = win32.DispatchEx('Excel.Application')
xlapp.DisplayAlerts = False
xlapp.Visible = True

path = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Общая\Реализация НП\рейтинги отчет\period.xlsx"
xlbook = xlapp.Workbooks.Open(path)

xlbook.RefreshAll()
time.sleep(30)

xlbook.Save()
time.sleep(20)
xlbook.Close()
xlapp.Quit()

