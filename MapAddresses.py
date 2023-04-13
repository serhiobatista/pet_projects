#!/usr/bin/env python
# coding: utf-8

# ## Скрипт ищет по адресу координаты АЗС конкурента с помощью Яндекс.Карт

# In[ ]:


from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import os
import glob
from os import path
from datetime import datetime
from datetime import datetime, timedelta
import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import time
from selenium.webdriver.common.by import By

import pandas as pd
from selenium import webdriver
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



address = pd.read_excel(r'C:\Users\skolesov\Desktop\Конкуренты ЦНП.xlsx')
path_to_save = 'address_azs_cnp.xlsx'

l_address = address['АЗС_конкурента'].to_list()

col_name_ad = 'address'
col_name_lat = 'latitude'
col_name_lon = 'longitude'

latitude = []
longitude = []



driver_service = Service(executable_path=r"C:\Users\skolesov\Desktop\chromedriver.exe")
driver = webdriver.Chrome(service=driver_service)

driver.maximize_window()
#driver.set_window_size(1400, 1100)
driver.get('https://snipp.ru/tools/address-coord')


#Web elements
search_field = '/html/body/div[2]/main/div[2]/div/ymaps/ymaps[5]/ymaps/ymaps[1]/ymaps/ymaps/ymaps[1]/ymaps/ymaps/ymaps/ymaps[1]/ymaps/ymaps[2]/input'
click_search = '/html/body/div[2]/main/div[2]/div/ymaps/ymaps[5]/ymaps/ymaps[1]/ymaps/ymaps/ymaps[1]/ymaps/ymaps/ymaps/ymaps[2]/ymaps/ymaps[2]/ymaps'
clear_field = '/html/body/div[2]/main/div[2]/div/ymaps/ymaps[5]/ymaps/ymaps[1]/ymaps/ymaps/ymaps[1]/ymaps/ymaps/ymaps/ymaps[1]/ymaps/ymaps[2]/ymaps'

n = 0
for i in l_address:
    n += 1
    
    if i == 'нет конкурентов':
        latitude.append(0)
        longitude.append(0)
    else:
        time.sleep(2)
        driver.find_element(By.XPATH,search_field).send_keys(i)

        time.sleep(1.5)
        driver.find_element(By.XPATH,click_search).click()

        time.sleep(2)
        flag = True
        while flag:
            try:
                time.sleep(1)
                driver.find_element(By.XPATH,clear_field).click()
                flag=False
            except ElementNotInteractableException:
                time.sleep(1.5)
                print('Не нашел нужную кнопку, попробую еще раз')
                

        time.sleep(1.5)
        x=driver.find_element(By.XPATH,'/html/body/div[2]/main/div[4]/div/input').get_attribute('value')
        time.sleep(1.5)
        latitude.append(x.split(',')[0])

        longitude.append(x.split(',')[1])
        print(n)




data = pd.DataFrame({'address':l_address, 'latitude':latitude, 'longitude':longitude})

