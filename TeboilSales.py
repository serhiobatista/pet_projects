#!/usr/bin/env python
# coding: utf-8

# ## Скрипт для автоматического заполнения справки по реализации по сети Тебойл

# In[ ]:


import pandas as pd
import glob
import time
import openpyxl
import shutil
import win32com.client as win32

from datetime import datetime, timedelta, date
from pathlib import Path
from os import path


# In[208]:




def clear_folder(path):
    
    list_of_files = glob.glob(path)
    latest_file = max(list_of_files, key=os.path.getctime)

    last_modification = os.path.getctime(latest_file) #timestamp
    last_date_mod = datetime.fromtimestamp(last_modification).date() #date дата последнего изменения файла в директории

    if (last_date_mod == datetime.today().date() - timedelta(days=1) or last_date_mod == datetime.today().date() - timedelta(days=3)) and datetime.today().date().strftime('%d.%m.%Y') != datetime.today().date().strftime('02.%m.%Y'):
        os.remove(latest_file)
        print('Файл удален')
    


# ### Удаляем старые файлы за отчетный период из папок с do и rba

# In[209]:


path_rba = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\выгрузки\*"
path_do = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\выгрузки фрайнчайзинг\*"


# In[210]:


clear_folder(path_rba)
clear_folder(path_do)


# ### Копируем файл с текущим месяцем в папку выгрузки, на основе которой формируется модель данных по тебойлу.

# In[211]:


path = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Общая\Реализация НП\План-факт\Тебойл\*"
dst_rba = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\выгрузки"
path_teb_prices = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Общая\Реализация НП\Цены на АЗС\цены тебойл текущий день\*"

for _ in range(1000):
    list_of_files = glob.glob(path)
    latest_file = max(list_of_files, key=os.path.getctime) #находим самый свежий файл

    last_modification = os.path.getctime(latest_file)
    last_date_mod = datetime.fromtimestamp(last_modification).date() #определяем время обновления самого последнего файла
    
    time.sleep(5)
    if last_date_mod == datetime.now().date():
        
        src = path[:-1] + latest_file.split("\\")[-1]
        dst = dst_rba + "\\" + latest_file.split("\\")[-1]
        print(dst)
        
        shutil.copyfile(src, dst)
        time.sleep(5)
        dataframe = openpyxl.load_workbook(dst)

        ws_line = dataframe["Sheet1"]
        ws_line.delete_cols(11)
        ws_line.delete_cols(6,2)
        ws_line.delete_cols(4)
        dataframe.save(dst)
        
        # Копируем файл c ценами на сегодня по тебойлу в папку, где лежат цена для обновления bi по динамике изменения цен
        list_of_files = glob.glob(path_teb_prices)
        latest_file = max(list_of_files, key=os.path.getctime)
        os.remove(latest_file)
        shutil.copyfile(dst, path_teb_prices[:-1]+dst.split('\\')[-1])
        
        break
    else:
        print("nope")
        continue


# ### Скачиваем франчайзинг и закидываем в папку

# In[212]:


log = 'login'
password = 'password'

driver = webdriver.Chrome(executable_path=r"C:\Users\skolesov\Desktop\chromedriver.exe")
driver.maximize_window()
#driver.set_window_size(1400, 1100)
driver.get('https://sales.teboil-azs.ru/stat')

time.sleep(3)

login = driver.find_element(By.XPATH, "/html/body/app/ng-component/div/div/div[1]/login/form/div[1]/input")
login.send_keys(log)

passs = driver.find_element(By.XPATH, "/html/body/app/ng-component/div/div/div[1]/login/form/div[2]/input")
passs.send_keys(password)

time.sleep(5)

button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div[1]/login/form/div[5]/p[1]/button')
button.click() 
time.sleep(3)

#Выбрать DO в выпадающем списке
button = driver.find_element(By.XPATH,'/html/body/app/ng-component/div/div/div/main/header/div[1]/span/main-filter/div/form/div/div[2]/data-sharing/form/select/option[3]')
button.click()
time.sleep(3)

# Подтвердить выбранные фильтры
button = driver.find_element(By.XPATH,'/html/body/app/ng-component/div/div/div/main/header/div[1]/span/main-filter/div/form/div/div[3]/button')
button.click()
time.sleep(2)

#шестеренка
button = driver.find_element(By.XPATH,'/html/body/app/ng-component/div/div/div/main/div/ng-component/tab-indicators/div/ul/li[6]')
button.click()
time.sleep(3)

#отчеты
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/revise-tabs/div/ul/li[2]')
button.click()
time.sleep(3)

#выпадающий список
time.sleep(3)
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[2]/div/report-create/div/div/angular2-multiselect/div')
button.click()

#проджажи по топливу
time.sleep(2)
button = driver.find_element(By.XPATH,'/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[2]/div/report-create/div/div/angular2-multiselect/div/div[2]/div[3]/div[2]/ul/li[1]/label')
button.click()

#cнять выделение с типа отчета
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[2]/div/report-create/div/div[2]/report-by-fuel/table/tbody/tr[3]/td[2]/label/input')
button.click()
time.sleep(1)

#выбрать тип данных
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[2]/div/report-create/div/div[2]/report-by-fuel/table/tbody/tr[2]/td[2]/label/input')
button.click()
time.sleep(1)



#нажать выгрузить
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[2]/div/report-create/div/div[1]/button')
button.click()
time.sleep(8)

#история запросов
button = driver.find_element(By.XPATH, '/html/body/app/ng-component/div/div/div/main/div/app-control/div/report/div[1]/div/div/ul/li[2]/a')
button.click()
time.sleep(1)

for i in range(1,2000):
        if not("Формируется" in driver.page_source):
            time.sleep(5)
            GetData = driver.find_element(By.XPATH, "//a[@title='Скачать']")
            GetData.click()
            time.sleep(5)
            print("Отчет захвачен")
            break
        else: 
            time.sleep(60)
time.sleep(10)

path = r"C:\Users\skolesov\Downloads\*"
dst = r'\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\выгрузки фрайнчайзинг'
list_of_files = glob.glob(path)
latest_file = max(list_of_files, key=os.path.getctime)

file_name = latest_file.split('\\')[-1]
shutil.move(f'{path[:-1]}{file_name}', f'{dst}\\{file_name}')


# ### Обновляем модель данных по Тебойл

# In[213]:


time.sleep(10)
xlapp = win32.DispatchEx('Excel.Application')
xlapp.DisplayAlerts = False
xlapp.Visible = True

path = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\Реализация Тебойл с 2019 upd.xlsm"
xlbook = xlapp.Workbooks.Open(path)

time.sleep(4)


xlbook.RefreshAll()
time.sleep(40)

xlapp.RUN("RBA")
xlapp.RUN("franch")
time.sleep(10)

xlbook.Save()
time.sleep(40)
xlbook.Close()
xlapp.Quit()


# ### Вставляем значения в справку за прошлый день 

# In[214]:


xlapp = win32.DispatchEx('Excel.Application')
xlapp.DisplayAlerts = False
xlapp.Visible = True

path_md = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Реализация с 2019 года\Реализация Тебойл с 2019 upd.xlsm"
path_reports = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Ежедневный отчет\04_2023\*"
#путь выше надо будет обновлять каждый месяц
#list_of_files = glob.glob(path_reports)
#path_sparvka = max(list_of_files, key=os.path.getctime)

xlbook_svod = xlapp.Workbooks.Open(path_md)


# In[215]:


svod = xlbook_svod.Worksheets(u'свод')
svodnyi = xlbook_svod.Worksheets(u'сводные')



ul_fact = svod.Range("B10:D11").value
fl_fact = svod.Range("B13:D14").value

franch_prev = svod.Range("N9:O10").value #франчайзинг 1
franch_curr = svod.Range("P9:P10").value #франчайзинг 2

if datetime.now().day != 1:
    r2 = datetime.today().day + 9 - 2
    r1 = datetime.today().day + 10 - 1
else:
    date = datetime.today() - timedelta(days=1)
    r2 = date.day + 9 - 1
    r1 = date.day + 10
    

ul_2022 = svodnyi.Range(f"B9:B{r2}").value
ul_2023 = svodnyi.Range(f"I9:I{r2}").value
ul_2023_lc = svodnyi.Range(f"L9:L{r2}").value

ab_lc = svodnyi.Cells(r1,15).value
dt_lc = svodnyi.Cells(r1,16).value

xlbook_svod.Close()


# In[216]:


#объем продаж за отчетный период


path_sparvka = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Ежедневный отчет\шаблон заполнения\Тебойл.xlsx"

xlbook_spravka = xlapp.Workbooks.Open(path_sparvka)
spravka = xlbook_spravka.Worksheets(u'План линейный')

spravka.Range("E11:F11").value = fl_fact[0][:2] #аб
spravka.Range("E12:F12").value = fl_fact[1][:2] #дт

spravka.Range("H11:H11").value = fl_fact[0][2] # 2023аб
spravka.Range("H12:H12").value = fl_fact[1][2] # 2023дт


spravka.Range("E14:F14").value = ul_fact[0][:2] #аб
spravka.Range("E15:F15").value = ul_fact[1][:2] #дт

spravka.Range("H14:H14").value = ul_fact[0][2] # 2023аб
spravka.Range("H15:H15").value = ul_fact[1][2] # 2023дт

spravka.Range("E32:F33").value = franch_prev
spravka.Range("H32:H33").value = franch_curr


spravka.Range("H17:H17").value = ab_lc #тк лукойл аб
spravka.Range("H18:H18").value = dt_lc #тк лукойл дт


# In[217]:


#сравнение с предыдущими периодами


spravka.Range(f"Q9:Q{r2}").value = ul_2022
spravka.Range(f"T9:T{r2}").value = ul_2023
spravka.Range(f"U9:U{r2}").value = ul_2023_lc

if isinstance(ul_2022, float) == True:
    spravka.Cells(9,24).formula = "=T" + str(9) + "/Q" +str(9) +"-1"
else:
    for i in range(9, len (ul_2022)+9):
        spravka.Cells(i,24).formula = "=T" + str(i) + "/Q" +str(i) +"-1"


# In[218]:


#добавляем планы
#сначала линейный

plans = xlbook_spravka.Worksheets(u'планы')

planl_fl = plans.Range("B4:B5").value
planl_ul = plans.Range("B7:B8").value

plans_fl = plans.Range("C4:C5").value
plans_ul = plans.Range("C7:C8").value

spravka_linear = xlbook_spravka.Worksheets(u'План линейный')
spravka_linear.Range("G11:G12").value = planl_fl
spravka_linear.Range("G14:G15").value = planl_ul

spravka_seasoned = xlbook_spravka.Worksheets(u'План сезонированный')
spravka_seasoned.Range("G11:G12").value = plans_fl
spravka_seasoned.Range("G14:G15").value = plans_ul


full_list_linear = spravka_linear.Range("B1:X40").value
full_list_seasoned = spravka_seasoned.Range("B1:X40").value

xlbook_spravka.Save()
time.sleep(15)
xlbook_spravka.Close()


# In[219]:


#Закидываем инфу в итоговый файл

prev_day = datetime.today() - timedelta(days=1)
file_name = prev_day.strftime('%d.%m.%Y')

path_shablon = r"\\Msk01-fileshare\pub\ЗГД по экономике и финансам\Шелл\Реализация\Ежедневный отчет\шаблон заполнения\Шаблон.xlsx"
dst = f'//Msk01-fileshare/pub/ЗГД по экономике и финансам/Шелл/Реализация/Ежедневный отчет/04_2023/Тебойл на ({file_name}).xlsx'



xlbook = xlapp.Workbooks.Open(path_shablon)

spravka_linear = xlbook.Worksheets(u'План линейный')
spravka_linear.Range("B1:X40").value = full_list_linear

spravka_seasoned = xlbook.Worksheets(u'План сезонированный')
spravka_seasoned.Range("B1:X40").value = full_list_seasoned

xlbook.Save()
xlbook.Close()
xlapp.Quit()

shutil.copyfile(path_shablon, dst)


# In[ ]:





# In[ ]:




