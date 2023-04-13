#!/usr/bin/env python
# coding: utf-8

# ## Функция для обновления экселя

# In[ ]:


def ExcelRefresh(CurrentTry,excelFile,FirstWS, SecondWS, ReqColumn):
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.DisplayAlerts = False
    Excel.Visible = True
    CurrentFact = Excel.Workbooks.Open(excelFile)
    time.sleep(30)
    # Вычисления формул (работают только на лист)
    Test = CurrentFact.Worksheets(FirstWS)
    Data = CurrentFact.Worksheets(SecondWS)
    time.sleep(5)

    Data.EnableCalculation = True
    Test.EnableCalculation = True
    Data.Calculate()
    Test.Calculate()
    time.sleep(5)

    CurrentFact.RefreshAll()

    time_wait = 600 + 120 * CurrentTry
    if Data.Cells(6, ReqColumn).value != "Проверка пройдена" and time_wait < 1500:
        Data.Calculate()
        time.sleep(time_wait)

    else: print("Файл обновлен, сохраняюсь")

    # Проверям, можно ли выйти из отчета
    CheckDate = Data.Cells(6, ReqColumn).value

    if CheckDate == "Проверка пройдена":
        print("Отчет обновлен, двигаемся дальше")
        time.sleep(10)
        CurrentFact.Close(True)
        Excel.Quit()
        Excel = None
    else:
        print("Прошло 20 минут, отчет не обновился, двигаемся дальше")
        CurrentFact.Close(False)
        Excel.Quit()
        Excel = None
        CurrentTry = CurrentTry + 1
        ExcelRefresh(CurrentTry,excelFile,FirstWS, SecondWS, ReqColumn)

