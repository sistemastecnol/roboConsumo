from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from pathlib import Path
import openpyxl
import os
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

import time
import csv
driver = webdriver.Firefox()
driver.get("https://safedataanalytics.com.br/report/financial")
driver.maximize_window()
delay = 3

login = driver.find_element(By.CSS_SELECTOR, "div.form-group:nth-child(1)>input:nth-child(2)")
login.send_keys('lucivando.santos@sistemastecnol.com.br')
time.sleep(60)

relatorios=driver.find_element(By.CSS_SELECTOR, "li.nav-item:nth-child(6)>span:nth-child(1)")
relatorios.click()

time.sleep(3)

faturamento=driver.find_element(By.CSS_SELECTOR, "ul.show>li:nth-child(3)>a:nth-child(1)")
faturamento.click()

time.sleep(3)

filtrar=driver.find_element(By.CSS_SELECTOR, ".btn")
filtrar.click()

time.sleep(5)

qtdregistros= Select(driver.find_element(By.NAME, "financialTable_length"))
qtdregistros.select_by_index(3)

time.sleep(3)

registros= driver.find_element(By.CSS_SELECTOR, ".fi-rr-file-delete")
registros.click()
time.sleep(3)
user=os.getlogin()
path_to_file = 'C:/Users/'+ user +'/Downloads/Relatório Faturamento.xlsx'
path = Path(path_to_file)

if path.is_file():
    print(f'The file {path_to_file} exists')
    df=pd.read_excel(path,skiprows=2,skipfooter=1,header=None,index_col=0,usecols=None,dtype={'Preço':float},engine="openpyxl")
    ultimalinha=int(len(df))
    ultimalinha+1

    with pd.ExcelWriter(path, mode="a",if_sheet_exists='overlay',engine="openpyxl") as f:
        df.to_excel(f,sheet_name='Base',header=None,startrow=ultimalinha)

else:
    print(f'The file {path_to_file} does not exist')
