from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta
from Self_function import *

def html_IO_table(table):
    data=[]

    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    # for idx,row in enumerate(rows):
    #     print(idx,row)
    drainage=rows[58]
    
    drainage_table=drainage.find('table')
    drainage_table=drainage_table.find('tbody')
    drainage_rows = drainage_table.find_all('tr')

    drainage_data=[]
    for drainage_row in drainage_rows:
        cols = drainage_row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        drainage_data.append(cols)
    # print(drainage_data)
    print(drainage_data)
    df = pd.DataFrame(drainage_data,columns=["項目","白班","小夜","大夜","總量"])

    # for row in rows:
    #  cols = row.find_all('td')
    #  cols = [ele.text.strip() for ele in cols]
    #  one_col=[ele for ele in cols if ele]
    #  # if \"New\" in one_col[1]:
    #  #     one_col[1]=one_col[1][4:]\
    #  one_col=one_col[0:5]
    #  # print(one_col)
    #  if not one_col==[]:
    #      data.append(one_col) # Get rid of empty values
    # df = pd.DataFrame(data[1:],columns=data[0])
    return df


def get_IO(driver, ID):
    adminID=get_adminID(driver,ID)
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=goNIS&hisid="+ID+"&caseno="+adminID)
    date=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
    # date="20240924"
    driver.get("https://web9.vghtpe.gov.tw/NIS/report/IORpt/details.do?gaugeDate1="+date)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    soup=soup.find(id="divshow_0")
    IOtable=soup.table.table.findAll('table')[1]
    df=html_IO_table(IOtable)
    return df