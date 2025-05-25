from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta

from io import BytesIO
from PIL import Image
import random

# split the html table
def html_table(table):
    data=[]
    table_head = table.find('thead')
    t_head = table_head.find_all('th')
    t_head = [ele.text for ele in t_head]
    
    
    table_body = table.find('tbody')

    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        one_col=[ele for ele in cols if ele]
        # if "New" in one_col[1]:
        #     one_col[1]=one_col[1][4:]
        data.append(one_col) # Get rid of empty values
    df = pd.DataFrame(data,columns=t_head)
    
    return df
#======================================
# Get TPR
def get_adminID(driver,ID):
    TPR_url="https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPbv&histno="+ID
    driver.get(TPR_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    adminID=soup.option['value'].split("=")[-1]
    return adminID

def get_TPR(driver,ID, adminID=None):
    if not adminID:
        adminID=get_adminID(driver,ID)
    
    TPR_url="https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findTpr&caseno="+adminID
    driver.get(TPR_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    soup.find(id="tprlist")
    data=html_table(soup)
    return data
#==========================================================
## Get TPR image

def get_TPR_img(driver,ID, adminID=None):
    if not adminID:
        adminID=get_adminID(driver,ID)
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findTpr&caseno="+adminID+"&pbvtype=tpr")
    temp = BytesIO(driver.get_screenshot_as_png())
    element = driver.find_element(By.TAG_NAME,"img")
    location = element.location
    size = element.size
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    image = Image.open(temp)
    image = image.crop((int(x), int(y), int(width), int(height)))
    return image
# =======================================================================
## Get BW_BL

def get_BW_BL(driver,ID, adminID="all"):
    if not adminID:
        adminID=get_adminID(ID)
    
    BW_BL_url="https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findVts&histno="+ID+"&caseno="+adminID+"&pbvtype=HWS"
    driver.get(BW_BL_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    # soup.find(id=Height)
    data=html_table(soup)
    
    return data

#==================================================================
## Get Lab value

def get_Lab_value(driver,ID, Lab_value):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=DCHEM&histno="+ID+"&resdtmonth=24")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    header_element = soup.find(id=Lab_value)
    time_list=header_element.text.split('|')
    Lab_data=[]
    for one_time in time_list:
        Lab_data.append(one_time.split("/"))
    return Lab_data
#=================================================================
## get latest admission note

def get_last_admission(driver,ID):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findAdm&histno="+ID)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    admnote = soup.find(title="admnote")
    root_url="https://web9.vghtpe.gov.tw/"
    admin_url=root_url+admnote['href']
    time.sleep(0.5)
    driver.get(admin_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    return soup.pre

# =====================================================
## get current drug

def get_drug(driver,ID):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findUd&histno="+ID)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    drug_url_list=soup.find_all("a")
    adminID=get_adminID(driver,ID)
    drug_url=drug_url_list[0]["href"]
    for a_url in drug_url_list:
        if adminID in a_url["href"]:
            drug_url=a_url["href"]

    root_url="https://web9.vghtpe.gov.tw/"
    driver.get(root_url+drug_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table=soup.find(id="udorder")
    drug_table=html_table(table)
    return drug_table

#=========================================
# split the html table
## get res report

def html_res_table(table):
    data=[]
    table_head = table.find('thead')
    t_head = table_head.find_all('th')
    t_head = [ele.text for ele in t_head]

    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    for row in rows[:-1]:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        
        # one_col=[ele for ele in cols if ele]
        data.append(cols)
    # print(data, len(t_head))
    df = pd.DataFrame(data,columns=t_head)
    return df

def get_res_report(driver, ID, resdtype="SMAC", resdtmonth="00"):
    report_dict={
        "SMAC":"DCHEM",
        "CBC":"DCBC",
        "Urine":"DURIN",
        "Cancer":"DNM1",
        
    }
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype="+report_dict[resdtype]+"&histno="+ID+"&resdtmonth="+resdtmonth)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table=soup.find(id="resdtable")
    # print(table)
    report_table=html_res_table(table)
    return report_table  

#=================

## get_progress_note

def get_progress_note(driver,ID,num=5):
    adminID=get_adminID(driver,ID)
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPrg&histno="+ID+"&caseno="+adminID)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    note_url=soup.find("a")["href"]
    root_url="https://web9.vghtpe.gov.tw/"
    driver.get(root_url+note_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    table=soup.find("table")
    table_body=table.find('tbody')
    rows = table_body.find_all('tr')
    
    # b_note=rows[13:26]
    
    prog_note_list=[]
    progress_title=["Description","Subjective","Objective", "Assessment", "Plan"]

    row_idx=0
    
    while len(prog_note_list)<num:
        progress_note={}
        row=rows[row_idx].text
        if "Note" in row or "Summary" in row:
            progress_note["date"]=row
            row_idx=row_idx+1
            for title in progress_title:
                row=rows[row_idx].text
                if title in row:
                    row_idx=row_idx+1
                    progress_note[title]=rows[row_idx].pre.text
                row_idx=row_idx+1
            prog_note_list.append(progress_note)
        if row_idx<len(rows)-1:    
            row_idx=row_idx+1
        else:
            break
            
    return prog_note_list



#============================================
def get_my_patient(driver):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    header_element = soup.find(id="patlist")
    
    data = []
    table = soup.find(id="patlist")
    table_body = table.find('tbody')
    
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        one_col=[ele for ele in cols if ele]
        if "New" in one_col[1]:
            one_col[1]=one_col[1][4:]
        data.append(one_col) 
    return data

#==============================
# get recent report

def html_report_table(table):
    data=[]
    # table_head = table.find('thead')
    # t_head = table_head.find_all('th')
    # t_head = [ele.text for ele in t_head]
    
    
    table_body = table.find('tbody')

    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        # one_col=[ele for ele in cols if ele]
        # if "New" in one_col[1]:
        #     one_col[1]=one_col[1][4:]
        # print(cols)
        if not cols==['']:
            data.append(cols)
    df = pd.DataFrame(data)
    df=df.dropna()

    
    return df

def get_recent_report(driver, ID, report_num=3):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findRes&tdept=ALL&histno="+ID)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    reslist=soup.find(id="reslist")
    table_body=reslist.tbody
    rows = table_body.find_all('tr')
    root_url="https://web9.vghtpe.gov.tw/"
    
    report_name_list=[]
    fin_report={}
    for row in rows[:report_num]:
        report = row.find("a")
        # if not 
        Report_name=report.text
        print(Report_name)
        report_name_list.append(Report_name)
        # report_url=report["href"]
        # time.sleep(random.random()*2)
        # driver.get(root_url+report_url)
        
        
        # soup = BeautifulSoup(driver.page_source, 'html.parser')
        # report_res=soup.find(id="RSCONTENT")
        # table=report_res.find("table")
        # table=html_report_table(table)
        # fin_report[Report_name]=table
        fin_report=None
    return report_name_list, fin_report

# ============================================

def get_serarched_patient(driver,ward="0",patID="",docID=""):
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&wd="+ward+"&histno="+patID+"&pidno=&namec=&drid="+docID+"&er=0&bilqrta=0&bilqrtdt=&bildurdt=0&other=0&nametype=")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    data = []
    table = soup.find("table")
    table_head = table.find('thead')
    t_head = table_head.find_all('th')
    t_head = [ele.text for ele in t_head]
    
    table_body = table.find('tbody')
    
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if "(N)" in cols[2]:
            cols[2]=cols[2][4:].replace('\xa0', '')
        if not ward=="0":
            cols[1]=cols[1].split("[")[0]
        # one_col=[ele for ele in cols if ele]
        cols=cols[1:]
        data.append(cols) 
        # df = pd.DataFrame(data,columns=t_head)
    return data

# ================================================
# get Drainage (IO)
def html_IO_table(table):
    data=[]

    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    for idx,row in enumerate(rows):

        if row.find('td').text=="引流":
            # print(idx,row)
            drainage=row
            break
    
    try:
        drainage_table=drainage.find('table')
        drainage_table=drainage_table.find('tbody')
        drainage_rows = drainage_table.find_all('tr')

        drainage_data=[]
        for drainage_row in drainage_rows:
            cols = drainage_row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            drainage_data.append(cols)
        # print(drainage_data)
        df = pd.DataFrame(drainage_data,columns=["項目","白班","小夜","大夜","總量"])
    except:
        df=None

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


def get_drainage(driver, ID):
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