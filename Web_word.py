#!/usr/bin/env python
# coding: utf-8

# In[1]:


# !conda install anaconda::pillow
# !conda install conda-forge::matplotlib
# !conda install anaconda::beautifulsoup4
# !conda install conda install conda-forge::selenium
# !conda install anaconda::pandas


# In[1]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from bs4 import BeautifulSoup
import time
import random

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from PIL import Image
from docx.oxml.ns import qn
import os
from Self_function import *
from datetime import datetime
import pwinput

# 配置 WebDriver
chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument("--headless=new")  # 如果不需要顯示瀏覽器界面，可以啟用 headless 模式
chrome_options.add_argument("--window-position=-2400,-2400")
chrome_options.add_argument('--log-level=3')

# chrome_options.add_argument("--no-sandbox")
# WINDOW_SIZE = "0,0"
# chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
# chrome_options.add_argument("screenshot")
# chrome_options.add_argument("--disable-dev-shm-usage")

# WebDriver 路徑
# webdriver_service = Service(r'C:\Users\reguser\Downloads\chrome-win64')  # 替換成你的 chromedriver 路徑
service = Service(executable_path=r'chromedriver.exe')
driver = webdriver.Chrome(service=service,options=chrome_options)


# In[2]:


username=input("帳號 : ")

password = pwinput.pwinput(prompt='密碼: ', mask='*')


# In[7]:


# 打開登入頁面
login_url = 'https://eip.vghtpe.gov.tw/login.php'  #
driver.get(login_url)

# 找到用戶名和密碼輸入框
username_field = driver.find_element(By.ID, 'login_name')  # 替換成實際的字段名稱
password_field = driver.find_element(By.ID, 'password')  # 替換成實際的字段名稱

# 輸入用戶名和密碼
username_field.send_keys(username)  # 替換成實際的用戶名
password_field.send_keys(password)  # 替換成實際的密碼

# 提交表單
password_field.send_keys(Keys.RETURN)

time.sleep(0.5)

driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=50687768")
soup = BeautifulSoup(driver.page_source, 'html.parser')
# print(soup)


# In[8]:


## Get my patient data

# def get_my_patient(driver):
#     driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&")
#     soup = BeautifulSoup(driver.page_source, 'html.parser')
#     header_element = soup.find(id="patlist")
    
#     data = []
#     table = soup.find(id="patlist")
#     table_body = table.find('tbody')
    
#     rows = table_body.find_all('tr')
#     for row in rows:
#         cols = row.find_all('td')
#         cols = [ele.text.strip() for ele in cols]
#         one_col=[ele for ele in cols if ele]
#         if "New" in one_col[1]:
#             one_col[1]=one_col[1][4:]
#         data.append(one_col) 
#     return data
# pat_data=get_my_patient(driver)
# print(pat_data)
# time.sleep(2)
docID=input("燈號(四碼)")
pat_data=get_serarched_patient(driver,ward="0",patID="",docID=docID)


# In[5]:


def set_paragraph_spacing(doc, spacing=0):
    """Set paragraph spacing for all paragraphs in the document."""
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.line_spacing = Pt(spacing)
        paragraph.paragraph_format.space_before = Pt(spacing)
        paragraph.paragraph_format.space_after = Pt(spacing)

def set_font_size(doc, size):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)

def add_table(doc, df):
    table = doc.add_table(rows=1, cols=len(df.columns))
    
    # 設置表頭的字體大小
    hdr_cells = table.rows[0].cells
    for i, column_name in enumerate(df.columns):
        hdr_cells[i].text = str(column_name)
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
            paragraph.paragraph_format.line_spacing = Pt(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
        
    
    # 添加數據行
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(8)
                paragraph.paragraph_format.line_spacing = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                
    for col in table.columns:
        max_length = max(len(cell.text) for cell in col.cells)
        # You can adjust the multiplier for a better fit
        col_width = Inches(max_length)  # Adjust factor as needed
        for cell in col.cells:
            cell.width = col_width

def convert_date(date_str):
    date_str=date_str[3:8]
    return date_str

def generate_report(driver,doc, ID):
    print(ID)
    
    # BW_BL=get_BW_BL(driver,ID)
    # time.sleep(0.1)
    
    
    # get_last_admission(driver,ID)
    
    # get_res_report(driver,ID)
    
    # reportname, report=get_recent_report(driver, ID, report_num=5)

    # 添加標題
    
    # 設置最小行高
    try:
        TPR=get_TPR(driver,ID)
        time.sleep(3*random.random())
        doc.add_paragraph(" |".join(list(TPR[["體溫","心跳","呼吸","收縮壓","舒張壓"]].iloc[0])))
    except:
        pass
    
    # 添加圖片
    try:
        TPR_img=get_TPR_img(driver,ID)
        time.sleep(3*random.random())
        image_path = 'temp_image.png'
        TPR_img.save(image_path)
        doc.add_picture(image_path, width=Inches(3))  # 插入圖片
        os.remove(image_path)
    except:
        pass

    try:
        progress_note=get_progress_note(driver,ID,num=5)
        time.sleep(3*random.random())

        for i in range(len(progress_note)):
            assessment=progress_note[i]["Assessment"]
            if len(assessment)>5:
                break
        doc.add_paragraph(assessment).paragraph_format.line_spacing = Pt(0)  # 
    except:
        pass

    
    doc.add_paragraph("-----------------------------------------------------------")
    # add_table(doc, BW_BL[["身高","體重"]] )
    # doc.add_paragraph("-----------------------------------------------------------")
    try:
        report_num=3
        report_name,recent_report=get_recent_report(driver, ID, report_num=report_num)
        time.sleep(3*random.random())
        for i in range(report_num):
            doc.add_paragraph(report_name[i])
            add_table(doc, recent_report[report_name[i]])
            doc.add_paragraph("-----------------------------------------------------------")
    except:
        pass

    try:
        SMAC=get_res_report(driver,ID,resdtype="SMAC")
        SMAC["日期"]=SMAC["日期"].apply(convert_date)
        time.sleep(3*random.random())
        add_table(doc, SMAC[["日期","NA","K","BUN","CREA","ALT","BILIT","CRP"]] )
        doc.add_paragraph("-----------------------------------------------------------")
    except:
        pass

    try:
        CBC=get_res_report(driver,ID,resdtype="CBC")
        time.sleep(3*random.random())
        CBC["日期"]=CBC["日期"].apply(convert_date)
        add_table(doc, CBC[["日期","WBC","HGB","PLT",'BAND', 'SEG', 'LYM']] )
        doc.add_paragraph("-----------------------------------------------------------")
    except:
        pass

    
    # add_table(doc, report[reportname[3]] )
    # doc.add_paragraph("-----------------------------------------------------------")
    try:
        drug=get_drug(driver,ID)
        time.sleep(3*random.random())
        add_table(doc, drug[drug["狀態"]=="使用中"][["學名","劑量","途徑","頻次","開始日"] ])
    except:
        pass
    doc.add_paragraph("=================================================")


doc = Document()
# set two column
section = doc.sections[0]
sectPr = section._sectPr
cols = sectPr.xpath('./w:cols')[0]
cols.set(qn('w:num'),'2')

# 設定邊界
section.top_margin = Pt(30)   # 0.5 inch
section.bottom_margin = Pt(30) # 0.5 inch
section.left_margin = Pt(30)   # 0.5 inch
section.right_margin = Pt(30)  # 0.5 inch

header = section.header
paragraph=header.paragraphs[0]
run = paragraph.add_run("日期:"+datetime.now().strftime('%Y-%m-%d')+" 醫師: "+docID)
run.font.size = Pt(10)


for pat in pat_data:
    paragraph =doc.add_paragraph()
    run = paragraph.add_run('/'.join(pat))
    run.bold = True
    run.underline = True
    ID=pat[1]
    generate_report(driver=driver,doc=doc,ID=ID)

# 設置所有文本字體為 8 號
set_font_size(doc, 8)
set_paragraph_spacing(doc, spacing=0)

# 保存 Word 文件
doc.save(docID+'.docx')
print("儲存為"+docID+'.docx')

driver.quit()

