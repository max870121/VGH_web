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
from docx.enum.section import WD_ORIENT
from PIL import Image
from docx.oxml.ns import qn
import os
from Self_function import *
from datetime import datetime, timedelta
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

docID=input("燈號(四碼)")
ward=input("病房(Ex A101)")
pat_data=get_serarched_patient(driver,ward=ward,patID="",docID=docID)


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
                run.font.size = Pt(6)
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
                    run.font.size = Pt(6)
                paragraph.paragraph_format.line_spacing = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                
    for col in table.columns:
        max_length = max(len(cell.text) for cell in col.cells)
        col_width = Inches(max_length)
        for cell in col.cells:
            cell.width = col_width

def convert_date(date_str):
    date_str=date_str[3:8]
    return date_str

def add_line(doc):
    doc.add_paragraph("-------------------------------")

def convert_drug(data_drug):
    data_drug=data_drug.split(" ")[:2]
    data_drug=" ".join(data_drug)
    return data_drug


def generate_table_report(driver,doc, ID, row_cells,pat):
    print(ID)
    
    info_cell=row_cells[0]
    paragraph = info_cell.paragraphs[0]
    paragraph.add_run("\n".join(pat))

    try:
        TPR=get_TPR(driver,ID)
        time.sleep(3*random.random())
        run=paragraph.add_run("\n")
        paragraph.add_run("\n".join(list(TPR[["體溫","心跳","呼吸","收縮壓","舒張壓"]].iloc[0])))
        
    except:
        pass
    
    try:
        run=paragraph.add_run()
        TPR_img=get_TPR_img(driver,ID)
        time.sleep(3*random.random())
        image_path = 'temp_image.png'
        TPR_img.save(image_path)
        run.add_picture(image_path, width=Inches(1))  # 插入圖片
        os.remove(image_path)
    except:
        pass

    try:
        BW_BL=get_BW_BL(driver,ID, adminID="all")
        BW_BL=BW_BL[["身高","體重"]]
        add_table(info_cell, BW_BL.head(2) )
    except:
        pass
 
    assessment_cell=row_cells[1]
    paragraph = assessment_cell.paragraphs[0]
    try:
        progress_note=get_progress_note(driver,ID,num=5)
        time.sleep(3*random.random())

        for i in range(len(progress_note)):
            assessment=progress_note[i]["Assessment"]
            if "Ditto" in assessment:
                continue
            else:
                break

        paragraph.add_run(assessment)
    except:
        pass

    Lab_cells = row_cells[2]

    try:
        patIO=get_drainage(driver, ID)
        add_table(Lab_cells,patIO[["項目","總量"]])
        # add_line(Lab_cells)
    except:
        pass
    

    try:
        report_num=3
        report_name,recent_report=get_recent_report(driver, ID, report_num=report_num)
        time.sleep(3*random.random())
        for i in range(report_num):
            Lab_cells.add_paragraph(report_name[i])
            # add_table(doc, recent_report[report_name[i]])
    except:
        pass


    try:
        SMAC=get_res_report(driver,ID,resdtype="SMAC")
        SMAC["日期"]=SMAC["日期"].apply(convert_date)
        SMAC=SMAC[["日期","NA","K","BUN","CREA","ALT","BILIT","CRP"]]
        SMAC = SMAC.loc[~(SMAC[["日期","NA","K","BUN","CREA","ALT","BILIT","CRP"]] == '-').all(axis=1)]
        time.sleep(3*random.random())
        add_table(Lab_cells, SMAC.tail(3) )
    except:
        pass

    try:
        CBC=get_res_report(driver,ID,resdtype="CBC")
        time.sleep(3*random.random())
        CBC["日期"]=CBC["日期"].apply(convert_date)
        CBC=CBC[["日期","WBC","HGB","PLT",'SEG', 'PT', 'APTT']]
        CBC = CBC.loc[~(CBC[["日期","WBC","HGB","PLT",'SEG', 'PT', 'APTT']] == '-').all(axis=1)]
        add_table(Lab_cells, CBC.tail(3) )
        # add_line(Lab_cells)
    except:
        pass

    try:

        def convert_drug(data_drug):
            data_drug=data_drug.split(" ")[:2]
            data_drug=" ".join(data_drug)
            return data_drug
        def convert_drug_date(data_drug_date):
            data_drug_date=data_drug_date[5:10]
            return data_drug_date
        drug=get_drug(driver,ID)
        drug["學名"]=drug["學名"].apply(convert_drug)
        drug["開始日"]=drug["開始日"].apply(convert_drug_date)
        time.sleep(3*random.random())
        add_table(Lab_cells, drug[drug["狀態"]=="使用中"][["學名","劑量","途徑","頻次","開始日"] ])
    except:
        pass

doc = Document()



section = doc.sections[0]
new_width, new_height = section.page_height, section.page_width
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width=new_width
section.page_height=new_height

# 設定邊界
section.top_margin = Pt(30)   # 0.5 inch
section.bottom_margin = Pt(30) # 0.5 inch
section.left_margin = Pt(30)   # 0.5 inch
section.right_margin = Pt(30)  # 0.5 inch

header = section.header
paragraph=header.paragraphs[0]
run = paragraph.add_run("日期:"+datetime.now().strftime('%Y-%m-%d')+" 醫師: "+docID)
run.font.size = Pt(6)

table = doc.add_table(rows=1, cols=3)
table.style = 'Table Grid'

hdr_cells = table.rows[0].cells
hdr_cells[0].text = '病人資料'
hdr_cells[1].text = 'Assessment'
hdr_cells[2].text = 'Lab Data+drug'
for cell in hdr_cells:
    set_font_size(cell, 6)


for pat in pat_data:
    row_cells = table.add_row().cells
    if len(pat)<3:
        continue
    if docID=="":
        ID=pat[2]
    else:
        ID=pat[1]
    generate_table_report(driver=driver,doc=doc, ID=ID, row_cells=row_cells,pat=pat)
    for cell in row_cells:
        set_font_size(cell, 6)
    input("Wait a while and press enter")

for idx,col in enumerate(table.columns):
    max_length = max(len(cell.text) for cell in col.cells)
    col_width = Inches(max_length)
    if idx==2:
        col_width = Inches(max_length*0.8)
    for cell in col.cells:
        cell.width = col_width


# 設置所有文本字體為 6 號
set_font_size(doc, 6)
set_paragraph_spacing(doc, spacing=0)

# 保存 Word 文件
filename=datetime.now().strftime('%Y%m%d')+"_"+docID+"_"+"patient_list"+'.docx'
doc.save(filename)
print("儲存為"+filename)

driver.quit()

