import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta
import random
import os

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
        data.append(one_col) # Get rid of empty values
    df = pd.DataFrame(data,columns=t_head)
    
    return df

#======================================
# Get TPR
def get_adminID(vgh, ID):
    url="https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno="+ID
    page_content = vgh.get_page_after_login(url)
    TPR_url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPbv&histno=" + ID
    page_content = vgh.get_page_after_login(TPR_url)
    soup = BeautifulSoup(page_content, 'html.parser')
    adminID = soup.option['value'].split("=")[-1]
    return adminID

def get_TPR(vgh, ID, adminID=None):
    if not adminID:
        adminID = get_adminID(vgh, ID)
    
    TPR_url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findTpr&caseno=" + adminID
    page_content = vgh.get_page_after_login(TPR_url)
    soup = BeautifulSoup(page_content, 'html.parser')
    soup.find(id="tprlist")
    data = html_table(soup)
    
    return data

#==========================================================
## Get TPR image (Note: Image capture functionality needs to be handled differently)
def get_TPR_img(vgh, ID, adminID=None):
    """
    Note: Image capture functionality cannot be directly converted.
    You'll need to implement screenshot capability in your vgh module
    or use a different approach for capturing TPR images.
    """
    if not adminID:
        adminID = get_adminID(vgh, ID)
    
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findTpr&caseno=" + adminID + "&pbvtype=tpr"
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    root_url = "https://web9.vghtpe.gov.tw"
    img_tag = soup.find('img')  # You can also use soup.find_all('img') for multiple images
    
    if img_tag and img_tag.get('src'):
        # Construct the full image URL (handle relative URLs)
        img_url = img_tag['src']
        img_url = root_url+ img_url
        # Fetch the image
        img_response = vgh.get_img_after_login(img_url)
        
        # Check if the image request was successful
        if img_response.status_code == 200:
            # Save the image to a local file
            with open("downloaded_image.jpg", "wb") as file:
                file.write(img_response.content)
            # print("Image downloaded successfully!")
        else:
            print(f"Failed to retrieve image. Status code: {img_response.status_code}")
    # return vgh.get_screenshot(url)
    # raise NotImplementedError("Image capture needs to be implemented in vgh module")

# =======================================================================
## Get BW_BL
def get_BW_BL(vgh, ID, adminID="all"):
    if not adminID:
        adminID = get_adminID(vgh, ID)
    
    BW_BL_url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findVts&histno=" + ID + "&caseno=" + adminID + "&pbvtype=HWS"
    page_content = vgh.get_page_after_login(BW_BL_url)
    soup = BeautifulSoup(page_content, 'html.parser')
    data = html_table(soup)
    
    return data

#==================================================================
## Get Lab value
def get_Lab_value(vgh, ID, Lab_value):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=DCHEM&histno=" + ID + "&resdtmonth=24"
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    header_element = soup.find(id=Lab_value)
    time_list = header_element.text.split('|')
    Lab_data = []
    for one_time in time_list:
        Lab_data.append(one_time.split("/"))
    return Lab_data

#=================================================================
## get latest admission note
def get_last_admission(vgh, ID):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findAdm&histno=" + ID
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    admnote = soup.find(title="admnote")
    root_url = "https://web9.vghtpe.gov.tw/"
    admin_url = root_url + admnote['href']
    time.sleep(0.5)
    
    page_content = vgh.get_page_after_login(admin_url)
    soup = BeautifulSoup(page_content, 'html.parser')
    return soup.pre

# =====================================================
## get current drug
def get_drug(vgh, ID):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findUd&histno=" + ID
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    drug_url_list = soup.find_all("a")
    adminID = get_adminID(vgh, ID)
    drug_url = drug_url_list[0]["href"]
    for a_url in drug_url_list:
        if adminID in a_url["href"]:
            drug_url = a_url["href"]

    root_url = "https://web9.vghtpe.gov.tw/"
    page_content = vgh.get_page_after_login(root_url + drug_url)
    soup = BeautifulSoup(page_content, 'html.parser')
    table = soup.find(id="udorder")
    drug_table = html_table(table)
    return drug_table

#=========================================
# split the html table
## get res report
def html_res_table(table):
    data = []
    table_head = table.find('thead')
    t_head = table_head.find_all('th')
    t_head = [ele.text for ele in t_head]

    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    for row in rows[:-1]:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append(cols)
    df = pd.DataFrame(data, columns=t_head)
    return df

def get_res_report(vgh, ID, resdtype="SMAC", resdtmonth="00"):
    report_dict = {
        "SMAC": "DCHEM",
        "CBC": "DCBC",
        "Urine": "DURIN",
        "Cancer": "DNM1",
    }
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=" + report_dict[resdtype] + "&histno=" + ID + "&resdtmonth=" + resdtmonth
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    table = soup.find(id="resdtable")
    report_table = html_res_table(table)
    return report_table  

#=================
## get_progress_note
def get_progress_note(vgh, ID, num=5):
    adminID = get_adminID(vgh, ID)
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPrg&histno=" + ID + "&caseno=" + adminID
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    note_url = soup.find("a")["href"]
    root_url = "https://web9.vghtpe.gov.tw/"
    
    page_content = vgh.get_page_after_login(root_url + note_url)
    soup = BeautifulSoup(page_content, 'html.parser')

    table = soup.find("table")
    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    
    prog_note_list = []
    progress_title = {"病情描述(Description):":"Description", "主觀資料(Subjective):":"Subjective", "客觀資料(Objective):":"Objective", "診斷(Assessment):":"Assessment", "治療計畫(Plan):":"Plan"}

    row_idx = 0
    
    while len(prog_note_list) < num:
        progress_note = {}
        row = rows[row_idx].text
        if "Note" in row or "Summary" in row:
            progress_note["date"] = row
            row_idx = row_idx + 1
            
            for title in progress_title.keys():
                row = rows[row_idx].text
                while not title in row:

                    if row_idx > len(rows) - 2:
                        break
                    row_idx = row_idx + 1
                    row = rows[row_idx].text
                else:
                    row_idx = row_idx + 1
                    progress_note[progress_title[title]] = rows[row_idx].pre.text

            prog_note_list.append(progress_note)
        if row_idx < len(rows) - 1:    
            row_idx = row_idx + 1
        else:
            break
            
    return prog_note_list

#============================================
def get_my_patient(vgh):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&"
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    header_element = soup.find(id="patlist")
    
    data = []
    table = soup.find(id="patlist")
    table_body = table.find('tbody')
    
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        one_col = [ele for ele in cols if ele]
        if "New" in one_col[1]:
            one_col[1] = one_col[1][4:]
        data.append(one_col) 
    return data

#==============================
# get recent report
def html_report_table(table):
    data = []
    table_body = table.find('tbody')

    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if not cols == ['']:
            data.append(cols)
    df = pd.DataFrame(data)
    df = df.dropna()
    
    return df

def get_recent_report(vgh, ID, report_num=3):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findRes&tdept=ALL&histno=" + ID
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    reslist = soup.find(id="reslist")
    table_body = reslist.tbody
    rows = table_body.find_all('tr')
    root_url = "https://web9.vghtpe.gov.tw/"
    
    report_name_list = []
    fin_report = {}
    for row in rows[:report_num]:
        report = row.find("a")
        Report_name = report.text
        print(Report_name)
        report_name_list.append(Report_name)
        # Note: If you need to fetch individual reports, uncomment and modify below
        # report_url = report["href"]
        # time.sleep(random.random()*2)
        # page_content = vgh.get_page_after_login(root_url + report_url)
        # soup = BeautifulSoup(page_content, 'html.parser')
        # report_res = soup.find(id="RSCONTENT")
        # table = report_res.find("table")
        # table = html_report_table(table)
        # fin_report[Report_name] = table
        fin_report = None
    return report_name_list, fin_report

# ============================================
def get_searched_patient(vgh, ward="0", patID="", docID=""):
    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&wd=" + ward + "&histno=" + patID + "&pidno=&namec=&drid=" + docID + "&er=0&bilqrta=0&bilqrtdt=&bildurdt=0&other=0&nametype="
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
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
            cols[2] = cols[2][4:].replace('\xa0', '')
        if not ward == "0":
            cols[1] = cols[1].split("[")[0]
        cols = cols[1:]
        data.append(cols) 
    return data

# ================================================
# get Drainage (IO)
def html_IO_table(table):
    data = []

    # table_body = table.find('tbody')
    rows = table.find_all('tr')
    for idx, row in enumerate(rows):
        if row.find('td').text == "引流":
            drainage = row
            break
    
    try:
        drainage_table = drainage.find('table')
        # drainage_table = drainage_table.find('tbody')
        drainage_rows = drainage_table.find_all('tr')

        drainage_data = []
        for drainage_row in drainage_rows:
            cols = drainage_row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            drainage_data.append(cols)
        df = pd.DataFrame(drainage_data, columns=["項目", "白班", "小夜", "大夜", "總量"])
    except:
        df = None

    return df

def get_drainage(vgh, ID):
    adminID = get_adminID(vgh, ID)

    url = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=goNIS&hisid=" + ID + "&caseno=" + adminID
    page_content = vgh.get_page_after_login(url)
    
    date = (datetime.now() - timedelta(1)).strftime('%Y%m%d')
    url = "https://web9.vghtpe.gov.tw/NIS/report/IORpt/details.do?gaugeDate1=" + date
    page_content = vgh.get_page_after_login(url)
    soup = BeautifulSoup(page_content, 'html.parser')
    soup = soup.find(id="divshow_0")
    IOtable = soup.table.table.findAll('table')[1]
    # breakpoint()
    df = html_IO_table(IOtable)
    return df