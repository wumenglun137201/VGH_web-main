import re
import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta
import random
import os
from urllib.parse import urlencode


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


def get_latest_nursing_page(vgh, ID, max_days_back=2):
    """
    走的流程：
    1. goNIS 建立 NIS session
    2. 打 main.jsp 設定查詢條件
    3. 用 POST 打 /NIS/daily/DaiLoopNb/tprReport.do (照 cURL)
    4. 回傳 tprReport.do 回來的 HTML/文字
    """
    adminID = get_adminID(vgh, ID)

    # 1) 先進 NIS
    go_nis_url = (
        "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?"
        f"action=goNIS&hisid={ID}&caseno={adminID}"
    )
    vgh.get_page_after_login(go_nis_url)

    last_page = None

    for delta in range(max_days_back + 1):
        date_str = (datetime.now() - timedelta(delta)).strftime('%Y%m%d')

        # 2) 先打 main.jsp，設定查詢條件到 session 裡
        main_params = {
            "fromDate": date_str,
            "toDate": date_str,
            "hisid": ID,
            "ser_numb": adminID,
            "inhdate": date_str,
            "disDate": "",
            "gaugeDate1": date_str,
            "timeCondition1": "0000",   # 你之後可以改成想要的起始時間
            "gaugeDate2": date_str,
            "timeCondition2": "2359",   # 結束時間
            "classType": "0",
            "reportType": "tpr",
            "range": "1",
            "action": "specify",
        }
        main_url = (
            "https://web9.vghtpe.gov.tw/NIS/report/RIR/main.jsp?"
            + urlencode(main_params)
        )
        main_html = vgh.get_page_after_login(main_url)
        print(f"[DEBUG] main.jsp URL: {main_url}")
        print(f"[DEBUG] main.jsp 長度: {len(main_html or '')}")
        if not main_html:
            continue

        # 3) 照 cURL 改：POST 到 /NIS/daily/DaiLoopNb/tprReport.do
        #    參數放在 query string，body 留空
        tpr_params = {
            "gaugeDate1": date_str,
            "timeCondition1": "0000",   # 如果要模仿 cURL 就填 '1030'
            "gaugeDate2": date_str,
            "timeCondition2": "2359",   # 如果要模仿 cURL 就填 '1350'
            "classType": "0",
            "reportType": "tpr",
        }
        tpr_url = (
            "https://web9.vghtpe.gov.tw/NIS/daily/DaiLoopNb/tprReport.do?"
            + urlencode(tpr_params)
        )

        headers = {
            "Accept": "text/plain, */*; q=0.01",
            "Origin": "https://web9.vghtpe.gov.tw",
            "Referer": main_url,
            "X-Requested-With": "XMLHttpRequest",
            "User-Agent": vgh.session.headers.get("User-Agent", ""),
        }

        # ⭐ 關鍵：用 session.post，params 放在 URL，不送 body（data=None）
        resp = vgh.session.post(tpr_url, headers=headers, data=None)
        tpr_html = resp.text if resp is not None else ""

        print(f"[DEBUG] tprReport.do URL: {tpr_url}")
        print(f"[DEBUG] tprReport.do status: {getattr(resp, 'status_code', 'N/A')}")
        print(f"[DEBUG] tprReport.do 長度: {len(tpr_html or '')}")

        if not tpr_html:
            continue

        # 粗略檢查一下是不是錯誤頁
        txt = BeautifulSoup(tpr_html, "html.parser").get_text().strip()
        print(f"[DEBUG] tprReport.do 純文字前 100 字：{txt[:100]}")

        if "Error" in txt[:50] or "錯誤" in txt[:50] or len(txt) < 200:
            # 太短或錯誤就往前一天找
            continue

        # 看起來像正常內容，就用這一頁
        last_page = tpr_html
        break

    return last_page


# ================================================
# 從 IORpt 畫面抽出文字

def extract_nursing_text_from_page(page_html):
    if not page_html:
        return ""
    soup = BeautifulSoup(page_html, "html.parser")
    # 有些報表會放在特定 div 裡，先嘗試找 id="divshow_0"
    div = soup.find(id="divshow_0")
    if not div:
        return soup.get_text("\n", strip=True)
    return div.get_text("\n", strip=True)


def _search(pattern, text, flags=re.IGNORECASE):
    """小工具：回傳第一個 regex group，沒抓到就回空字串"""
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


# ================================================
# 解析護理紀錄文字成 V / I / P / ECMO / F / S / N

def parse_nursing_text_to_vipfsn(text):
    """
    把護理紀錄文字 parse 成指定結構：
    V (Ventilation) / I (Infection) / P (Pressure)
    ECMO / F (Fluid) / S (Sedation) / N (Nutrition)
    同時保留 raw_text 方便你 debug
    """

    # 預設空框架，避免沒資料時還是有固定 key
    base_result = {
        "V": {"mode": "", "FiO2": "", "PC": "", "PEEP": ""},
        "I": {"abx": "", "culture": "", "infection_status": ""},
        "P": {
            "Perdipine_ml_hr": "",
            "NTG_ml_hr": "",
            "BP_goal": "",
            "Levophed_mcg_kg_min": "",
            "Epinephrine_mcg_kg_min": "",
            "Dopamine_mcg_kg_min": "",
            "Vasopressin_U_hr": "",
        },
        "ECMO": {
            "mode": "",
            "FiO2": "",
            "gas_flow": "",
            "rate": "",
            "VAD_ratio": "",
            "IABP_ratio": "",
        },
        "F": {
            "CVVH": "",
            "HD": "",
            "dry_weight": "",
            "diuretics": "",
            "IO_summary": "",
        },
        "S": {
            "DORMICUM": "",
            "PROPOFOL": "",
            "FENTANYL": "",
        },
        "N": {"route": ""},
        "raw_text": text or "",
    }

    if not text:
        return base_result

    # ------- V (Ventilation) -------
    V = {
        # 你實際護理紀錄上如果是「模式：SIMV」或「Ventilation Mode: SIMV」可以直接抓到
        "mode": _search(r"(?:Vent(?:ilator)? *Mode|模式)\s*[:：]\s*([A-Za-z0-9/+_-]+)", text),
        "FiO2": _search(r"FiO2\s*[:：]\s*([0-9]+%?)", text),
        "PC": _search(r"(?:PC|壓控)\s*[:：]\s*([0-9.]+)", text),
        "PEEP": _search(r"PEEP\s*[:：]\s*([0-9.]+)", text),
    }

    # ------- I (Infection) -------
    I = {
        # 例如「Abx: ceftriaxone, vancomycin」
        "abx": _search(r"(?:Abx|ABX|抗生素)\s*[:：]\s*([^\n\r]+)", text),
        # 例如「Culture: blood cx pending」
        "culture": _search(r"(?:Cx|Culture|培養)\s*[:：]\s*([^\n\r]+)", text),
        # 例如「Infection status: improving / fulminant sepsis」
        "infection_status": _search(r"(?:infection\s*status|感染狀態)\s*[:：]\s*([^\n\r]+)", text),
    }

    # ------- P (Pressure / pump) -------
    P = {
        "Perdipine_ml_hr": _search(r"(?:Perdipine|Nicardipine)\s*[:： ]\s*([0-9.]+)\s*ml/?hr", text),
        "NTG_ml_hr": _search(r"NTG\s*[:： ]\s*([0-9.]+)\s*ml/?hr", text),
        "BP_goal": _search(r"BP\s*goal\s*[:：]\s*([^\n\r]+)", text),
        "Levophed_mcg_kg_min": _search(
            r"(?:Levophed|Norepi(?:nephrine)?)\s*[:： ]\s*([0-9.]+)\s*mcg/kg/min", text
        ),
        "Epinephrine_mcg_kg_min": _search(
            r"Epi(?:nephrine)?\s*[:： ]\s*([0-9.]+)\s*mcg/kg/min", text
        ),
        "Dopamine_mcg_kg_min": _search(
            r"Dopamine\s*[:： ]\s*([0-9.]+)\s*mcg/kg/min", text
        ),
        "Vasopressin_U_hr": _search(
            r"Vasopressin\s*[:： ]\s*([0-9.]+)\s*U/hr", text
        ),
    }

    # ------- ECMO / VAD / IABP -------
    ECMO = {
        "mode": _search(r"ECMO\s*Mode\s*[:：]\s*([^\s,]+)", text),
        "FiO2": _search(r"ECMO\s*FiO2\s*[:：]\s*([0-9]+%?)", text),
        "gas_flow": _search(r"(?:Gas\s*flow|Sweep)\s*[:：]\s*([0-9.]+)", text),
        "rate": _search(r"ECMO\s*rate\s*[:：]\s*([0-9.]+)", text),
        "VAD_ratio": _search(r"VAD\s*ratio\s*[:：]\s*([0-9.:/]+)", text),
        "IABP_ratio": _search(r"IABP\s*ratio\s*[:：]\s*([0-9.:/]+)", text),
    }

    # ------- F (Fluid / 腎臟替代 / I/O) -------
    upper = text.upper()
    F = {
        "CVVH": "CVVH" if "CVVH" in upper else "",
        "HD": "HD" if " HD" in upper or "HEMODIALYSIS" in upper else "",
        "dry_weight": _search(r"(?:dry\s*weight|DW)\s*[:：]\s*([0-9.]+ *kg)", text),
        "diuretics": _search(r"(?:Diuretics|利尿劑)\s*[:：]\s*([^\n\r]+)", text),
        # 例如「I/O: +500 ml/day」
        "IO_summary": _search(r"(?:I/O|Input/Output)\s*[:：]\s*([^\n\r]+)", text),
    }

    # ------- S (Sedation) -------
    S = {
        # 例如「Dormicum: 2 mg/hr IV drip」或「Midazolam: 2 mg/hr」
        "DORMICUM": _search(r"(?:Dormicum|Midazolam)\s*[:： ]\s*([0-9.]+[^\n\r]*)", text),
        "PROPOFOL": _search(r"Propofol\s*[:： ]\s*([0-9.]+[^\n\r]*)", text),
        "FENTANYL": _search(r"Fentanyl\s*[:： ]\s*([0-9.]+[^\n\r]*)", text),
    }

    # ------- N (Nutrition) -------
    route = ""
    if re.search(r"\bNPO\b", text, re.IGNORECASE):
        route = "NPO"
    elif re.search(r"\bNG\b", text, re.IGNORECASE):
        route = "NG"
    elif re.search(r"\bNJ\b", text, re.IGNORECASE):
        route = "NJ"
    N = {"route": route}

    result = {
        "V": V,
        "I": I,
        "P": P,
        "ECMO": ECMO,
        "F": F,
        "S": S,
        "N": N,
        "raw_text": text,
    }
    return result


# ================================================
# 高階封裝：外面只要丟 vgh + 病歷號就好

def get_latest_nursing_data(vgh, ID):
    page = get_latest_nursing_page(vgh, ID)

    if not page:
        # 抓不到任何頁面，就回傳空殼
        return parse_nursing_text_to_vipfsn("")

    text = extract_nursing_text_from_page(page)

    # 這兩行是 debug，在 terminal 看得到
    print(f"[DEBUG] 抓到的護理紀錄文字長度：{len(text)}")
    print("[DEBUG] 護理紀錄前 300 字預覽：", text[:300].replace("\n", " "))

    return parse_nursing_text_to_vipfsn(text)


def get_latest_icu_note_page(vgh, ID, max_days_back=2):
    adminID = get_adminID(vgh, ID)

    # 先進 NIS
    url = (
        "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?"
        f"action=goNIS&hisid={ID}&caseno={adminID}"
    )
    vgh.get_page_after_login(url)

    last_page = None
    for delta in range(max_days_back + 1):
        date = (datetime.now() - timedelta(delta)).strftime('%Y%m%d')
        # TODO: 把下面這條 換成你在瀏覽器看到的「ICU 护理紀錄」URL
        icu_url = (
            "https://web9.vghtpe.gov.tw/NIS/你的ICU紀錄report.do?"
            f"gaugeDate1={date}"
        )
        page = vgh.get_page_after_login(icu_url)
        if page:
            last_page = page
            break

    return last_page

def debug_fetch_tpr_once(vgh):
    """
    專門用來測一次 ICU TPR 的 XHR，照 DevTools 那一筆 tprReport.do 抄過來。
    在命令列跑這支，先確定真的拿得到資料。
    """
    # 1. 先建立 NIS session（跟其他函式一樣）
    ID = "51509707"  # 測試用病歷號，之後再改成參數
    adminID = get_adminID(vgh, ID)

    go_nis_url = (
        "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?"
        f"action=goNIS&hisid={ID}&caseno={adminID}"
    )
    vgh.get_page_after_login(go_nis_url)

    # 2. 先打 main.jsp，照你現在的 URL 就好
    date = datetime.now().strftime('%Y%m%d')
    main_url = (
        "https://web9.vghtpe.gov.tw/NIS/report/RIR/main.jsp"
        f"?fromDate={date}"
        f"&toDate={date}"
        f"&hisid={ID}"
        f"&ser_numb={adminID}"
        f"&inhdate={date}"
        f"&gaugeDate1={date}&timeCondition1=0000"
        f"&gaugeDate2={date}&timeCondition2=2359"
        "&classType=0"
        "&reportType=tpr"
        "&range=1"
        "&action=specify"
    )
    vgh.get_page_after_login(main_url)

    # 3. ⚠️ 下面這一段請「照 DevTools 那一筆 tprReport.do 貼」
    # 假設 DevTools 顯示：
    #   Request URL: https://web9.../NIS/report/RIR/tprReport.do?AAA=...&BBB=...
    #   Method: POST
    #   Form data: CCC=...&DDD=...
    url = "https://web9.vghtpe.gov.tw/NIS/report/RIR/tprReport.do?AAA=...&BBB=..."  # ← 把這行改成 DevTools 的 Request URL

    headers = {
        # 至少要有 User-Agent，其他像 X-Requested-With、Origin、Referer
        # 看 DevTools 的 Request Headers 有哪些，就抄幾個非 Cookie 的。
        "User-Agent": vgh.session.headers.get("User-Agent", ""),
        "X-Requested-With": "XMLHttpRequest",   # 如果 DevTools 有這一行，就加
        "Origin": "https://web9.vghtpe.gov.tw", # 如果有就加
        "Referer": main_url,                    # 通常是 main.jsp 那頁
    }

    # 如果 Method 是 GET，下面這段 data 可以留空，改成 vgh.session.get(url, headers=headers)
    data = {
        # 這裡把 DevTools > Form Data 裡的欄位一個一個貼上來
        # 例如: "gaugeDate1": date, "timeCondition1": "1030", ...
        # "CCC": "...",
        # "DDD": "...",
    }

    # 依 DevTools 的 Method 選一個：
    # resp = vgh.session.get(url, headers=headers)          # 如果是 GET
    resp = vgh.session.post(url, headers=headers, data=data)  # 如果是 POST

    txt = resp.text
    print("[DEBUG] tprReport 回傳長度:", len(txt))
    print("[DEBUG] tprReport 前 300 字:", txt[:300].replace("\n", " "))

