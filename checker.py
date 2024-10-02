import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import openpyxl
import argparse
from openpyxl.worksheet.hyperlink import Hyperlink
import requests
from bs4 import BeautifulSoup
import itertools
import numpy as np
import json
from datetime import datetime
from selenium.webdriver import ChromeOptions
from seleniumrequests import Chrome
import logging
import re
import os
import urllib.parse
from functools import partial
import yaml
try:
    from yaml import CLoader as Loader, CDumper as Dumper
except ImportError:
    from yaml import Loader, Dumper
    
    
# Create a custom logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # Set the overall logging level

# Create handlers
console_handler = logging.StreamHandler()  # Handler for console output
file_handler = logging.FileHandler("debug.log")  # Handler for file output

# Set logging levels for handlers
console_handler.setLevel(logging.INFO)  # Only INFO and above for console
file_handler.setLevel(logging.DEBUG)  # DEBUG and above for file

# Create formatters and add it to handlers
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
console_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)

# Add handlers to the logger
logger.addHandler(console_handler)
logger.addHandler(file_handler)

session = requests.Session()
session.headers.update(
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",  # The version of requests will vary
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Connection": "keep-alive",
        "Referer": "https://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/SearchBSALicenseeContent.aspx",
        "Accept-Language": "en-US,en;q=0.8",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
    }
)

def enum_rows(sheet):

    headers = list()

    for r in sheet.rows:
        values = [f"{c.value}".strip() for c in r]

        if not headers:
            headers = [v.lower() for v in values]
            continue

        item = dict()
        for h, c in zip(headers, r):
            item[h] = f"{c.value}".strip() if c else ""

        yield r, item

def parse_qbcc_response(html):
    soup = BeautifulSoup(html, "html.parser")

    business_name_element = soup.select_one(
        "#ctl00_generalContentPlaceHolder_LicenceInfoControl1_lbLicenceName"
    )
    trading_name_element = soup.select_one(
        "ctl00_generalContentPlaceHolder_LicenceInfoControl1_lbTradingName"
    )

    items = [
        td.text.strip("\r\n\t ")
        for td in soup.select(
            "table[id='ctl00_generalContentPlaceHolder_LicenceInfoControl1_gvLicenceClass'] td"
        )
    ]

    items_array = np.array(items)

    if len(items_array) > 0 and (len(items_array) % 4) == 0:
        for part in np.split(items_array, len(items_array) / 4):
            yield [p.strip() for p in part]

def parse_surveyor_response(html_content):
    # Parse the HTML using BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    element = soup.select_one(".search-results")
    if not element:
        return

    # Extract name
    name = element.find('h4').get_text(strip=False).split('<br>')[0]
    lic_class = element.find('h4').get_text(strip=False).split('<br>')[-1]
    # Extract phone number
    phone_element = element.find('td', string='Phone ')
    phone = ''
    if phone_element:        
        phone = phone_element.find_next_sibling('td').get_text()

    # Extract email
    email_element = element.find('td', string='Email ')
    email = ''
    if email_element:
        email = email_element.find_next_sibling('td').get_text()
        
    # Extract address
    address_element = element.find('td', string='Address')
    address = ''
    if address_element:
        address = address_element.find_next_sibling('td').get_text(separator=', ')
    
    types = []
    # Extract types
    types = [span.get_text() for span in element.find_all('div', class_='types')[0].find_all('span')]

    return {
        "name" : re.sub(r"\s+"," ",name),
        "license_class" :  re.sub(r"\s+"," ",lic_class),
        "phone" : phone,
        "email" : email,
        "address" : address,
        "expertise" : '; '.join(types),
    }

def query_surveyor_license(search_text):
    search_text = re.sub(r"\s+", " ", f"{search_text}".strip())
    logger.info(f"Looking up surveyor info: {search_text}")
    url = "https://sbq.com.au/find-a-surveyor/search-cadastral/"
    session.get(url)

    params = {
        "surveyor-type": "individual",
        "search-type": "name",
        "title": search_text,
        "postcode": "",
        "radius": "5",
        "q": "Search",
    }

    response = session.get(url, params=params)
    if response.status_code!=200:
        return
    
    return parse_surveyor_response(response.text)


def query_qbcc_license(license_no):
    
    session.get('https://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/SearchBSALicenseeContent.aspx')

    license_no = license_no.strip("\r\n\t ")
    url = f"http://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/ShowDetailResultContent.aspx"
    params = {
        "LicNO": f"{license_no}",
        "licCat": "LIC",
        "name": "",
        "firstName": "",
        "searchType": "Contractor",
        "FromPage": "SearchContr",
    }
    response = requests.get(url, params=params)
    for r in parse_qbcc_response(response.text):
        yield r


def query_qbcc_certifier_license(license_no):
    
    session.get('https://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/SearchBuildingCertifierContent.aspx')

    license_no = license_no.strip("\r\n\t ")
    url = f"https://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/ShowDetailResultContent.aspx"
    params = {
        "LicNO": f"{license_no}",
        "licCat": "LIC",
        "name": "",
        "firstName": "",
        "searchType": "Certifier",
        "FromPage": "SearchContr",
    }
        
    response = session.get(url, params=params)
    for r in parse_qbcc_response(response.text):
        yield r


def query_engr_registration(registration_no, driver: Chrome):
    try:
        url1 = "https://portal.bpeq.qld.gov.au/BPEQPortal/Search_for_a_RPEQ/BPEQPortal/Engineer_Search.aspx"
        driver.get(url1)
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Sheet0_Input3_TextBox1",
                )
            )
        )
        element.send_keys(registration_no)
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Sheet0_SubmitButton",
                )
            )
        )
        element.click()
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Grid1_ctl00__0"]/td[1]/a',
                )
            )
        )
        registration_no = element.get_attribute("href").split("=")[1]

        url = f"https://portal.bpeq.qld.gov.au/Party.aspx?ID={registration_no}"
        response = driver.request("GET", url, verify=False)
        soup = BeautifulSoup(response.text, "html.parser")
        parts = [
            p.text.strip("\r\n\t ") for p in soup.select(".PanelFieldValue > span")
        ]
        logger.info(f"Num Parts: {len(parts)}")

        if len(parts) == 12:
            name = parts[0]
            company = parts[1]
            date_joined = parts[2]
            job_type = parts[3]
            status = parts[4]
            date_registered = parts[5]

            return name, company, date_joined, job_type, status, date_registered
        elif len(parts) == 11:
            name = parts[0]
            date_joined = parts[1]
            job_type = parts[2]
            status = parts[3]
            date_registered = parts[4]

            return name, None, date_joined, job_type, status, date_registered
        elif len(parts) == 16:
            name = parts[0]
            company = parts[1]
            date_joined = parts[2]
            job_type = parts[3]
            _ = parts[4]
            status = parts[5]
            date_registered = parts[6]

            return name, company, date_joined, job_type, status, date_registered
        elif len(parts) == 15:
            name = parts[0]
            date_joined = parts[1]
            job_type = parts[2]
            _ = parts[3]
            status = parts[4]
            date_registered = parts[5]

            return name, None, date_joined, job_type, status, date_registered

    except Exception as e:
        logger.info(e)


def query_arch_registration(registration_no, driver: Chrome):
    try:
        url1 = "https://www.boaq.qld.gov.au/BOAQ/Search_Register/BOAQ/Search_Register/Architect_Search.aspx"
        driver.get(url1)
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Sheet0_Input3_TextBox1",
                )
            )
        )
        element.send_keys(registration_no)
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.NAME,
                    "ctl01$TemplateBody$WebPartManager1$gwpciArchitectsearch$ciArchitectsearch$ResultsGrid$Sheet0$SubmitButton",
                )
            )
        )
        element.click()
        element = WebDriverWait(driver, 16).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Grid1_ctl00__0"]/td[1]/a',
                )
            )
        )
        registration_no = element.get_attribute("href").split("=")[1]
        url = f"https://www.boaq.qld.gov.au/Shared_Content/ContactManagement/Profile.aspx?ID={registration_no}"
        response = driver.request("GET", url)
        soup = BeautifulSoup(response.text, "html.parser")
        parts = [
            p.text.strip("\r\n\t ") for p in soup.select(".PanelFieldValue > span")
        ]
        logger.info(f"Num Parts: {len(parts)}")

        if len(parts) == 12:
            name = parts[0]
            company = parts[1]
            date_joined = parts[2]
            job_type = parts[3]
            status = parts[4]
            date_registered = parts[5]

            return name, company, date_joined, job_type, status, date_registered
        elif len(parts) == 11:
            name = parts[0]
            date_joined = parts[1]
            job_type = parts[2]
            status = parts[3]
            date_registered = parts[4]

            return name, None, date_joined, job_type, status, date_registered

    except Exception as e:
        logger.info(e)


def process_sheet_arch(wb, sheetname, args, config,sheet_config):
    
    orig_filename = args.input
    
    if not "architects" in sheetname.lower() or not args.arch:
        return

    logger.info("Processing Architects Tab...")
    sheet = wb[sheetname]

    config = read_config()
    count = 0

    options = ChromeOptions()
    # options.add_argument("headless")
    driver = Chrome(options=options)

    for idx, (row, data) in enumerate(enum_rows(sheet)):

        if count > 0 and (count % config["numrec_before_save"]) == 0:
            logger.info(
                "===============================================\n \
                Saving progress to excel file...\n=================================="
            )
            wb.save(orig_filename)

        if not "Registration" in data:
            logger.info("Registration No. Column not found !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "No Registration Number Column found!"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        if not data["Registration"] or data["Registration"].strip() == "":
            logger.info("Registration No. is BLANK in the excel file !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "Registration No. is BLANK !"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        registration_no = data["Registration"]
        logger.info(f"Feetching Registration info of {registration_no}:")
        reg_status = query_arch_registration(registration_no, driver)

        if reg_status:
            logger.info("Registration info found!")

            name, company, date_joined, job_type, status, date_registered = reg_status

            # lic_class, _, _, lic_status = lic_statuses[0]
            logger.info(f"\tName: {name}")
            logger.info(f"\tCompany: {company}")
            logger.info(f"\tDate Joined: {date_joined}")
            logger.info(f"\tType: {job_type}")
            logger.info(f"\tStatus: {status}")
            logger.info(f"\tDate Registered: {date_registered}")

            row[config["sheets_config"][sheetname]["status_index"]].value = (
                status.strip().title()
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
        else:
            logger.info("Registration info not found in online register !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "Missing in Register"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )

        count = count + 1


def process_sheet_engr(wb, sheetname, orig_filename, config, sheet_config):
    orig_filename = args.input
        
    if not all(keyword in sheetname.lower() for keyword in ["engineers"]) or not args.engr:
        return

    logger.info("Processing Engineers Tab...")
    sheet = wb[sheetname]

    config = read_config()
    count = 0

    options = ChromeOptions()
    # options.add_argument("headless")
    driver = Chrome(options=options)

    for row, data in enum_rows(sheet):

        if count > 0 and (count % config["numrec_before_save"]) == 0:
            logger.info(
                "===============================================\n \
                Saving progress to excel file...\n=================================="
            )
            wb.save(orig_filename)

        if not "Registration" in data:
            logger.info("Registration No. Column not found !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "No Registration Number Column found!"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        if not data["Registration"] or data["Registration"].strip() == "":
            logger.info("Registration No. is BLANK in the excel file !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "Registration No. is BLANK !"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        registration_no = data["Registration"]
        logger.info(f"Feetching Registration info of {registration_no}:")
        reg_status = query_engr_registration(registration_no, driver)

        if reg_status:
            logger.info("Registration info found!")

            name, company, date_joined, job_type, status, date_registered = reg_status

            # lic_class, _, _, lic_status = lic_statuses[0]
            logger.info(f"\tName: {name}")
            logger.info(f"\tCompany: {company}")
            logger.info(f"\tDate Joined: {date_joined}")
            logger.info(f"\tType: {job_type}")
            logger.info(f"\tStatus: {status}")
            logger.info(f"\tDate Registered: {date_registered}")

            row[config["sheets_config"][sheetname]["status_index"]].value = (
                status.strip().title()
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )
        else:
            logger.info("Registration info not found in online register !")
            row[config["sheets_config"][sheetname]["status_index"]].value = (
                "Missing in Register"
            )
            row[config["sheets_config"][sheetname]["last_checked_index"]].value = (
                datetime.now().date()
            )

        count = count + 1
        
def update_license_status(row, status, sheet_config):
    """Helper function to update status and last checked date for a row."""
    row[sheet_config["status_index"]].value = status
    row[sheet_config["last_checked_index"]].value = datetime.now().date()
    
def handle_surveyor_license_query(row, search_text, sheet_config):
    """Query surveyor's license and update the status in the row."""
    result = query_surveyor_license(search_text)
    if result:
        update_license_status(row, "Active", sheet_config)
        return True
    else:
        update_license_status(row, "License Not Found", sheet_config)
        return False

def process_sheet_surveyor(wb, sheetname, args, config, sheet_config):
    orig_filename = args.input
    
    if not "surveyor" in sheetname.lower() or not args.surv:
        return
    
    logger.info("Processing Surveyor Tab: {}...".format(sheetname))
    session.headers.clear()
    session.headers.update({
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'accept-language': 'en-US,en;q=0.7',
        'cache-control': 'no-cache',
        'pragma': 'no-cache',
        'referer': 'https://sbq.com.au/find-a-surveyor/search-cadastral/',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    })
    
    sheet = wb[sheetname]
    count = 0
    
    for idx, (row, data) in enumerate(enum_rows(sheet)):
        if count > 0 and (count % config["numrec_before_save"]) == 0:
            logger.info(
                "===============================================\n \
                Saving progress to excel file...\n=================================="
            )
            wb.save(orig_filename)  
        
        # Skip rows with a recent "last checked" date
        if should_skip_row(row, sheet_config):
            continue
        
        first_name = f"{row[sheet_config['first_name_index']].value or ''}".strip() 
        surname = f"{row[sheet_config['surname_index']].value or ''}".strip()
        search_text = re.sub(r'\s+',' ',f"{first_name} {surname}".strip()) 

        if first_name == '' or surname=='':
            company = f"{row[sheet_config['company_index']].value or ''}".strip() 
            handle_surveyor_license_query(row, company, sheet_config)    
        else:        
            # First try with name, fallback to company if name fails
            if not handle_surveyor_license_query(row, search_text, sheet_config):
                company = f"{row[sheet_config['company_index']].value or ''}".strip()
                handle_surveyor_license_query(row, company, sheet_config)
            
        count += 1

def should_skip_row(row, sheet_config):
    """Check if the row should be skipped based on last checked date."""
    
    last_date_checked = row[sheet_config['last_checked_index']].value
    return (last_date_checked and isinstance(last_date_checked, datetime) 
            and last_date_checked.date() >= datetime.now().date())


def query_pool_safety_license(lic_no):
    default_headers = {
        'accept': '*/*',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'origin': 'https://my.qbcc.qld.gov.au',
        'pragma': 'no-cache',
        'priority': 'u=1, i',
        'referer': 'https://my.qbcc.qld.gov.au/s/pool-safety-inspector-search',
        'sec-ch-ua': '"Brave";v="129", "Not=A?Brand";v="8", "Chromium";v="129"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'sec-gpc': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
        # 'x-sfdc-page-scope-id': '3bdc58c4-042b-42d8-8bcd-b875c8aa3bfb',
        # 'x-sfdc-request-id': '415150000004b0f99f',
    }
    
    s = requests.Session()
    response = s.get("https://my.qbcc.qld.gov.au/s/pool-safety-inspector-search")
    
    cookies = s.cookies.get_dict()
    context = cookies.get("renderCtx")
    context_decoded = json.loads(urllib.parse.unquote(context))
    sfdc_req_id = response.headers.get("x-sfdc-request-id")
    x_sfdc_page_scope_id = context_decoded.get('pageId')
    
    s.headers.update(default_headers)
    s.headers.update({
        #"x-request-id":req_id,
        "X-Sfdc-Request-Id":sfdc_req_id,
        "X-Sfdc-Page-Scope-Id":x_sfdc_page_scope_id
    })
    url = 'https://my.qbcc.qld.gov.au/s/sfsites/aura?other.PSISearch.searchInspectors=1'
    
    data = {
        'message': '{"actions":[{"id":"175;a","descriptor":"apex://PSISearchController/ACTION$searchInspectors","callingDescriptor":"markup://c:PSI_Search","params":{"searchBy":"licence","firstName":"","lastName":"","businessName":"","licenceNumber":"%s","distanceInKm":5,"batchSize":1000,"offset":0}}]}' % (lic_no,),
        'aura.context': '{"mode":"PROD","fwuid":"eGx3MHlRT1lEMUpQaWVxbGRUM1h0Z2hZX25NdHFVdGpDN3BnWlROY1ZGT3cyNTAuOC40LTYuNC41","app":"siteforce:communityApp","loaded":{"APPLICATION@markup://siteforce:communityApp":"wi0I2YUoyrm6Lo80fhxdzA","COMPONENT@markup://instrumentation:o11ySecondaryLoader":"1JitVv-ZC5qlK6HkuofJqQ"},"dn":[],"globals":{},"uad":false}',
        'aura.pageURI': '/s/pool-safety-inspector-search',
        'aura.token': 'null',
    }
    
    response = s.post(url,data=data)

    results = response.json()['actions'][0]['returnValue']
    if not results or len(results)<=0:
        return
    
    results0 = results[0]
    expiry_date = datetime.strptime(results0['expiryDate'], "%Y-%m-%d")
    if datetime.now() > expiry_date:
        results0['expired'] = True
        
    return results0
    
def process_sheet_qbcc_pool_safety(wb, sheetname, args, config, sheet_config):
    orig_filename = args.input
    
    if not all(keyword in sheetname.lower() for keyword in ["qbcc", "pool", "safety"]) or not args.qbcc:
        return

    logger.info("Processing QBCC Pool Safety Tab <QBCC>...")
    sheet = wb[sheetname]
    orig_filename = args.input
    save_interval = config["numrec_before_save"]
    
    for count, (row, data) in enumerate(enum_rows(sheet), start=1):
        logger.info(f"Processing Line #{count}")
        
        if count % save_interval == 0:
            logger.info("Saving progress to excel file...")
            wb.save(orig_filename)

        license_no = data.get("licence number", "").strip()
                # Skip rows with a recent "last checked" date
        
        if not license_no:
            message = "License No. Column not found!" if "licence number" not in data else "License No is BLANK !"
            logger.info(message)
            row[sheet_config["status_index"]].value = message
            row[sheet_config["last_checked_index"]].value = datetime.now().date()
            continue
        
        if should_skip_row(row, sheet_config):
            continue
        
        logger.info(f"Fetching License info of {license_no}:")
        lic_status = query_pool_safety_license(license_no)
        if lic_status:
            expired = lic_status.get('expired',False)
            
            if expired:
                row[sheet_config["status_index"]].value = "License Expired"
            else:
                row[sheet_config["status_index"]].value = "Active"
        else:
            logger.info("License not found in online register!")
            row[sheet_config["status_index"]].value = "Missing in Register"
        
        row[sheet_config["last_checked_index"]].value = datetime.now().date()
    
    wb.save(orig_filename)
        
        
                
def process_sheet_qbcc_individual(wb, sheetname, args, config, sheet_config, license_querier=query_qbcc_license, keywords=["qbcc", "individual"]):
    orig_filename = args.input
    if not all(keyword in sheetname.lower() for keyword in keywords) or not args.qbcc:
        return

    logger.info("Processing QBCC Individual...")

    sheet = wb[sheetname]
    count = 0
    
    for idx, (row, data) in enumerate(enum_rows(sheet)):
        logger.info(f"Processing Line #{idx + 1}")
        
        if should_skip_row(row,sheet_config):
            continue
        
        if count > 0 and (count % config["numrec_before_save"]) == 0:
            logger.info(
                "===============================================\n \
                Saving progress to excel file...\n=================================="
            )
            wb.save(orig_filename)

        if not "licence number" in data:
            logger.info("License No. Column not found !")
            row[sheet_config["status_index"]].value = (
                "No License Number Column found!"
            )
            row[sheet_config["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        if not data["licence number"] or data["licence number"].strip() == "":
            logger.info("License No. is BLANK in the excel file !")
            row[sheet_config["status_index"]].value = (
                "License No is BLANK !"
            )
            row[sheet_config["last_checked_index"]].value = (
                datetime.now().date()
            )
            count = count + 1
            continue

        license_no = data["licence number"]
        logger.info(f"Fetching License info of {license_no}:")
        lic_statuses = list(license_querier(license_no))
        if len(lic_statuses) > 0:
            logger.info("License info found!")
            lic_class, _, _, lic_status = lic_statuses[0]
            logger.info(f"\tLicense Class: {lic_class}")
            logger.info(f"\tStatus: {lic_status}")

            row[sheet_config["status_index"]].value = (
                lic_status.title().strip()
            )
            row[sheet_config["last_checked_index"]].value = (
                datetime.now().date()
            )
        else:
            logger.info("License not found in online register !")
            row[sheet_config["status_index"]].value = (
                "Missing in Register"
            )
            row[sheet_config["last_checked_index"]].value = (
                datetime.now().date()
            )

        count = count + 1

def read_config():

    with open("./config.yml", "rt", errors="ignore") as fp:
        conf = yaml.load(fp,Loader=Loader)

    return conf


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("--qbcc", action="store_true")
    parser.add_argument("--engr", action="store_true")
    parser.add_argument("--arch", action="store_true")
    parser.add_argument("--surv", action="store_true")

    args = parser.parse_args()
    
    
    if not os.path.isfile(args.input):
        logger.info("ERROR: Cannot open input file {}!".format(args.input))
        exit(1)
        
    try:
        wb = openpyxl.load_workbook(args.input)

        logger.info(f"Found {len(wb.sheetnames)} sheets:")
        logger.info("\n".join([f"\t{s}" for s in wb.sheetnames]))

        config = read_config()
        process_qbcc_certifier = partial(process_sheet_qbcc_individual, license_querier=query_qbcc_certifier_license,keywords=['qbcc','certifier'])
        process_qbcc_company = partial(process_sheet_qbcc_individual,keywords=['qbcc','company'])

        processors = [
            process_sheet_qbcc_individual,
            process_qbcc_company,
            process_qbcc_certifier,
            process_sheet_qbcc_pool_safety
            # process_sheet_arch,
            # process_sheet_engr,
            # process_sheet_surveyor
        ]
        
        for sheetname in wb.sheetnames:
            for sheetname_filter in config["sheets_config"].keys():
                if sheetname_filter.lower() == sheetname.lower():
                    logger.info(f"Processing SHEET: {sheetname}")
                    sheet_config = config["sheets_config"][sheetname]
                    
                    
                    for processor in processors:
                        processor(wb, sheetname, args, config, sheet_config)
                                
        logger.info(f"Process done. Saving to {args.input}")
        wb.save(args.input)
        
    except Exception as e:
        logger.exception(e)
    finally:
        wb.close()
