
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

# def get_column_indexes(row):
#     surname = None

#     for idx, c in enumerate(row):
#         if "sur" in str(c.value).lower():
#             surname = idx + 1


def parse_response(html):
    soup = BeautifulSoup(html, 'html.parser')

    business_name_element = soup.select_one(
        "#ctl00_generalContentPlaceHolder_LicenceInfoControl1_lbLicenceName")
    trading_name_element = soup.select_one(
        'ctl00_generalContentPlaceHolder_LicenceInfoControl1_lbTradingName')

    items = [td.text.strip("\r\n\t ") for td in soup.select(
        "table[id='ctl00_generalContentPlaceHolder_LicenceInfoControl1_gvLicenceClass'] td")]

    items_array = np.array(items)

    if len(items_array) > 0 and (len(items_array) % 4) == 0:
        for part in np.split(items_array, len(items_array)/4):
            yield [p.strip() for p in part]


def query_license(license_no):
    license_no = license_no.strip("\r\n\t ")
    url = f'http://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/ShowDetailResultContent.aspx?LicNO={license_no}&licCat=LIC&name=&firstName=&searchType=Contractor&FromPage=SearchContr'
    response = requests.get(url)
    for r in parse_response(response.text):
        yield r


def query_engr_registration(registration_no, driver: Chrome):
    try:
        url1 = 'https://portal.bpeq.qld.gov.au/BPEQPortal/Search_for_a_RPEQ/BPEQPortal/Engineer_Search.aspx'
        driver.get(url1)
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.ID, 'ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Sheet0_Input3_TextBox1')))
        element.send_keys(registration_no)
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.ID, 'ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Sheet0_SubmitButton')))
        element.click()
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciEngineersearch_ciEngineersearch_ResultsGrid_Grid1_ctl00__0"]/td[1]/a')))
        registration_no = element.get_attribute('href').split("=")[1]

        url = f'https://portal.bpeq.qld.gov.au/Party.aspx?ID={registration_no}'
        response = driver.request("GET", url, verify=False)
        soup = BeautifulSoup(response.text, 'html.parser')
        parts = [p.text.strip("\r\n\t ")
                 for p in soup.select('.PanelFieldValue > span')]
        print(f"Num Parts: {len(parts)}")

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
        print(e)


def query_arch_registration(registration_no, driver: Chrome):
    try:
        url1 = 'https://www.boaq.qld.gov.au/BOAQ/Search_Register/BOAQ/Search_Register/Architect_Search.aspx'
        driver.get(url1)
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.ID, 'ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Sheet0_Input3_TextBox1')))
        element.send_keys(registration_no)
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.NAME, 'ctl01$TemplateBody$WebPartManager1$gwpciArchitectsearch$ciArchitectsearch$ResultsGrid$Sheet0$SubmitButton')))
        element.click()
        element = WebDriverWait(driver, 16).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Grid1_ctl00__0"]/td[1]/a')))
        registration_no = element.get_attribute('href').split("=")[1]
        url = f'https://www.boaq.qld.gov.au/Shared_Content/ContactManagement/Profile.aspx?ID={registration_no}'
        response = driver.request("GET", url)
        soup = BeautifulSoup(response.text, 'html.parser')
        parts = [p.text.strip("\r\n\t ")
                 for p in soup.select('.PanelFieldValue > span')]
        print(f"Num Parts: {len(parts)}")

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
        print(e)


def process_sheet_arch(wb, sheetname, orig_filename, config):

    if not 'architects' in sheetname.lower():
        return

    print("Processing Architects Tab...")
    sheet = wb[sheetname]

    config = read_config()
    count = 0

    options = ChromeOptions()
    # options.add_argument("headless")
    driver = Chrome(options=options)

    for row, data in enum_rows(sheet):

        if count > 0 and (count % config['numrec_before_save']) == 0:
            print("===============================================\n \
                Saving progress to excel file...\n==================================")
            wb.save(orig_filename)

        if not 'Registration' in data:
            print("Registration No. Column not found !")
            row[config['sheets_config'][sheetname]['status_index']
                ].value = "No Registration Number Column found!"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
            count = count + 1
            continue

        if not data['Registration'] or data['Registration'].strip() == '':
            print("Registration No. is BLANK in the excel file !")
            row[config['sheets_config'][sheetname]
                ['status_index']].value = "Registration No. is BLANK !"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
            count = count + 1
            continue

        registration_no = data['Registration']
        print(f"Feetching Registration info of {registration_no}:")
        reg_status = query_arch_registration(registration_no, driver)

        if reg_status:
            print("Registration info found!")

            name, company, date_joined, job_type, status, date_registered = reg_status

            # lic_class, _, _, lic_status = lic_statuses[0]
            print(f"\tName: {name}")
            print(f"\tCompany: {company}")
            print(f"\tDate Joined: {date_joined}")
            print(f"\tType: {job_type}")
            print(f"\tStatus: {status}")
            print(f"\tDate Registered: {date_registered}")

            row[config['sheets_config'][sheetname]['status_index']
                ].value = status.strip().title()
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
        else:
            print("Registration info not found in online register !")
            row[config['sheets_config'][sheetname]
                ['status_index']].value = "Missing in Register"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()

        count = count + 1


def process_sheet_engr(wb, sheetname, orig_filename, config):

    if not 'engineers' in sheetname.lower():
        return

    print("Processing Engineers Tab...")
    sheet = wb[sheetname]

    config = read_config()
    count = 0

    options = ChromeOptions()
    # options.add_argument("headless")
    driver = Chrome(options=options)

    for row, data in enum_rows(sheet):

        if count > 0 and (count % config['numrec_before_save']) == 0:
            print("===============================================\n \
                Saving progress to excel file...\n==================================")
            wb.save(orig_filename)

        if not 'Registration' in data:
            print("Registration No. Column not found !")
            row[config['sheets_config'][sheetname]['status_index']
                ].value = "No Registration Number Column found!"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
            count = count + 1
            continue

        if not data['Registration'] or data['Registration'].strip() == '':
            print("Registration No. is BLANK in the excel file !")
            row[config['sheets_config'][sheetname]
                ['status_index']].value = "Registration No. is BLANK !"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
            count = count + 1
            continue

        registration_no = data['Registration']
        print(f"Feetching Registration info of {registration_no}:")
        reg_status = query_engr_registration(registration_no, driver)

        if reg_status:
            print("Registration info found!")

            name, company, date_joined, job_type, status, date_registered = reg_status

            # lic_class, _, _, lic_status = lic_statuses[0]
            print(f"\tName: {name}")
            print(f"\tCompany: {company}")
            print(f"\tDate Joined: {date_joined}")
            print(f"\tType: {job_type}")
            print(f"\tStatus: {status}")
            print(f"\tDate Registered: {date_registered}")

            row[config['sheets_config'][sheetname]['status_index']
                ].value = status.strip().title()
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()
        else:
            print("Registration info not found in online register !")
            row[config['sheets_config'][sheetname]
                ['status_index']].value = "Missing in Register"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()

        count = count + 1


def process_sheet_qbcc(wb, sheetname, orig_filename, config):
    if not 'qbcc' in sheetname.lower():
        return

    print("Processing Architects Tab <QBCC>...")

    sheet = wb[sheetname]

    count = 0
    idx = 0
    for row, data in enum_rows(sheet):
        idx = idx + 1
        print(f"Processing Line #L{idx}")
        if count > 0 and (count % config['numrec_before_save']) == 0:
            print("===============================================\n \
                Saving progress to excel file...\n==================================")
            wb.save(orig_filename)

        if not 'licence number' in data:
            print("License No. Column not found !")
            row[config['sheets_config'][sheetname]['status_index']
                ].value = "No License Number Column found!"
            row[config['sheets_config'][sheetname]['last_checked_index']
                ].value = datetime.now().date()
            count = count + 1
            continue

        if not data['licence number'] or data['licence number'].strip() == '':
            print("License No. is BLANK in the excel file !")
            row[config['sheets_config'][sheetname]['status_index']
                ].value = "Licens No is BLANK !"
            row[config['sheets_config'][sheetname]['last_checked_index']
                ].value = datetime.now().date()
            count = count + 1
            continue

        license_no = data['licence number']
        print(f"Feetching License info of {license_no}:")
        lic_statuses = list(query_license(license_no))

        if len(lic_statuses) > 0:
            print("License info found!")
            lic_class, _, _, lic_status = lic_statuses[0]
            print(f"\tLicense Class: {lic_class}")
            print(f"\tStatus: {lic_status}")

            row[config['sheets_config'][sheetname]['status_index']
                ].value = lic_status.title().strip()
            row[config['sheets_config'][sheetname]['last_checked_index']
                ].value = datetime.now().date()
        else:
            print("License not found in online register !")
            row[config['sheets_config'][sheetname]
                ['status_index']].value = "Missing in Register"
            row[config['sheets_config'][sheetname]
                ['last_checked_index']].value = datetime.now().date()

        count = count + 1


def enum_rows(sheet):
    # link = sheet['A1'].value
    # print(type(link))
    # row_start_index = None
    # if link:
    #     for i in range(20):

    #         if str(sheet['A%d' % (i + 1, )].value).lower() == 'surname' or str(sheet['A%d' % (i + 1, )].value).lower() == 'sur name':
    #             row_start_index = i + 1
    #             print(f"Data Row Start Index { row_start_index}")
    #             break

    #     if row_start_index:
    #         sheet[f'A{row_start_index}'].value

    headers = list()

    for i, r in enumerate(sheet.rows):

        print(f"Processing Line # {i+1}...")

        values = [str(c.value).strip() for c in r]

        if not 'surname' == values[0].lower() and not headers:
            continue

        if not headers:
            headers = values
            continue

        item = dict()
        for h, c in zip(headers, r):
            item[h] = str(c.value) if c.value else ''

        yield r, item

    # for itm in items:
    #     pass


def read_config():

    with open('./config.json', 'rt', errors='ignore') as fp:
        conf = json.loads(fp.read())

    return conf


if __name__ == "__main__":

    # with open('./result.html', 'rt') as fp:
    #     for license_class, _, _, status in parse_response(fp.read()):
    #         pass

    # exit(0)

    parser = argparse.ArgumentParser()
    parser.add_argument("input")

    args = parser.parse_args()

    try:
        wb = openpyxl.load_workbook(args.input)

        print(f"Found {len(wb.sheetnames)} sheets:")
        print('\n'.join([f'\t{s}' for s in wb.sheetnames]))

        config = read_config()

        for sheetname in wb.sheetnames:
            for filtsheet in config["sheets_config"].keys():
                if filtsheet.lower() == sheetname.lower():
                    print(f"Processing {sheetname}")

                    # process_sheet_qbcc(wb, sheetname, args.input, config)

                    # process_sheet_arch(
                    #     wb, sheetname, args.input, config)

                    process_sheet_engr(wb, sheetname, args.input, config)

        wb.save(args.input)

    except Exception as e:
        print(e)
    finally:
        wb.close()
