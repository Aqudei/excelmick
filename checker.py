
import openpyxl
import argparse
from openpyxl.worksheet.hyperlink import Hyperlink
import requests
from bs4 import BeautifulSoup
import itertools
import numpy as np
import json
from datetime import datetime
# def get_column_indexes(row):
#     surname = None

#     for idx, c in enumerate(row):
#         if "sur" in str(c.value).lower():
#             surname = idx + 1


def parse_response(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = [td.text.strip("\r\n\t ") for td in soup.select(
        "table[id='ctl00_generalContentPlaceHolder_LicenceInfoControl1_gvLicenceClass'] td")]

    items_array = np.array(items)

    if len(items_array) > 0 and (len(items_array) % 4) == 0:
        for part in np.split(items_array, len(items_array)/4):
            yield [p.strip() for p in part]


def query_license(license_no):
    url = f'http://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/ShowDetailResultContent.aspx?LicNO={license_no}&licCat=LIC&name=&firstName=&searchType=Contractor&FromPage=SearchContr'
    response = requests.get(url)
    for r in parse_response(response.text):
        yield r


def process_sheet(sheet, wb, orig_filename):
    config = read_config()
    count = 0

    for row, data in enum_rows(sheet):

        if count > 0 and (count % config['numrec_before_save']) == 0:
            print("===============================================\n \
                Saving progress to excel file...\n==================================")
            wb.save(orig_filename)

        if not 'licence number' in data:
            print("License is BLANK in the excel file !")
            row[config['status_index']].value = "License Number is Blank!"
            row[config['last_checked_index']].value = datetime.now().date()
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
            row[config['status_index']].value = lic_status.title().strip()
            row[config['last_checked_index']].value = datetime.now().date()
        else:
            print("License not found in online register !")
            row[config['status_index']].value = "Missing in Register"
            row[config['last_checked_index']].value = datetime.now().date()

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

        if i == 0:
            continue
        elif i == 1:
            headers = [str(c.value).strip() for c in r]
            continue
        else:
            item = dict()
            for i in range(len(headers)):
                item[headers[i]] = str(r[i].value) if r[i].value else ''
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

    wb = openpyxl.load_workbook(args.input)

    print(f"Found {len(wb.sheetnames)} sheets:")
    print('\n'.join([f'\t{s}' for s in wb.sheetnames]))

    config = read_config()

    for sheetname in wb.sheetnames:
        for filtsheet in config["sheets_filter"]:
            if filtsheet.lower() in sheetname.lower():
                print(f"Processing {sheetname}")
                process_sheet(wb[sheetname], wb, args.input)

    wb.save(args.input)
    wb.close()
