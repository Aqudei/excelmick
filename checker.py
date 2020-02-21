
import openpyxl
import argparse
from openpyxl.worksheet.hyperlink import Hyperlink
import requests
from bs4 import BeautifulSoup
import itertools
import numpy as np

sheets_filter = [
    'QBCC'
]


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


def process_sheet(sheet):
    for row in enum_rows(sheet):
        if not 'licence number' in row:
            raise Exception("The column 'licence number' was not found")

        license_no = row['licence number']
        print(f"Lic info of {license_no}:")
        lic_statuses = query_license(license_no)
        for lic_class, _, _, lic_status in lic_statuses:
            print(f"\t{lic_class} - {lic_status}")


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
    items = list()

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
            yield item

    # for itm in items:
    #     pass


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

    for sheetname in wb.sheetnames:
        for filtsheet in sheets_filter:
            if filtsheet.lower() in sheetname.lower():
                print(f"Processing {sheetname}")
                process_sheet(wb[sheetname])

    wb.close()
