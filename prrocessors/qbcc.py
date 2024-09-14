
import logging

from bs4 import BeautifulSoup
import requests
import numpy as np

logger = logging.getLogger('micklogger')

class QBCCProcessor(object):

    def __init__(self, workbook, args,config):
        
        self.session = requests.Session()
        self.workbook = workbook
        self.args = args
        self.config = config
        
    def __parse_response(self,html):
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
 
    def __query_license(self,license_no):
        self.session.get('https://www.onlineservices.qbcc.qld.gov.au/OnlineLicenceSearch/VisualElements/SearchBSALicenseeContent.aspx')

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
        response = requests.get(url)
        for r in self.__parse_response(response.text):
            yield r
        
    def __enum_rows(self, sheet):

        headers = list()
        counter = 0

        for r in sheet.rows:
            values = [f"{c.value or ''}".strip() for c in r]

            if not headers:
                headers = [v.lower() for v in values]
                continue

            item = dict()
            for h, c in zip(headers, r):
                item[h] = f"{c.value or ''}".strip() if c else ""

            yield counter,r, item
            counter += 1
            
    def process(self, sheet):
        for idx,row,data in self.__enum_rows(sheet):
            pass