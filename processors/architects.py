from bs4 import BeautifulSoup
from playwright.async_api import Page
import requests
import logging
from db import architect_set_status, architect_find
import openpyxl
from datetime import datetime

logger = logging.getLogger(__name__)

async def architect_get_status(resource):
    url = "https://www.boaq.qld.gov.au" + resource

    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")
    status_element = soup.select_one(
        "span#ctl01_TemplateBody_WebPartManager1_gwpciNewContactMiniProfileCommon_ciNewContactMiniProfileCommon_contactStatus_memberStatus"
    )
    if not status_element:
        return "Not Found"
    status = status_element.get_text(strip=True)
    return status


async def boaq_search_reg_no(page: Page, reg_no):
    rows = []

    try:
        await page.goto(
            "https://www.boaq.qld.gov.au/Web/Consumers/Search_the_Register/Web/Architect_Search.aspx?hkey=f493b110-1ad9-4ec8-a830-f9a1f70e16b5"
        )

        await page.fill(
            "#ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Sheet0_Input3_TextBox1",
            reg_no,
        )
        await page.click(
            "#ctl01_TemplateBody_WebPartManager1_gwpciArchitectsearch_ciArchitectsearch_ResultsGrid_Sheet0_SubmitButton",
        )

        await page.wait_for_load_state("networkidle")
        soup = BeautifulSoup(await page.content(), "html.parser")
        table = soup.find("table", class_="rgMasterTable")

        # ---- get visible headers ----
        headers = []
        for th in table.select("thead th"):
            # skip hidden headers
            if "display:none" in (th.get("style") or ""):
                continue
            headers.append(th.get_text(strip=True))

        # ---- parse rows ----

        for tr in table.select("tbody tr"):
            cells = tr.find_all("td")

            row = {}
            col_index = 0

            for td in cells:
                # skip hidden cells
                if "display:none" in (td.get("style") or ""):
                    continue

                # extract link if present
                a = td.find("a")
                if a:
                    value = a.get_text(strip=True)
                    link = a.get("href")
                    row[headers[col_index]] = {"text": value, "url": link}
                else:
                    row[headers[col_index]] = td.get_text(strip=True)

                col_index += 1

            rows.append(row)
    except Exception as e:
        logger.info(f"Error occurred: {e}")

    return rows


async def handle_sheet(df, sheet_name, input_file):
    # Example processing for 'archi' sheets
    print("Fetching architect updates...")
    await fetch_updates(df, sheet_name)
    print("Applying architect updates...")
    await apply_updates(sheet_name, input_file)


async def apply_updates(sheet_name, input_file):
    wb = openpyxl.load_workbook(input_file)
    ws = wb[sheet_name]

    # 1-based index
    key_column = 3
    value_column = 7
    timestamp_column = 8

    for row_idx, _row in enumerate(ws.iter_rows(min_row=2), start=1):
        value = _row[key_column - 1].value  # column A
        if value is None:
            continue

        archs = architect_find(value)

        ws.cell(row=row_idx + 1, column=value_column).value = archs[0]["status"]
        ws.cell(row=row_idx + 1, column=timestamp_column).value = datetime.now()

    wb.save(filename=input_file)
    wb.close()


async def fetch_updates(df, sheet_name):
    logger.info("Processing 'archi' sheet <{}>...".format(sheet_name))

    for index, row in df.iterrows():
        status = await architect_get_status(
            "/Party.aspx?ID={}".format(row.licence_number)
        )
        architect_set_status(row, status)
