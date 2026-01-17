# https://sbq.com.au/find-a-surveyor/search-sbq-registrant-list/?type=all&search-type=name&title=Paul+Leonard+Heald&postcode=&radius=10&q=Search

import os
from bs4 import BeautifulSoup
from playwright.async_api import Page
import logging
from db import upsert, find_by_key
import openpyxl
from datetime import datetime
from playwright.async_api import async_playwright

logger = logging.getLogger(__name__)

DB_NAME = "qbcc-company"


def parse_response(html):
    # Assuming your HTML is stored in a variable called `html`
    soup = BeautifulSoup(html, "html.parser")

    # Extract licensee name
    name_tag = soup.find("a", class_="slds-card__header-link")
    name = name_tag.get_text(strip=True) if name_tag else None
    if not name:
        return {}, []

    # Extract License number, Type, Address, ABN, MR category
    def get_detail(label):
        div_label = soup.find("div", text=label)
        if div_label:
            value_div = div_label.next_sibling
            if value_div:
                return value_div.get_text(strip=True)
        return None

    license_number = get_detail("Licence Number:")
    license_type = get_detail("Type:")
    address = get_detail("Address:")
    abn = get_detail("ABN:")
    mr_category = get_detail("MR category:")

    # Select all table rows
    rows = soup.select("tbody c-virtual-record-list-row tr")

    licence_classes = []

    for row in rows:
        cols = row.find_all("div", class_="slds-truncate")
        if len(cols) >= 4:  # Ensure we have at least 4 columns
            licence_classes.append(
                {
                    "class": cols[0].get_text(strip=True),
                    "type": cols[1].get_text(strip=True),
                    "condition": cols[2].get_text(strip=True),
                    "status": cols[3].get_text(strip=True),
                }
            )

    # Output
    for lc in licence_classes:
        print(lc)

    return {
        "name": name,
        "license_number": license_number,
        "license_type": license_type,
        "address": address,
        "abn": abn,
        "mr_category": mr_category,
    }, licence_classes


async def fetch_item(page: Page, lic_no: str):
    try:
        url = "https://my.qbcc.qld.gov.au/myQBCC/s/qbcc-licensee-register"
        await page.goto(url)
        # Open combobox
        await page.get_by_role("combobox").click()

        # Wait for dropdown to appear
        await page.get_by_role("listbox").wait_for()

        # Select option
        await page.get_by_role("option", name="Licence number").click()
        await page.get_by_placeholder("Licence number").fill(lic_no)
        await page.get_by_role("button", name="Search").click()

        await page.get_by_role("button", name="Licensee Info").click()
        await page.get_by_title("LICENCE DETAILS").wait_for()

        content = await page.content()
        return parse_response(content)

    except Exception as e:
        return {}, []


async def boaq_search_reg_no(page: Page, reg_no):
    pass


async def handle_sheet(df, sheet_name, input_file, args=None):
    if args and args.fetch:
        logger.info("Fetching qbcc_company updates...")
        await fetch_updates(df, sheet_name)

    if args and args.apply:
        logger.info("Applying qbcc_company updates...")
        await apply_updates(sheet_name, input_file)


async def apply_updates(sheet_name, input_filename):
    wb = openpyxl.load_workbook(input_filename)
    ws = wb[sheet_name]

    value_column = 8
    timestamp_column = value_column + 2

    for row_idx, _row in enumerate(ws.iter_rows(min_row=2), start=1):
        key = str(_row[2].value or "").strip()

        if key in [None, ""]:
            continue

        qbcc = find_by_key(DB_NAME, key)

        ws.cell(row=row_idx + 1, column=value_column).value = qbcc.get("status")
        ws.cell(row=row_idx + 1, column=timestamp_column).value = datetime.now()

    base, ext = os.path.splitext(input_filename)
    final_output_filename = f"{base}_output{ext}"

    wb.save(filename=final_output_filename)
    wb.close()


async def fetch_updates(df, sheet_name):
    async with async_playwright() as playwright:
        chromium = playwright.chromium  # or "firefox" or "webkit".
        browser = await chromium.launch(headless=False)
        page = await browser.new_page()

        # other actions...

        logger.info("Processing 'qbcc' sheet <{}>...".format(sheet_name))

        for index, row in df.iterrows():
            lic_no = str(row.licence_number or "").strip()
            logger.info("Processing 'qbcc' with lic_no <{}>...".format(lic_no))

            if lic_no in [None, ""]:
                upsert(DB_NAME, {"status": "Not Found", "key": lic_no}, lic_no)
                continue

            info, lic_classes = await fetch_item(page, lic_no)

            if not info:
                upsert(DB_NAME, {"status": "Not Found", "key": lic_no}, lic_no)

            else:
                is_active = any(s["status"].lower() == "active" for s in lic_classes)
                status = "Active" if is_active else "Not Active"
                upsert(DB_NAME, {**info, "status": status, "key": lic_no}, lic_no)

        await browser.close()
