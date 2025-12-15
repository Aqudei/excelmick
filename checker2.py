import asyncio
import pdb
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from playwright.async_api import async_playwright
from playwright.async_api import Page


async def get_sheets_as_df(file_path):
    # Load the Excel file
    file_path = "processing\\24.10.03 - Competent Person Register.xlsx"
    workbook = openpyxl.load_workbook(file_path)

    for sheet in workbook.sheetnames:
        print(f"Sheet name: {sheet}")
        sheet = workbook[sheet]
        # Read data into a DataFrame
        data = pd.DataFrame(sheet.values)
        data.columns = data.iloc[0]
        data = data[1:]
        yield data, sheet.title

    workbook.close()


async def boaq_search_reg_no(page: Page, reg_no):
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
    rows = []
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
                row[headers[col_index]] = {
                    "text": value,
                    "url": link
                }
            else:
                row[headers[col_index]] = td.get_text(strip=True)

            col_index += 1

        rows.append(row)

    return rows

async def handle_archi_sheet(df, sheet_name):
    # Example processing for 'archi' sheets
    print("Processing 'archi' sheet...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        for index, row in df.iterrows():
            result = await boaq_search_reg_no(page, row.licence_number)

        await browser.close()


async def main():
    file_path = "processing\\24.10.03 - Competent Person Register.xlsx"
    dfs = get_sheets_as_df(file_path)

    async for df, sheet_name in dfs:
        df.columns = df.columns.str.lower().str.replace(" ", "_")
        if "archi" in sheet_name.lower():
            await handle_archi_sheet(df, sheet_name)


if __name__ == "__main__":
    asyncio.run(main())
