import asyncio
from processors import surveyor
import pandas as pd
import openpyxl
import logging
from logging_config import setup_logging

setup_logging()

logger = logging.getLogger(__name__)


def get_sheets_as_df(file_path):
    logger.info("Opening file: <{}>".format(file_path))
    workbook = openpyxl.load_workbook(file_path, read_only=True)

    sheets_data = {}

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]

        df = pd.DataFrame(ws.values)
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)

        sheets_data[sheet_name] = df

    workbook.close()

    return sheets_data


async def main():
    file_path = r"C:\dev\excelmick\processing\24.09.27 - Competent Person Register.xlsx"
    dfs = get_sheets_as_df(file_path)

    for sheet_name, df in dfs.items():
        df.columns = df.columns.str.lower().str.replace(" ", "_")
        # if "archi" in sheet_name.lower():
        #     await architects.handle_sheet(df, sheet_name, file_path)

        if "surveyor" in sheet_name.lower():
            await surveyor.handle_sheet(df, sheet_name, file_path)


if __name__ == "__main__":
    asyncio.run(main())
