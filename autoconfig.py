import argparse
import openpyxl
import openpyxl.worksheet

def individual(sheet_name, columns):
    """
    docstring
    """
    pass

def extract_columns(worksheet):
    """
    docstring
    """
    rows = worksheet.iter_rows()
    for row in rows:
        print(row)
def main():
    """
    docstring
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    args = parser.parse_args()

    workbook = openpyxl.load_workbook(args.input, read_only=True)
    for worksheet in workbook:
        columns = extract_columns(worksheet)
        print(columns)

if __name__ == "__main__":
    main()
