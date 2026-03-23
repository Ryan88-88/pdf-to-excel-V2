import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook

def process_pdf_file(pdf_path, output_path):

    temp_excel = "temp_output.xlsx"

    with pdfplumber.open(pdf_path) as pdf:
        writer = pd.ExcelWriter(temp_excel, engine='openpyxl')
        table_count = 0

        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()

            for table in tables:
                df = pd.DataFrame(table)
                sheet_name = f"Page{page_number}_Table{table_count+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                table_count += 1

        writer.close()

    all_sheets = pd.read_excel(temp_excel, sheet_name=None, header=None)
    sheet_names = list(all_sheets.keys())

    if len(sheet_names) < 2:
        raise Exception("PDF table not found")

    df1 = all_sheets[sheet_names[0]]
    df2 = all_sheets[sheet_names[1]]

    wb = load_workbook(temp_excel)

    ws_ns = wb.create_sheet("NS")
    ws_ds = wb.create_sheet("DS")

    ws_ns.append([df1.iloc[8,3], df1.iloc[8,8], df1.iloc[8,13]])
    ws_ds.append([df2.iloc[9,3], df2.iloc[9,8], df2.iloc[9,13]])

    wb.save(output_path)
    os.remove(temp_excel)
