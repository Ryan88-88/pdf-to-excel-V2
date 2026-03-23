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

    # 自动读取所有表（防报错版本）
    all_sheets = pd.read_excel(temp_excel, sheet_name=None, header=None)
    sheet_names = list(all_sheets.keys())

    if len(sheet_names) < 2:
        raise Exception("Not enough tables found in PDF")

    df1 = all_sheets[sheet_names[0]]
    df2 = all_sheets[sheet_names[1]]

    wb = load_workbook(temp_excel)

    for sheet in ["NS", "DS"]:
        if sheet in wb.sheetnames:
            del wb[sheet]

    ws_ns = wb.create_sheet("NS")
    ws_ds = wb.create_sheet("DS")

    col_D = 3; col_E = 4; col_F = 5; col_G = 6; col_I = 8
    col_J = 9; col_L = 11; col_N = 13; col_Q = 16

    ws_ns.append([df1.iloc[8, col_D], df1.iloc[8, col_I], df1.iloc[8, col_N]])
    ws_ds.append([df2.iloc[9, col_D], df2.iloc[9, col_I], df2.iloc[9, col_N]])

    for r in range(10, 65):
        if pd.notna(df1.iloc[r, col_L]) and df1.iloc[r, col_L] != 0:
            ws_ns.append([df1.iloc[r, col_D], df1.iloc[r, col_E], df1.iloc[r, col_F],
                          df1.iloc[r, col_G], df1.iloc[r, col_J], df1.iloc[r, col_L],
                          df1.iloc[r, col_Q]])

        if pd.notna(df1.iloc[r, col_N]) and df1.iloc[r, col_N] != 0:
            ws_ns.append([df1.iloc[r, col_D], df1.iloc[r, col_E], df1.iloc[r, col_F],
                          df1.iloc[r, col_G], df1.iloc[r, col_J], df1.iloc[r, col_N],
                          df1.iloc[r, col_Q]])

    for r in range(11, 66):
        if pd.notna(df2.iloc[r, col_L]) and df2.iloc[r, col_L] != 0:
            ws_ds.append([df2.iloc[r, col_D], df2.iloc[r, col_E], df2.iloc[r, col_F],
                          df2.iloc[r, col_G], df2.iloc[r, col_J], df2.iloc[r, col_L],
                          df2.iloc[r, col_Q]])

        if pd.notna(df2.iloc[r, col_N]) and df2.iloc[r, col_N] != 0:
            ws_ds.append([df2.iloc[r, col_D], df2.iloc[r, col_E], df2.iloc[r, col_F],
                          df2.iloc[r, col_G], df2.iloc[r, col_J], df2.iloc[r, col_N],
                          df2.iloc[r, col_Q]])

    wb.save(output_path)
    os.remove(temp_excel)


if __name__ == "__main__":
    process_pdf_file("input.pdf", "final_output.xlsx")
