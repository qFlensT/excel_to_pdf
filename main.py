# Import Module
import pathlib
import sys
import os
from typing import Any
from win32com import client
from win32.win32api import MessageBox
from win32.lib.win32con import MB_ICONERROR

def open_excel() -> Any:
    excel = client.Dispatch("Excel.Application")
    return excel

def open_workbook(excel: Any, path: str) -> Any:
    return excel.Workbooks.Open(path)

def open_worksheet(workbook: Any, index: int) -> Any:
    return workbook.Worksheets[index]

def sheets_count(workbook: Any) -> int:
    return workbook.Sheets.Count

def apply_params(worksheet: Any):
    worksheet.PageSetup.Zoom = False
    worksheet.PageSetup.FitToPagesTall = 1
    worksheet.PageSetup.FitToPagesWide = 1

def select_worksheets(workbook: Any, index_list: list):
    workbook.WorkSheets(list(map(lambda x: x + 1, index_list))).Select()

def export_to_pdf(workbook: Any, path: str):
    workbook.ActiveSheet.ExportAsFixedFormat(0, path)

def close(excel: Any, workbook: Any):
    workbook.Close()

def generate_pdf_path(excel_file_path: str) -> str:
    working_dir = os.path.normpath(excel_file_path + os.sep + os.pardir)
    excel_file_name = pathlib.Path(excel_file_path).name
    pdf_file_name = list(excel_file_name)
    while True:
        if pdf_file_name[-1] != ".":
            del pdf_file_name[-1]
        else:
            pdf_file_name.append("pdf")
            pdf_file_name = "".join(pdf_file_name)
            break

    return os.path.join(working_dir + os.sep + pdf_file_name)

def main():
    if len(sys.argv) == 1:
        MessageBox(0, f"sys.argv error (len == 1). Arguments: {sys.argv}", "Error", MB_ICONERROR)
    excel_file_path = "".join([path + " " for path in sys.argv[1::]])
    try:
        pdf_file_path = generate_pdf_path(excel_file_path)
    except Exception as err:
        MessageBox(0, f"Generate PDF Path Error: {err}.\nExcel path: {excel_file_path}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        excel = open_excel()
    except Exception as err:
        MessageBox(0, f"Open Excel Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        workbook = open_workbook(excel, excel_file_path)
    except Exception as err:
        MessageBox(0, f"Open Workbook Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        worksheets = [open_worksheet(workbook, i) for i in range(sheets_count(workbook))]
    except Exception as err:
        MessageBox(0, f"Getting Worksheets Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        for worksheet in worksheets:
            apply_params(worksheet)
    except Exception as err:
        MessageBox(0, f"Params Apply Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        select_worksheets(workbook, [i for i in range(sheets_count(workbook))])
    except Exception as err:
        MessageBox(0, f"Select Worksheets Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        export_to_pdf(workbook, pdf_file_path)
    except Exception as err:
        MessageBox(0, f"PDF Export Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)
    try:
        close(excel, workbook)
        pass
    except Exception as err:
        MessageBox(0, f"Excel Close Error: {err}", "Error", MB_ICONERROR)
        sys.exit(0)

if __name__ == "__main__":
    main()