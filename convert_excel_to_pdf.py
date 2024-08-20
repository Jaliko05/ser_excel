import win32com.client as win32
import os
import time

def convert_excel_to_pdf(file_path, retries=5, wait_time=5):
    xlsx_file = file_path
    pdf_file = xlsx_file.replace('.xlsx', '.pdf')

    if os.path.exists(pdf_file):
        os.remove(pdf_file)

    attempt = 0
    while attempt < retries:
        try:
            excel = win32.Dispatch('Excel.Application')
            workbook = excel.Workbooks.Open(xlsx_file)
            workbook.ExportAsFixedFormat(0, pdf_file)
            workbook.Close(False)
            excel.Quit()
            break  # Si todo sale bien, salimos del loop
        except Exception as e:
            attempt += 1
            if attempt >= retries:
                raise e  # Si se superan los intentos, lanzar el error
            time.sleep(wait_time)  # Espera antes de reintentar
