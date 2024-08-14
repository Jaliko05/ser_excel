import win32com.client as win32
def convert_excel_to_pdf(file_path):
    # Ruta al archivo .xlsx
    xlsx_file = file_path
    pdf_file = xlsx_file.replace('.xlsx', '.pdf')

    # Abre Excel usando COM
    excel = win32.Dispatch('Excel.Application')

    # Carga el archivo .xlsx
    workbook = excel.Workbooks.Open(xlsx_file)

    # Exporta a PDF
    workbook.ExportAsFixedFormat(0, pdf_file)

    # Cierra el archivo y Excel
    workbook.Close(False)
    excel.Quit()