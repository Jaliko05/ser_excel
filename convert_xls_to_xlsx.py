import os
import win32com.client as client

def convert_xls_to_xlsx(route_template_xls, route_template_xlsx):
    excel = client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(os.path.abspath(route_template_xls))
    wb.SaveAs(os.path.abspath(route_template_xlsx), FileFormat=51)
    wb.Close()
    excel.Quit()

