import time
import sys
import os
from pathlib import Path
from get_data_txt import get_routs, get_name_template, get_data_report
from log import log
from convert_xls_to_xlsx import convert_xls_to_xlsx
from create_report_excel import create_report_excel
from get_data_template_excel import get_data_template_excel
from delete_sheets import delete_sheets

def main():
    start_time = time.time()

    #optain the parameters command line
    params = sys.argv

    separator = '' 
    # name_report_txt = params[1]
    # number_session = params[2]
    name_report_txt = '000007477'

    rout_aplication = str(Path(__file__).parent.absolute())# ruta SIIFNET
    print("ruta SIIFNET: ", rout_aplication)

    # rout_environment = os.path.dirname(rout_aplication) #ruta del ambiente IDEA
    # print(rout_environment)
    rout_environment = "C:\\Users\\javier.puentes\\ser_excel"

    rout_log = rout_environment + "\\" + get_routs(rout_aplication,11).strip()
    print("ruta log: ", rout_log)

    #Variables iniciales para log
    nameAplication = "ser_excel"
    message = "Iniciando aplicación ser_excel \n"
 
    # if params.count > 1:
    rout_fiel_txt = rout_environment + "\\" + get_routs(rout_aplication,24).strip() + name_report_txt + '.txt' #ruta del reporte txt
    message = message + "Ruta del archivo de reporte: " + rout_fiel_txt + "\n"
    
    if os.path.exists(rout_fiel_txt):
        name_template = get_name_template(rout_fiel_txt,separator)
        print("name_template: ", name_template)
        rout_template_excel = rout_environment + "\\" + get_routs(rout_aplication,4).strip() + name_template + '.xlsx' #ruta del template excel
        print('ruta plantilla: ', rout_template_excel)
        if not os.path.exists(rout_template_excel):
            rout_template_excel_xls = rout_environment + "\\" + get_routs(rout_aplication,4).strip() + name_template + '.xls' 
            message = message + " Plantilla con extension xls, se convierte a xlsx \n"
            try:
                convert_xls_to_xlsx(rout_template_excel_xls, rout_template_excel)
            except Exception as e:
                print(e)
                message = message + "Error al convertir la plantilla xls a xlsx" + "\n"
                message = message + "Error: " + str(e) + "\n"
        
        if os.path.exists(rout_template_excel):
            print("ruta plantilla: ", rout_template_excel)
            message = message + "Ruta del archivo de plantilla: " + rout_template_excel + "\n"

            data_report, messageCall = get_data_report(rout_fiel_txt, separator)
            message = message + messageCall + "\n"

            messageCall, rout_report_excel, principal_sheet = create_report_excel(data_report, rout_template_excel)
            message = message + messageCall + "\n"

            #delete_sheets(rout_report_excel, principal_sheet)

            finish_time = time.time()
            message = message + "Tiempo de ejecución: " + str(finish_time - start_time) + " segundos" + "\n"
        else:
            message = message + "No existe el archivo de la plantilla xlsx" + rout_template_excel +  "\n"
            print(message)

    else:
        message = message + "No existe el archivo de reporte"+ rout_fiel_txt + "\n"
        print(message)

    log(rout_log, nameAplication, message)

if __name__ == "__main__":
    main()