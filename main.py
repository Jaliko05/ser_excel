import time
import sys
import os

from pathlib import Path
from convert_excel_to_pdf import convert_excel_to_pdf
from get_data_txt import get_routs, get_name_template, get_data_report
from log import log
from convert_xls_to_xlsx import convert_xls_to_xlsx
from create_report_excel import create_report_excel

def main():
    start_time = time.time()
    separator = '' 

    #comentar hata la linea 42 para ejecutar en local con python main.py,
    #optain the parameters command line
    param = sys.argv
    print(param)
    params = param[1].split(separator)
    if params.count == 1:
       params = param.split(" ")
    path = os.path.abspath(__file__)  
    print("real path: ", path)

    name_report_txt = params[0][1:]  
    utilita = params[0][0]  
    number_session = params[1]
    print("number_session: ", number_session)
    print("name_report_txt: ", name_report_txt)
    print("utilita: ", utilita)

    #Obtener la ruta completa del archivo ejecutable
    ruta_exe = sys.executable

    rout_aplication =  os.path.dirname(ruta_exe)#ruta SIIFNET
    print("ruta SIIFNET: ", rout_aplication)

    partes_ruta = os.path.normpath(rout_aplication).split(os.sep)

    rout_environment = '\\'.join(partes_ruta[0:-2]) #ruta del ambiente "IDEA"
    print("ruta ambiente: ", rout_environment)


    #descomentar hata la linea 50 para ejecutar en local con python main.py,
    # number_session = "0090106"
    # name_report_txt = '000022899'
    # utilita = 'P'
    # rout_environment = "C:\\Users\\javier.puentes\\ser_excel"
    # rout_aplication =  "C:\\Users\\javier.puentes\\ser_excel\\ser_excel"

    rout_log = rout_environment + "\\" + get_routs(rout_aplication,11).strip()
    print("ruta log: ", rout_log)

    route_report_xlsx = rout_environment + "\\" + get_routs(rout_aplication,2).strip() + "E" + number_session + '.xlsx' #ruta del reporte txt
    print("ruta reporte xlsx: ", route_report_xlsx)

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
            message = message + "Plantilla con extension xls, se convierte a xlsx \n"
            try:
                convert_xls_to_xlsx(rout_template_excel_xls, rout_template_excel)
                message = message + "Plantilla convertida a xlsx exitosamente" + "\n"
                os.remove(rout_template_excel_xls)
                message = message + "Plantilla xls eliminada exitosamente" + "\n"
            except Exception as e:
                print(e)
                message = message + "Error al convertir la plantilla xls a xlsx" + "\n"
                message = message + "Error: " + str(e) + "\n"
                log(rout_log, nameAplication, message)
        
        if os.path.exists(rout_template_excel):
            print("ruta plantilla: ", rout_template_excel)
            message = message + "Ruta del archivo de plantilla: " + rout_template_excel + "\n"

            data_report, messageCall = get_data_report(rout_fiel_txt, separator)
            message = message + messageCall 

            messageCall, rout_report_excel, principal_sheet = create_report_excel(data_report, rout_template_excel, route_report_xlsx , rout_log)
            message = message + messageCall + "\n"
            if utilita == 'P':
                try:
                    message = message + "Generando archivo PDF" + "\n"
                    message = message + convert_excel_to_pdf(rout_report_excel)
                    message = message + "Archivo PDF generado exitosamente" + "\n"
                except Exception as e:
                    message = message + "Error al generar archivo PDF" + "\n"
                    message = message + "Error: " + str(e) + "\n"
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