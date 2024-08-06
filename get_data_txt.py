def get_routs(routAplication, pos):
    rout_rutas = routAplication + '\\rutas.txt'
    pos = int(pos)
    with open(rout_rutas, 'r') as archivo:
        for index, line in enumerate(archivo, start=1):
            if index == pos:
                return line.strip()  # Devuelve la línea en la posición solicitada sin espacios adicionales
        return None


def get_name_template(ruta_archivo,separator):
    with open(ruta_archivo, 'r') as archivo:
        primera_linea = archivo.readline().strip()  
        datos = primera_linea.split(separator)  
        return datos[0]


def get_data_report(ruta_archivo,separator):
    message = "inicio lectura archivo de reportetxt: " + ruta_archivo +"\n"
    try:
        with open(ruta_archivo, 'r') as archivo:
            lines = archivo.readlines()
            lines.pop(0)
            data_report = []
            for line in lines:
                datos_line = line.strip().split(separator)
                hoja = datos_line.pop(0)
                data = {
                    hoja : datos_line
                }
                data_report.append(data)
    except Exception as e:
        message = message + "Error al leer el archivo de reporte: " + ruta_archivo + "\n"
        message = message + "Error: " + str(e) + "\n"
    return data_report, message
