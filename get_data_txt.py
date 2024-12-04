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


def get_data_report(ruta_archivo, separator):
    message = f"inicio lectura archivo de reportetxt: {ruta_archivo}\n"
    data_report = []
    try:
        # Intenta con UTF-8; si falla, usa una alternativa
        with open(ruta_archivo, 'r', encoding='utf-8', errors='replace') as archivo:
            lines = archivo.readlines()
            if lines:
                lines.pop(0)  # Elimina la primera línea
            for line in lines:
                datos_line = line.strip().split(separator)
                hoja = datos_line.pop(0)
                data = {hoja: datos_line}
                # print("data: ", data, "\n")
                data_report.append(data)
    except Exception as e:
        message += f"Error al leer el archivo de reporte: {ruta_archivo}\n"
        message += f"Error: {str(e)}\n"
    return data_report, message

def get_path(ruta_archivo):
    with open(ruta_archivo, 'r') as archivo:
        primera_linea = archivo.readline().strip()   
        return primera_linea

