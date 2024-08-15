from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
import win32com.client as win32
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import os
import random
import copy

def obtener_info_excel(ruta_excel):
    workbook = load_workbook(ruta_excel)
    info_excel = {}

    for sheet in workbook.worksheets:
        sheet_info = {}
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    font_color = cell.font.color.rgb if cell.font.color and cell.font.color.type == 'rgb' else None
                    fill_color = cell.fill.fgColor.rgb if cell.fill.fgColor and cell.fill.fgColor.type == 'rgb' else None

                    # Obtener la configuración de la página
                    page_setup = {
                        'orientation': sheet.page_setup.orientation,
                        'paper_size': sheet.page_setup.paperSize,
                        'fit_to_width': sheet.page_setup.fitToWidth,
                        'fit_to_height': sheet.page_setup.fitToHeight,
                        'scale': sheet.page_setup.scale,
                        'margin_top': sheet.page_margins.top,
                        'margin_bottom': sheet.page_margins.bottom,
                        'margin_left': sheet.page_margins.left,
                        'margin_right': sheet.page_margins.right,
                        'print_area': sheet.print_area
                    }

                    # Obtener el ancho de la columna y la altura de la fila
                    #column_width = sheet.column_dimensions[cell.column_letter].width
                    row_height = sheet.row_dimensions[cell.row].height

                    cell_info = {
                        'value': cell.value,
                        'font': {
                            'name': cell.font.name,
                            'size': cell.font.sz,
                            'bold': cell.font.b,
                            'italic': cell.font.i,
                            'underline': cell.font.u,
                            'color': font_color
                        },
                        'fill': {
                            'fgColor': fill_color,
                            'patternType': cell.fill.patternType
                        },
                        'border': {
                            'left': cell.border.left.style,
                            'right': cell.border.right.style,
                            'top': cell.border.top.style,
                            'bottom': cell.border.bottom.style
                        },
                        'alignment': {
                            'horizontal': cell.alignment.horizontal,
                            'vertical': cell.alignment.vertical,
                            'wrap_text': cell.alignment.wrap_text
                        },
                        'number_format': cell.number_format,
                        'row': cell.row,
                        'column': cell.column,
                        'merge_cells': cell.coordinate in sheet.merged_cells,
                        # 'column_width': column_width,
                        'row_height': row_height
                    }
                    sheet_info[cell.coordinate] = cell_info

        merge_info = []
        for merged_cell in sheet.merged_cells.ranges:
            merge_info.append(str(merged_cell))

        info_excel[sheet.title] = {'cells': sheet_info, 'merges': merge_info, 'page_setup': page_setup}

    return info_excel


def es_numero(valor):
    # Verificar si el valor es un número que se pueda convertir a float
    try:
        # Intentar convertir a float, reemplazando coma por punto si es necesario
        float(valor.replace(',', '.'))
        return True
    except ValueError:
        return False

def reemplazar_vars(sheet_info, data):
    # Hacer una copia profunda del sheet_info original
    sheet_info_copia = copy.deepcopy(sheet_info)
    
    for var_counter, value in enumerate(data, start=1):
        var_placeholder = f'<VAR{var_counter:03}>'
        
        # Verificar si el valor es un número
        if es_numero(value):
            # Diferenciar entre enteros y flotantes
            if value.isdigit():
                value = int(value)  # Convertir a entero si es un número entero
            else:
                try:
                    value = float(value.replace(',', '.'))  # Convertir a float si es un número decimal
                except ValueError:
                    pass  # Si no se puede convertir, dejar el valor como está
        
        for cell_info in sheet_info_copia['cells'].values():
            if isinstance(cell_info['value'], str) and var_placeholder in cell_info['value']:
                # Reemplazar solo la variable específica sin afectar el resto del contenido del campo
                cell_info['value'] = cell_info['value'].replace(var_placeholder, str(value))
                
    return sheet_info_copia

def aplicar_info_a_hoja(sheet, sheet_info, start_row, sheet_name):
    max_row = start_row
    for coord, cell_info in sheet_info['cells'].items():
        col_letter = ''.join(filter(str.isalpha, coord))
        row_number = int(''.join(filter(str.isdigit, coord)))
        new_coord = f"{col_letter}{start_row + row_number - 1}"

        # Aplicar la altura de la fila
        if 'row_height' in cell_info and cell_info['row_height'] is not None:
            sheet.row_dimensions[start_row + row_number - 1].height = cell_info['row_height']


        cell = sheet[new_coord]
        if cell_info['value'] != '??FIN??':  # Evitar escribir ??FIN??
            cell.value = cell_info['value']
            cell.font = Font(
                name=cell_info['font']['name'],
                size=cell_info['font']['size'],
                bold=cell_info['font']['bold'],
                italic=cell_info['font']['italic'],
                underline=cell_info['font']['underline'],
                color=Color(rgb=cell_info['font']['color']) if cell_info['font']['color'] else None
            )
            cell.fill = PatternFill(
                fgColor=Color(rgb=cell_info['fill']['fgColor']) if cell_info['fill']['fgColor'] else None,
                patternType=cell_info['fill']['patternType']
            )
            cell.border = Border(
                left=Side(style=cell_info['border']['left']),
                right=Side(style=cell_info['border']['right']),
                top=Side(style=cell_info['border']['top']),
                bottom=Side(style=cell_info['border']['bottom'])
            )
            cell.alignment = Alignment(
                horizontal=cell_info['alignment']['horizontal'],
                vertical=cell_info['alignment']['vertical'],
                wrap_text=cell_info['alignment']['wrap_text']
            )
            cell.number_format = cell_info['number_format']

        if start_row + row_number - 1 > max_row:
            max_row = start_row + row_number - 1

    for merge_range in sheet_info['merges']:
        merge_start, merge_end = merge_range.split(':')
        start_col_letter = ''.join(filter(str.isalpha, merge_start))
        start_row_number = int(''.join(filter(str.isdigit, merge_start)))
        end_col_letter = ''.join(filter(str.isalpha, merge_end))
        end_row_number = int(''.join(filter(str.isdigit, merge_end)))

        new_merge_start = f"{start_col_letter}{start_row + start_row_number - 1}"
        new_merge_end = f"{end_col_letter}{start_row + end_row_number - 1}"
        sheet.merge_cells(f"{new_merge_start}:{new_merge_end}")

    return max_row

def find_next_start_row(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == '??FIN??':
                return cell.row + 1
    return 1

def get_image_position(rout_template_excel):
    excel = win32.Dispatch("Excel.Application")
    archivo_excel = os.path.abspath(rout_template_excel)
    wb_win32 = excel.Workbooks.Open(archivo_excel)

    # Diccionario para almacenar las posiciones de las imágenes
    posiciones_imagenes = {}

    # Recorrer cada hoja para obtener las posiciones de las imágenes
    for sheet in wb_win32.Sheets:
        posiciones_imagenes[sheet.Name] = []
        for shape in sheet.Shapes:
            if shape.Type == 13:  # El tipo 13 es msoPicture
                # Obtener la posición de la imagen
                left = shape.Left
                top = shape.Top

                # Guardar la posición y el nombre del archivo
                posiciones_imagenes[sheet.Name].append({
                    "name": shape.Name,
                    "left": left,
                    "top": top
                })
    
    # Cerrar el libro original y la aplicación Excel
    wb_win32.Close(False)
    excel.Quit()
    return posiciones_imagenes

def copy_column_widths(origen, destino):
    # Obtener la última columna con datos en la hoja de origen
    max_col = origen.max_column
    ajuste = 0.5  # Ajustar el ancho de las columnas según el contenido
    # Iterar sobre cada columna hasta la última columna con datos
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        # Obtener el ancho de la columna en la hoja de origen
        origen_ancho = origen.column_dimensions[col_letter].width
         # Ajustar el ancho restando un valor específico
        if origen_ancho is not None:
            destino.column_dimensions[col_letter].width = max(0, origen_ancho - ajuste)

def create_report_excel(datos_report, ruta_template_excel, ruta_report_excel):
    message = "Inicio de la copia del archivo: " + ruta_template_excel + "\n"
    try: 
        wb = load_workbook(ruta_template_excel)

        # Buscar la hoja "PRINCIPAL" o variantes, 
        principal_sheet = None
        for sheet_name in ["PRINCIPAL", "principal", "Principal"]:
            if sheet_name in wb.sheetnames:
                principal_sheet = wb[sheet_name]
                break
        if not principal_sheet:
            principal_sheet = wb.create_sheet(title="PRINCIPAL")
        message = message + "Hoja principal: " + str(principal_sheet) + "\n"

        # Obtener la fila de inicio
        start_row = find_next_start_row(principal_sheet)
        message = message + "Obtener fila inicial exitosamente: "  + "\n"


        # Obtener la información de cada hoja de la plantilla
        info_excel = obtener_info_excel(ruta_template_excel)
        message = message + "Obtener informacion de plantilla exitosamente: "  + "\n"

        #Obtener las posiciones de las imagenes
        posiciones_imagenes = get_image_position(ruta_template_excel)
        message = message + "Obtener posiciones de las imagenes exitosamente: "  + "\n"
        
        #iniciar libro de excel
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = principal_sheet.title
        
        #copiar ancho de columnas de la hoja 001 al reporte
        copy_column_widths(wb["001"], sheet)
        message = message + "Copiar ancho de columnas de la hoja 001 al reporte exitosamente: "  + "\n"

        #aplicar informacion de la hoja principal
        a = aplicar_info_a_hoja(principal_sheet, info_excel[principal_sheet.title], 1, principal_sheet.title)   
        message = message + "Aplicar informacion de la hoja principal exitosamente: "  + "\n"

        #aplicar imagenes de la hoja principal a la nueva hoja principal
        if posiciones_imagenes[principal_sheet.title]:
            sheet_template = wb[principal_sheet.title]
            for img_info, image in zip(posiciones_imagenes[principal_sheet.title], sheet_template._images):
                new_image = Image(image.ref)

                # Convertir las posiciones a celdas aproximadas
                cell_row = int(img_info['top'] / 18)  # Ajustar según la altura de la fila
                cell_col = int(img_info['left'] / 64)  # Ajustar según el ancho de la columna
                cell_position = f"{chr(65 + cell_col)}{cell_row + start_row}" 
                # Insertar la imagen en la nueva hoja
                sheet.add_image(new_image, cell_position)
        message = message + "Aplicar imagenes de la hoja principal exitosamente: "  + "\n"

        # Iterar sobre los datos de reporte y las hojas correspondientes
        for data in datos_report:
            for sheet_name, values in data.items():
                if sheet_name in info_excel:
                    sheet_info = info_excel[sheet_name]
                    # Reemplazar las variables en la hoja
                    sheet_info_modificada = reemplazar_vars(sheet_info, values)
                    # Aplicar la información modificada a la hoja "PRINCIPAL"
                    max_row = aplicar_info_a_hoja(sheet, sheet_info_modificada, start_row, sheet_name)

                    # Aplicar las imagenes a la hoja
                    if posiciones_imagenes[sheet_name]:
                        sheet_template = wb[sheet_name]
                        if sheet_name in posiciones_imagenes:
                            for img_info, image in zip(posiciones_imagenes[sheet_name], sheet_template._images):
                                new_image = Image(image.ref)

                                # Convertir las posiciones a celdas aproximadas
                                cell_row = int(img_info['top'] / 18)  # Ajustar según la altura de la fila
                                cell_col = int(img_info['left'] / 64)  # Ajustar según el ancho de la columna
                                cell_position = f"{chr(65 + cell_col)}{cell_row + start_row}"  
                                # Insertar la imagen en la nueva hoja
                                sheet.add_image(new_image, cell_position) 


                    start_row = max_row
        #aplicar formato de las hojas
        sheet_info = info_excel['001'] 
        if 'page_setup' in sheet_info:
            page_setup = sheet_info['page_setup']
            sheet.page_setup.orientation = page_setup['orientation']
            sheet.page_setup.paperSize = page_setup['paper_size']
            sheet.page_setup.fitToWidth = page_setup['fit_to_width']
            sheet.page_setup.fitToHeight = page_setup['fit_to_height']
            sheet.page_setup.scale = page_setup['scale']
            sheet.page_margins.top = page_setup['margin_top']
            sheet.page_margins.bottom = page_setup['margin_bottom']
            sheet.page_margins.left = page_setup['margin_left']
            sheet.page_margins.right = page_setup['margin_right']
            sheet.print_area = page_setup['print_area']
            
        message = message + "aplicar formato de las hojas exitosamente: "  + "\n"
        workbook.save(ruta_report_excel)
        message += "Archivo creado exitosamente: " + ruta_report_excel + "\n"
    except Exception as e:
        message += "Error al crear el reporte: " + ruta_template_excel + "\n"
        message += "Error: " + str(e) + "\n"
    return message, ruta_report_excel, principal_sheet.title
