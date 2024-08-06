from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
import shutil
import os
import random

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
                        'merge_cells': cell.coordinate in sheet.merged_cells
                    }
                    sheet_info[cell.coordinate] = cell_info

        info_excel[sheet.title] = sheet_info

    return info_excel

def reemplazar_vars(sheet_info, data):
    var_counter = 0
    for cell_info in sheet_info.values():
        if isinstance(cell_info['value'], str) and '<VAR' in cell_info['value']:
            if var_counter < len(data):
                cell_info['value'] = data[var_counter]
                var_counter += 1
    return sheet_info

def aplicar_info_a_hoja(sheet, sheet_info, start_row):
    max_row = start_row
    for coord, cell_info in sheet_info.items():
        col_letter = ''.join(filter(str.isalpha, coord))
        row_number = int(''.join(filter(str.isdigit, coord)))
        new_coord = f"{col_letter}{start_row + row_number - 1}"

        cell = sheet[new_coord]
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

    return max_row

def find_next_start_row(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == '??FIN??':
                return cell.row + 1
    return 1

def create_report_excel(datos_report, ruta_template_excel):
    message = "Inicio de la copia del archivo: " + ruta_template_excel + "\n"
    try:
        # Generar un nombre aleatorio para el nuevo archivo
        random_number = random.randint(0, 1000000)
        name_archive = os.path.splitext(ruta_template_excel)[0]
        rout_report_excel = f"{name_archive}_{random_number}.xlsx"   

        # Realizar la copia del archivo
        shutil.copy(ruta_template_excel, rout_report_excel)

        workbook = load_workbook(rout_report_excel)

        # Buscar la hoja "PRINCIPAL" o variantes, o crear una si no existe
        principal_sheet = None
        for sheet_name in ["PRINCIPAL", "principal", "Principal"]:
            if sheet_name in workbook.sheetnames:
                principal_sheet = workbook[sheet_name]
                break
        if not principal_sheet:
            principal_sheet = workbook.create_sheet(title="PRINCIPAL")

        # Obtener la información de cada hoja
        info_excel = obtener_info_excel(ruta_template_excel)

        # Obtener la fila de inicio
        start_row = find_next_start_row(principal_sheet)

        # Iterar sobre los datos de reporte y las hojas correspondientes
        for data in datos_report:
            for sheet_name, values in data.items():
                if sheet_name in info_excel:
                    sheet_info = info_excel[sheet_name]
                    # Reemplazar las variables en la hoja
                    sheet_info = reemplazar_vars(sheet_info, values)
                    # Aplicar la información modificada a la hoja "PRINCIPAL"
                    max_row = aplicar_info_a_hoja(principal_sheet, sheet_info, start_row)
                    start_row = max_row + 1

        workbook.save(rout_report_excel)
        message += "Archivo creado exitosamente: " + rout_report_excel + "\n"
    except Exception as e:
        message += "Error al copiar el archivo: " + ruta_template_excel + "\n"
        message += "Error: " + str(e) + "\n"
    return message
