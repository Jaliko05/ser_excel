def aplicar_info_a_hoja(sheet, sheet_info):
    for coord, cell_info in sheet_info.items():
        cell = sheet[coord]
        cell.value = cell_info['value']
        cell.font = cell.font.copy(
            name=cell_info['font']['name'],
            size=cell_info['font']['size'],
            bold=cell_info['font']['bold'],
            italic=cell_info['font']['italic'],
            underline=cell_info['font']['underline'],
            color=cell_info['font']['color']
        )
        cell.fill = cell.fill.copy(
            fgColor=cell_info['fill']['fgColor'],
            patternType=cell_info['fill']['patternType']
        )
        cell.border = cell.border.copy(
            left=cell_info['border']['left'],
            right=cell_info['border']['right'],
            top=cell_info['border']['top'],
            bottom=cell_info['border']['bottom']
        )
        cell.alignment = cell.alignment.copy(
            horizontal=cell_info['alignment']['horizontal'],
            vertical=cell_info['alignment']['vertical'],
            wrap_text=cell_info['alignment']['wrap_text']
        )
        cell.number_format = cell_info['number_format']