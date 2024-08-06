def reemplazar_vars(sheet_info, data):
    var_counter = 0
    for cell_info in sheet_info.values():
        if isinstance(cell_info['value'], str) and '<VAR' in cell_info['value']:
            if var_counter < len(data):
                cell_info['value'] = data[var_counter]
                var_counter += 1
    return sheet_info