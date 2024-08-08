import win32com.client
import pythoncom
import os

def extract_images_from_excel(file_path):
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)
    images = []

    for sheet in workbook.Sheets:
        for shape in sheet.Shapes:
            if shape.Type == 13:  # msoPicture
                # Guardar la imagen en un archivo temporal
                temp_path = f"C:\\Temp\\{shape.Name}.jpg"
                shape.CopyPicture()
                sheet.PasteSpecial(Format=13)
                chart = sheet.ChartObjects(1).Chart
                chart.Export(temp_path)
                
                # Leer la imagen en memoria
                with open(temp_path, "rb") as img_file:
                    img_data = img_file.read()
                    images.append({
                        'sheet': sheet.Name,
                        'name': shape.Name,
                        'image': img_data
                    })
                
                # Eliminar el archivo temporal
                os.remove(temp_path)
                sheet.ChartObjects(1).Delete()

    workbook.Close(SaveChanges=False)
    excel.Quit()
    pythoncom.CoUninitialize()
    print("imágenes extraídas", images)
    return images

# Ejemplo de uso
file_path = "Libro2.xlsx"
images = extract_images_from_excel(file_path)

# Aquí puedes manipular las imágenes en memoria
for img_info in images:
    print(f"Hoja: {img_info['sheet']}, Nombre de la imagen: {img_info['name']}, Tamaño de la imagen: {len(img_info['image'])} bytes")
