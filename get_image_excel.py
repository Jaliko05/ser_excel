from openpyxl import Workbook
from openpyxl.drawing.image import Image
import xlwings as xw
import os

# Paso 1: Abrir el libro de Excel original
archivo_excel = "PSRH004.xlsx"
wb = xw.Book(archivo_excel)

# Variable para almacenar las imágenes encontradas y su posición
imagenes_encontradas = {}

# Paso 2: Recorrer cada hoja para encontrar imágenes
for hoja in wb.sheets:
    # Obtener todas las imágenes en la hoja
    imagenes = hoja.pictures

    # Verificar si hay imágenes y almacenar la imagen junto con su posición
    for imagen in imagenes:
        if hoja.name not in imagenes_encontradas:
            imagenes_encontradas[hoja.name] = []

        # Almacenar la imagen y su posición
        imagenes_encontradas[hoja.name].append({
            "imagen": imagen,
            "posicion_x": imagen.left,
            "posicion_y": imagen.top
        })

# Cerrar el libro original
wb.close()

# Paso 3: Crear un nuevo archivo de Excel para guardar las imágenes
new_workbook = Workbook()
new_sheet = new_workbook.active
new_sheet.title = "Imágenes Encontradas"

# Paso 4: Insertar imágenes encontradas en el nuevo archivo Excel
for hoja, imagenes in imagenes_encontradas.items():
    for idx, img in enumerate(imagenes):
        # Guardar la imagen temporalmente
        temp_image_path = f"temp_image_{hoja}_{idx}.png"
        img['imagen'].export(temp_image_path)

        # Insertar la imagen en la nueva hoja de Excel
        img_to_insert = Image(temp_image_path)
        cell_position = f"A{idx + 1}"  # Cambia la celda según sea necesario
        new_sheet.add_image(img_to_insert, cell_position)

        # Eliminar la imagen temporal
        os.remove(temp_image_path)

# Paso 5: Guardar el nuevo archivo de Excel
new_workbook.save("imagenes_encontradas.xlsx")
