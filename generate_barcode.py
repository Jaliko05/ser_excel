from barcode import Code128
from barcode.writer import ImageWriter


def generate_barcode(valor, ruta_imagen):
    barcode = Code128(valor, writer=ImageWriter())
    barcode.save(ruta_imagen)