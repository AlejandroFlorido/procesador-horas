"""
BRIE › Adelantos de sueldo del personal.

Sección preparada, todavía sin lógica. Se carga en el navegador igual que
lectorExcel_IA.py, así que acá se puede usar pandas y openpyxl.

Cuando esté definida:
  1. Escribir la lógica en procesar() y exportar_excel().
  2. Armar la pantalla en index.html, dentro de <div id="contenido-adelantos">.
  3. Poner ACTIVA: true para "adelantos" en SECCIONES_BRIE (index.html).
"""
import json


def procesar(ruta):
    """Recibe la ruta del archivo cargado y devuelve un JSON para mostrar en la página."""
    raise NotImplementedError("Sección adelantos todavía no implementada")


def exportar_excel(ruta_salida):
    """Genera el Excel para descargar."""
    raise NotImplementedError("Sección adelantos todavía no implementada")
