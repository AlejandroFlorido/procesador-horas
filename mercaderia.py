"""
BRIE › Mercadería retirada por el personal.

Sección preparada, todavía sin lógica. Se carga en el navegador igual que
lectorExcel_IA.py, así que acá se puede usar pandas y openpyxl.

Reglas ya definidas:
  - El adelanto puede superar el sueldo: se avisa y queda en amarillo (a cobrar negativo).
  - Adelanto + mercadería también puede superar el sueldo: el "A cobrar" queda
    negativo (el empleado debe la diferencia). No bloquear; avisar igual que en Adelantos.

Cuando esté definida:
  1. Escribir la lógica en procesar() y exportar_excel().
  2. Armar la pantalla en index.html, dentro de <div id="contenido-mercaderia">.
  3. Poner ACTIVA: true para "mercaderia" en SECCIONES_BRIE (index.html).
"""
import json


def liquidacion():
    """DataFrame de la liquidación de horas ya procesada en el paso 1."""
    import web
    return web._df


def procesar(ruta):
    """
    Paso del flujo BRIE: Liquidar → Adelantos → Mercadería.
    Trabaja sobre la liquidación (paso 1) y los adelantos (paso 2).
    Recibe la ruta del archivo cargado (si hace falta uno) y devuelve un JSON para la página.
    """
    raise NotImplementedError("Sección mercaderia todavía no implementada")


def exportar_excel(ruta_salida):
    """Genera el Excel para descargar."""
    raise NotImplementedError("Sección mercaderia todavía no implementada")
