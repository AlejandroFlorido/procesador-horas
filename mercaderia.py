"""
BRIE › Mercadería retirada por el personal (paso 3 del flujo Liquidar → Adelantos → Mercadería).

Lee el Excel "Merc.a Personal ... CON DESCUENTO" (siempre el mismo formato):
  - Bloques por empleado: fila "Empleado:" (col A) con el nombre (col B) y el código E000xx (col J).
  - Líneas de productos: Fecha (B), Artículo, Descripción, Rubro, Cantidad, SubTotal (G),
    % Desc. (H), Descuento (I), SubTotal c/Desc. (J).
  - Fila "SubTotal Empleado" (col E) y al final "Total General:".
Lo que importa es el nombre del empleado y su SubTotal c/Desc. (los descuentos se cargan a mano
en el archivo, por eso se respeta lo que diga cada línea).

Reglas:
  - El adelanto puede superar el sueldo: se avisa y queda en amarillo (a cobrar negativo).
  - Adelanto + mercadería también puede superar el sueldo: el "A cobrar" queda negativo.
    No bloquear; avisar igual que en Adelantos.
"""
import json
import re

from openpyxl import load_workbook

COL_A, COL_B, COL_C, COL_E, COL_G, COL_H, COL_I, COL_J = 0, 1, 2, 4, 6, 7, 8, 9


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _porcentaje(formula, pct_general):
    """Resuelve la columna % Desc.: un número, o una referencia a la celda del descuento general."""
    n = _num(formula)
    if n is not None:
        return n
    if isinstance(formula, str) and re.fullmatch(r"=\$?C\$?4", formula.strip(), re.I):
        return pct_general
    return None


def procesar(ruta):
    """Devuelve un JSON con el total con descuento de cada empleado del archivo."""
    wb_val = load_workbook(ruta, data_only=True)    # valores calculados (si Excel los guardó)
    wb_for = load_workbook(ruta)                    # fórmulas, por si faltan los valores
    hv, hf = wb_val.worksheets[0], wb_for.worksheets[0]
    filas_v = list(hv.iter_rows(values_only=True))
    filas_f = list(hf.iter_rows(values_only=True))

    titulo = next((str(c) for c in (filas_v[0] if filas_v else []) if c), "")
    m = re.search(r"Desde el (\S+) Hasta el (\S+)", titulo)
    periodo = f"{m.group(1)} al {m.group(2)}" if m else ""
    local = str(filas_v[2][0]) if len(filas_v) > 2 and filas_v[2][0] else ""
    pct_general = 0.0
    for fila in filas_v[:10]:
        if fila and str(fila[COL_A] or "").startswith("Descuento aplicado") and _num(fila[COL_C]) is not None:
            pct_general = _num(fila[COL_C])

    empleados, actual, avisos, total_general = [], None, [], None

    def cerrar(subtotal_fila_v):
        if not actual:
            return
        calculado = round(sum(actual.pop("_lineas")), 2)
        guardado = _num(subtotal_fila_v[COL_J]) if subtotal_fila_v else None
        actual["total"] = round(guardado if guardado is not None else calculado, 2)
        actual["original"] = round(_num(subtotal_fila_v[COL_G]) if subtotal_fila_v and _num(subtotal_fila_v[COL_G]) is not None
                                   else actual.pop("_original", 0), 2)
        actual.pop("_original", None)
        actual["descuento"] = round(actual["original"] - actual["total"], 2)
        if guardado is not None and abs(guardado - calculado) > 0.05:
            avisos.append(f"{actual['nombre']}: el subtotal del archivo ({guardado:.2f}) no coincide con la suma de sus líneas ({calculado:.2f}); se usó el del archivo")
        empleados.append(actual)

    for fv, ff in zip(filas_v, filas_f):
        a = str(fv[COL_A] or "").strip()
        e = str(fv[COL_E] or "").strip()
        if a.startswith("Empleado:"):
            if actual and "_lineas" in actual:   # bloque anterior sin fila "SubTotal Empleado"
                cerrar(None)
            actual = {"codigo": str(fv[COL_J] or "").strip(), "nombre": str(fv[COL_B] or "").strip(),
                      "items": 0, "_lineas": [], "_original": 0.0}
        elif e.startswith("SubTotal Empleado") and actual and "_lineas" in actual:
            cerrar(fv)
            actual = None
        elif e.startswith("Total General"):
            total_general = _num(fv[COL_J])
        elif actual and hasattr(fv[COL_B], "year") and _num(fv[COL_G]) is not None:
            sub = _num(fv[COL_G])
            con_desc = _num(fv[COL_J])
            if con_desc is None:     # sin valor guardado: calcular igual que la fórmula del archivo
                desc = _num(fv[COL_I])
                if desc is None:
                    pct = _porcentaje(ff[COL_H], pct_general) or 0
                    desc = round(sub * pct, 2)
                con_desc = sub - desc
            actual["_lineas"].append(con_desc)
            actual["_original"] += sub
            actual["items"] += 1
    if actual and "_lineas" in actual:
        cerrar(None)

    suma = round(sum(x["total"] for x in empleados), 2)
    if total_general is not None and abs(total_general - suma) > 0.05:
        avisos.append(f"La suma por empleado ({suma:.2f}) no coincide con el Total General del archivo ({total_general:.2f})")
    if not empleados:
        raise ValueError("No se encontraron empleados. ¿Es el archivo de Mercadería a Personal con descuento?")

    return json.dumps({"periodo": periodo, "local": local, "descuento_general": pct_general,
                       "total": suma, "empleados": empleados, "avisos": avisos}, ensure_ascii=False)
