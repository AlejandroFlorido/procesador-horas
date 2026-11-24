import csv
import os
import re
from datetime import datetime, date, time, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
#  CONFIGURACIÓN (podés ajustar estos valores)
# =====================================================================
HORAS_TURNO_EXTENDIDO = 8      # más de esto -> "Turno extendido"
MINUTOS_TOLERANCIA_EXTENDIDO = 90 # tolerancia: con 90 recién se marca como extendido pasadas las 9:30 h
HORAS_TURNO_SOSPECHOSO = 12    # más de esto -> "Turno sospechoso"
HORAS_MAX_PAR = 16             # más de esto entre dos marcas -> no son el mismo turno
MINUTOS_MARCA_REPETIDA = 10    # dos marcas separadas por <= esto -> se ignora la segunda
MINUTOS_TURNO_CORTO = 30       # turno de menos de esto -> "Turno muy corto"

MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
         "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12}

# Niveles de alerta (se usan para colores en pantalla y en el Excel)
OK, EXTENDIDO, SOSPECHOSO, CORTO, SIN_PAREJA, REPETIDA = (
    "ok", "extendido", "sospechoso", "corto", "sin_pareja", "repetida")
NIVELES_A_REVISAR = {SOSPECHOSO, CORTO, SIN_PAREJA}


# =====================================================================
#  LECTURA Y PARSEO
# =====================================================================
# Nombres de columna del formato nuevo del reloj (reporte "Transaction"), en minúscula
COLUMNAS_TABLA = {
    "id": ("id",),
    "nombre": ("first name", "nombre"),
    "apellido": ("last name", "apellido"),
    "fecha": ("date", "fecha"),
    "hora": ("time", "hora"),
}


def _filas_crudas(ruta):
    """Devuelve las filas del archivo como listas de celdas (texto), sea .xlsx, .xls o .csv."""
    if ruta.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(ruta, header=None, dtype=str)
        return [[str(v).strip() if pd.notna(v) else "" for v in fila] for fila in df.itertuples(index=False)]
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, encoding=codificacion) as f:
                texto = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("No se pudo leer la codificación del archivo")
    separador = ";" if texto.count(";") > texto.count(",") else ","
    return [next(csv.reader([l], delimiter=separador)) if l.strip() else [] for l in texto.splitlines()]


def _buscar_cabecera(filas):
    """Si el archivo es una tabla con columnas (formato nuevo), devuelve (fila_cabecera, {campo: columna})."""
    for n, fila in enumerate(filas[:30]):
        celdas = [c.strip().lower() for c in fila]
        cols = {}
        for campo, nombres in COLUMNAS_TABLA.items():
            for i, c in enumerate(celdas):
                if c in nombres:
                    cols[campo] = i
                    break
        if {"id", "fecha", "hora"} <= cols.keys():
            return n, cols
    return None, None


def _limpio(txt):
    """Solo saca espacios de más: el nombre queda tal cual se registró en el reloj."""
    return " ".join(str(txt).split())


def leer_lineas(ruta):
    """
    Devuelve una lista con una entrada por fichada:
      - formato nuevo (tabla con columnas First Name, Last Name, ID, Department, Date, Time):
        dict ya interpretado, o ("error", nro_linea, texto) si la fila no se pudo leer;
      - formato viejo ('74 Nicolas Aguirre 22-oct.-2026 16:05): texto de la línea.
    Las líneas de encabezado del formato nuevo ("Transaction", "Export Time: ...") se ignoran.
    """
    filas = _filas_crudas(ruta)
    cab, cols = _buscar_cabecera(filas)
    if cab is None:   # formato viejo: una sola columna con todo junto
        return [" ".join(c for c in fila if c) for fila in filas]

    registros = []
    for n, fila in enumerate(filas[cab + 1:], start=cab + 2):
        if not any(c.strip() for c in fila):
            continue
        celda = lambda campo: fila[cols[campo]].strip() if campo in cols and cols[campo] < len(fila) else ""
        try:
            fecha = parsear_fecha(celda("fecha"))
            hora = parsear_hora(celda("hora"))
            id_txt = celda("id").lstrip("'")
            if not id_txt:
                raise ValueError("sin ID")
            registros.append({
                "id": int(id_txt) if id_txt.isdigit() else id_txt,
                "nombre": _limpio(celda("nombre")),
                "apellido": _limpio(celda("apellido").strip("-")),
                "fecha": fecha, "hora": hora, "dt": datetime.combine(fecha, hora),
            })
        except Exception:
            registros.append(("error", n, ", ".join(fila)))
    return registros


def parsear_fecha(txt):
    t = txt.lower().replace(".", "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", t)          # 2026-11-01 (formato nuevo)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{1,2})-([a-zñ]+)-(\d{4})", t)
    if m and m.group(2)[:3] in MESES:
        return date(int(m.group(3)), MESES[m.group(2)[:3]], int(m.group(1)))
    return pd.to_datetime(txt, dayfirst=True).date()


def parsear_hora(txt):
    partes = [int(p) for p in txt.strip().split(" ")[-1].split(":")]
    return time(partes[0], partes[1], partes[2] if len(partes) > 2 else 0)


def parsear_linea(linea):
    """'74 Nicolas Aguirre 22-oct.-2026 16:05 -> dict. Lanza ValueError si no se puede."""
    tok = linea.split()
    if len(tok) < 4:
        raise ValueError("faltan datos")
    fecha, hora = parsear_fecha(tok[-2]), parsear_hora(tok[-1])
    id_txt = tok[0].lstrip("'")
    nombres = tok[1:-2]
    return {
        "id": int(id_txt) if id_txt.isdigit() else id_txt,
        "nombre": nombres[0] if nombres else "",
        "apellido": " ".join(nombres[1:]),
        "fecha": fecha,
        "hora": hora,
        "dt": datetime.combine(fecha, hora),
    }


# =====================================================================
#  LÓGICA DE TURNOS
# =====================================================================
def _costo_par(horas):
    """Costo de emparejar dos marcas como un turno. None = imposible."""
    if horas > HORAS_MAX_PAR:
        return None
    if horas > HORAS_TURNO_SOSPECHOSO:
        return 6
    if horas * 60 < MINUTOS_TURNO_CORTO:
        return 3
    return 0


COSTO_SIN_PAREJA = 10


def emparejar(marcas):
    """
    Decide qué marcas forman turnos (entrada -> salida) minimizando errores.
    Así, si alguien se olvida UNA marca, no se corren todos los turnos siguientes:
    esa marca queda sola y el resto se empareja bien.
    Devuelve lista de tuplas: ('par', i, j) o ('solo', i).
    """
    n = len(marcas)
    dp = [0] * (n + 2)
    eleccion = [None] * (n + 1)
    for i in range(n - 1, -1, -1):
        mejor, opc = COSTO_SIN_PAREJA + dp[i + 1], "solo"
        if i + 1 < n:
            c = _costo_par((marcas[i + 1]["dt"] - marcas[i]["dt"]).total_seconds() / 3600)
            if c is not None and c + dp[i + 2] <= mejor:
                mejor, opc = c + dp[i + 2], "par"
        dp[i], eleccion[i] = mejor, opc
    resultado, i = [], 0
    while i < n:
        if eleccion[i] == "par":
            resultado.append(("par", i, i + 1))
            i += 2
        else:
            resultado.append(("solo", i))
            i += 1
    return resultado


def _min_del_dia(t):
    return t.hour * 60 + t.minute


def _distancia_circular(a, b):
    d = abs(a - b) % 1440
    return min(d, 1440 - d)


def _fmt_horas(td):
    total_min = int(round(td.total_seconds() / 60))
    return f"{total_min // 60}:{total_min % 60:02d}"


def procesar_empleado(marcas):
    """Recibe las marcas de UN empleado y devuelve las filas con sus observaciones."""
    marcas = sorted(marcas, key=lambda m: m["dt"])

    # 1) Marcas repetidas
    validas, filas = [], []
    for m in marcas:
        fila = dict(m, tipo="", horas=None, nivel=OK, obs="", par=None)
        if validas and (m["dt"] - validas[-1]["dt"]) <= timedelta(minutes=MINUTOS_MARCA_REPETIDA):
            dif = int((m["dt"] - validas[-1]["dt"]).total_seconds() // 60)
            fila["tipo"], fila["nivel"] = "repetida", REPETIDA
            fila["obs"] = ("Marca duplicada exacta (ignorada)" if dif == 0
                           else f"Marca repetida a los {dif} min (ignorada)")
        else:
            validas.append(fila)
        filas.append(fila)

    # 2) Emparejar entradas y salidas
    pares = emparejar(validas)
    entradas_tipicas, salidas_tipicas = [], []
    for num, p in enumerate(pares):
        if p[0] != "par":
            continue
        e, s = validas[p[1]], validas[p[2]]
        dur = s["dt"] - e["dt"]
        h = dur.total_seconds() / 3600
        e["tipo"], s["tipo"] = "entrada", "salida"
        e["par"] = s["par"] = num
        s["horas"] = dur
        entradas_tipicas.append(_min_del_dia(e["hora"]))
        salidas_tipicas.append(_min_del_dia(s["hora"]))
        if h > HORAS_TURNO_SOSPECHOSO:
            nivel, obs = SOSPECHOSO, f"Turno sospechoso ({_fmt_horas(dur)} h) - revisar marcas"
        elif h > HORAS_TURNO_EXTENDIDO + MINUTOS_TOLERANCIA_EXTENDIDO / 60:
            nivel, obs = EXTENDIDO, f"Turno extendido ({_fmt_horas(dur)} h)"
        elif h * 60 < MINUTOS_TURNO_CORTO:
            nivel, obs = CORTO, f"Turno muy corto ({_fmt_horas(dur)} h) - revisar marcas"
        else:
            nivel, obs = OK, ""
        e["nivel"] = s["nivel"] = nivel
        s["obs"] = obs

    # 3) Marcas sin pareja: adivinamos si parece entrada o salida según sus horarios habituales
    for p in pares:
        if p[0] != "solo":
            continue
        f = validas[p[1]]
        f["tipo"], f["nivel"] = "sin_pareja", SIN_PAREJA
        obs = "Marca sin pareja - falta entrada o salida"
        if entradas_tipicas:
            mm = _min_del_dia(f["hora"])
            d_e = min(_distancia_circular(mm, x) for x in entradas_tipicas)
            d_s = min(_distancia_circular(mm, x) for x in salidas_tipicas)
            if d_e + 60 < d_s:
                obs = "Marca sin pareja - parece una ENTRADA, falta la salida"
            elif d_s + 60 < d_e:
                obs = "Marca sin pareja - parece una SALIDA, falta la entrada"
        f["obs"] = obs
    return filas


def procesar(lineas):
    """Devuelve (DataFrame con todas las filas procesadas, lista de líneas que no se pudieron leer)."""
    marcas_por_id, orden, invalidas = {}, [], []
    for n, linea in enumerate(lineas, start=1):
        if isinstance(linea, tuple):          # fila del formato nuevo que no se pudo leer
            invalidas.append((linea[1], linea[2]))
            continue
        if isinstance(linea, dict):           # fila del formato nuevo ya interpretada
            m = linea
        else:
            if not linea.strip():
                continue
            try:
                m = parsear_linea(linea)
            except Exception:
                invalidas.append((n, linea.strip()))
                continue
        if m["id"] not in marcas_por_id:
            marcas_por_id[m["id"]] = []
            orden.append(m["id"])
        marcas_por_id[m["id"]].append(m)

    filas = []
    for id_ in orden:
        filas.extend(procesar_empleado(marcas_por_id[id_]))
    return pd.DataFrame(filas), invalidas


def resumen(df):
    """Totales y conteo de alertas por empleado."""
    res = []
    for id_, g in df.groupby("id", sort=False):
        horas = g["horas"].dropna()
        res.append({
            "id": id_, "nombre": g["nombre"].iloc[0], "apellido": g["apellido"].iloc[0],
            "total": sum(horas, timedelta()),
            "turnos": len(horas),
            "extendidos": int((g["nivel"].eq(EXTENDIDO) & g["tipo"].eq("salida")).sum()),
            "sospechosos": int((g["nivel"].eq(SOSPECHOSO) & g["tipo"].eq("salida")).sum()),
            "cortos": int((g["nivel"].eq(CORTO) & g["tipo"].eq("salida")).sum()),
            "sin_pareja": int(g["tipo"].eq("sin_pareja").sum()),
            "repetidas": int(g["tipo"].eq("repetida").sum()),
        })
    return pd.DataFrame(res)


# =====================================================================
#  EXPORTAR A EXCEL (con fórmulas, formato [h]:mm:ss y colores)
# =====================================================================
COLORES = {
    EXTENDIDO:  (PatternFill("solid", fgColor="FFC7CE"), Font(color="9C0006")),
    SOSPECHOSO: (PatternFill("solid", fgColor="C00000"), Font(color="FFFFFF", bold=True)),
    CORTO:      (PatternFill("solid", fgColor="FFEB9C"), Font(color="7F6000")),
    SIN_PAREJA: (PatternFill("solid", fgColor="F4B084"), Font(color="833C0B", bold=True)),
    REPETIDA:   (PatternFill("solid", fgColor="E7E6E6"), Font(color="808080", italic=True, strike=True)),
}
CABECERA_FILL = PatternFill("solid", fgColor="1F3864")
CABECERA_FONT = Font(color="FFFFFF", bold=True)
BORDE = Border(bottom=Side(style="thin", color="BFBFBF"))


def _a_dias(td):
    """Excel guarda las horas como fracción de día: 12 h = 0.5"""
    return td.total_seconds() / 86400


LINEA_CUADRO = Side(style="medium", color="1F3864")
LINEA_INTERNA = Side(style="hair", color="BFBFBF")


def encuadrar(ws, fila_ini, fila_fin, col_ini=1, col_fin=8):
    """Encierra un bloque de celdas en un cuadro (borde grueso afuera, líneas finas adentro)."""
    for f in range(fila_ini, fila_fin + 1):
        for c in range(col_ini, col_fin + 1):
            ws.cell(f, c).border = Border(
                left=LINEA_CUADRO if c == col_ini else LINEA_INTERNA,
                right=LINEA_CUADRO if c == col_fin else LINEA_INTERNA,
                top=LINEA_CUADRO if f == fila_ini else LINEA_INTERNA,
                bottom=LINEA_CUADRO if f == fila_fin else LINEA_INTERNA,
            )


def exportar_excel(df, ruta, sueldo_hora=None, adelantos=None, mercaderia=None):
    """
    sueldo_hora, adelantos (dict id -> monto) y mercaderia (dict id -> {"monto", "cuenta"})
    son opcionales: los usa el flujo BRIE de la web.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Horas"
    cab = ["ID", "Nombre", "Apellido", "Fecha", "Hora", "Horas turno", "Total semanal", "Observación"]
    ws.append(cab)
    for c in ws[1]:
        c.fill, c.font, c.alignment = CABECERA_FILL, CABECERA_FONT, Alignment(horizontal="center")

    fila_excel = 2
    for id_, g in df.groupby("id", sort=False):
        inicio = fila_excel
        for _, r in g.iterrows():
            ws.cell(fila_excel, 1, r["id"])
            ws.cell(fila_excel, 2, r["nombre"])
            ws.cell(fila_excel, 3, r["apellido"])
            ws.cell(fila_excel, 4, r["fecha"]).number_format = "dd/mm/yyyy"
            ws.cell(fila_excel, 5, r["hora"]).number_format = "hh:mm"
            if r["tipo"] == "salida":
                # Valor explícito (no fórmula): se ve en cualquier visor, ya calculado con fecha + hora
                c = ws.cell(fila_excel, 6, _a_dias(r["horas"]))
                c.number_format = "[h]:mm:ss"
            if r["obs"]:
                ws.cell(fila_excel, 8, r["obs"])
            if r["nivel"] in COLORES:
                fill, font = COLORES[r["nivel"]]
                for col in range(1, 9):
                    ws.cell(fila_excel, col).fill = fill
                    ws.cell(fila_excel, col).font = font
            fila_excel += 1
        fin = fila_excel - 1
        tot = ws.cell(fin, 7, _a_dias(sum(g["horas"].dropna(), timedelta())))
        tot.number_format = "[h]:mm:ss"
        tot.font = Font(bold=True, color=tot.font.color)
        encuadrar(ws, inicio, fin)   # cuadro alrededor de cada empleado
        fila_excel += 1        # fila vacía entre empleados

    for col, ancho in zip("ABCDEFGH", [7, 14, 16, 12, 8, 13, 14, 52]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"

    # --- Hoja Resumen ---
    wr = wb.create_sheet("Resumen")
    cab_r = ["ID", "Nombre", "Apellido", "Total horas", "Turnos", "Extendidos",
             "Sospechosos", "Muy cortos", "Sin pareja", "Repetidas ignoradas", "Estado"]
    if sueldo_hora:
        cab_r += ["Sueldo por hora", "Sueldo"]
        if adelantos is not None:
            cab_r += ["Adelanto"]
        if mercaderia is not None:
            cab_r += ["Mercadería", "Cuenta mercadería"]
        if adelantos is not None or mercaderia is not None:
            cab_r += ["A cobrar"]
    wr.append(cab_r)
    for c in wr[1]:
        c.fill, c.font, c.alignment = CABECERA_FILL, CABECERA_FONT, Alignment(horizontal="center", wrap_text=True)
    for i, r in enumerate(resumen(df).itertuples(index=False), start=2):
        revisar = r.sospechosos or r.cortos or r.sin_pareja
        valores = [r.id, r.nombre, r.apellido, _a_dias(r.total), r.turnos,
                   r.extendidos, r.sospechosos, r.cortos, r.sin_pareja, r.repetidas,
                   "REVISAR" if revisar else "OK"]
        for col, v in enumerate(valores, start=1):
            c = wr.cell(i, col, v)
            c.border = BORDE
            if col >= 5:
                c.alignment = Alignment(horizontal="center")
        wr.cell(i, 4).number_format = "[h]:mm:ss"
        for col, nivel in ((6, EXTENDIDO), (7, SOSPECHOSO), (8, CORTO), (9, SIN_PAREJA)):
            if valores[col - 1]:
                wr.cell(i, col).fill, wr.cell(i, col).font = COLORES[nivel]
        if sueldo_hora:
            sueldo = round(r.total.total_seconds() / 3600 * sueldo_hora, 2)
            montos, a_cobrar = [sueldo_hora, sueldo], sueldo
            if adelantos is not None:
                adelanto = float(adelantos.get(str(r.id), 0) or 0)
                montos.append(adelanto)
                a_cobrar -= adelanto
            if mercaderia is not None:
                m = mercaderia.get(str(r.id)) or {}
                monto_m = float(m.get("monto", 0) or 0)
                montos += [monto_m, m.get("cuenta", "")]
                a_cobrar -= monto_m
            if adelantos is not None or mercaderia is not None:
                montos.append(round(a_cobrar, 2))
            for k, v in enumerate(montos, start=12):
                c = wr.cell(i, k, v)
                if not isinstance(v, str):
                    c.number_format = '"$" #,##0.00'
                c.border = BORDE
            ultima = wr.cell(i, 11 + len(montos))
            ultima.font = Font(bold=True, color="9C5700" if a_cobrar < 0 else "000000")
            if a_cobrar < 0:      # a cobrar negativo: amarillo, igual que en la página
                ultima.fill = PatternFill("solid", fgColor="FFEB9C")
        estado = wr.cell(i, 11)
        estado.fill, estado.font = (COLORES[SOSPECHOSO] if revisar else
                                    (PatternFill("solid", fgColor="C6EFCE"), Font(color="006100", bold=True)))

    for col, ancho in zip("ABCDEFGHIJKLMNOPQ", [7, 14, 16, 12, 8, 11, 12, 11, 10, 12, 10, 14, 15, 14, 15, 30, 15]):
        wr.column_dimensions[col].width = ancho
    wr.row_dimensions[1].height = 30
    wr.freeze_panes = "A2"

    # Leyenda de colores
    fila_ley = wr.max_row + 3
    wr.cell(fila_ley, 2, "Referencias").font = Font(bold=True)
    leyenda = [
        (EXTENDIDO, f"Turno extendido (más de {_fmt_horas(timedelta(hours=HORAS_TURNO_EXTENDIDO, minutes=MINUTOS_TOLERANCIA_EXTENDIDO))} h, ya con tolerancia)"),
        (SOSPECHOSO, f"Turno sospechoso (más de {HORAS_TURNO_SOSPECHOSO} h) - probable marca faltante"),
        (CORTO, f"Turno muy corto (menos de {MINUTOS_TURNO_CORTO} min)"),
        (SIN_PAREJA, "Marca sin pareja - falta entrada o salida (no suma horas)"),
        (REPETIDA, f"Marca repetida dentro de {MINUTOS_MARCA_REPETIDA} min - ignorada"),
    ]
    for k, (nivel, texto) in enumerate(leyenda, start=1):
        c = wr.cell(fila_ley + k, 2, "")
        c.fill = COLORES[nivel][0]
        wr.cell(fila_ley + k, 3, texto)

    wb.save(ruta)


# =====================================================================
#  INTERFAZ
# =====================================================================
def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinterdnd2 import DND_FILES, TkinterDnD

    estado = {"procesado": None, "actual": None, "ruta": None}

    # ---------- Acciones ----------
    def seleccionar_archivo():
        archivo = filedialog.askopenfilename(title="Selecciona un archivo",
                                             filetypes=[("Archivos Excel o CSV", "*.xlsx *.xls *.csv")])
        if archivo:
            leer_archivo(archivo)

    def manejar_arrastre(data):
        # Las rutas con espacios llegan entre llaves: {C:/mis archivos/reloj.csv}
        rutas = ventana.tk.splitlist(data)
        if rutas:
            leer_archivo(rutas[0])

    def leer_archivo(ruta):
        try:
            df, invalidas = procesar(leer_lineas(ruta))
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo: {e}")
            return
        if df.empty:
            messagebox.showwarning("!!!", "No se encontraron fichadas válidas en el archivo")
            return
        estado.update(procesado=df, actual=df, ruta=ruta)
        entrada_busqueda.delete(0, tk.END)
        var_solo_revisar.set(False)
        mostrar_datos(df)
        if invalidas:
            detalle = "\n".join(f"Línea {n}: {t}" for n, t in invalidas[:15])
            messagebox.showwarning("Líneas ignoradas",
                                   f"No se pudieron leer {len(invalidas)} línea(s):\n\n{detalle}")

    def mostrar_datos(df):
        texto_datos.config(state="normal")
        texto_datos.delete(1.0, tk.END)
        anchos = [6, 14, 16, 12, 7, 8, 9]
        cab = ["ID", "Nombre", "Apellido", "Fecha", "Hora", "Horas", "Total", "Observación"]
        # Ancho interno del cuadro: columnas fijas + la observación más larga
        ancho = sum(anchos) + max(len("Observación"), int(df["obs"].str.len().max() or 0))
        texto_datos.insert(tk.END, "  " + "".join(f"{c:<{w}}" for c, w in zip(cab, anchos)) + cab[-1] + "\n",
                           "cabecera")

        for _, g in df.groupby("id", sort=False):
            total = sum(g["horas"].dropna(), timedelta())
            # Borde superior con el nombre del empleado
            titulo = f" {g['id'].iloc[0]} · {g['nombre'].iloc[0]} {g['apellido'].iloc[0]} "
            texto_datos.insert(tk.END, "┌─" + titulo + "─" * max(0, ancho - len(titulo) - 1) + "┐\n", "borde")
            for k, (_, r) in enumerate(g.iterrows()):
                vals = [r["id"], r["nombre"], r["apellido"], r["fecha"].strftime("%d/%m/%Y"),
                        r["hora"].strftime("%H:%M"),
                        _fmt_horas(r["horas"]) if pd.notna(r["horas"]) else "",
                        _fmt_horas(total) if k == len(g) - 1 else ""]
                contenido = "".join(f"{str(v):<{w}}" for v, w in zip(vals, anchos)) + r["obs"]
                texto_datos.insert(tk.END, "│ ", "borde")
                texto_datos.insert(tk.END, f"{contenido:<{ancho - 1}}", r["nivel"])
                texto_datos.insert(tk.END, "│\n", "borde")
            texto_datos.insert(tk.END, "└" + "─" * ancho + "┘\n\n", "borde")
        texto_datos.config(state="disabled")
        actualizar_estado(df)

    def actualizar_estado(df):
        r = resumen(df)
        txt = (f"{len(r)} empleados  ·  {r['turnos'].sum()} turnos  ·  "
               f"{r['extendidos'].sum()} extendidos  ·  {r['sospechosos'].sum()} sospechosos  ·  "
               f"{r['cortos'].sum()} muy cortos  ·  {r['sin_pareja'].sum()} sin pareja  ·  "
               f"{r['repetidas'].sum()} repetidas ignoradas")
        lbl_estado.config(text=txt)

    def aplicar_filtros(event=None):
        df = estado["procesado"]
        if df is None:
            return
        mascara = pd.Series(True, index=df.index)
        for termino in entrada_busqueda.get().split():
            mascara &= (df["id"].astype(str).str.contains(termino, case=False, regex=False)
                        | df["nombre"].str.contains(termino, case=False, regex=False)
                        | df["apellido"].str.contains(termino, case=False, regex=False))
        filtrado = df[mascara]
        if var_solo_revisar.get():
            ids = filtrado.loc[filtrado["nivel"].isin(NIVELES_A_REVISAR), "id"].unique()
            filtrado = filtrado[filtrado["id"].isin(ids)]
        # Si un término coincide solo con parte de las filas, traemos todas las del empleado
        filtrado = df[df["id"].isin(filtrado["id"].unique())]
        estado["actual"] = filtrado
        if filtrado.empty:
            texto_datos.config(state="normal")
            texto_datos.delete(1.0, tk.END)
            texto_datos.insert(tk.END, "Sin resultados\n")
            texto_datos.config(state="disabled")
            lbl_estado.config(text="")
        else:
            mostrar_datos(filtrado)

    def restaurar_vista_original():
        if estado["procesado"] is not None:
            entrada_busqueda.delete(0, tk.END)
            var_solo_revisar.set(False)
            estado["actual"] = estado["procesado"]
            mostrar_datos(estado["procesado"])

    def guardar_excel():
        df = estado["actual"]
        if df is None or df.empty:
            messagebox.showwarning("!!!", "No data for process")
            return
        sugerido = os.path.splitext(os.path.basename(estado["ruta"] or "horas"))[0] + "_procesado.xlsx"
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=sugerido,
                                               filetypes=[("Excel File", "*.xlsx")])
        if archivo:
            try:
                exportar_excel(df, archivo)
                messagebox.showinfo("Saved", f"Succesfully saved at: {archivo}")
            except Exception as e:
                messagebox.showerror("Error", f"The file could not be saved {e}")

    def limpiar_interfaz():
        estado.update(procesado=None, actual=None, ruta=None)
        texto_datos.config(state="normal")
        texto_datos.delete(1.0, tk.END)
        texto_datos.insert(tk.END, "Drag the file or select it from the button\n")
        texto_datos.config(state="disabled")
        entrada_busqueda.delete(0, tk.END)
        var_solo_revisar.set(False)
        lbl_estado.config(text="")

    # ---------- Estilos ----------
    bg_dark, bg_darker, fg_light = "#2b2b2b", "#1e1e1e", "#e0e0e0"
    accent, btn_bg, btn_active = "#4a90e2", "#3c3f41", "#505050"

    ventana = TkinterDnD.Tk()
    ventana.title("Excel and CSV Files Reader")
    ventana.geometry("1150x720")
    ventana.configure(bg=bg_dark)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", background=bg_dark, foreground=fg_light, font=("Segoe UI", 10))
    style.configure("TFrame", background=bg_dark)
    style.configure("TButton", font=("Segoe UI", 10, "bold"), background=btn_bg,
                    foreground=fg_light, borderwidth=1, padding=6)
    style.map("TButton", background=[("active", btn_active), ("pressed", accent)],
              foreground=[("active", "#ffffff")])
    style.configure("TLabel", background=bg_dark, foreground=fg_light, font=("Segoe UI", 10))
    style.configure("Estado.TLabel", foreground="#9cdcfe", font=("Segoe UI", 9))
    style.configure("TEntry", fieldbackground=bg_darker, foreground=fg_light, padding=5)
    style.configure("TCheckbutton", background=bg_dark, foreground=fg_light)
    style.map("TCheckbutton", background=[("active", bg_dark)])
    style.configure("Vertical.TScrollbar", gripcount=0, background=btn_bg, troughcolor=bg_dark,
                    borderwidth=0, arrowsize=14)
    style.configure("Horizontal.TScrollbar", gripcount=0, background=btn_bg, troughcolor=bg_dark,
                    borderwidth=0, arrowsize=14)

    ventana.columnconfigure(0, weight=1)
    ventana.rowconfigure(2, weight=1)

    # ---------- Controles ----------
    frame_controles = ttk.Frame(ventana, padding="15 15 15 5")
    frame_controles.grid(row=0, column=0, sticky="ew")

    frame_botones = ttk.Frame(frame_controles)
    frame_botones.pack(side=tk.LEFT, fill=tk.X)
    ttk.Button(frame_botones, text="📂 Select File", command=seleccionar_archivo).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Button(frame_botones, text="💾 Save", command=guardar_excel).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Button(frame_botones, text="⬅ Back", command=restaurar_vista_original).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Button(frame_botones, text="↺ Clear", command=limpiar_interfaz).pack(side=tk.LEFT, padx=(0, 10))

    var_solo_revisar = tk.BooleanVar(value=False)
    ttk.Checkbutton(frame_botones, text="⚠ Solo para revisar", variable=var_solo_revisar,
                    command=aplicar_filtros).pack(side=tk.LEFT)

    frame_busqueda = ttk.Frame(frame_controles)
    frame_busqueda.pack(side=tk.RIGHT)
    ttk.Label(frame_busqueda, text="🔍 Search:").pack(side=tk.LEFT, padx=(0, 5))
    entrada_busqueda = ttk.Entry(frame_busqueda, width=30)
    entrada_busqueda.pack(side=tk.LEFT, padx=(0, 5))
    entrada_busqueda.bind("<Return>", aplicar_filtros)
    ttk.Button(frame_busqueda, text="Go", command=aplicar_filtros, width=4).pack(side=tk.LEFT)

    frame_info = ttk.Frame(ventana, padding="15 0 15 0")
    frame_info.grid(row=1, column=0, sticky="ew")
    lbl_estado = ttk.Label(frame_info, text="", style="Estado.TLabel")
    lbl_estado.pack(side=tk.LEFT)

    # Leyenda de colores de la vista previa
    frame_leyenda = ttk.Frame(frame_info)
    frame_leyenda.pack(side=tk.RIGHT)
    lim_ext = _fmt_horas(timedelta(hours=HORAS_TURNO_EXTENDIDO, minutes=MINUTOS_TOLERANCIA_EXTENDIDO))
    for texto, bg, fg in ((f"Extendido >{lim_ext}h", "#5a2a2a", "#ffb3b3"),
                          (f"Sospechoso >{HORAS_TURNO_SOSPECHOSO}h", "#a01818", "#ffffff"),
                          ("Muy corto", "#5a4a1a", "#ffe699"),
                          ("Sin pareja", "#7a4310", "#ffd8b0"),
                          ("Repetida", bg_darker, "#7a7a7a")):
        tk.Label(frame_leyenda, text=f" {texto} ", bg=bg, fg=fg,
                 font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=2)

    # ---------- Área de datos ----------
    frame_datos = ttk.Frame(ventana, padding="15 5 15 15")
    frame_datos.grid(row=2, column=0, sticky="nsew")
    scroll_y = ttk.Scrollbar(frame_datos, orient="vertical")
    scroll_x = ttk.Scrollbar(frame_datos, orient="horizontal")
    texto_datos = tk.Text(frame_datos, wrap="none", font=("Consolas", 10), bg=bg_darker, fg="#d4d4d4",
                          selectbackground="#264f78", selectforeground="#ffffff", insertbackground="white",
                          bd=0, highlightthickness=1, highlightbackground=btn_bg, highlightcolor=accent,
                          yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    scroll_y.config(command=texto_datos.yview)
    scroll_x.config(command=texto_datos.xview)
    scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    texto_datos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Colores por tipo de alerta en pantalla
    texto_datos.tag_configure("cabecera", foreground="#4a90e2", font=("Consolas", 10, "bold"))
    texto_datos.tag_configure("borde", foreground="#4a90e2")
    texto_datos.tag_configure(EXTENDIDO, background="#5a2a2a", foreground="#ffb3b3")
    texto_datos.tag_configure(SOSPECHOSO, background="#a01818", foreground="#ffffff",
                              font=("Consolas", 10, "bold"))
    texto_datos.tag_configure(CORTO, background="#5a4a1a", foreground="#ffe699")
    texto_datos.tag_configure(SIN_PAREJA, background="#7a4310", foreground="#ffd8b0",
                              font=("Consolas", 10, "bold"))
    texto_datos.tag_configure(REPETIDA, foreground="#7a7a7a", overstrike=True)

    texto_datos.insert(tk.END, "Drag the file or select it from the button\n")
    texto_datos.config(state="disabled")

    texto_datos.drop_target_register(DND_FILES)
    texto_datos.dnd_bind("<<Drop>>", lambda e: manejar_arrastre(e.data))

    ventana.mainloop()


if __name__ == "__main__":
    main()
