"""
Puente entre la página web (JavaScript) y la lógica de lectorExcel_IA.py.
Corre dentro del navegador con Pyodide: los archivos nunca salen de la compu del usuario.
"""
import json
from datetime import timedelta

import lectorExcel_IA as L

_df = None


def _fmt(td):
    return L._fmt_horas(td) if td is not None and td == td else ""


def _empleados(df):
    res = L.resumen(df).set_index("id")
    salida = []
    for id_, g in df.groupby("id", sort=False):
        r = res.loc[id_]
        salida.append({
            "id": str(id_),
            "nombre": g["nombre"].iloc[0],
            "apellido": g["apellido"].iloc[0],
            "total": _fmt(r["total"]),
            "turnos": int(r["turnos"]),
            "extendidos": int(r["extendidos"]),
            "sospechosos": int(r["sospechosos"]),
            "cortos": int(r["cortos"]),
            "sin_pareja": int(r["sin_pareja"]),
            "repetidas": int(r["repetidas"]),
            "revisar": bool(g["nivel"].isin(L.NIVELES_A_REVISAR).any()),
            "filas": [{
                "fecha": f["fecha"].strftime("%d/%m/%Y"),
                "hora": f["hora"].strftime("%H:%M"),
                "horas": _fmt(f["horas"]),
                "nivel": f["nivel"],
                "obs": f["obs"],
            } for _, f in g.iterrows()],
        })
    return salida


def cargar(ruta):
    """Procesa el archivo y devuelve todo en JSON para que la página lo dibuje."""
    global _df
    df, invalidas = L.procesar(L.leer_lineas(ruta))
    _df = df
    return json.dumps({
        "empleados": _empleados(df) if not df.empty else [],
        "invalidas": [{"linea": n, "texto": t} for n, t in invalidas],
        "config": {
            "extendido": L._fmt_horas(timedelta(hours=L.HORAS_TURNO_EXTENDIDO,
                                                minutes=L.MINUTOS_TOLERANCIA_EXTENDIDO)),
            "sospechoso": L.HORAS_TURNO_SOSPECHOSO,
            "corto": L.MINUTOS_TURNO_CORTO,
            "repetida": L.MINUTOS_MARCA_REPETIDA,
        },
    }, ensure_ascii=False)


def exportar(ids_json, ruta):
    """Genera el Excel solo con los empleados visibles (respeta búsqueda y filtros)."""
    ids = set(json.loads(ids_json))
    L.exportar_excel(_df[_df["id"].astype(str).isin(ids)], ruta)
