# Procesador de Horas

Procesa el archivo que exporta el reloj de fichadas, calcula las horas de cada turno y el total semanal por empleado, y detecta errores de marcación. Devuelve un Excel listo para liquidar.

**Versión web:** https://alejandroflorido.github.io/procesador-horas/

## Menú

- **Bares** → liquidar horas.
- **BRIE** → flujo en 3 pasos: **Liquidar → Adelantos → Mercadería**. Desde Liquidar, el botón "Siguiente paso" lleva el archivo ya procesado a Adelantos, y de ahí a Mercadería (en construcción).
  - **Liquidar:** además del archivo, se completa el **sueldo por hora** (igual para todos). Calcula el sueldo de cada empleado con sus horas totales y lo agrega al Excel (hoja Resumen).
  - **Adelantos:** se elige el empleado en un desplegable, se escribe el monto y se resta. El empleado queda en verde para no restarlo dos veces ("Deshacer" por si hubo un error). Con "Continuar" se pasa a Mercadería; no hace falta que todos tengan adelanto. **El adelanto no puede superar el sueldo** del empleado.
  - **Mercadería** (en definición): la suma de adelanto + mercadería **sí puede** superar el sueldo; en ese caso el "A cobrar" queda negativo.

## Qué hace la liquidación

- Separa ID, nombre, apellido, fecha y hora.
- Empareja entrada y salida usando **fecha + hora**, así que los turnos que pasan la medianoche (ej. 19:00 → 02:00) se calculan bien.
- Escribe las horas de cada turno y el **total semanal** (formato `[h]:mm:ss`).
- Antes de procesar pide el **nombre del bar**; el Excel se descarga con ese nombre.
- Encierra a cada empleado en un cuadro y agrega una hoja **Resumen**.
- Marca con colores:
  - **Turno extendido**: más de 8 h + 1:30 de tolerancia (desde 9:30 h).
  - **Turno sospechoso**: más de 12 h (probable marca faltante).
  - **Turno muy corto**: menos de 30 min.
  - **Marca sin pareja**: falta la entrada o la salida (no suma horas).
  - **Marca repetida**: dentro de 10 min de la anterior; se ignora.

Los límites se ajustan al principio de `lectorExcel_IA.py`.

## Privacidad

La versión web corre Python dentro del navegador ([Pyodide](https://pyodide.org)). El archivo **no se sube a ningún servidor** y no se guarda nada: se procesa en la compu del usuario y se descarga el resultado.

## Archivos

| Archivo | Para qué |
|---|---|
| `lectorExcel_IA.py` | Lógica de procesamiento + versión de escritorio (Tkinter) |
| `web.py` | Puente entre la página y la lógica |
| `index.html` | Interfaz web |

## Versión de escritorio

```bash
pip install -r requirements.txt
python lectorExcel_IA.py
```

## Cómo activar Adelantos o Mercadería

Las dos secciones de BRIE ya tienen su lugar preparado:

1. **Lógica:** completar `adelantos.py` o `mercaderia.py` (funciones `procesar` y `exportar_excel`). La liquidación del paso 1 está disponible con `liquidacion()`. En la página, cada paso lee y guarda sus resultados en el objeto `flujo` (`flujo.liquidacion`, `flujo.adelantos`, `flujo.mercaderia`).
2. **Pantalla:** en `index.html`, armar el contenido dentro de `<div id="contenido-adelantos">` o `<div id="contenido-mercaderia">`.
3. **Activar:** en `index.html`, en `SECCIONES_BRIE`, cambiar `ACTIVA: false` por `ACTIVA: true`. Con eso se quita el cartel "Próximamente" del menú y se muestra el contenido.
