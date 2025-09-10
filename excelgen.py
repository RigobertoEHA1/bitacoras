# -*- coding: utf-8 -*-
"""
Archivo: excelgen.py
Descripción: Generación y actualización del Excel con dashboard, registro de incidencias y contador de faltas.
"""

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import PieChart, Reference
from openpyxl.utils import get_column_letter

EXCEL_PATH = os.path.join("data", "bitacoras.xlsx")


def autosize_sheet(ws, min_width=8):
    """Ajusta ancho de columnas basado en el contenido (aproximado)."""
    dims = {}
    for row in ws.iter_rows(values_only=True):
        for idx, cell in enumerate(row, start=1):
            if cell is None:
                length = 0
            else:
                length = len(str(cell))
            dims[idx] = max(dims.get(idx, 0), length)
    for idx, width in dims.items():
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = max(min_width, width + 2)


def inicializar_excel():
    """
    Crea el archivo Excel con las hojas necesarias y los encabezados correctos si no existe.
    """
    if not os.path.exists("data"):
        os.makedirs("data")

    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()

        # Dashboard (página principal)
        ws_dash = wb.active
        ws_dash.title = "Dashboard"
        ws_dash["A1"] = "Dashboard de Incidencias"
        ws_dash["A1"].font = Font(size=14, bold=True)
        ws_dash["A1"].alignment = Alignment(horizontal="center")
        ws_dash.merge_cells("A1:D1")

        # Hoja de Incidencias (simplificada)
        ws_inc = wb.create_sheet("Incidencias")
        ws_inc.append([
            "Fecha", "Hora", "Lugar", "Gravedad", "Participantes", "Link al Documento"
        ])

        # Hoja para el Registro de Faltas con columnas de severidad
        ws_faltas = wb.create_sheet("Registro de Faltas")
        ws_faltas.append(["Alumno", "Total de Faltas", "Leve", "Moderada", "Grave"])

        autosize_sheet(ws_inc)
        autosize_sheet(ws_faltas)
        wb.save(EXCEL_PATH)


def _update_faltas_en_wb(wb, participantes, gravedad):
    """
    Actualiza la hoja 'Registro de Faltas' dentro del workbook abierto.
    Incrementa Total y la columna de la gravedad correspondiente para cada participante.
    """
    if not participantes:
        return

    ws = wb["Registro de Faltas"]

    # Mapa alumno -> fila (si existe)
    alumnos_existentes = {}
    for row in ws.iter_rows(min_row=2, max_col=5):
        nombre_cell = row[0]
        if nombre_cell.value:
            alumnos_existentes[str(nombre_cell.value)] = nombre_cell.row

    # Determinar la columna de la gravedad
    gravedad = gravedad or ""
    gravedad_col_map = {"Leve": 3, "Moderada": 4, "Grave": 5}  # columnas en la hoja: 1=Alumno,2=Total,3=Leve,...

    for alumno in participantes:
        if alumno in alumnos_existentes:
            row_idx = alumnos_existentes[alumno]
            # Total
            cell_total = ws.cell(row=row_idx, column=2)
            try:
                current_total = int(cell_total.value or 0)
            except:
                current_total = 0
            cell_total.value = current_total + 1

            # Gravedad específica
            if gravedad in gravedad_col_map:
                col_idx = gravedad_col_map[gravedad]
                cell_g = ws.cell(row=row_idx, column=col_idx)
                try:
                    current_g = int(cell_g.value or 0)
                except:
                    current_g = 0
                cell_g.value = current_g + 1
        else:
            # Nueva fila: inicializamos columnas: Alumno, Total, Leve, Moderada, Grave
            vale_leve = vale_mod = vale_grave = 0
            if gravedad == "Leve":
                vale_leve = 1
            elif gravedad == "Moderada":
                vale_mod = 1
            elif gravedad == "Grave":
                vale_grave = 1
            ws.append([alumno, 1, vale_leve, vale_mod, vale_grave])

    # Actualizar resumen/totales en la misma hoja (columna G/H)
    _refresh_faltas_summary(ws)


def _refresh_faltas_summary(ws):
    """
    Escribe un resumen de totales (clase) en la misma hoja 'Registro de Faltas',
    en las columnas G/H para no interferir con la tabla principal.
    """
    # Remover contenido previo del área de resumen (G1:H6 por ejemplo)
    for row in range(1, 8):
        ws.cell(row=row, column=7).value = None
        ws.cell(row=row, column=8).value = None

    # Calcular sumas (ignorando la fila de encabezado)
    total_sum = 0
    suma_leve = 0
    suma_mod = 0
    suma_grave = 0

    for row in ws.iter_rows(min_row=2, min_col=1, max_col=5, values_only=True):
        if not row or not row[0]:
            continue
        try:
            t = int(row[1] or 0)
        except:
            t = 0
        try:
            l = int(row[2] or 0)
        except:
            l = 0
        try:
            m = int(row[3] or 0)
        except:
            m = 0
        try:
            g = int(row[4] or 0)
        except:
            g = 0

        total_sum += t
        suma_leve += l
        suma_mod += m
        suma_grave += g

    # Escribir resumen compactado en columnas G/H
    ws.cell(row=1, column=7).value = "Resumen (clase)"
    ws.cell(row=2, column=7).value = "Total (suma Totales)"
    ws.cell(row=2, column=8).value = total_sum
    ws.cell(row=3, column=7).value = "Leve (suma por alumnos)"
    ws.cell(row=3, column=8).value = suma_leve
    ws.cell(row=4, column=7).value = "Moderada (suma por alumnos)"
    ws.cell(row=4, column=8).value = suma_mod
    ws.cell(row=5, column=7).value = "Grave (suma por alumnos)"
    ws.cell(row=5, column=8).value = suma_grave

    # Negrita para títulos
    ws.cell(row=1, column=7).font = Font(bold=True)
    ws.cell(row=2, column=7).font = Font(bold=True)
    ws.cell(row=3, column=7).font = Font(bold=True)
    ws.cell(row=4, column=7).font = Font(bold=True)
    ws.cell(row=5, column=7).font = Font(bold=True)


def registrar_incidencia(datos):
    """
    Registra una nueva incidencia en la hoja 'Incidencias'.
    'datos' es un dict con claves:
    [fecha, hora, lugar, gravedad, participantes, link]
    Se añade el link como hipervínculo clicable (ruta absoluta, file:///).
    Además actualiza la hoja 'Registro de Faltas' para los participantes.
    """
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Incidencias"]

    ws.append([
        datos["fecha"], datos["hora"], datos["lugar"], datos["gravedad"],
        ", ".join(datos["participantes"]), datos.get("link", "")
    ])

    last_row = ws.max_row
    link_cell = ws.cell(row=last_row, column=6)  # columna 6 = Link al Documento

    ruta = datos.get("link", "")
    if ruta:
        abs_path = os.path.abspath(ruta)
        href = "file:///" + abs_path.replace("\\", "/")
        link_cell.hyperlink = href
        link_cell.value = os.path.basename(abs_path)
        link_cell.font = Font(color="0000FF", underline="single")

    # Actualizar registro de faltas para los participantes
    try:
        participantes = datos.get("participantes", []) or []
        gravedad = datos.get("gravedad", "")
        _update_faltas_en_wb(wb, participantes, gravedad)
    except Exception:
        # No queremos que un error en faltas impida guardar la incidencia
        pass

    autosize_sheet(ws)
    # También autosize la hoja de faltas por si se actualizó
    try:
        autosize_sheet(wb["Registro de Faltas"])
    except Exception:
        pass

    wb.save(EXCEL_PATH)


def registrar_falta(alumnos_con_falta, gravedad=None):
    """
    Registra una o varias faltas en la hoja 'Registro de Faltas'.
    'alumnos_con_falta' es una lista de nombres de alumnos.
    Si 'gravedad' se proporciona (Leve/Moderada/Grave), también se incrementa la columna correspondiente.
    """
    wb = load_workbook(EXCEL_PATH)
    _update_faltas_en_wb(wb, alumnos_con_falta, gravedad)
    autosize_sheet(wb["Registro de Faltas"])
    wb.save(EXCEL_PATH)


def actualizar_dashboard():
    """
    Actualiza la hoja Dashboard con el resumen de gravedad (sin duplicados).
    """
    wb = load_workbook(EXCEL_PATH)
    ws_dash = wb["Dashboard"]
    ws_inc = wb["Incidencias"]

    # Limpiar contenido anterior del dashboard (desde fila 3 en adelante)
    for row in ws_dash.iter_rows(min_row=3):
        for cell in row:
            cell.value = None
    # Eliminar gráficos antiguos para no sobreponerlos
    ws_dash._charts = []

    # --- Lógica para contar incidencias por gravedad SIN duplicados ---
    incidencias_unicas = set()
    for row in ws_inc.iter_rows(min_row=2, max_col=5, values_only=True):
        if not row or not row[0]:
            continue

        fecha = row[0]
        gravedad = row[3]
        participantes_raw = row[4] or ""
        participantes = tuple(sorted(p.strip() for p in participantes_raw.split(',') if p.strip()))

        incidencia_unica = (fecha, gravedad, participantes)
        incidencias_unicas.add(incidencia_unica)

    total_gravedad = {"Leve": 0, "Moderada": 0, "Grave": 0}
    for incidencia in incidencias_unicas:
        gravedad = incidencia[1]
        if gravedad in total_gravedad:
            total_gravedad[gravedad] += 1

    # Escribir la tabla de resumen de gravedad
    ws_dash["A3"] = "Gravedad"
    ws_dash["B3"] = "Cantidad"
    ws_dash["A3"].font = Font(bold=True)
    ws_dash["B3"].font = Font(bold=True)

    fila = 4
    for g, val in total_gravedad.items():
        ws_dash[f"A{fila}"] = g
        ws_dash[f"B{fila}"] = val
        fila += 1

    # Crear el gráfico de pastel
    pie = PieChart()
    pie.title = "Distribución de Incidencias por Gravedad"

    data = Reference(ws_dash, min_col=2, min_row=3, max_row=fila - 1)
    labels = Reference(ws_dash, min_col=1, min_row=4, max_row=fila - 1)

    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    ws_dash.add_chart(pie, "D3")

    autosize_sheet(ws_dash)
    wb.save(EXCEL_PATH)
