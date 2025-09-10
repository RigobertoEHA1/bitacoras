# -*- coding: utf-8 -*-
"""
Archivo: excelgen.py
Descripción: Generación y actualización del Excel con dashboard y registro de incidencias.
"""

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import BarChart, PieChart, Reference


EXCEL_PATH = os.path.join("data", "bitacoras.xlsx")


def inicializar_excel():
    """
    Crea el archivo Excel con las hojas necesarias si no existe.
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
        ws_dash.merge_cells("A1:E1")

        # Hoja incidencias
        ws = wb.create_sheet("Incidencias")
        ws.append([
            "Fecha", "Hora", "Lugar", "Actividad", "Tipo",
            "Gravedad", "Participantes", "Narración", "Medidas", "Seguimiento"
        ])

        # Hoja recursos
        ws_rec = wb.create_sheet("Recursos")
        ws_rec["A1"] = "Lugares"
        ws_rec["B1"] = "Tipos de Incidencia"
        ws_rec["C1"] = "Gravedad"
        ws_rec["A2"] = "Aula"
        ws_rec["A3"] = "Patio"
        ws_rec["A4"] = "Dirección"
        ws_rec["B2"] = "Agresión verbal"
        ws_rec["B3"] = "Agresión física"
        ws_rec["B4"] = "Falta de respeto"
        ws_rec["C2"] = "Leve"
        ws_rec["C3"] = "Moderada"
        ws_rec["C4"] = "Grave"

        wb.save(EXCEL_PATH)


def registrar_incidencia(datos):
    """
    Registra una nueva incidencia en la hoja 'Incidencias'.
    datos = dict con claves:
    [fecha, hora, lugar, actividad, tipo, gravedad, participantes, narracion, medidas, seguimiento]
    """
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Incidencias"]

    ws.append([
        datos["fecha"], datos["hora"], datos["lugar"], datos["actividad"],
        datos["tipo"], datos["gravedad"], ", ".join(datos["participantes"]),
        datos["narracion"], datos["medidas"], datos["seguimiento"]
    ])

    wb.save(EXCEL_PATH)


def actualizar_dashboard():
    """
    Actualiza la hoja Dashboard con resumen y gráficos.
    """
    wb = load_workbook(EXCEL_PATH)

    if "Dashboard" not in wb.sheetnames:
        wb.create_sheet("Dashboard")
    ws_dash = wb["Dashboard"]

    # Limpiar contenido excepto título
    for row in ws_dash["A3:Z100"]:
        for cell in row:
            cell.value = None

    ws_inc = wb["Incidencias"]

    # Construir tabla resumen
    conteo = {}
    for row in ws_inc.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        alumnos = row[6].split(", ")
        gravedad = row[5]
        for alumno in alumnos:
            if alumno not in conteo:
                conteo[alumno] = {"Leve": 0, "Moderada": 0, "Grave": 0}
            conteo[alumno][gravedad] += 1

    # Escribir tabla
    ws_dash["A3"] = "Alumno"
    ws_dash["B3"] = "Leves"
    ws_dash["C3"] = "Moderadas"
    ws_dash["D3"] = "Graves"
    ws_dash["E3"] = "Total"
    fila = 4
    for alumno, datos in conteo.items():
        ws_dash[f"A{fila}"] = alumno
        ws_dash[f"B{fila}"] = datos["Leve"]
        ws_dash[f"C{fila}"] = datos["Moderada"]
        ws_dash[f"D{fila}"] = datos["Grave"]
        ws_dash[f"E{fila}"] = sum(datos.values())
        fila += 1

    # Gráfico de barras por alumno
    chart = BarChart()
    chart.title = "Incidencias por Alumno"
    data = Reference(ws_dash, min_col=2, max_col=4, min_row=3, max_row=fila - 1)
    cats = Reference(ws_dash, min_col=1, min_row=4, max_row=fila - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.y_axis.title = "Número de incidencias"
    chart.x_axis.title = "Alumno"
    ws_dash.add_chart(chart, "G3")

    # Gráfico de pastel por gravedad total
    total_gravedad = {"Leve": 0, "Moderada": 0, "Grave": 0}
    for d in conteo.values():
        for k in total_gravedad.keys():
            total_gravedad[k] += d[k]

    ws_dash["A20"] = "Gravedad"
    ws_dash["B20"] = "Cantidad"
    fila = 21
    for g, val in total_gravedad.items():
        ws_dash[f"A{fila}"] = g
        ws_dash[f"B{fila}"] = val
        fila += 1

    pie = PieChart()
    pie.title = "Distribución por Gravedad"
    data = Reference(ws_dash, min_col=2, min_row=20, max_row=fila - 1)
    labels = Reference(ws_dash, min_col=1, min_row=21, max_row=fila - 1)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    ws_dash.add_chart(pie, "G20")

    wb.save(EXCEL_PATH)
