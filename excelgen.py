# -*- coding: utf-8 -*-
"""
Archivo: excelgen.py
Descripción: Generación y actualización del Excel con dashboard, registro de incidencias y contador de faltas.
"""

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import PieChart, Reference

EXCEL_PATH = os.path.join("data", "bitacoras.xlsx")


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
        
        # Hoja para el Registro de Faltas
        ws_faltas = wb.create_sheet("Registro de Faltas")
        ws_faltas.append(["Alumno", "Total de Faltas"])

        wb.save(EXCEL_PATH)


def registrar_incidencia(datos):
    """
    Registra una nueva incidencia en la hoja 'Incidencias'.
    'datos' es un dict con claves:
    [fecha, hora, lugar, gravedad, participantes, link]
    """
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Incidencias"]

    ws.append([
        datos["fecha"], datos["hora"], datos["lugar"], datos["gravedad"],
        ", ".join(datos["participantes"]), datos["link"]
    ])

    wb.save(EXCEL_PATH)


def registrar_falta(alumnos_con_falta):
    """
    Registra una o varias faltas en la hoja 'Registro de Faltas'.
    'alumnos_con_falta' es una lista de nombres de alumnos.
    """
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Registro de Faltas"]

    # Crear un diccionario para buscar filas existentes rápidamente
    alumnos_existentes = {}
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_col=2), start=2):
        alumno_cell = row[0]
        conteo_cell = row[1]
        alumnos_existentes[alumno_cell.value] = conteo_cell

    for alumno in alumnos_con_falta:
        if alumno in alumnos_existentes:
            # Si el alumno existe, se incrementa su contador
            alumnos_existentes[alumno].value += 1
        else:
            # Si no existe, se añade una nueva fila
            ws.append([alumno, 1])
    
    wb.save(EXCEL_PATH)


def actualizar_dashboard():
    """
    Actualiza la hoja Dashboard con el resumen de gravedad (sin duplicados).
    """
    wb = load_workbook(EXCEL_PATH)
    ws_dash = wb["Dashboard"]
    ws_inc = wb["Incidencias"]

    # Limpiar contenido anterior del dashboard
    for row in ws_dash.iter_rows(min_row=3):
        for cell in row:
            cell.value = None
    # Eliminar gráficos antiguos para no sobreponerlos
    ws_dash._charts = []

    # --- Lógica para contar incidencias por gravedad SIN duplicados ---
    incidencias_unicas = set()
    # Se considera duplicado si coincide: Fecha, Gravedad y Participantes
    for row in ws_inc.iter_rows(min_row=2, max_col=5, values_only=True):
        if not row[0]: # Ignorar filas vacías
            continue
        
        fecha = row[0]
        gravedad = row[3]
        # Ordenar participantes para que "A, B" sea igual que "B, A"
        participantes = tuple(sorted(p.strip() for p in row[4].split(',')))
        
        incidencia_unica = (fecha, gravedad, participantes)
        incidencias_unicas.add(incidencia_unica)

    # Contar totales desde el conjunto de incidencias únicas
    total_gravedad = {"Leve": 0, "Moderada": 0, "Grave": 0}
    for incidencia in incidencias_unicas:
        gravedad = incidencia[1] # La gravedad es el segundo elemento de la tupla
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
    ws_dash.add_chart(pie, "D3") # Posicionar el gráfico

    wb.save(EXCEL_PATH)