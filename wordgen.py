# -*- coding: utf-8 -*-
"""
Archivo: wordgen.py
Descripción: Generación del documento Word para incidencias.
"""

import os
import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import SCHOOL_NAME, LOCATION, DIRECTOR_NAME, TEACHER_NAME, GRADE, GROUP


def generar_word(fecha, hora, lugar, actividad, participantes, tipo_inc,
                 gravedad, narracion, medidas, seguimiento, padres_dict,
                 alumnos_seleccionados, output_path):
    """
    Genera el documento Word de la bitácora.
    """

    # Crear documento
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # ----- Encabezado con logos -----
    header = sec.header
    table = header.add_table(rows=1, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Logo 1 (izquierda, cuadrado)
    cell_left = table.cell(0, 0)
    logo1 = os.path.join("recursos", "logo1.png")
    if os.path.exists(logo1):
        cell_left.paragraphs[0].add_run().add_picture(logo1, width=Inches(1))

    # Logo 2 (derecha, rectangular más grande)
    cell_right = table.cell(0, 1)
    logo2 = os.path.join("recursos", "logo2.png")
    if os.path.exists(logo2):
        cell_right.paragraphs[0].add_run().add_picture(logo2, width=Inches(2.5))

    # ----- Título -----
    t = doc.add_paragraph()
    r = t.add_run(f"BITÁCORA DE INCIDENCIA - {SCHOOL_NAME}\n")
    r.bold = True
    r.font.size = Pt(14)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph(LOCATION)
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].italic = True

    doc.add_paragraph("")

    # ----- Narración en prosa -----
    narr = (
        f"Siendo las {hora} horas del día {fecha}, durante {actividad} "
        f"en {lugar}, se presentó una incidencia del tipo {tipo_inc}. "
        f"En ella participaron {', '.join(participantes)} del grupo {GRADE} {GROUP}. "
        f"La gravedad fue evaluada como {gravedad.lower()}. "
        f"Los hechos se describen de la siguiente manera: {narracion}. "
    )

    if medidas:
        narr += f"Las medidas tomadas fueron: {medidas}. "

    if seguimiento:
        narr += f"Para su seguimiento se determinó: {seguimiento}. "

    doc.add_paragraph(narr)

    doc.add_paragraph("")

    # ----- Firmas dinámicas -----
    firmas = []

    # Siempre Maestro y alumnos implicados + testigo
    firmas.append(f"Maestro de grupo: {TEACHER_NAME}")
    for alumno in alumnos_seleccionados:
        firmas.append(f"Alumno: {alumno}")

    firmas.append("Testigo: __________________________")

    # Moderada y Grave → también Director
    if gravedad in ["Moderada", "Grave"]:
        firmas.insert(0, f"Director: {DIRECTOR_NAME}")

    # Grave → también Padres de los alumnos
    if gravedad == "Grave":
        for alumno in alumnos_seleccionados:
            padre = padres_dict.get(alumno, "Padre de familia")
            firmas.append(f"Padre de familia ({alumno}): {padre}")

    # Crear tabla de firmas
    table = doc.add_table(rows=len(firmas), cols=1)
    for i, text in enumerate(firmas):
        cell = table.cell(i, 0)
        p = cell.paragraphs[0]
        run = p.add_run(text + "\n\nFirma: __________________________")
        run.font.size = Pt(10)

    # Guardar documento
    doc.save(output_path)
    return output_path
