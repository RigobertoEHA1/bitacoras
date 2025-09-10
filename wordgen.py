# -*- coding: utf-8 -*-
"""
Archivo: wordgen.py
Descripción: Generación del documento Word para incidencias.
"""

import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL

from config import SCHOOL_NAME, LOCATION, DIRECTOR_NAME, TEACHER_NAME, GRADE, GROUP

def set_cell_border(cell, **kwargs):
    """
    Establece los bordes de una celda de tabla.
    Cada borde debe ser un dict: {'sz': 12, 'val': 'single', 'color': '#000000'}.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    borders = {
        'top': {'sz': 6, 'val': 'single', 'color': '000000'},
        'bottom': {'sz': 6, 'val': 'single', 'color': '000000'},
        'left': {'sz': 6, 'val': 'single', 'color': '000000'},
        'right': {'sz': 6, 'val': 'single', 'color': '000000'},
    }

    # Sobrescribimos solo si kwargs tiene dicts válidos
    for key, val in kwargs.items():
        if val is None:
            borders[key] = None
        elif isinstance(val, dict):
            borders[key] = val
        else:
            # Ignorar valores inválidos
            continue

    for border_name, border_props in borders.items():
        if border_props is not None:
            border_elm = OxmlElement(f'w:{border_name}')
            for prop, val in border_props.items():
                border_elm.set(qn(f'w:{prop}'), str(val))
            tcPr.append(border_elm)



def generar_word(fecha, hora, lugar, actividad, participantes, tipo_inc,
                 gravedad, narracion, medidas, seguimiento, padres_dict,
                 alumnos_seleccionados, output_path):
    """
    Genera el documento Word de la bitácora de manera segura.
    """

    # 🔹 Aseguramos que padres_dict sea un diccionario
    if not isinstance(padres_dict, dict):
        padres_dict = {}

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # ----- Encabezado con logos -----
    try:
        header = sec.header
        header_table = header.add_table(rows=1, cols=3, width=Inches(6.5))
        header_table.autofit = False

        # Logo izquierdo
        cell_logo1 = header_table.cell(0, 0)
        cell_logo1.width = Inches(1.2)
        cell_logo1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        logo1_path = os.path.join("recursos", "logo1.png")
        if os.path.exists(logo1_path):
            run = cell_logo1.paragraphs[0].add_run()
            run.add_picture(logo1_path, width=Inches(0.8))
        else:
            cell_logo1.paragraphs[0].add_run("Logo Izquierdo")

        # Columna central
        cell_center = header_table.cell(0, 1)
        cell_center.width = Inches(2.8)

        # Logo derecho
        cell_logo2 = header_table.cell(0, 2)
        cell_logo2.width = Inches(2.5)
        cell_logo2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        logo2_path = os.path.join("recursos", "logo2.png")
        if os.path.exists(logo2_path):
            run = cell_logo2.paragraphs[0].add_run()
            run.add_picture(logo2_path, width=Inches(2.5))
        else:
            cell_logo2.paragraphs[0].add_run("Logo Derecho")

    except Exception as e:
        print(f"Advertencia: No se pudieron agregar los logos al encabezado. Causa: {e}")

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

    # ----- Narración -----
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
    p_narr = doc.add_paragraph(narr)
    p_narr.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph("")
    doc.add_paragraph("")

    # ----- Firmas -----
    firmas_data = []
    if gravedad in ["Moderada", "Grave"]:
        firmas_data.append(("Director", DIRECTOR_NAME))
    firmas_data.append(("Maestro de Grupo", TEACHER_NAME))
    for alumno in alumnos_seleccionados:
        firmas_data.append(("Alumno", alumno))
    if gravedad == "Grave":
        for alumno in alumnos_seleccionados:
            padre = str(padres_dict.get(alumno, "Padre de familia"))
            firmas_data.append((f"Padre/Madre de familia ({alumno})", padre))
    firmas_data.append(("Testigo", "____________________"))

    num_firmas = len(firmas_data)
    num_cols = 2
    num_rows = (num_firmas + num_cols - 1) // num_cols

    signatures_table = doc.add_table(rows=num_rows, cols=num_cols)
    signatures_table.autofit = False
    signatures_table.width = Inches(6.5)
    signatures_table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    col_width = Inches(signatures_table.width.emu / num_cols)
    for col in signatures_table.columns:
        col.width = col_width

    firma_idx = 0
    for r_idx in range(num_rows):
        for c_idx in range(num_cols):
            cell = signatures_table.cell(r_idx, c_idx)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if firma_idx < num_firmas:
                title, name = firmas_data[firma_idx]
                cell.text = ''
                set_cell_border(cell, top=None, bottom=None, left=None, right=None)
                set_cell_border(cell, sz=12, val='single', color='#000000')

                p_title = cell.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_title = p_title.add_run(title)
                run_title.bold = True
                run_title.font.size = Pt(10)

                p_line = cell.add_paragraph()
                p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_line.paragraph_format.space_after = Pt(0)
                p_line.add_run("\n\n____________________")

                p_name = cell.add_paragraph()
                p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_name.paragraph_format.space_before = Pt(0)
                run_name = p_name.add_run(name)
                run_name.font.size = Pt(9)

                firma_idx += 1
            else:
                cell.text = ''
                set_cell_border(cell, top=None, bottom=None, left=None, right=None)

    # Guardar documento
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)
    return output_path
