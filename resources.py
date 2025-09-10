# -*- coding: utf-8 -*-
"""
Archivo: resources.py
Descripción: Carga de recursos (alumnos, tutores, ubicaciones, tipos de incidencia).
"""

import os

RECURSOS_DIR = "recursos"

def load_students():
    """
    Carga alumnos y sus padres desde students.txt
    Formato esperado: Alumno$Padre
    """
    students_file = os.path.join(RECURSOS_DIR, "students.txt")
    alumnos = []
    padres  = {}
    if os.path.exists(students_file):
        with open(students_file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                if "$" in line:
                    alumno, padre = line.split("$", 1)
                    alumno, padre = alumno.strip(), padre.strip()
                    alumnos.append(alumno)
                    padres[alumno] = padre
                else:
                    # Si no hay "$", lo consideramos solo alumno
                    alumnos.append(line.strip())
                    padres[line.strip()] = "N/A"
    return alumnos, padres

def load_locations():
    """Carga ubicaciones desde locations.txt"""
    loc_file = os.path.join(RECURSOS_DIR, "locations.txt")
    if os.path.exists(loc_file):
        with open(loc_file, encoding="utf-8") as f:
            return [l.strip() for l in f if l.strip()]
    return []

def load_tipo_incidencia():
    """Carga tipos de incidencia desde tipoIncidencia.txt"""
    tipo_file = os.path.join(RECURSOS_DIR, "tipoIncidencia.txt")
    if os.path.exists(tipo_file):
        with open(tipo_file, encoding="utf-8") as f:
            return [l.strip() for l in f if l.strip()]
    return ["Accidente", "Pelea", "Indisciplina", "Bullying", "Otro"]

def load_all_resources():
    """
    Devuelve alumnos, padres, ubicaciones y tipos de incidencia.
    """
    alumnos, padres = load_students()
    locations = load_locations()
    tipos     = load_tipo_incidencia()
    return alumnos, padres, locations, tipos
