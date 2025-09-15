# -*- coding: utf-8 -*-
"""
Archivo: resources.py
Descripción: Carga de recursos desde archivos JSON.
"""

import json
import os

DATA_DIR = "data"
ALUMNOS_FILE = os.path.join(DATA_DIR, "alumnos.json")
LOCATIONS_FILE = os.path.join(DATA_DIR, "ubicaciones.json")
TIPOS_FILE = os.path.join(DATA_DIR, "tipos_incidencia.json")

def load_json_data(filepath, default_value=None):
    """Carga datos desde un archivo JSON de forma segura."""
    if default_value is None:
        default_value = []
    try:
        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # Si el archivo no existe, lo creamos con el valor por defecto
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(default_value, f, ensure_ascii=False, indent=2)
            return default_value
    except (json.JSONDecodeError, IOError) as e:
        print(f"Error al leer o crear {filepath}: {e}")
        return default_value

def load_all_resources():
    """
    Carga todos los recursos necesarios para la aplicación desde archivos JSON.
    Devuelve:
        - Una tupla con (lista de nombres de alumnos, diccionario de padres).
        - Una lista de ubicaciones.
        - Una lista de tipos de incidencia.
    """
    alumnos_data = load_json_data(ALUMNOS_FILE, default_value=[])
    
    # Procesar datos de alumnos para separar nombres y padres
    alumnos_nombres = [alumno.get("nombre", "") for alumno in alumnos_data]
    padres_dict = {alumno.get("nombre"): alumno.get("padre", "") for alumno in alumnos_data}
    
    locations = load_json_data(LOCATIONS_FILE, default_value=["El patio"])
    tipos = load_json_data(TIPOS_FILE, default_value=["Indisciplina"])
    
    return alumnos_nombres, padres_dict, locations, tipos
