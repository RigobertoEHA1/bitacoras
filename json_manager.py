# -*- coding: utf-8 -*-
"""
Archivo: json_manager.py
Descripción: Funciones para gestionar (leer/escribir) los archivos de datos JSON.
"""

import json
import os

DATA_DIR = "data"
ALUMNOS_FILE = os.path.join(DATA_DIR, "alumnos.json")
LOCATIONS_FILE = os.path.join(DATA_DIR, "ubicaciones.json")
TIPOS_FILE = os.path.join(DATA_DIR, "tipos_incidencia.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")

# Asegurarse de que el directorio de datos exista
os.makedirs(DATA_DIR, exist_ok=True)

def leer_json(filepath, default_value=None):
    """Lee un archivo JSON y devuelve su contenido."""
    if default_value is None:
        default_value = []
    if not os.path.exists(filepath):
        escribir_json(filepath, default_value)
        return default_value
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return default_value

def escribir_json(filepath, data):
    """Escribe datos en un archivo JSON."""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except IOError:
        return False

# --- Funciones específicas para cada tipo de dato ---

def obtener_alumnos():
    return leer_json(ALUMNOS_FILE, [])

def guardar_alumnos(data):
    return escribir_json(ALUMNOS_FILE, data)

def obtener_ubicaciones():
    return leer_json(LOCATIONS_FILE, [])

def guardar_ubicaciones(data):
    return escribir_json(LOCATIONS_FILE, data)

def obtener_tipos_incidencia():
    return leer_json(TIPOS_FILE, [])

def guardar_tipos_incidencia(data):
    return escribir_json(TIPOS_FILE, data)

def obtener_config():
    default_config = {
        "teacher_name": "Maestro Titular", "grade": "1", "group": "A",
        "director_name": "Director/a", "school_name": "Nombre Escuela", "location": "Ubicación Escuela"
    }
    return leer_json(CONFIG_FILE, default_config)

def guardar_config(data):
    return escribir_json(CONFIG_FILE, data)
