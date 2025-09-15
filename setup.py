# -*- coding: utf-8 -*-
"""
Archivo: setup.py
Descripción: Módulo de configuración inicial para la aplicación.
             Verifica e instala dependencias, y crea la estructura de archivos necesaria.
"""

import subprocess
import sys
import os
import json

# --- Dependencias Requeridas ---
REQUIRED_PACKAGES = [
    "openpyxl",
    "python-docx"
]

# --- Estructura de Archivos y Carpetas ---
REQUIRED_DIRS = ["data", "incidencias", "recursos"]
DEFAULT_DATA_FILES = {
    os.path.join("data", "alumnos.json"): [
        {"nombre": "Juan Ejemplo", "padre": "Pedro Ejemplo", "grado": "1", "grupo": "A"},
        {"nombre": "Maria Muestra", "padre": "Ana Muestra", "grado": "1", "grupo": "A"}
    ],
    os.path.join("data", "ubicaciones.json"): [
        "Patio de juegos", "Salón de clases", "Comedor", "Biblioteca"
    ],
    os.path.join("data", "tipos_incidencia.json"): [
        "Indisciplina", "Agresión verbal", "Agresión física"
    ],
    os.path.join("data", "config.json"): {
        "teacher_name": "Nombre del Maestro Titular",
        "grade": "1",
        "group": "A",
        "director_name": "Nombre del Director/a",
        "school_name": "Nombre de la Escuela",
        "location": "Ubicación de la Escuela"
    }
}

def check_and_install_packages():
    """
    Verifica si los paquetes requeridos están instalados.
    Si no lo están, intenta instalarlos usando pip.
    """
    print("Verificando dependencias...")
    try:
        import pkg_resources
        installed_packages = {pkg.key for pkg in pkg_resources.working_set}
        missing_packages = [pkg for pkg in REQUIRED_PACKAGES if pkg.lower() not in installed_packages]

        if missing_packages:
            print(f"Faltan los siguientes paquetes: {', '.join(missing_packages)}")
            print("Intentando instalarlos ahora...")
            
            for package in missing_packages:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            
            print("Dependencias instaladas correctamente.")
        else:
            print("Todas las dependencias ya están instaladas.")
            
    except ImportError:
        print("Advertencia: 'pkg_resources' no encontrado. No se pueden verificar las dependencias automáticamente.")
        print("Asegúrate de tener instalados: pip install openpyxl python-docx")
    except Exception as e:
        print(f"Ocurrió un error durante la instalación de dependencias: {e}")
        print("Por favor, instala manualmente los paquetes requeridos: pip install openpyxl python-docx")

def create_project_structure():
    """
    Verifica y crea la estructura de carpetas y archivos de datos por defecto
    si no existen.
    """
    print("Verificando estructura del proyecto...")
    # Crear carpetas
    for dir_name in REQUIRED_DIRS:
        if not os.path.exists(dir_name):
            print(f"Creando directorio: {dir_name}")
            os.makedirs(dir_name)

    # Crear archivos de datos JSON por defecto
    for filepath, default_content in DEFAULT_DATA_FILES.items():
        if not os.path.exists(filepath):
            print(f"Creando archivo de datos por defecto: {filepath}")
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(default_content, f, ensure_ascii=False, indent=4)
            except IOError as e:
                print(f"Error al crear el archivo {filepath}: {e}")
    print("Estructura del proyecto verificada.")

def run_setup():
    """
    Ejecuta todas las tareas de configuración inicial.
    """
    print("--- Iniciando configuración de la aplicación ---")
    check_and_install_packages()
    create_project_structure()
    print("--- Configuración completada ---")

if __name__ == '__main__':
    # Permite ejecutar este script directamente para la configuración
    run_setup()
