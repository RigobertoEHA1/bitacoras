# -*- coding: utf-8 -*-
"""
Archivo: requirements.py
Descripción: Manejo automático de dependencias y estructura inicial.
"""

import os
import sys
import subprocess

# 🔽 Paquetes necesarios
REQUIRED_PACKAGES = ["python-docx", "openpyxl", "matplotlib"]

def install_missing_packages():
    """Instala automáticamente los paquetes que falten."""
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package.replace("-", "_"))
        except ImportError:
            print(f"⚠️ Instalando paquete faltante: {package} ...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def ensure_directories():
    """Crea las carpetas necesarias si no existen."""
    os.makedirs("recursos", exist_ok=True)
    os.makedirs("incidencias", exist_ok=True)

def ensure_resource_files():
    """Crea archivos de recursos si no existen."""
    base_files = {
        "students.txt": "Rigo$Flor\nDiego$Cristobal\n",
        "locations.txt": "Aula\nPatio\nCancha\nDirección\n",
        "tipoIncidencia.txt": "Accidente\nPelea\nIndisciplina\nBullying\nOtro\n"
    }

    for fname, content in base_files.items():
        fpath = os.path.join("recursos", fname)
        if not os.path.exists(fpath):
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(content)

def setup_environment():
    """Ejecuta todas las configuraciones iniciales."""
    install_missing_packages()
    ensure_directories()
    ensure_resource_files()
    print("✅ Entorno listo.")

# Si se ejecuta directamente
if __name__ == "__main__":
    setup_environment()
