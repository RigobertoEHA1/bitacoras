# -*- coding: utf-8 -*-
"""
Archivo: programa.py
Interfaz principal para registrar, eliminar y gestionar incidencias escolares.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os

from wordgen import generar_word
from excelgen import registrar_incidencia, actualizar_dashboard, inicializar_excel, eliminar_incidencia
from resources import load_all_resources
from config import INCIDENCIAS_DIR, SCHOOL_NAME, LOCATION, DIRECTOR_NAME, TEACHER_NAME, GRADE, GROUP

# ===================== CARGA DE RECURSOS =====================
alumnos, padres, locations, tipos = load_all_resources()

# Valores por defecto si no se cargó nada desde recursos/
if not alumnos:
    alumnos = ["Rigo", "Diego", "Juan", "Pedro"]
if not locations:
    locations = ["El patio", "El salón", "Los baños"]
if not tipos:
    tipos = ["Indisciplina", "Agresión física", "Agresión verbal"]

# ===================== INICIALIZACIÓN =====================
# Inicializar Excel (crea data/bitacoras.xlsx si no existe)
inicializar_excel()

# Crear el directorio para los documentos si no existe
_output_dir = INCIDENCIAS_DIR if INCIDENCIAS_DIR else "documentos"
INCIDENCIAS_DIR = _output_dir
os.makedirs(INCIDENCIAS_DIR, exist_ok=True)

gravedades = ["Leve", "Moderada", "Grave"]

# ===================== FUNCIONES =====================
def generar_doc():
    fecha = entry_fecha.get()
    hora = entry_hora.get()
    lugar = combo_lugar.get()
    actividad = entry_actividad.get()
    tipo = combo_tipo.get()
    gravedad = combo_gravedad.get()
    seleccion = listbox_alumnos.curselection()
    participantes = [alumnos[i] for i in seleccion]
    narracion = text_narracion.get("1.0", tk.END).strip()
    medidas = text_medidas.get("1.0", tk.END).strip()
    seguimiento = text_seguimiento.get("1.0", tk.END).strip()

    # Validación
    if not participantes or not tipo or not lugar or not gravedad:
        messagebox.showwarning("Falta información", "Debe seleccionar al menos un alumno y completar todos los menús desplegables.")
        return

    # 1. Preparamos el diccionario de datos para Excel
    datos_excel = {
        "fecha": fecha, "hora": hora, "lugar": lugar, "gravedad": gravedad,
        "participantes": participantes, "link": ""  # El link se añadirá después
    }

    # 2. Definimos una ruta única para el documento de Word
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombres_alumnos = "_".join(participantes).replace(" ", "")
    output_filename = f"Incidencia_{nombres_alumnos}_{timestamp}.docx"
    output_path = os.path.join(INCIDENCIAS_DIR, output_filename)

    try:
        # 3. Generamos el documento de Word
        ruta_generada = generar_word(
            fecha=fecha,
            hora=hora,
            lugar=lugar,
            actividad=actividad,
            participantes=participantes,
            tipo_inc=tipo,
            gravedad=gravedad,
            narracion=narracion,
            medidas=medidas,
            seguimiento=seguimiento,
            padres_dict=padres,
            alumnos_seleccionados=participantes,
            output_path=output_path
        )

        # 4. AÑADIMOS EL LINK AL DICCIONARIO
        datos_excel['link'] = ruta_generada

        # 5. Guardamos el registro en Excel ahora que el diccionario está completo
        registrar_incidencia(datos_excel)

        # 6. Actualizamos el dashboard automáticamente
        actualizar_dashboard()

        messagebox.showinfo("Éxito", f"Incidencia registrada.\nWord guardado en: {ruta_generada}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el documento o registrar en Excel:\n{e}")


def abrir_submenu_eliminar():
    """
    Abre un submenú para eliminar incidencias.
    """
    def eliminar():
        seleccion = listbox_incidencias.curselection()
        if not seleccion:
            messagebox.showwarning("Seleccione una incidencia", "Debe seleccionar una incidencia para eliminarla.")
            return

        # Obtener el índice
        indice = seleccion[0]
        incidencia = incidencias[indice]

        # Confirmar eliminación
        confirm = messagebox.askyesno("Confirmar eliminación", f"¿Está seguro de que desea eliminar la incidencia?\n\n{incidencia}")
        if confirm:
            try:
                eliminar_incidencia(indice)
                actualizar_dashboard()
                messagebox.showinfo("Éxito", "Incidencia eliminada correctamente.")
                submenu.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar la incidencia:\n{e}")

    incidencias = obtener_incidencias()  # Obtener lista de incidencias desde Excel

    submenu = tk.Toplevel(root)
    submenu.title("Eliminar Incidencias")
    submenu.geometry("600x400")

    ttk.Label(submenu, text="Seleccione una incidencia para eliminar:").pack(pady=10)

    listbox_incidencias = tk.Listbox(submenu, height=15, width=80)
    for inc in incidencias:
        listbox_incidencias.insert(tk.END, inc)
    listbox_incidencias.pack(pady=10)

    btn_eliminar = ttk.Button(submenu, text="Eliminar", command=eliminar)
    btn_eliminar.pack(pady=10)


# ===================== INTERFAZ =====================
root = tk.Tk()
root.title("Bitácora de Incidencias")
root.geometry("900x650")

frm = ttk.Frame(root, padding=15)
frm.pack(fill="both", expand=True)

# ---- FILA 1: Fecha y hora
ttk.Label(frm, text="Fecha:").grid(row=0, column=0, sticky="w")
entry_fecha = ttk.Entry(frm)
entry_fecha.insert(0, datetime.now().strftime("%Y-%m-%d"))
entry_fecha.grid(row=0, column=1, sticky="ew", padx=5)

ttk.Label(frm, text="Hora:").grid(row=0, column=2, sticky="w")
entry_hora = ttk.Entry(frm)
entry_hora.insert(0, datetime.now().strftime("%H:%M"))
entry_hora.grid(row=0, column=3, sticky="ew", padx=5)

# ---- FILA 2: Lugar, Tipo, Gravedad
ttk.Label(frm, text="Lugar:").grid(row=1, column=0, sticky="w")
combo_lugar = ttk.Combobox(frm, values=locations, state="readonly")
combo_lugar.grid(row=1, column=1, sticky="ew", padx=5)

ttk.Label(frm, text="Tipo de incidencia:").grid(row=1, column=2, sticky="w")
combo_tipo = ttk.Combobox(frm, values=tipos, state="readonly")
combo_tipo.grid(row=1, column=3, sticky="ew", padx=5)

ttk.Label(frm, text="Gravedad:").grid(row=1, column=4, sticky="w")
combo_gravedad = ttk.Combobox(frm, values=gravedades, state="readonly")
combo_gravedad.grid(row=1, column=5, sticky="ew", padx=5)

# ---- FILA 3: Actividad
ttk.Label(frm, text="Actividad:").grid(row=2, column=0, sticky="w")
entry_actividad = ttk.Entry(frm, width=80)
entry_actividad.grid(row=2, column=1, columnspan=5, sticky="ew", pady=5)

# ---- FILA 4: Lista de alumnos
ttk.Label(frm, text="Alumnos implicados:").grid(row=3, column=0, sticky="nw")
listbox_alumnos = tk.Listbox(frm, selectmode="multiple", height=8, exportselection=False)
for alumno in alumnos:
    listbox_alumnos.insert(tk.END, alumno)
listbox_alumnos.grid(row=3, column=1, columnspan=5, sticky="ew", pady=5)

# ---- FILA 5: Narración
ttk.Label(frm, text="Narración:").grid(row=4, column=0, sticky="nw")
text_narracion = tk.Text(frm, height=5, width=70)
text_narracion.grid(row=4, column=1, columnspan=5, sticky="ew", pady=5)

# ---- FILA 6: Medidas
ttk.Label(frm, text="Medidas:").grid(row=5, column=0, sticky="nw")
text_medidas = tk.Text(frm, height=4, width=70)
text_medidas.grid(row=5, column=1, columnspan=5, sticky="ew", pady=5)

# ---- FILA 7: Seguimiento
ttk.Label(frm, text="Seguimiento:").grid(row=6, column=0, sticky="nw")
text_seguimiento = tk.Text(frm, height=4, width=70)
text_seguimiento.grid(row=6, column=1, columnspan=5, sticky="ew", pady=5)

# ---- FILA 8: Botones
btn_word = ttk.Button(frm, text="Generar Word + Guardar", command=generar_doc)
btn_word.grid(row=7, column=1, pady=15, sticky="ew")

btn_eliminar = ttk.Button(frm, text="Eliminar Incidencia", command=abrir_submenu_eliminar)
btn_eliminar.grid(row=7, column=3, pady=15, sticky="ew")

# Ajuste de columnas
for i in range(6):
    frm.columnconfigure(i, weight=1)

root.mainloop()
