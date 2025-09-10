# -*- coding: utf-8 -*-
"""
Archivo: main.py
Interfaz principal para registrar incidencias escolares.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os  # 👈 Added os import for file path handling

from resources import load_all_resources
from wordgen import generar_word
from excelgen import registrar_incidencia, actualizar_dashboard, inicializar_excel
# 👇 CORRECTED: Using the correct variable names from config.py
from config import SCHOOL_NAME, LOCATION, DIRECTOR_NAME, TEACHER_NAME, GRADE, GROUP, INCIDENCIAS_DIR


# ===================== INICIALIZACIÓN =====================
inicializar_excel()

# Cargamos recursos
alumnos, padres, locations, tipos = load_all_resources()
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
    if not participantes:
        messagebox.showwarning("Falta información", "Debe seleccionar al menos un alumno.")
        listbox_alumnos.focus_set()
        return
    if not tipo or not lugar or not gravedad:
        messagebox.showwarning("Falta información", "Complete todos los menús desplegables.")
        return

    # --- 👇 CORRECTED: Function call logic ---

    # 1. Prepare data dictionary for Excel
    datos_excel = {
        "fecha": fecha, "hora": hora, "lugar": lugar, "actividad": actividad,
        "tipo": tipo, "gravedad": gravedad, "participantes": participantes,
        "narracion": narracion, "medidas": medidas, "seguimiento": seguimiento
    }
    
    # 2. Define a unique filename for the Word document
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombres_alumnos = "_".join(participantes).replace(" ", "")
    output_filename = f"Incidencia_{nombres_alumnos}_{timestamp}.docx"
    output_path = os.path.join(INCIDENCIAS_DIR, output_filename)

    # 3. Call generar_word with all required arguments, not just one dictionary
    try:
        generar_word(
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
            padres_dict=padres,  # Pass the loaded parents dictionary
            alumnos_seleccionados=participantes, # Pass the list of selected students
            output_path=output_path # Pass the full path for the output file
        )
        
        # 4. Guardar en Excel only after Word generation is successful
        registrar_incidencia(datos_excel)

        messagebox.showinfo("Éxito", f"Incidencia registrada correctamente.\nWord guardado en: {output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el documento: {e}")


def actualizar_excel():
    try:
        actualizar_dashboard()
        messagebox.showinfo("Excel", "Dashboard actualizado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo actualizar el dashboard: {e}")


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

btn_excel = ttk.Button(frm, text="Actualizar Dashboard Excel", command=actualizar_excel)
btn_excel.grid(row=7, column=3, pady=15, sticky="ew")


# Ajuste de columnas
for i in range(6):
    frm.columnconfigure(i, weight=1)


root.mainloop()