# -*- coding: utf-8 -*-
"""
Archivo: programa.py
Interfaz principal para registrar y gestionar incidencias escolares.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import json # Necesario para la configuración

import setup  # Importar el nuevo módulo de configuración
setup.run_setup()  # Ejecutar la configuración inicial

from wordgen import generar_word
from excelgen import registrar_incidencia, actualizar_dashboard, inicializar_excel
from resources import load_all_resources
import json_manager as jm

# --- Cargar configuración global ---
CONFIG = jm.obtener_config()
INCIDENCIAS_DIR = CONFIG.get("incidencias_dir", "incidencias") # Usar valor de config o default
SCHOOL_NAME = CONFIG.get("school_name", "Nombre Escuela")
LOCATION = CONFIG.get("location", "Ubicación Escuela")
DIRECTOR_NAME = CONFIG.get("director_name", "Nombre Director")
TEACHER_NAME = CONFIG.get("teacher_name", "Nombre Maestro")
GRADE = CONFIG.get("grade", "1")
GROUP = CONFIG.get("group", "A")

# ===================== VARIABLES GLOBALES =====================
alumnos_data_global, padres_data_global, locations_data_global, tipos_data_global = [], {}, [], []
alumnos_externos = []
maestros_externos = []

# ===================== INICIALIZACIÓN =====================
def inicializar_sistema():
    """Inicializa Excel y crea directorios necesarios."""
    inicializar_excel()
    os.makedirs(INCIDENCIAS_DIR, exist_ok=True)

# ===================== FUNCIONES DE LA APLICACIÓN =====================

# --- Funciones de la Pestaña de Registro ---

def recargar_recursos_y_actualizar_ui():
    """Recarga los datos desde los JSON y actualiza los widgets."""
    global alumnos_data_global, padres_data_global, locations_data_global, tipos_data_global
    global CONFIG, INCIDENCIAS_DIR, SCHOOL_NAME, LOCATION, DIRECTOR_NAME, TEACHER_NAME, GRADE, GROUP
    
    # Recargar configuración
    CONFIG = jm.obtener_config()
    INCIDENCIAS_DIR = CONFIG.get("incidencias_dir", "incidencias")
    SCHOOL_NAME = CONFIG.get("school_name", "Nombre Escuela")
    LOCATION = CONFIG.get("location", "Ubicación Escuela")
    DIRECTOR_NAME = CONFIG.get("director_name", "Nombre Director")
    TEACHER_NAME = CONFIG.get("teacher_name", "Nombre Maestro")
    GRADE = CONFIG.get("grade", "1")
    GROUP = CONFIG.get("group", "A")

    # Cargar otros recursos
    alumnos_data_global, padres_data_global, locations_data_global, tipos_data_global = load_all_resources()

    # Actualizar comboboxes
    combo_lugar['values'] = locations_data_global
    combo_tipo['values'] = tipos_data_global

    # Actualizar listbox de alumnos (solo los del grupo actual)
    actualizar_lista_alumnos_grupo()
    
    # Actualizar la pestaña de administración
    poblar_treeview_alumnos()
    poblar_listbox_ubicaciones()
    poblar_listbox_tipos()
    poblar_campos_config() # Llenar campos de configuración

def actualizar_lista_alumnos_grupo():
    """Filtra y muestra solo los alumnos del grupo actual."""
    listbox_alumnos.delete(0, tk.END)
    if alumnos_data_global: # Asegurarse de que alumnos no esté vacío
        for alumno in alumnos_data_global:
            if alumno.get("grado") == GRADE and alumno.get("grupo") == GROUP:
                listbox_alumnos.insert(tk.END, alumno.get("nombre", ""))

def toggle_alumnos_externos():
    if var_check_externos.get():
        frame_externos.pack(fill="x", expand=True, padx=10, pady=5)
    else:
        frame_externos.pack_forget()

def toggle_maestros_externos():
    if var_check_maestros.get():
        frame_maestros_externos.pack(fill="x", expand=True, padx=10, pady=5)
    else:
        frame_maestros_externos.pack_forget()

def agregar_alumno_externo():
    nombre = entry_nombre_externo.get()
    grado = entry_grado_externo.get()
    grupo = entry_grupo_externo.get()
    if not all([nombre, grado, grupo]):
        messagebox.showwarning("Datos incompletos", "Debe rellenar nombre, grado y grupo.")
        return
    alumnos_externos.append({"nombre": nombre, "grado": grado, "grupo": grupo})
    listbox_externos.insert(tk.END, f"{nombre} ({grado}° '{grupo}')")
    for entry in [entry_nombre_externo, entry_grado_externo, entry_grupo_externo]:
        entry.delete(0, tk.END)

def quitar_alumno_externo():
    seleccion = listbox_externos.curselection()
    if not seleccion: return
    del alumnos_externos[seleccion[0]]
    listbox_externos.delete(seleccion[0])

def agregar_maestro_externo():
    nombre = entry_nombre_maestro.get()
    grupo = entry_grupo_maestro.get()
    if not all([nombre, grupo]):
        messagebox.showwarning("Datos incompletos", "Debe rellenar nombre y grupo/asignatura.")
        return
    maestros_externos.append({"nombre": nombre, "grupo": grupo})
    listbox_maestros.insert(tk.END, f"{nombre} ({grupo})")
    entry_nombre_maestro.delete(0, tk.END)
    entry_grupo_maestro.delete(0, tk.END)

def quitar_maestro_externo():
    seleccion = listbox_maestros.curselection()
    if not seleccion: return
    del maestros_externos[seleccion[0]]
    listbox_maestros.delete(seleccion[0])

def limpiar_formulario():
    entry_fecha.delete(0, tk.END)
    entry_fecha.insert(0, datetime.now().strftime("%Y-%m-%d"))
    entry_hora.delete(0, tk.END)
    entry_hora.insert(0, datetime.now().strftime("%H:%M"))
    for combo in [combo_lugar, combo_tipo, combo_gravedad]:
        combo.set('')
    entry_actividad.delete(0, tk.END)
    listbox_alumnos.selection_clear(0, tk.END)
    for text_widget in [text_narracion, text_medidas, text_seguimiento]:
        text_widget.delete("1.0", tk.END)
    
    alumnos_externos.clear()
    listbox_externos.delete(0, tk.END)
    var_check_externos.set(False)
    toggle_alumnos_externos()

    maestros_externos.clear()
    listbox_maestros.delete(0, tk.END)
    var_check_maestros.set(False)
    toggle_maestros_externos()
    
def generar_doc():
    # Recopilación de datos
    datos = {
        "fecha": entry_fecha.get(), "hora": entry_hora.get(), "lugar": combo_lugar.get(),
        "actividad": entry_actividad.get(), "tipo_inc": combo_tipo.get(), "gravedad": combo_gravedad.get(),
        "narracion": text_narracion.get("1.0", tk.END).strip(),
        "medidas": text_medidas.get("1.0", tk.END).strip(),
        "seguimiento": text_seguimiento.get("1.0", tk.END).strip()
    }

    # Recopilar participantes
    participantes = []
    for i in listbox_alumnos.curselection():
        # Buscar el alumno completo en alumnos_data_global para obtener grado y grupo
        nombre_alumno_seleccionado = listbox_alumnos.get(i)
        for alumno_completo in alumnos_data_global:
            if alumno_completo.get("nombre") == nombre_alumno_seleccionado:
                participantes.append({"nombre": alumno_completo.get("nombre"), 
                                      "grado": alumno_completo.get("grado"), 
                                      "grupo": alumno_completo.get("grupo")})
                break
    participantes.extend(alumnos_externos)

    if not all([participantes, datos["tipo_inc"], datos["lugar"], datos["gravedad"]]):
        messagebox.showwarning("Falta información", "Debe seleccionar al menos un alumno y completar todos los menús desplegables.")
        return

    # Preparar y generar documentos
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombres_alumnos = "_".join(p["nombre"] for p in participantes).replace(" ", "")
    output_filename = f"Incidencia_{nombres_alumnos}_{timestamp}.docx"
    output_path = os.path.join(INCIDENCIAS_DIR, output_filename)

    try:
        generar_word(
            **datos, participantes=participantes, padres_dict=padres_data_global,
            output_path=output_path, maestros_externos=maestros_externos,
            school_name=SCHOOL_NAME, director_name=DIRECTOR_NAME,
            teacher_name=TEACHER_NAME, grade=GRADE, group=GROUP
        )
        datos_excel = {k: v for k, v in datos.items() if k in ["fecha", "hora", "lugar", "gravedad"]}
        datos_excel["participantes"] = participantes
        datos_excel["link"] = output_path
        registrar_incidencia(datos_excel)
        actualizar_dashboard()
        messagebox.showinfo("Éxito", f"Incidencia registrada.\nWord guardado en: {output_path}")
        limpiar_formulario()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el documento o registrar en Excel:\n{e}")

# --- Funciones de la Pestaña de Administración ---

def poblar_treeview_alumnos():
    tree_alumnos.delete(*tree_alumnos.get_children())
    for alumno in jm.obtener_alumnos():
        tree_alumnos.insert("", "end", values=(alumno.get('nombre', ''), alumno.get('padre', ''), alumno.get('grado', ''), alumno.get('grupo', '')))

def on_alumno_select(event):
    if not tree_alumnos.selection(): return
    item = tree_alumnos.selection()[0]
    values = tree_alumnos.item(item, "values")
    
    # Llenar campos y deshabilitar
    entry_admin_nombre.delete(0, tk.END); entry_admin_nombre.insert(0, values[0])
    entry_admin_padre.delete(0, tk.END); entry_admin_padre.insert(0, values[1])
    entry_admin_grado.delete(0, tk.END); entry_admin_grado.insert(0, values[2])
    entry_admin_grupo.delete(0, tk.END); entry_admin_grupo.insert(0, values[3])
    
    set_state_admin_alumnos(tk.DISABLED)
    btn_modificar_alumno.config(state=tk.NORMAL)
    btn_guardar_alumno.config(state=tk.DISABLED)

def set_state_admin_alumnos(state):
    """Establece el estado (normal/disabled) de los campos de edición de alumnos."""
    for widget in [entry_admin_nombre, entry_admin_padre, entry_admin_grado, entry_admin_grupo]:
        widget.config(state=state)

def habilitar_edicion_alumno():
    set_state_admin_alumnos(tk.NORMAL)
    btn_modificar_alumno.config(state=tk.DISABLED)
    btn_guardar_alumno.config(state=tk.NORMAL)
    # Opcional: Deshabilitar selección en el Treeview mientras se edita
    # tree_alumnos.config(state=tk.DISABLED)

def guardar_cambios_alumno():
    if not tree_alumnos.selection():
        messagebox.showwarning("Sin selección", "Seleccione un alumno para guardar cambios.")
        return
    
    nombre_original = tree_alumnos.item(tree_alumnos.selection()[0], "values")[0]
    alumnos_data = jm.obtener_alumnos()
    for alumno in alumnos_data:
        if alumno["nombre"] == nombre_original:
            alumno["nombre"] = entry_admin_nombre.get()
            alumno["padre"] = entry_admin_padre.get()
            alumno["grado"] = entry_admin_grado.get()
            alumno["grupo"] = entry_admin_grupo.get()
            break
    jm.guardar_alumnos(alumnos_data)
    recargar_recursos_y_actualizar_ui()
    limpiar_campos_admin_alumnos()
    set_state_admin_alumnos(tk.DISABLED) # Volver a deshabilitar
    btn_modificar_alumno.config(state=tk.NORMAL)
    btn_guardar_alumno.config(state=tk.DISABLED)
    # tree_alumnos.config(state=tk.NORMAL) # Habilitar selección de nuevo

def eliminar_alumno():
    if not tree_alumnos.selection():
        messagebox.showwarning("Sin selección", "Seleccione un alumno para eliminar.")
        return
    nombre_a_eliminar = tree_alumnos.item(tree_alumnos.selection()[0], "values")[0]
    alumnos_data = [a for a in jm.obtener_alumnos() if a["nombre"] != nombre_a_eliminar]
    jm.guardar_alumnos(alumnos_data)
    recargar_recursos_y_actualizar_ui()
    limpiar_campos_admin_alumnos()

def limpiar_campos_admin_alumnos():
    for entry in [entry_admin_nombre, entry_admin_padre, entry_admin_grado, entry_admin_grupo]:
        entry.delete(0, tk.END)
    if tree_alumnos.selection():
        tree_alumnos.selection_remove(tree_alumnos.selection())
    # Asegurarse de que los campos queden deshabilitados si se limpian
    set_state_admin_alumnos(tk.DISABLED)
    btn_modificar_alumno.config(state=tk.NORMAL)
    btn_guardar_alumno.config(state=tk.DISABLED)

def poblar_listbox_ubicaciones():
    listbox_ubicaciones.delete(0, tk.END)
    for item in jm.obtener_ubicaciones():
        listbox_ubicaciones.insert(tk.END, item)

def agregar_ubicacion():
    nueva = entry_admin_ubicacion.get()
    if not nueva: return
    data = jm.obtener_ubicaciones()
    if nueva not in data:
        data.append(nueva)
        jm.guardar_ubicaciones(data)
        recargar_recursos_y_actualizar_ui()
    entry_admin_ubicacion.delete(0, tk.END)

def eliminar_ubicacion():
    seleccion = listbox_ubicaciones.curselection()
    if not seleccion: return
    a_eliminar = listbox_ubicaciones.get(seleccion[0])
    data = [u for u in jm.obtener_ubicaciones() if u != a_eliminar]
    jm.guardar_ubicaciones(data)
    recargar_recursos_y_actualizar_ui()

def poblar_listbox_tipos():
    listbox_tipos.delete(0, tk.END)
    for item in jm.obtener_tipos_incidencia():
        listbox_tipos.insert(tk.END, item)

def agregar_tipo():
    nuevo = entry_admin_tipo.get()
    if not nuevo: return
    data = jm.obtener_tipos_incidencia()
    if nuevo not in data:
        data.append(nuevo)
        jm.guardar_tipos_incidencia(data)
        recargar_recursos_y_actualizar_ui()
    entry_admin_tipo.delete(0, tk.END)

def eliminar_tipo():
    seleccion = listbox_tipos.curselection()
    if not seleccion: return
    a_eliminar = listbox_tipos.get(seleccion[0])
    data = [t for t in jm.obtener_tipos_incidencia() if t != a_eliminar]
    jm.guardar_tipos_incidencia(data)
    recargar_recursos_y_actualizar_ui()

# --- Funciones para la Pestaña de Configuración General ---
def poblar_campos_config():
    config = jm.obtener_config()
    entry_config_teacher.delete(0, tk.END); entry_config_teacher.insert(0, config.get("teacher_name", ""))
    entry_config_grade.delete(0, tk.END); entry_config_grade.insert(0, config.get("grade", ""))
    entry_config_group.delete(0, tk.END); entry_config_group.insert(0, config.get("group", ""))
    entry_config_director.delete(0, tk.END); entry_config_director.insert(0, config.get("director_name", ""))
    entry_config_school.delete(0, tk.END); entry_config_school.insert(0, config.get("school_name", ""))
    entry_config_location.delete(0, tk.END); entry_config_location.insert(0, config.get("location", ""))

def habilitar_edicion_config():
    for widget in [entry_config_teacher, entry_config_grade, entry_config_group, entry_config_director, entry_config_school, entry_config_location]:
        widget.config(state=tk.NORMAL)
    btn_modificar_config.config(state=tk.DISABLED)
    btn_guardar_config.config(state=tk.NORMAL)

def guardar_configuracion():
    nueva_config = {
        "teacher_name": entry_config_teacher.get(),
        "grade": entry_config_grade.get(),
        "group": entry_config_group.get(),
        "director_name": entry_config_director.get(),
        "school_name": entry_config_school.get(),
        "location": entry_config_location.get(),
        "incidencias_dir": INCIDENCIAS_DIR # Mantener el directorio de incidencias
    }
    if jm.guardar_config(nueva_config):
        messagebox.showinfo("Guardado", "Configuración guardada exitosamente.")
        recargar_recursos_y_actualizar_ui() # Recargar todo para aplicar cambios
        # Volver a deshabilitar los campos después de guardar
        for widget in [entry_config_teacher, entry_config_grade, entry_config_group, entry_config_director, entry_config_school, entry_config_location]:
            widget.config(state=tk.DISABLED)
        btn_modificar_config.config(state=tk.NORMAL)
        btn_guardar_config.config(state=tk.DISABLED)
    else:
        messagebox.showerror("Error", "No se pudo guardar la configuración.")

# ===================== INTERFAZ GRÁFICA =====================
root = tk.Tk()
root.title("Bitácora de Incidencias")
root.geometry("950x700")
style = ttk.Style(root)
style.theme_use("clam")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# --- Pestaña de Registro ---
tab_registro = ttk.Frame(notebook)
notebook.add(tab_registro, text="Registrar Nueva Incidencia")

canvas = tk.Canvas(tab_registro)
scrollbar = ttk.Scrollbar(tab_registro, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# --- Widgets de la Pestaña de Registro ---
frame_info = ttk.LabelFrame(scrollable_frame, text="Información del Evento", padding=10)
frame_info.pack(fill="x", expand=True, padx=10, pady=5)
ttk.Label(frame_info, text="Fecha:").grid(row=0, column=0, sticky="w", pady=2)
entry_fecha = ttk.Entry(frame_info); entry_fecha.grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(frame_info, text="Hora:").grid(row=0, column=2, sticky="w", padx=10)
entry_hora = ttk.Entry(frame_info); entry_hora.grid(row=0, column=3, sticky="ew", padx=5)
ttk.Label(frame_info, text="Lugar:").grid(row=1, column=0, sticky="w", pady=2)
combo_lugar = ttk.Combobox(frame_info, state="readonly"); combo_lugar.grid(row=1, column=1, sticky="ew", padx=5)
ttk.Label(frame_info, text="Tipo:").grid(row=1, column=2, sticky="w", padx=10)
combo_tipo = ttk.Combobox(frame_info, state="readonly"); combo_tipo.grid(row=1, column=3, sticky="ew", padx=5)
ttk.Label(frame_info, text="Gravedad:").grid(row=1, column=4, sticky="w", padx=10)
combo_gravedad = ttk.Combobox(frame_info, values=["Leve", "Moderada", "Grave"], state="readonly"); combo_gravedad.grid(row=1, column=5, sticky="ew", padx=5)
ttk.Label(frame_info, text="Actividad:").grid(row=2, column=0, sticky="w", pady=2)
entry_actividad = ttk.Entry(frame_info); entry_actividad.grid(row=2, column=1, columnspan=5, sticky="ew", padx=5, pady=5)
for i in [1, 3, 5]: frame_info.columnconfigure(i, weight=1)

frame_alumnos = ttk.LabelFrame(scrollable_frame, text="Personas Implicadas", padding=10)
frame_alumnos.pack(fill="x", expand=True, padx=10, pady=5)
ttk.Label(frame_alumnos, text=f"Alumnos de {GRADE}° '{GROUP}':").pack(anchor="w")
listbox_alumnos = tk.Listbox(frame_alumnos, selectmode="multiple", height=6, exportselection=False); listbox_alumnos.pack(fill="x", expand=True, pady=5)
var_check_externos = tk.BooleanVar()
ttk.Checkbutton(frame_alumnos, text="¿Incluir alumno de otro grupo?", variable=var_check_externos, command=toggle_alumnos_externos).pack(anchor="w", pady=5)
frame_externos = ttk.Frame(frame_alumnos, padding=5)
ttk.Label(frame_externos, text="Nombre:").grid(row=0, column=0, sticky="w")
entry_nombre_externo = ttk.Entry(frame_externos); entry_nombre_externo.grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(frame_externos, text="Grado:").grid(row=0, column=2, sticky="w", padx=5)
entry_grado_externo = ttk.Entry(frame_externos, width=5); entry_grado_externo.grid(row=0, column=3, sticky="w")
ttk.Label(frame_externos, text="Grupo:").grid(row=0, column=4, sticky="w", padx=5)
entry_grupo_externo = ttk.Entry(frame_externos, width=5); entry_grupo_externo.grid(row=0, column=5, sticky="w")
btn_agregar_externo = ttk.Button(frame_externos, text="Agregar", command=agregar_alumno_externo); btn_agregar_externo.grid(row=0, column=6, padx=10)
frame_externos.columnconfigure(1, weight=1)
ttk.Label(frame_externos, text="Alumnos externos añadidos:").grid(row=1, column=0, columnspan=7, sticky="w", pady=(10, 2))
listbox_externos = tk.Listbox(frame_externos, height=4); listbox_externos.grid(row=2, column=0, columnspan=6, sticky="ew")
btn_quitar_externo = ttk.Button(frame_externos, text="Quitar", command=quitar_alumno_externo); btn_quitar_externo.grid(row=2, column=6, padx=10, sticky="n")

var_check_maestros = tk.BooleanVar()
ttk.Checkbutton(frame_alumnos, text="¿Incluir maestro de otro grupo?", variable=var_check_maestros, command=toggle_maestros_externos).pack(anchor="w", pady=5)
frame_maestros_externos = ttk.Frame(frame_alumnos, padding=5)
ttk.Label(frame_maestros_externos, text="Nombre:").grid(row=0, column=0, sticky="w")
entry_nombre_maestro = ttk.Entry(frame_maestros_externos); entry_nombre_maestro.grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(frame_maestros_externos, text="Grupo/Asignatura:").grid(row=0, column=2, sticky="w", padx=5)
entry_grupo_maestro = ttk.Entry(frame_maestros_externos, width=15); entry_grupo_maestro.grid(row=0, column=3, sticky="w")
btn_agregar_maestro = ttk.Button(frame_maestros_externos, text="Agregar", command=agregar_maestro_externo); btn_agregar_maestro.grid(row=0, column=4, padx=10)
frame_maestros_externos.columnconfigure(1, weight=1)
ttk.Label(frame_maestros_externos, text="Maestros externos añadidos:").grid(row=1, column=0, columnspan=5, sticky="w", pady=(10, 2))
listbox_maestros = tk.Listbox(frame_maestros_externos, height=3); listbox_maestros.grid(row=2, column=0, columnspan=4, sticky="ew")
btn_quitar_maestro = ttk.Button(frame_maestros_externos, text="Quitar", command=quitar_maestro_externo); btn_quitar_maestro.grid(row=2, column=4, padx=10, sticky="n")

frame_desc = ttk.LabelFrame(scrollable_frame, text="Descripción de los Hechos y Acciones", padding=10)
frame_desc.pack(fill="x", expand=True, padx=10, pady=5)
ttk.Label(frame_desc, text="Narración:").pack(anchor="w")
text_narracion = tk.Text(frame_desc, height=5); text_narracion.pack(fill="x", expand=True, pady=2)
ttk.Label(frame_desc, text="Medidas:").pack(anchor="w")
text_medidas = tk.Text(frame_desc, height=3); text_medidas.pack(fill="x", expand=True, pady=2)
ttk.Label(frame_desc, text="Seguimiento:").pack(anchor="w")
text_seguimiento = tk.Text(frame_desc, height=3); text_seguimiento.pack(fill="x", expand=True, pady=2)

frame_botones = ttk.Frame(scrollable_frame, padding=10)
frame_botones.pack(fill="x")
btn_generar = ttk.Button(frame_botones, text="Generar y Registrar", command=generar_doc); btn_generar.pack(side="right", padx=5)
btn_limpiar = ttk.Button(frame_botones, text="Limpiar Formulario", command=limpiar_formulario); btn_limpiar.pack(side="right")

# --- Pestaña de Administración ---
tab_admin = ttk.Frame(notebook)
notebook.add(tab_admin, text="Administrar Datos")
admin_notebook = ttk.Notebook(tab_admin)
admin_notebook.pack(fill="both", expand=True, padx=5, pady=5)

# Sub-pestaña Alumnos
tab_admin_alumnos = ttk.Frame(admin_notebook)
admin_notebook.add(tab_admin_alumnos, text="Alumnos")
frame_tree = ttk.Frame(tab_admin_alumnos); frame_tree.pack(fill="both", expand=True, pady=5)
cols = ("Nombre", "Padre/Madre", "Grado", "Grupo")
tree_alumnos = ttk.Treeview(frame_tree, columns=cols, show="headings")
for col in cols:
    tree_alumnos.heading(col, text=col)
tree_alumnos.pack(side="left", fill="both", expand=True)
tree_scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree_alumnos.yview)
tree_alumnos.configure(yscrollcommand=tree_scrollbar.set)
tree_scrollbar.pack(side="right", fill="y")
tree_alumnos.bind("<<TreeviewSelect>>", on_alumno_select)

frame_form_admin = ttk.LabelFrame(tab_admin_alumnos, text="Datos del Alumno", padding=10); frame_form_admin.pack(fill="x", pady=5)
ttk.Label(frame_form_admin, text="Nombre:").grid(row=0, column=0, sticky="w")
entry_admin_nombre = ttk.Entry(frame_form_admin); entry_admin_nombre.grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(frame_form_admin, text="Padre/Madre:").grid(row=0, column=2, sticky="w", padx=10)
entry_admin_padre = ttk.Entry(frame_form_admin); entry_admin_padre.grid(row=0, column=3, sticky="ew", padx=5)
ttk.Label(frame_form_admin, text="Grado:").grid(row=1, column=0, sticky="w")
entry_admin_grado = ttk.Entry(frame_form_admin); entry_admin_grado.grid(row=1, column=1, sticky="ew", padx=5)
ttk.Label(frame_form_admin, text="Grupo:").grid(row=1, column=2, sticky="w", padx=10)
entry_admin_grupo = ttk.Entry(frame_form_admin); entry_admin_grupo.grid(row=1, column=3, sticky="ew", padx=5)
frame_form_admin.columnconfigure(1, weight=1); frame_form_admin.columnconfigure(3, weight=1)

frame_botones_admin = ttk.Frame(tab_admin_alumnos); frame_botones_admin.pack(fill="x", pady=5)
ttk.Button(frame_botones_admin, text="Agregar", command=agregar_alumno).pack(side="left", padx=5)
btn_modificar_alumno = ttk.Button(frame_botones_admin, text="Modificar", command=habilitar_edicion_alumno)
btn_modificar_alumno.pack(side="left", padx=5)
btn_guardar_alumno = ttk.Button(frame_botones_admin, text="Guardar Cambios", command=guardar_cambios_alumno, state=tk.DISABLED)
btn_guardar_alumno.pack(side="left", padx=5)
ttk.Button(frame_botones_admin, text="Eliminar", command=eliminar_alumno).pack(side="left", padx=5)
ttk.Button(frame_botones_admin, text="Limpiar Campos", command=limpiar_campos_admin_alumnos).pack(side="right", padx=5)

# Sub-pestaña Ubicaciones y Tipos
for tab_frame, getter_func, adder_func, remover_func, label_text in [
    (ttk.Frame(admin_notebook), jm.obtener_ubicaciones, agregar_ubicacion, eliminar_ubicacion, "Ubicación:"),
    (ttk.Frame(admin_notebook), jm.obtener_tipos_incidencia, agregar_tipo, eliminar_tipo, "Tipo de Incidencia:")
]:
    admin_notebook.add(tab_frame, text=label_text.replace(":", ""))
    
    listbox_current = tk.Listbox(tab_frame); listbox_current.pack(fill="both", expand=True, padx=5, pady=5)
    frame_add_current = ttk.Frame(tab_frame); frame_add_current.pack(fill="x", padx=5, pady=5)
    ttk.Label(frame_add_current, text=label_text).pack(side="left")
    entry_current = ttk.Entry(frame_add_current); entry_current.pack(side="left", fill="x", expand=True, padx=5)
    ttk.Button(frame_add_current, text="Agregar", command=adder_func).pack(side="left")
    ttk.Button(tab_frame, text="Eliminar Selección", command=remover_func).pack(pady=5)
    
    if "Ubicación" in label_text:
        listbox_ubicaciones = listbox_current
        entry_admin_ubicacion = entry_current
    else:
        listbox_tipos = listbox_current
        entry_admin_tipo = entry_current

# --- Pestaña de Configuración General ---
tab_config = ttk.Frame(admin_notebook)
admin_notebook.add(tab_config, text="Configuración General")

frame_config_general = ttk.LabelFrame(tab_config, text="Ajustes Generales", padding=10)
frame_config_general.pack(fill="x", expand=True, padx=10, pady=5)

ttk.Label(frame_config_general, text="Nombre Maestro:").grid(row=0, column=0, sticky="w", pady=2)
entry_config_teacher = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_teacher.grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(frame_config_general, text="Grado:").grid(row=0, column=2, sticky="w", padx=10)
entry_config_grade = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_grade.grid(row=0, column=3, sticky="ew", padx=5)
ttk.Label(frame_config_general, text="Grupo:").grid(row=0, column=4, sticky="w", padx=10)
entry_config_group = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_group.grid(row=0, column=5, sticky="ew", padx=5)

ttk.Label(frame_config_general, text="Director:").grid(row=1, column=0, sticky="w", pady=2)
entry_config_director = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_director.grid(row=1, column=1, sticky="ew", padx=5)
ttk.Label(frame_config_general, text="Escuela:").grid(row=1, column=2, sticky="w", padx=10)
entry_config_school = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_school.grid(row=1, column=3, sticky="ew", padx=5)
ttk.Label(frame_config_general, text="Ubicación:").grid(row=1, column=4, sticky="w", padx=10)
entry_config_location = ttk.Entry(frame_config_general, state=tk.DISABLED); entry_config_location.grid(row=1, column=5, sticky="ew", padx=5)

for i in [1, 3, 5]: frame_config_general.columnconfigure(i, weight=1)

frame_config_botones = ttk.Frame(tab_config); frame_config_botones.pack(fill="x", pady=5)
btn_modificar_config = ttk.Button(frame_config_botones, text="Modificar", command=habilitar_edicion_config)
btn_modificar_config.pack(side="left", padx=5)
btn_guardar_config = ttk.Button(frame_config_botones, text="Guardar Cambios", command=guardar_configuracion, state=tk.DISABLED)
btn_guardar_config.pack(side="left", padx=5)

# --- Inicialización Final ---
inicializar_sistema()
recargar_recursos_y_actualizar_ui()
root.mainloop()
