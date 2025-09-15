# -- coding utf-8 --

Archivo menu_eliminar.py
Descripción Menú independiente para listar y eliminar incidencias.


import tkinter as tk
from tkinter import ttk, messagebox
from excelgen import obtener_incidencias, eliminar_incidencia, actualizar_dashboard


def abrir_menu_eliminar()
    
    Abre una ventana para listar y eliminar incidencias.
    
    def eliminar()
        seleccion = listbox_incidencias.curselection()
        if not seleccion
            messagebox.showwarning(Seleccione una incidencia, Debe seleccionar una incidencia para eliminarla.)
            return

        # Obtener el índice
        indice = seleccion[0]
        incidencia = incidencias[indice]

        # Confirmar eliminación
        confirm = messagebox.askyesno(Confirmar eliminación, f¿Está seguro de que desea eliminar la incidenciann{incidencia})
        if confirm
            try
                eliminar_incidencia(indice)
                actualizar_dashboard()
                messagebox.showinfo(Éxito, Incidencia eliminada correctamente.)
                listbox_incidencias.delete(indice)  # Eliminar de la lista
                incidencias.pop(indice)  # Eliminar del backend
            except Exception as e
                messagebox.showerror(Error, fOcurrió un error al eliminar la incidencian{e})

    incidencias = obtener_incidencias()

    root = tk.Tk()
    root.title(Gestión de Incidencias)
    root.geometry(600x400)

    ttk.Label(root, text=Seleccione una incidencia para eliminar).pack(pady=10)

    listbox_incidencias = tk.Listbox(root, height=15, width=80)
    for inc in incidencias
        listbox_incidencias.insert(tk.END, inc)
    listbox_incidencias.pack(pady=10)

    btn_eliminar = ttk.Button(root, text=Eliminar, command=eliminar)
    btn_eliminar.pack(pady=10)

    root.mainloop()


if __name__ == __main__
    abrir_menu_eliminar()
