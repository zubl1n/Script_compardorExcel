import pandas as pd
import tkinter as tk
from tkinter import filedialog

def buscar_todos_los_datos(tabla_origen, tabla_destino, columna_clave):
    df_origen = pd.read_excel(tabla_origen)
    df_destino = pd.read_excel(tabla_destino)
    datos_coincidentes = df_destino[df_destino[columna_clave].isin(df_origen[columna_clave])]
    datos_coincidentes.to_excel('tabla_destino_actualizada.xlsx', index=False)

def buscar_datos_btn_click():
    tabla_origen_path = filedialog.askopenfilename(title="Seleccionar tabla origen")
    tabla_destino_path = filedialog.askopenfilename(title="Seleccionar tabla destino")
    columna_clave = columna_clave_entry.get()

    if tabla_origen_path and tabla_destino_path and columna_clave:
        buscar_todos_los_datos(tabla_origen_path, tabla_destino_path, columna_clave)
        resultado_label.config(text="Búsqueda y actualización completada. Archivo Excel generado.")
    else:
        resultado_label.config(text="Por favor, ingresa todas las opciones.")

# Crear la ventana de la aplicación
app = tk.Tk()
app.title("Unificar Planillas")
app.geometry("400x200")

# Etiquetas y campos de entrada
#tk.Label(app, text="Ruta tabla origen:").pack()
#tabla_origen_entry = tk.Entry(app, width=40)
#tabla_origen_entry.pack()

#tk.Label(app, text="Ruta tabla destino:").pack()
#tabla_destino_entry = tk.Entry(app, width=40)
#tabla_destino_entry.pack()

tk.Label(app, text="Columna clave (SKU):").pack()
columna_clave_entry = tk.Entry(app, width=20)
columna_clave_entry.pack()

# Botón para buscar y unificar datos
buscar_btn = tk.Button(app, text="Buscar y Unificar", command=buscar_datos_btn_click)
buscar_btn.pack()

# Etiqueta para mostrar el resultado
resultado_label = tk.Label(app, text="")
resultado_label.pack()

# Ejecutar la aplicación
app.mainloop()
