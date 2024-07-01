import tkinter as tk
import re
from tkinter import filedialog, messagebox, ttk
import sv_ttk  # Importar el tema sv_ttk
import os
from bancos import get_codigos_bancos, get_bancos
import pandas as pd  # Importar pandas para manejar archivos CSV y Excel
import pdfplumber  # Para manejo de archivos PDF

# Función para seleccionar archivo PDF
from extractor import extract_data_from_pdf


def cargar_archivo():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        entry_var.set(file_path)
        actualizar_info_pdf(file_path)


# Función para actualizar la información del PDF
def actualizar_info_pdf(file_path):
    try:
        with pdfplumber.open(file_path) as pdf:
            num_paginas = len(pdf.pages)
        info_label_var.set(f"El PDF tiene {num_paginas} páginas.")
        max_pages_var.set(num_paginas)
    except Exception as e:
        info_label_var.set(f"Error al leer el PDF: {e}")
        max_pages_var.set(0)


# Función para convertir extracto
def convertir_extracto():
    pdf_path = entry_var.get()
    banco = banco_var.get()

    try:
        if not os.path.isfile(pdf_path):
            raise ValueError("Ruta del archivo no válida. Seleccione un archivo PDF válido.")

        pagina_inicio = int(start_page_var.get())
        pagina_fin = int(end_page_var.get())
        max_pages = max_pages_var.get()

        if pagina_inicio <= 0 or pagina_fin <= 0:
            raise ValueError("El número de página debe ser mayor a 0")
        if pagina_inicio > pagina_fin:
            raise ValueError("La página de inicio no puede ser mayor que la página final")
        if pagina_fin > max_pages:
            raise ValueError(f"El PDF solo tiene {max_pages} páginas. Ingrese un valor válido.")

    except ValueError as e:
        if "invalid literal" in str(e):
            messagebox.showwarning("Advertencia", "Existen datos ingresados no válidos para ejecutar la conversión.")
        else:
            messagebox.showwarning("Advertencia", str(e))
        return
    except Exception as e:
        messagebox.showwarning("Advertencia", f"Error en los números de página: {e}")
        return

    if not pdf_path or not banco:
        messagebox.showwarning("Advertencia", "Debe seleccionar un archivo y un banco.")
        return

    # Páginas a procesar
    pages = list(range(pagina_inicio, pagina_fin + 1))

    # Aquí llamamos a la función que procesa el PDF
    df = extract_data_from_pdf(pdf_path, pages)

    # Rellenar los NaN de la columna 'Crédito' con 0
    df['Crédito'] = df['Crédito'].fillna(0)
    df['Crédito'] = df['Crédito'].astype(float) # Convertir 'Crédito' a float

    # Cargar códigos de banco
    codigos_banco = get_codigos_bancos().get(banco, {"no_similares": [], "similares": []})

    # Función para limpiar códigos de banco y conceptos
    def limpiar_texto(texto):
        # Eliminar puntos y espacios en blanco adicionales
        return re.sub(r'\.| ', '', texto.lower())

    # DataFrame para las hojas del Excel
    df_hojas = {}

    # Iterar sobre los códigos no similares
    for codigo in codigos_banco["no_similares"]:
        codigo_limpiado = limpiar_texto(codigo)
        mask = df['Concepto'].str.lower().apply(limpiar_texto).str.contains(codigo_limpiado)
        df_filtered = df[mask]
        nombre_hoja = f"Mayor {codigo}"
        df_hojas[nombre_hoja] = pd.concat([df_hojas.get(nombre_hoja, pd.DataFrame(columns=df.columns)), df_filtered],
                                          ignore_index=True)
        df = df[~mask]  # Eliminar las filas filtradas del DataFrame original

        # Iterar sobre los conjuntos de códigos similares
    for idx, conjunto_similares in enumerate(codigos_banco["similares"], start=1):
        nombre_hoja_similares = f"Mayor " + max(conjunto_similares, key=len)
        for codigo_similar in conjunto_similares:
            df_hojas[nombre_hoja_similares] = pd.concat(
                [df_hojas.get(nombre_hoja_similares, pd.DataFrame(columns=df.columns)),
                 df[df['Concepto'].str.lower().apply(limpiar_texto).str.contains(
                     limpiar_texto(codigo_similar.lower()))]], ignore_index=True)
            df = df[
                ~df['Concepto'].str.lower().apply(limpiar_texto).str.contains(limpiar_texto(codigo_similar.lower()))]

    df_emitidos_matched = pd.DataFrame()  # DataFrame para guardar las coincidencias

    if df_emitidos_global is not None:
        df_emitidos = df_emitidos_global

        # Ajustar valores en df['Crédito'] a dos decimales y coma
        df['Crédito'] = df['Crédito'].apply(lambda x: f"{x:.2f}".replace('.', ','))

        # Buscar coincidencias y guardar en la hoja 'Coincidencias'
        for index, row in df_emitidos.iterrows():
            credit_value = row['Imp. Total']
            if credit_value in df['Crédito'].values:
                df_temp = df[df['Crédito'] == credit_value]  # Almacenar las filas coincidentes
                df_emitidos_matched = pd.concat([df_emitidos_matched, df_temp])
                df = df[df['Crédito'] != credit_value]  # Eliminar las filas coincidentes de df

    # Seleccionar la ruta y el nombre del archivo Excel
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        try:
            # Guardar los DataFrames en hojas separadas del archivo Excel
            with pd.ExcelWriter(output_file) as writer:
                # Guardar Extracto original
                df_original = extract_data_from_pdf(pdf_path, pages)
                df_original.to_excel(writer, sheet_name='Extracto Original', index=False)

                # Guardar Extracto modificado con 'Crédito' como valores numéricos
                df['Crédito'] = df['Crédito'].str.replace(',', '.').astype(float).round(2)
                # Guardar Extracto modificado
                df.to_excel(writer, sheet_name='Extracto modif', index=False)

                if not df_emitidos_matched.empty:
                    # Eliminar columna 'Saldo' si existe en Coincidencias Emitidos
                    if 'Saldo' in df_emitidos_matched.columns:
                        df_emitidos_matched.drop(columns=['Saldo'], inplace=True)
                    # Guardar Extracto modificado con 'Crédito' como valores numéricos
                    df_emitidos_matched['Crédito'] = df_emitidos_matched['Crédito'].str.replace(',', '.').astype(float).round(2)
                    df_emitidos_matched.to_excel(writer, sheet_name='Coincidencias Emitidos', index=False)

                # Guardar Hojas adicionales por cada hoja de DataFrame
                for nombre_hoja, df_hoja in df_hojas.items():
                    # Eliminar columna 'Saldo' si existe
                    if 'Saldo' in df_hoja.columns:
                        df_hoja.drop(columns=['Saldo'], inplace=True)

                    # Calcular suma de 'Debito' y 'Credito' y agregar fila 'TOTAL'
                    suma_debito = df_hoja['Débito'].sum()
                    suma_credito = df_hoja['Crédito'].sum()
                    total_row = pd.DataFrame(
                        {'Concepto': ['TOTAL'], 'Débito': [suma_debito], 'Crédito': [suma_credito]})
                    df_hoja = pd.concat([df_hoja, total_row], ignore_index=True)
                    df_hoja.to_excel(writer, sheet_name=f'{nombre_hoja}', index=False)

                messagebox.showinfo("Éxito", f"Archivo Excel creado con éxito: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")


# Función para cargar archivo comprobantes emitidos CSV o Excel
def cargar_emitidos():
    global df_emitidos_global
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"),
                                                      ("Excel files", "*.xlsx;*.xls;*.xlsm")])
    if file_path:
        comprobante_emitidos_entry.delete(0, tk.END)  # Limpiar entrada anterior si hay
        comprobante_emitidos_entry.insert(0, file_path)
        try:
            df_emitidos = pd.read_csv(file_path, sep=";")
            df_emitidos_global = df_emitidos
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")


# Función para cargar archivo comprobantes recibidos CSV o Excel
def cargar_recibidos():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"),
                                                      ("Excel files", "*.xlsx;*.xls;*.xlsm")])
    if file_path:
        comprobante_recibidos_entry.delete(0, tk.END)  # Limpiar entrada anterior si hay
        comprobante_recibidos_entry.insert(0, file_path)
        # Aquí puedes procesar el archivo CSV o Excel según sea necesario
        try:
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
                # Procesar el DataFrame de CSV (ejemplo)
                messagebox.showinfo("Éxito", f"Archivo CSV cargado correctamente")
            elif file_path.endswith((".xlsx", ".xls", ".xlsm")):
                df = pd.read_excel(file_path)
                # Procesar el DataFrame de Excel (ejemplo)
                messagebox.showinfo("Éxito", f"Archivo Excel cargado correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")


# Lista de bancos (puedes agregar más bancos aquí)
bancos = get_bancos()

# Crear ventana principal
root = tk.Tk()
root.title("CONVERTIDOR DE EXTRACTOS")

# Aplicar el tema sv_ttk
sv_ttk.set_theme("dark")  # Puedes cambiar a "light" si prefieres un tema claro

# Variable para el archivo seleccionado
entry_var = tk.StringVar()
max_pages_var = tk.IntVar()
info_label_var = tk.StringVar()

# Título centrado
title_label = tk.Label(root, text="CONVERTIDOR DE EXTRACTOS", font=("Helvetica", 18, "bold"))
title_label.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

# Campo de texto y botón para cargar archivo
tk.Label(root, text="Cargar extracto:", font=("Helvetica", 12)).grid(row=1, column=0, padx=10, pady=10)
entry = tk.Entry(root, textvariable=entry_var, width=50, font=("Helvetica", 12))
entry.grid(row=1, column=1, padx=10, pady=10)
cargar_button = tk.Button(root, text="Cargar extracto", command=cargar_archivo, font=("Helvetica", 12))
cargar_button.grid(row=1, column=2, padx=10, pady=10)

# Etiqueta para mostrar la información del PDF
info_label = tk.Label(root, textvariable=info_label_var, font=("Helvetica", 12))
info_label.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

# Desplegable para seleccionar banco
tk.Label(root, text="Seleccionar banco:", font=("Helvetica", 12)).grid(row=3, column=0, padx=10, pady=10)
banco_var = tk.StringVar()
banco_combobox = ttk.Combobox(root, textvariable=banco_var, values=bancos, state="readonly", font=("Helvetica", 12))
banco_combobox.grid(row=3, column=1, padx=10, pady=10)

# Campo de texto para la página de inicio
tk.Label(root, text="Convertir desde página:", font=("Helvetica", 12)).grid(row=4, column=0, padx=10, pady=10)
start_page_var = tk.StringVar()
start_page_entry = tk.Entry(root, textvariable=start_page_var, width=10, font=("Helvetica", 12))
start_page_entry.grid(row=4, column=1, padx=10, pady=10)

# Campo de texto para la página final
tk.Label(root, text="Convertir hasta página:", font=("Helvetica", 12)).grid(row=5, column=0, padx=10, pady=10)
end_page_var = tk.StringVar()
end_page_entry = tk.Entry(root, textvariable=end_page_var, width=10, font=("Helvetica", 12))
end_page_entry.grid(row=5, column=1, padx=10, pady=10)

# Campo para cargar comprobantes emitidos
tk.Label(root, text="Cargar comprobantes emitidos:", font=("Helvetica", 12)).grid(row=6, column=0, padx=10, pady=10)
comprobante_emitidos_entry = tk.Entry(root, width=50, font=("Helvetica", 12))
comprobante_emitidos_entry.grid(row=6, column=1, padx=10, pady=10)
cargar_comprobante_emitidos_button = tk.Button(root, text="Cargar emitidos", command=cargar_emitidos,
                                               font=("Helvetica", 12))
cargar_comprobante_emitidos_button.grid(row=6, column=2, padx=10, pady=10)

# Campo para cargar comprobantes recibidos
tk.Label(root, text="Cargar comprobantes recibidos:", font=("Helvetica", 12)).grid(row=7, column=0, padx=10, pady=10)
comprobante_recibidos_entry = tk.Entry(root, width=50, font=("Helvetica", 12))
comprobante_recibidos_entry.grid(row=7, column=1, padx=10, pady=10)
cargar_comprobante_recibidos_button = tk.Button(root, text="Cargar recibidos", command=cargar_recibidos,
                                                font=("Helvetica", 12))
cargar_comprobante_recibidos_button.grid(row=7, column=2, padx=10, pady=10)

# Botón para convertir extracto
convertir_button = tk.Button(root, text="Convertir extracto", command=convertir_extracto, font=("Helvetica", 12))
convertir_button.grid(row=8, column=0, columnspan=3, pady=20)

# Aplicar colores y efectos a los botones
button_bg = "#1E90FF"  # Dodger Blue
button_fg = "#ffffff"


def on_enter(e, button):
    button['background'] = button_fg
    button['foreground'] = button_bg


def on_leave(e, button):
    button['background'] = button_bg
    button['foreground'] = button_fg


for button in [cargar_button, convertir_button, cargar_comprobante_emitidos_button,
               cargar_comprobante_recibidos_button]:
    button.configure(bg=button_bg, fg=button_fg, activebackground=button_fg, activeforeground=button_bg)
    button.bind("<Enter>", lambda e, b=button: on_enter(e, b))
    button.bind("<Leave>", lambda e, b=button: on_leave(e, b))

# Iniciar el bucle principal de la aplicación
root.mainloop()
