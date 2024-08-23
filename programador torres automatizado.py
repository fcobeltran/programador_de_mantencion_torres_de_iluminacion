# -*- coding: utf-8 -*-
"""
Created on Fri jun 23 01:00:31 2023

@author: fcobeltran
"""
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime

def generar_programacion(file_path_inspecciones, file_path_base, output_path, fechas_mantencion):
    # Solicitar la fecha de la última inspección
    fecha_ultima_inspeccion_str = simpledialog.askstring("Fecha de Última Inspección", "Ingresa la fecha de la última inspección (YYYY-MM-DD):")
    fecha_ultima_inspeccion = pd.to_datetime(fecha_ultima_inspeccion_str, errors='coerce')

    # Fecha de hoy tomada del sistema
    fecha_hoy = pd.to_datetime(datetime.today().strftime('%Y-%m-%d'))

    # Cargar los datos de la tabla desde las filas correctas (A4:L48)
    df_inspecciones = pd.read_excel(file_path_inspecciones, sheet_name='Hoja1', skiprows=3, usecols="A:L", nrows=44)

    # Renombrar columnas para facilitar el manejo
    df_inspecciones = df_inspecciones.rename(columns={
        'ITEM': 'num_plan',
        'EQUIPOS': 'id_torre',
        'UBICACIÓN': 'ubicacion',
        'HOROMETRO ACTUAL': 'horometro_ultima_inspeccion'
    })

    # Asignar la fecha de la última inspección a todos los registros
    df_inspecciones['fecha_ultima_inspeccion'] = fecha_ultima_inspeccion

    # Convertir los IDs de df_inspecciones al formato correcto
    df_inspecciones['id_torre'] = df_inspecciones['id_torre'].apply(lambda x: f"TIM-{str(int(x.split('-')[-1])).zfill(3)}")

    # Cargar el archivo base (programacion_resultado_final.xlsx) que se generó anteriormente
    df_base = pd.read_excel(file_path_base)

    # Realizar la combinación para actualizar el archivo base con los nuevos datos de inspección
    df_combined = pd.merge(df_base, df_inspecciones[['id_torre', 'horometro_ultima_inspeccion', 'fecha_ultima_inspeccion', 'ubicacion']],
                           on='id_torre', how='left', suffixes=('_base', '_inspeccion'))

    # Asegurarse de que las columnas correctas se están utilizando
    df_combined['horometro_ultima_mantencion'] = df_combined['horometro_ultima_mantencion_inspeccion']
    df_combined['fecha_ultima_inspeccion'] = df_combined['fecha_ultima_inspeccion_inspeccion']
    df_combined['horometro_ultima_inspeccion'] = df_combined['horometro_ultima_inspeccion_inspeccion']

    # Decidir cuál columna de ubicación usar
    df_combined['ubicacion'] = df_combined['ubicacion_inspeccion'].combine_first(df_combined['ubicacion_base'])

    # Convertir el diccionario de fechas de mantención a un DataFrame
    df_fechas_mantencion = pd.DataFrame(list(fechas_mantencion.items()), columns=['id_torre', 'fecha_ultima_mantencion'])
    df_fechas_mantencion['fecha_ultima_mantencion'] = pd.to_datetime(df_fechas_mantencion['fecha_ultima_mantencion'])

    # Combinar con el DataFrame combinado para actualizar la fecha de última mantención
    df_combined = pd.merge(df_combined, df_fechas_mantencion, on='id_torre', how='left')

    # Usar la columna correcta para 'fecha_ultima_mantencion'
    df_combined['fecha_ultima_mantencion'] = df_combined['fecha_ultima_mantencion_y']

    # Asegurarse de que las fechas están en el formato datetime
    df_combined['fecha_ultima_inspeccion'] = pd.to_datetime(df_combined['fecha_ultima_inspeccion'], errors='coerce')
    df_combined['fecha_ultima_mantencion'] = pd.to_datetime(df_combined['fecha_ultima_mantencion'], errors='coerce')

    # Calcular el recorrido diario y la programación sugerida
    df_combined['dias_desde_ultima_mantencion'] = (df_combined['fecha_ultima_inspeccion'] - df_combined['fecha_ultima_mantencion']).dt.days
    df_combined['dias_desde_ultima_mantencion'] = df_combined['dias_desde_ultima_mantencion'].apply(lambda x: max(x, 1))

    df_combined['recorrido_diario'] = (df_combined['horometro_ultima_inspeccion'] - df_combined['horometro_ultima_mantencion']) / df_combined['dias_desde_ultima_mantencion']

    # Reemplazar valores infinitos o NaN por un valor adecuado (por ejemplo, 0 o un valor por defecto)
    df_combined['recorrido_diario'].replace([float('inf'), -float('inf')], 0, inplace=True)
    df_combined['recorrido_diario'].fillna(0, inplace=True)

    df_combined['horas_restantes'] = 250 - (df_combined['horometro_ultima_inspeccion'] - df_combined['horometro_ultima_mantencion'])
    df_combined['dias_hasta_prox_mantencion'] = df_combined['horas_restantes'] / df_combined['recorrido_diario']

    # Manejar posibles divisiones por cero o valores infinitos en 'dias_hasta_prox_mantencion'
    df_combined['dias_hasta_prox_mantencion'].replace([float('inf'), -float('inf')], 0, inplace=True)
    df_combined['dias_hasta_prox_mantencion'].fillna(0, inplace=True)

    df_combined['programacion_sugerida'] = df_combined['fecha_ultima_inspeccion'] + pd.to_timedelta(df_combined['dias_hasta_prox_mantencion'], unit='d')

    # Ajustar fechas pasadas a mañana, usando la fecha del sistema
    df_combined['programacion_sugerida'] = df_combined['programacion_sugerida'].apply(
        lambda x: fecha_hoy + pd.Timedelta(days=1) if (x < fecha_hoy and x != pd.NaT) else x
    )

    # Establecer el estado de todas las torres como "Operativa"
    df_combined['estado'] = "Operativa"

    # Guardar los resultados en un archivo Excel
    df_resultado_final = df_combined[['num_plan', 'id_torre', 'ubicacion', 'estado', 
                                    'horometro_ultima_mantencion', 'fecha_ultima_mantencion',
                                    'horometro_ultima_inspeccion', 'fecha_ultima_inspeccion',
                                    'recorrido_diario', 'programacion_sugerida', 'programacion_sap']]

    df_resultado_final.to_excel(output_path, index=False)

    print("Programación de mantenimiento generada exitosamente.")

def solicitar_fechas_mantencion(df_base):
    fechas_mantencion = {}
    
    for _, row in df_base.iterrows():
        torre = row['id_torre']
        fecha_actual = row['fecha_ultima_mantencion']
        fecha_actual_str = pd.to_datetime(fecha_actual).strftime('%Y-%m-%d')

        cambiar_fecha = messagebox.askyesno("Cambiar Fecha de Mantención",
                                            f"La fecha de última mantención de {torre} es {fecha_actual_str}. ¿Deseas cambiarla?")
        if cambiar_fecha:
            nueva_fecha = simpledialog.askstring("Ingresar Nueva Fecha",
                                                 f"Ingrese la nueva fecha de última mantención para {torre} (YYYY-MM-DD):")
            fechas_mantencion[torre] = nueva_fecha
        else:
            fechas_mantencion[torre] = fecha_actual_str
    
    return fechas_mantencion

def seleccionar_archivo():
    # Crear la ventana principal de Tkinter (no visible)
    root = tk.Tk()
    root.withdraw()

    # Abrir el cuadro de diálogo para seleccionar el archivo de inspección
    file_path_inspecciones = filedialog.askopenfilename(
        title="Selecciona el archivo de inspección",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    if file_path_inspecciones:
        file_path_base = 'C:/Users/CPU/Downloads/programacion_resultado_final.xlsx'
        
        # Solicitar el nombre del archivo de salida
        nombre_salida = simpledialog.askstring("Nombre del Archivo de Salida", "Ingresa el nombre del archivo de salida (sin extensión):")
        
        # Crear el path completo para el archivo de salida
        output_path = f'C:/Users/CPU/Downloads/{nombre_salida}.xlsx'
        
        # Cargar el archivo base para obtener las fechas de última mantención
        df_base = pd.read_excel(file_path_base)

        # Solicitar las fechas de última mantención
        fechas_mantencion = solicitar_fechas_mantencion(df_base)
        
        # Generar la programación de mantenimiento utilizando el archivo base y el nuevo archivo de inspección
        generar_programacion(file_path_inspecciones, file_path_base, output_path, fechas_mantencion)

# Ejecutar la función para seleccionar el archivo y generar la programación
seleccionar_archivo()