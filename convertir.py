import pandas as pd
import os

# Carpeta donde están los archivos Excel (raíz del proyecto)
carpeta_excel = os.path.dirname(os.path.abspath(__file__))

# Solicita solo el nombre del archivo Excel
nombre_archivo = input("Introduce el nombre del archivo Excel (ejemplo.xlsx): ")
excel_path = os.path.join(carpeta_excel, nombre_archivo)

# Lee el archivo Excel
df = pd.read_excel(excel_path)

# Selecciona columnas por posición: D (3), F (5), G (6), O (15)
columnas_posicion = [14, 7, 15, 8, 1]  # Python empieza en 0
df_seleccionado = df.iloc[:, columnas_posicion]

# Renombra las columnas a "a1", "a2", "a3", "a4"
df_seleccionado.columns = ['DNI', 'CONTRATO', 'TIPO DOC', 'KRUK CASE ID', 'BOX']

# Reemplaza valores vacíos por 'n/a'
df_seleccionado = df_seleccionado.fillna('NA')

# Solicita el nombre del archivo CSV de salida
nombre_csv = input("Introduce el nombre para el archivo CSV de salida (ejemplo.csv): ")
csv_path = os.path.join(carpeta_excel, nombre_csv)

# Guarda el DataFrame como CSV con punto y coma y comillas dobles
df_seleccionado.to_csv(csv_path, sep=';', index=False, quoting=1)

print("¡CSV generado correctamente en:", csv_path)