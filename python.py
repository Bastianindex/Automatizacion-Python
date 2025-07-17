import pandas as pd
import numpy as np # Importar numpy para np.nan

# --- CONFIGURACIÓN ---
# Ruta absoluta del archivo Excel
# Asegúrate de que esta ruta sea correcta en tu sistema
archivo = r"C:\Users\basti\OneDrive\Escritorio\proyecto CENCOSUD\Proyecto de Entrevista Practicante.xlsx"

# --- CARGA DE DATOS ---
try:
    excel_data = pd.ExcelFile(archivo)
except FileNotFoundError:
    print(f"Error: El archivo Excel no se encuentra en la ruta especificada: {archivo}")
    exit() # Termina el script si el archivo no existe

# Cargar hojas de datos mensuales
try:
    enero = pd.read_excel(excel_data, sheet_name='Enero')
    febrero = pd.read_excel(excel_data, sheet_name='Febrero')
    marzo = pd.read_excel(excel_data, sheet_name='Marzo')
except Exception as e:
    print(f"Error al cargar una de las hojas mensuales: {e}")
    exit()

# Unificar datos mensuales
df = pd.concat([enero, febrero, marzo], ignore_index=True)

# --- LIMPIEZA Y TRANSFORMACIÓN DE df PRINCIPAL ---
# Limpieza de nombres de columnas: elimina espacios al inicio/final y reemplaza espacios por guiones bajos
df.columns = df.columns.str.strip().str.replace(' ', '_')

# Conversión de tipos de datos para df
# 'errors='coerce'' convierte los valores no válidos a NaN (Not a Number) o NaT (Not a Time)
df['Sueldo_Base'] = pd.to_numeric(df['Sueldo_Base'], errors='coerce')
df['Bono_%'] = pd.to_numeric(df['Bono_%'], errors='coerce')
# Al convertir 'Mes', lo hacemos a datetime, lo cual es correcto para los cálculos.
df['Mes'] = pd.to_datetime(df['Mes'], errors='coerce')

# Eliminar duplicados para asegurar datos únicos
df = df.drop_duplicates()

# --- CÁLCULOS ADICIONALES PARA df PRINCIPAL ---
# Calcular el bono basado en el Sueldo_Base y el Bono_%
df['Bono_Calculado'] = df['Sueldo_Base'] * df['Bono_%']
# Calcular la compensación total
df['Compensación_Total'] = df['Sueldo_Base'] + df['Bono_Calculado']

# --- CARGA Y PREPARACIÓN DE DATOS DE LA HOJA 'BASE' ---
try:
    base = pd.read_excel(excel_data, sheet_name='Base')
except Exception as e:
    print(f"Error al cargar la hoja 'Base': {e}")
    exit()

# Limpieza de nombres de columnas para 'base'
# Esto convertirá "Fecha de Ingreso" a "Fecha_de_Ingreso" y "ID empeado" a "ID_empeado"
base.columns = base.columns.str.strip().str.replace(' ', '_')

# Conversión de 'Fecha_de_Ingreso' a tipo datetime
base['Fecha_de_Ingreso'] = pd.to_datetime(base['Fecha_de_Ingreso'], errors='coerce')

# --- UNIR DATAFRAMES Y CALCULAR ANTIGÜEDAD ---
# Unir df con la base de empleados por 'ID_empeado' para obtener la fecha de ingreso
# Usa 'how='left'' para mantener todas las filas de df
df = df.merge(base[['ID_empeado', 'Fecha_de_Ingreso']], on='ID_empeado', how='left')

# Función para calcular la diferencia de meses de forma precisa (contando meses completos)
def calculate_months_diff(end_date, start_date):
    if pd.isna(end_date) or pd.isna(start_date):
        return np.nan # Usar np.nan para valores numéricos faltantes
    # Calcula la diferencia en años * 12, más la diferencia en meses
    # Resta 1 si el día del mes final es menor que el día del mes inicial (el mes completo aún no se ha cumplido)
    return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) - (end_date.day < start_date.day)

# Aplicar la función para calcular la antigüedad en meses para cada fila
df['Antigüedad_meses'] = df.apply(lambda row: calculate_months_diff(row['Mes'], row['Fecha_de_Ingreso']), axis=1)

# Asegurarse de que la Antigüedad_meses sea un número entero
df['Antigüedad_meses'] = df['Antigüedad_meses'].fillna(0).astype(int) # Rellena NaN con 0 antes de convertir a int

# --- AJUSTE PARA EXPORTAR: FORMATEAR FECHAS SIN HORA ---
# Convertir la columna 'Mes' a formato de solo fecha (AAAA-MM-DD)
# El .dt.date convierte los Timestamps a objetos date, eliminando la parte de la hora.
# Esto solo afecta la representación en el archivo Excel, no el tipo de dato interno para cálculos.
df['Mes'] = df['Mes'].dt.date
df['Fecha_de_Ingreso'] = df['Fecha_de_Ingreso'].dt.date


# --- EXPORTAR RESULTADO ---
# Exportar el DataFrame final a un nuevo archivo Excel
output_path = "resultado_final.xlsx"
df.to_excel(output_path, index=False)

print(f"Archivo '{output_path}' generado correctamente.")
print("La automatización se ha completado con éxito.")