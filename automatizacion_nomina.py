import pandas as pd
import numpy as np
import configparser # Para leer el archivo .ini
import sys # Para salir del script en caso de errores críticos

# --- CONFIGURACIÓN ---
# Inicializa el parser de configuración
config = configparser.ConfigParser()

# Intenta leer el archivo de configuración
try:
    config.read('config.ini')
except Exception as e:
    print(f"Error crítico: No se pudo leer el archivo de configuración 'config.ini'. Asegúrate de que esté en la misma carpeta que el script. Detalle: {e}")
    sys.exit(1) # Sale del script con un código de error

# Obtiene los valores de configuración. Maneja errores si faltan secciones u opciones.
try:
    ruta_excel_entrada = config.get('Archivos', 'ruta_excel')
    nombre_archivo_salida = config.get('Archivos', 'nombre_salida')

    nombre_hoja_base = config.get('Hojas', 'hoja_base')
    meses_nomina_str = config.get('Hojas', 'meses_nomina')
    # Convierte la cadena de meses a una lista de cadenas, eliminando espacios extra
    meses_nomina = [m.strip() for m in meses_nomina_str.split(',')]
except configparser.Error as e:
    print(f"Error crítico en el archivo 'config.ini'. Revisa la sintaxis y los nombres de secciones/opciones. Detalle: {e}")
    sys.exit(1)

print(f"Configuración cargada:")
print(f"  Archivo Excel de entrada: {ruta_excel_entrada}")
print(f"  Archivo de salida: {nombre_archivo_salida}")
print(f"  Hoja base de empleados: {nombre_hoja_base}")
print(f"  Meses de nómina a procesar: {', '.join(meses_nomina)}")
print("-" * 50)

# --- CARGA DE DATOS ---
try:
    excel_data = pd.ExcelFile(ruta_excel_entrada)
    print("Archivo Excel de entrada cargado correctamente.")
except FileNotFoundError:
    print(f"Error: El archivo Excel no se encuentra en la ruta especificada en 'config.ini': {ruta_excel_entrada}")
    sys.exit(1) # Sale del script si el archivo principal no se encuentra
except Exception as e:
    print(f"Error inesperado al intentar abrir el archivo Excel: {e}")
    sys.exit(1)

# Lista para almacenar los DataFrames de los meses
df_list = []
for mes in meses_nomina:
    try:
        df_list.append(pd.read_excel(excel_data, sheet_name=mes))
        print(f"  Hoja '{mes}' cargada.")
    except Exception as e:
        print(f"  Advertencia: No se pudo cargar la hoja '{mes}'. Se ignorará esta hoja. Detalle: {e}")
        # Continúa procesando las otras hojas aunque una falle

# Verifica si se cargó al menos una hoja de nómina
if not df_list:
    print("Error: No se pudo cargar ninguna hoja de nómina mensual. Revisa los nombres de las hojas en 'config.ini' y el archivo Excel.")
    sys.exit(1)

# Concatena todos los DataFrames mensuales en uno solo
df_principal = pd.concat(df_list, ignore_index=True)
print("Datos mensuales unificados.")

# --- LIMPIEZA Y TRANSFORMACIÓN DEL DATAFRAME PRINCIPAL ---
# Limpiar nombres de columnas: eliminar espacios al inicio/final y reemplazar espacios por guiones bajos
df_principal.columns = df_principal.columns.str.strip().str.replace(' ', '_')
print("Nombres de columnas del DataFrame principal limpiados.")

# Convertir tipos de datos a numérico y datetime. 'errors='coerce'' convierte inválidos a NaN/NaT.
df_principal['Sueldo_Base'] = pd.to_numeric(df_principal['Sueldo_Base'], errors='coerce')
df_principal['Bono_%'] = pd.to_numeric(df_principal['Bono_%'], errors='coerce')
df_principal['Mes'] = pd.to_datetime(df_principal['Mes'], errors='coerce')
print("Tipos de datos convertidos (Sueldo_Base, Bono_%, Mes).")

# Eliminar filas duplicadas para asegurar la unicidad de los registros
initial_rows = len(df_principal)
df_principal.drop_duplicates(inplace=True)
if len(df_principal) < initial_rows:
    print(f"Se eliminaron {initial_rows - len(df_principal)} filas duplicadas.")
else:
    print("No se encontraron filas duplicadas.")

# --- CÁLCULOS ADICIONALES PARA EL DATAFRAME PRINCIPAL ---
df_principal['Bono_Calculado'] = df_principal['Sueldo_Base'] * df_principal['Bono_%']
df_principal['Compensación_Total'] = df_principal['Sueldo_Base'] + df_principal['Bono_Calculado']
print("Bono Calculado y Compensación Total calculados.")

# --- CARGA Y PREPARACIÓN DE DATOS DE LA HOJA 'BASE' ---
try:
    df_base_empleados = pd.read_excel(excel_data, sheet_name=nombre_hoja_base)
    print(f"Hoja '{nombre_hoja_base}' cargada.")
except Exception as e:
    print(f"Error crítico: No se pudo cargar la hoja de empleados '{nombre_hoja_base}'. Detalle: {e}")
    sys.exit(1)

# Limpiar nombres de columnas de la base de empleados
df_base_empleados.columns = df_base_empleados.columns.str.strip().str.replace(' ', '_')
print("Nombres de columnas de la base de empleados limpiados.")

# Convertir la columna 'Fecha_de_Ingreso' a datetime
df_base_empleados['Fecha_de_Ingreso'] = pd.to_datetime(df_base_empleados['Fecha_de_Ingreso'], errors='coerce')
print("Tipo de dato 'Fecha_de_Ingreso' convertido.")

# --- UNIR DATAFRAMES Y CALCULAR ANTIGÜEDAD ---
# Realiza la unión (merge) de los DataFrames
# Asegúrate que 'ID_empeado' es el nombre correcto en ambas hojas después de la limpieza
df_final = df_principal.merge(df_base_empleados[['ID_empeado', 'Fecha_de_Ingreso']], on='ID_empeado', how='left')
print("Datos de fecha de ingreso unidos al DataFrame principal.")

# Función para calcular la diferencia de meses de forma precisa
def calculate_months_diff(end_date, start_date):
    if pd.isna(end_date) or pd.isna(start_date):
        return np.nan # Retorna NaN si alguna fecha es nula
    # Calcula la diferencia total de meses (años * 12 + meses)
    # Resta 1 si el día del mes final es menor que el día del mes inicial,
    # indicando que el mes completo aún no se ha cumplido.
    return (end_date.year - start_date.year) * 12 + \
           (end_date.month - start_date.month) - \
           (end_date.day < start_date.day)

# Aplica la función para calcular la antigüedad en meses
df_final['Antigüedad_meses'] = df_final.apply(lambda row: calculate_months_diff(row['Mes'], row['Fecha_de_Ingreso']), axis=1)

# Rellena los valores NaN en 'Antigüedad_meses' con 0 y convierte a entero
df_final['Antigüedad_meses'] = df_final['Antigüedad_meses'].fillna(0).astype(int)
print("Antigüedad en meses calculada.")

# --- EXPORTAR RESULTADO FINAL ---
try:
    with pd.ExcelWriter(nombre_archivo_salida, engine='xlsxwriter', date_format='DD-MM-AAAA') as writer:
        df_final.to_excel(writer, index=False)
    print(f"\n¡Éxito! Archivo '{nombre_archivo_salida}' generado correctamente con formato de fecha DD-MM-AAAA.")
except Exception as e:
    print(f"Error al exportar el archivo Excel '{nombre_archivo_salida}'. Detalle: {e}")
    sys.exit(1)

print("\nLa automatización se ha completado. ¡Revisa el archivo de salida!")