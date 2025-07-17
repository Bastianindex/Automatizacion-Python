import pandas as pd
import numpy as np
import configparser
import sys

# --- CONFIGURACIÓN ---
config = configparser.ConfigParser()
try:
    config.read('config.ini')
except Exception as e:
    print(f"Error crítico: No se pudo leer el archivo de configuración 'config.ini'. Asegúrate de que esté en la misma carpeta que el script. Detalle: {e}")
    sys.exit(1)

try:
    ruta_excel_entrada = config.get('Archivos', 'ruta_excel')
    nombre_archivo_salida = config.get('Archivos', 'nombre_salida')

    nombre_hoja_base = config.get('Hojas', 'hoja_base')
    meses_nomina_str = config.get('Hojas', 'meses_nomina')
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
    sys.exit(1)
except Exception as e:
    print(f"Error inesperado al intentar abrir el archivo Excel: {e}")
    sys.exit(1)

df_list = []
for mes in meses_nomina:
    try:
        df_list.append(pd.read_excel(excel_data, sheet_name=mes))
        print(f"  Hoja '{mes}' cargada.")
    except Exception as e:
        print(f"  Advertencia: No se pudo cargar la hoja '{mes}'. Se ignorará esta hoja. Detalle: {e}")
        pass

if not df_list:
    print("Error: No se pudo cargar ninguna hoja de nómina mensual. Revisa los nombres de las hojas en 'config.ini' y el archivo Excel.")
    sys.exit(1)

df_principal = pd.concat(df_list, ignore_index=True)
print("Datos mensuales unificados.")

# --- LIMPIEZA Y TRANSFORMACIÓN DEL DATAFRAME PRINCIPAL ---
df_principal.columns = df_principal.columns.str.strip().str.replace(' ', '_')
print("Nombres de columnas del DataFrame principal limpiados.")

df_principal['Sueldo_Base'] = pd.to_numeric(df_principal['Sueldo_Base'], errors='coerce')
df_principal['Bono_%'] = pd.to_numeric(df_principal['Bono_%'], errors='coerce')
df_principal['Mes'] = pd.to_datetime(df_principal['Mes'], errors='coerce')
print("Tipos de datos convertidos (Sueldo_Base, Bono_%, Mes).")

initial_rows = len(df_principal)
df_principal.drop_duplicates(inplace=True)
if len(df_principal) < initial_rows:
    print(f"Se eliminaron {initial_rows - len(df_principal)} filas duplicadas.")
else:
    print("No se encontraron filas duplicadas.")

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

df_base_empleados.columns = df_base_empleados.columns.str.strip().str.replace(' ', '_')
print("Nombres de columnas de la base de empleados limpiados.")

df_base_empleados['Fecha_de_Ingreso'] = pd.to_datetime(df_base_empleados['Fecha_de_Ingreso'], errors='coerce')
print("Tipo de dato 'Fecha_de_Ingreso' convertido.")

# --- UNIR DATAFRAMES Y CALCULAR ANTIGÜEDAD ---
df_final = df_principal.merge(df_base_empleados[['ID_empeado', 'Fecha_de_Ingreso']], on='ID_empeado', how='left')
print("Datos de fecha de ingreso unidos al DataFrame principal.")

def calculate_months_diff(end_date, start_date):
    if pd.isna(end_date) or pd.isna(start_date):
        return np.nan
    return (end_date.year - start_date.year) * 12 + \
           (end_date.month - start_date.month) - \
           (end_date.day < start_date.day)

df_final['Antigüedad_meses'] = df_final.apply(lambda row: calculate_months_diff(row['Mes'], row['Fecha_de_Ingreso']), axis=1)
df_final['Antigüedad_meses'] = df_final['Antigüedad_meses'].fillna(0).astype(int)
print("Antigüedad en meses calculada.")

# --- CÁLCULO DE MÉTRICAS RESUMEN ---
print("\nCalculando métricas resumen...")
total_registros = len(df_final)
empleados_unicos = df_final['ID_empeado'].nunique()

sueldo_base_promedio = df_final['Sueldo_Base'].mean()
bono_porcentual_promedio = df_final['Bono_%'].mean()
bono_calculado_promedio = df_final['Bono_Calculado'].mean()
compensacion_total_promedio = df_final['Compensación_Total'].mean()

antiguedad_promedio = df_final['Antigüedad_meses'].mean()
antiguedad_minima = df_final['Antigüedad_meses'].min()
antiguedad_maxima = df_final['Antigüedad_meses'].max()

# Contar valores nulos en columnas clave
nulos_sueldo_base = df_final['Sueldo_Base'].isnull().sum()
nulos_bono_porcentual = df_final['Bono_%'].isnull().sum()
nulos_mes = df_final['Mes'].isnull().sum()
nulos_fecha_ingreso = df_final['Fecha_de_Ingreso'].isnull().sum()

# Crear un DataFrame para el resumen
resumen_data = {
    'Métrica': [
        'Total de Registros Procesados',
        'Número de Empleados Únicos',
        'Sueldo Base Promedio',
        'Bono Porcentual Promedio',
        'Bono Calculado Promedio',
        'Compensación Total Promedio',
        'Antigüedad Promedio (meses)',
        'Antigüedad Mínima (meses)',
        'Antigüedad Máxima (meses)',
        'Registros con Sueldo Base Nulo',
        'Registros con Bono % Nulo',
        'Registros con Mes Nulo',
        'Registros con Fecha de Ingreso Nula'
    ],
    'Valor': [
        total_registros,
        empleados_unicos,
        f"{sueldo_base_promedio:,.2f}", # Formato de moneda
        f"{bono_porcentual_promedio:.2%}", # Formato de porcentaje
        f"{bono_calculado_promedio:,.2f}",
        f"{compensacion_total_promedio:,.2f}",
        f"{antiguedad_promedio:.2f}",
        antiguedad_minima,
        antiguedad_maxima,
        nulos_sueldo_base,
        nulos_bono_porcentual,
        nulos_mes,
        nulos_fecha_ingreso
    ]
}
df_resumen = pd.DataFrame(resumen_data)
print("Métricas resumen calculadas.")

# --- EXPORTAR RESULTADO FINAL CON MÚLTIPLES HOJAS ---
try:
    with pd.ExcelWriter(nombre_archivo_salida, engine='xlsxwriter', date_format='DD-MM-AAAA') as writer:
        # Escribe el DataFrame principal en la primera hoja
        df_final.to_excel(writer, sheet_name='Datos Procesados', index=False)

        # Escribe el DataFrame de resumen en una segunda hoja
        df_resumen.to_excel(writer, sheet_name='Resumen Ejecutivo', index=False)

        # Opcional: Ajustar el ancho de las columnas en la hoja de resumen para que se vea mejor
        workbook = writer.book
        worksheet_resumen = writer.sheets['Resumen Ejecutivo']
        worksheet_resumen.set_column('A:A', 35) # Ancho para la columna 'Métrica'
        worksheet_resumen.set_column('B:B', 20) # Ancho para la columna 'Valor'

    print(f"\n¡Éxito! Archivo '{nombre_archivo_salida}' generado correctamente con dos hojas: 'Datos Procesados' y 'Resumen Ejecutivo'.")
except Exception as e:
    print(f"Error al exportar el archivo Excel '{nombre_archivo_salida}'. Detalle: {e}")
    sys.exit(1)

print("\nLa automatización se ha completado. ¡Revisa el archivo de salida!")