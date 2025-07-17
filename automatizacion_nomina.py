import pandas as pd
import numpy as np
import configparser
import sys
import logging
from datetime import datetime
import os # Importar la librería 'os' para manejar rutas y directorios

# --- CONFIGURACIÓN DESDE config.ini (TEMPORAL PARA LOGGING) ---
# Se lee config.ini aquí primero para obtener la carpeta de logs antes de configurar el logger
temp_config = configparser.ConfigParser()
try:
    temp_config.read('config.ini')
    carpeta_logs = temp_config.get('Carpetas', 'carpeta_logs', fallback='Logs') # Leer la carpeta de logs, con 'Logs' como fallback
except Exception:
    carpeta_logs = 'Logs' # Fallback si no se puede leer config.ini o la sección/opción

# --- CONFIGURACIÓN DEL LOGGING ---
# Obtener la ruta del directorio actual del script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Unir la ruta del script con el nombre de la carpeta de logs
ruta_completa_carpeta_logs = os.path.join(script_dir, carpeta_logs)

# Crear la carpeta de logs si no existe
try:
    os.makedirs(ruta_completa_carpeta_logs, exist_ok=True)
except OSError as e:
    # Si falla la creación de la carpeta de logs, se intenta guardar en el directorio del script
    print(f"Advertencia: No se pudo crear la carpeta de logs '{ruta_completa_carpeta_logs}'. Los logs se guardarán en el directorio del script. Detalle: {e}")
    ruta_completa_carpeta_logs = script_dir # Usar el directorio del script como fallback

# Obtener la fecha y hora actual para el nombre del archivo de log
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = os.path.join(ruta_completa_carpeta_logs, f"automatizacion_nomina_{timestamp}.log")

# Configurar el logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename), # Guarda los logs en el archivo dentro de la carpeta
        logging.StreamHandler(sys.stdout) # Muestra los logs en la consola
    ]
)

logging.info("--- INICIO DE LA AUTOMATIZACIÓN DE NÓMINA ---")

# --- CONFIGURACIÓN DESDE config.ini (LECTURA FINAL) ---
config = configparser.ConfigParser()
try:
    config.read('config.ini')
except Exception as e:
    logging.error(f"Error crítico: No se pudo leer el archivo de configuración 'config.ini'. Asegúrate de que esté en la misma carpeta que el script. Detalle: {e}")
    sys.exit(1)

try:
    ruta_excel_entrada = config.get('Archivos', 'ruta_excel')
    nombre_archivo_salida_base = config.get('Archivos', 'nombre_salida')
    
    nombre_hoja_base = config.get('Hojas', 'hoja_base')
    meses_nomina_str = config.get('Hojas', 'meses_nomina')
    meses_nomina = [m.strip() for m in meses_nomina_str.split(',')]

    carpeta_reportes = config.get('Carpetas', 'carpeta_reportes')
    # carpeta_logs ya se leyó arriba para configurar el logger
except configparser.Error as e:
    logging.error(f"Error crítico en el archivo 'config.ini'. Revisa la sintaxis y los nombres de secciones/opciones. Detalle: {e}")
    sys.exit(1)

# --- CREACIÓN DE LA CARPETA DE REPORTES ---
# (Este bloque ya lo tenías, se mantiene igual)
script_dir = os.path.dirname(os.path.abspath(__file__))
ruta_completa_carpeta_reportes = os.path.join(script_dir, carpeta_reportes)

try:
    os.makedirs(ruta_completa_carpeta_reportes, exist_ok=True)
    logging.info(f"Carpeta de reportes '{ruta_completa_carpeta_reportes}' asegurada/creada.")
except OSError as e:
    logging.error(f"Error al crear la carpeta de reportes '{ruta_completa_carpeta_reportes}'. Detalle: {e}")
    sys.exit(1)

# Construir la ruta completa del archivo de salida dentro de la carpeta de reportes
nombre_archivo_salida_completa = os.path.join(ruta_completa_carpeta_reportes, nombre_archivo_salida_base)


logging.info("Configuración cargada:")
logging.info(f"  Archivo Excel de entrada: {ruta_excel_entrada}")
logging.info(f"  Carpeta de reportes: {ruta_completa_carpeta_reportes}")
logging.info(f"  Carpeta de logs: {ruta_completa_carpeta_logs}") # Nuevo mensaje
logging.info(f"  Archivo de salida completo: {nombre_archivo_salida_completa}")
logging.info(f"  Hoja base de empleados: {nombre_hoja_base}")
logging.info(f"  Meses de nómina a procesar: {', '.join(meses_nomina)}")
logging.info("-" * 50)

# --- CARGA DE DATOS ---
try:
    excel_data = pd.ExcelFile(ruta_excel_entrada)
    logging.info("Archivo Excel de entrada cargado correctamente.")
except FileNotFoundError:
    logging.error(f"Error: El archivo Excel no se encuentra en la ruta especificada en 'config.ini': {ruta_excel_entrada}")
    sys.exit(1)
except Exception as e:
    logging.error(f"Error inesperado al intentar abrir el archivo Excel: {e}")
    sys.exit(1)

df_list = []
for mes in meses_nomina:
    try:
        df_list.append(pd.read_excel(excel_data, sheet_name=mes))
        logging.info(f"  Hoja '{mes}' cargada.")
    except Exception as e:
        logging.warning(f"  Advertencia: No se pudo cargar la hoja '{mes}'. Se ignorará esta hoja. Detalle: {e}")
        pass

if not df_list:
    logging.error("Error: No se pudo cargar ninguna hoja de nómina mensual. Revisa los nombres de las hojas en 'config.ini' y el archivo Excel.")
    sys.exit(1)

df_principal = pd.concat(df_list, ignore_index=True)
logging.info("Datos mensuales unificados.")

# --- LIMPIEZA Y TRANSFORMACIÓN DEL DATAFRAME PRINCIPAL ---
df_principal.columns = df_principal.columns.str.strip().str.replace(' ', '_')
logging.info("Nombres de columnas del DataFrame principal limpiados.")

df_principal['Sueldo_Base'] = pd.to_numeric(df_principal['Sueldo_Base'], errors='coerce')
df_principal['Bono_%'] = pd.to_numeric(df_principal['Bono_%'], errors='coerce')
df_principal['Mes'] = pd.to_datetime(df_principal['Mes'], errors='coerce')
logging.info("Tipos de datos convertidos (Sueldo_Base, Bono_%, Mes).")

initial_rows = len(df_principal)
df_principal.drop_duplicates(inplace=True)
if len(df_principal) < initial_rows:
    logging.info(f"Se eliminaron {initial_rows - len(df_principal)} filas duplicadas.")
else:
    logging.info("No se encontraron filas duplicadas.")

df_principal['Bono_Calculado'] = df_principal['Sueldo_Base'] * df_principal['Bono_%']
df_principal['Compensación_Total'] = df_principal['Sueldo_Base'] + df_principal['Bono_Calculado']
logging.info("Bono Calculado y Compensación Total calculados.")

# --- CARGA Y PREPARACIÓN DE DATOS DE LA HOJA 'BASE' ---
try:
    df_base_empleados = pd.read_excel(excel_data, sheet_name=nombre_hoja_base)
    logging.info(f"Hoja '{nombre_hoja_base}' cargada.")
except Exception as e:
    logging.error(f"Error crítico: No se pudo cargar la hoja de empleados '{nombre_hoja_base}'. Detalle: {e}")
    sys.exit(1)

df_base_empleados.columns = df_base_empleados.columns.str.strip().str.replace(' ', '_')
logging.info("Nombres de columnas de la base de empleados limpiados.")

df_base_empleados['Fecha_de_Ingreso'] = pd.to_datetime(df_base_empleados['Fecha_de_Ingreso'], errors='coerce')
logging.info("Tipo de dato 'Fecha_de_Ingreso' convertido.")

# --- UNIR DATAFRAMES Y CALCULAR ANTIGÜEDAD ---
df_final = df_principal.merge(df_base_empleados[['ID_empeado', 'Fecha_de_Ingreso']], on='ID_empeado', how='left')
logging.info("Datos de fecha de ingreso unidos al DataFrame principal.")

def calculate_months_diff(end_date, start_date):
    if pd.isna(end_date) or pd.isna(start_date):
        return np.nan
    return (end_date.year - start_date.year) * 12 + \
           (end_date.month - start_date.month) - \
           (end_date.day < start_date.day)

df_final['Antigüedad_meses'] = df_final.apply(lambda row: calculate_months_diff(row['Mes'], row['Fecha_de_Ingreso']), axis=1)
df_final['Antigüedad_meses'] = df_final['Antigüedad_meses'].fillna(0).astype(int)
logging.info("Antigüedad en meses calculada.")

# --- CÁLCULO DE MÉTRICAS RESUMEN ---
logging.info("Calculando métricas resumen...")
total_registros = len(df_final)
empleados_unicos = df_final['ID_empeado'].nunique()

sueldo_base_promedio = df_final['Sueldo_Base'].mean()
bono_porcentual_promedio = df_final['Bono_%'].mean()
bono_calculado_promedio = df_final['Bono_Calculado'].mean()
compensacion_total_promedio = df_final['Compensación_Total'].mean()

antiguedad_promedio = df_final['Antigüedad_meses'].mean()
antiguedad_minima = df_final['Antigüedad_meses'].min()
antiguedad_maxima = df_final['Antigüedad_meses'].max()

nulos_sueldo_base = df_final['Sueldo_Base'].isnull().sum()
nulos_bono_porcentual = df_final['Bono_%'].isnull().sum()
nulos_mes = df_final['Mes'].isnull().sum()
nulos_fecha_ingreso = df_final['Fecha_de_Ingreso'].isnull().sum()

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
        f"{sueldo_base_promedio:,.2f}",
        f"{bono_porcentual_promedio:.2%}",
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
logging.info("Métricas resumen calculadas.")

# --- EXPORTAR RESULTADO FINAL CON MÚLTIPLES HOJAS ---
try:
    with pd.ExcelWriter(nombre_archivo_salida_completa, engine='xlsxwriter', date_format='DD-MM-AAAA') as writer:
        df_final.to_excel(writer, sheet_name='Datos Procesados', index=False)
        df_resumen.to_excel(writer, sheet_name='Resumen Ejecutivo', index=False)

        workbook = writer.book
        worksheet_resumen = writer.sheets['Resumen Ejecutivo']
        worksheet_resumen.set_column('A:A', 35)
        worksheet_resumen.set_column('B:B', 20)

    logging.info(f"¡Éxito! Archivo '{nombre_archivo_salida_completa}' generado correctamente con dos hojas: 'Datos Procesados' y 'Resumen Ejecutivo'.")
except Exception as e:
    logging.error(f"Error al exportar el archivo Excel '{nombre_archivo_salida_completa}'. Detalle: {e}")
    sys.exit(1)

logging.info("La automatización se ha completado. ¡Revisa el archivo de salida y el log!")
logging.info("--- FIN DE LA AUTOMATIZACIÓN DE NÓMINA ---")