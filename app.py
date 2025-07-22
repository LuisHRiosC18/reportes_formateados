import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- Funciones de Procesamiento de Datos ---

def quitar_numeros(texto):
    """Elimina los dígitos de una cadena de texto."""
    if isinstance(texto, str):
        return re.sub(r'\d+', '', texto)
    return texto

def generate_report(excel_1, excel_2, proyecciones, ecobro):
    """
    Procesa los cuatro DataFrames de entrada para generar el reporte final.
    Esta función contiene la lógica de negocio principal proporcionada por el usuario.
    """
    # --- Inicio del código de procesamiento del usuario ---

    excel_1['estatus'] = excel_1['estatus'].str.lower()
    pabs_act = excel_1[excel_1['estatus'] == 'activo']

    excel_2['estatus'] = excel_2['estatus'].str.lower().str.strip()
    sigo_act = excel_2[excel_2['estatus'] == 'activo']

    activos = pd.merge(pabs_act, sigo_act, on='contrato', how='inner')
    reporte = activos[['contrato', 'fecha_contrato', 'forma_pago_x', 'estatus_x', 'estatus_y', 'cliente_x', 'domicilio_x', 'colonia_x', 'localidad_x', 'telefono_x', 'promotor_x', 'cobrador_y', 'Dia de visita semanal', 'monto_pago_actual']].copy()

    # Añadir columnas para cada día de la semana
    days_of_week = ['Jueves', 'Viernes', 'Sabado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    for day in days_of_week:
        reporte[day] = None

    # Ordenar el DataFrame por el día de visita semanal
    orden_dias = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    # Corregir 'Sabado' a 'Sábado' si es necesario en la columna original
    reporte['Dia de visita semanal'] = reporte['Dia de visita semanal'].replace({'Sabado': 'Sábado'})
    reporte['Dia de visita semanal'] = pd.Categorical(reporte['Dia de visita semanal'], categories=orden_dias, ordered=True)
    reporte = reporte.sort_values('Dia de visita semanal')

    # Crear la columna cobrador sin segmento
    reporte['COBRADOR SIN SEGMENTO'] = reporte['cobrador_y'].apply(quitar_numeros)

    # Marcar contratos con proyección
    reporte['Proyeccion'] = np.where(reporte['contrato'].isin(proyecciones['Contrato']), 'Proyeccion', 'Sin proyecciones')

    # Procesar datos de Ecobro
    ecobro['Fecha'] = pd.to_datetime(ecobro['Fecha'])
    ecobro['Dia de visita semanal'] = ecobro['Fecha'].dt.day_name()
    day_mapping = {
        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
        'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
    }
    ecobro['Dia de visita semanal'] = ecobro['Dia de visita semanal'].map(day_mapping)
    ecobro = ecobro.rename(columns={'No. de Contrato': 'Contrato'})

    ecobro_filtered = ecobro.copy()
    ecobro_filtered['is_cobro'] = (ecobro_filtered['Detalle'] == 'Cobro').astype(int)
    ecobro_filtered = ecobro_filtered.sort_values(by=['Contrato', 'is_cobro'], ascending=[True, False])
    ecobro_filtered = ecobro_filtered.drop_duplicates(subset=['Contrato'], keep='first')
    ecobro_filtered['Monto'] = ecobro_filtered['Monto'].astype(str).str.replace('[$,]', '', regex=True)
    ecobro_filtered['Monto'] = pd.to_numeric(ecobro_filtered['Monto'], errors='coerce').fillna(0)

    ecobro_processed = ecobro_filtered[['Contrato', 'Detalle', 'Monto', 'Dia de visita semanal']]

    # Fusionar reporte con datos de ecobro
    reporte_merged = pd.merge(reporte, ecobro_processed, how='left', left_on='contrato', right_on='Contrato')

    reporte['Resultado'] = reporte_merged['Detalle']
    reporte['Aportacion actual'] = reporte_merged['Monto'].fillna(0.0)
    reporte['Aportacion'] = (reporte['Aportacion actual'] > 0).astype(int)
    reporte['Dia Cobro'] = reporte_merged['Dia de visita semanal_y']
    reporte['Dia Visita Correcto'] = (reporte['Dia de visita semanal'] == reporte['Dia Cobro']).map({True: 'Correcto', False: 'Incorrecto'})

    # Llenar los días de la semana con el resultado
    for index, row in reporte.iterrows():
        if pd.notna(row['Dia Cobro']):
            day = row['Dia Cobro']
            if day in reporte.columns:
                reporte.loc[index, day] = row['Resultado']

    # Mapeo de salas
    sala_dict = {
        'A0': 'ESPARTANOS', 'B0': 'AGUILAS', 'C0': '0', 'D0': 'SOLES', 'DD': 'SOLES',
        'E0': '0', 'F0': 'ESPARTANOS SLRC', 'FA': 'ESPARTANOS SLRC', 'G0': 'TORRE FUERTE',
        'H0': 'LOBOS', 'I0': 'VICTORIA', 'IN': 'VICTORIA', 'J0': 'DIAMANTE', 'JJ': 'DIAMANTE',
        'K0': 'AGUILAS SLRC', 'L0': 'INNOVA', 'LA': 'INNOVA', 'LB': 'INNOVA', 'LL': 'INNOVA',
        'M0': 'LEGIONARIOS SLRC', 'MA': '0', 'MF': '0', 'N0': 'HALCONES SLRC', 'O0': 'ALFAS',
        'OO': 'ALFAS', 'OP': 'ALFAS', 'P0': 'EMPLEADO', 'Q0': 'EAR OLIMPO', 'R0': 'ELITE',
        'S0': 'ALFAS MXL', 'k0': 'AGUILAS SLRC'
    }
    reporte['contrato_prefix_reporte'] = reporte['contrato'].astype(str).str[:2].str.upper()
    reporte['SALA'] = reporte['contrato_prefix_reporte'].map(sala_dict)
    reporte = reporte.drop(columns=['contrato_prefix_reporte'])

    # Selección y renombrado final de columnas
    columnas_finales = [
        'contrato', 'fecha_contrato', 'forma_pago_x', 'estatus_x', 'estatus_y', 'SALA',
        'cliente_x', 'domicilio_x', 'colonia_x', 'localidad_x', 'telefono_x',
        'promotor_x', 'cobrador_y', 'Dia de visita semanal', 'monto_pago_actual',
        'Jueves', 'Viernes', 'Sabado', 'Domingo', 'Lunes', 'Martes', 'Miércoles',
        'COBRADOR SIN SEGMENTO', 'Proyeccion', 'Resultado', 'Aportacion actual',
        'Aportacion', 'Dia Cobro', 'Dia Visita Correcto'
    ]
    reporte = reporte[columnas_finales]

    # Renombrar columnas
    rename_dict = {
        'contrato': 'CONTRATO', 'fecha_contrato': 'FECHA CONTRATO', 'forma_pago_x': 'FORMA PAGO',
        'estatus_x': 'ESTATUS PABS', 'estatus_y': 'ESTATUS SIGGO', 'cliente_x': 'CLIENTE',
        'domicilio_x': 'DOMICILIO', 'colonia_x': 'COLONIA', 'localidad_x': 'LOCALIDAD',
        'telefono_x': 'TELEFONO', 'promotor_x': 'PROMOTOR', 'cobrador_y': 'COBRADOR',
        'monto_pago_actual': 'MONTO PAGO ACTUAL'
    }
    reporte = reporte.rename(columns=rename_dict)
    
    # --- Fin del código de procesamiento del usuario ---
    return reporte

def to_excel(df):
    """Convierte un DataFrame a un objeto de bytes en formato Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    processed_data = output.getvalue()
    return processed_data

# --- Interfaz de Usuario de Streamlit ---

st.set_page_config(page_title="Generador de Reportes", layout="wide")

st.title("📄 Generador de Reporte de Análisis de Datos")
st.write("Cargue los 4 archivos de Excel necesarios para generar el reporte consolidado.")

# Inicializar el estado de la sesión para el reporte
if 'reporte_generado' not in st.session_state:
    st.session_state.reporte_generado = None

# Contenedores para la carga de archivos
col1, col2 = st.columns(2)
with col1:
    excel_1_file = st.file_uploader("1. Cargar Base PABS", type=["xlsx", "xls"])
    excel_2_file = st.file_uploader("2. Cargar Base SIGGO", type=["xlsx", "xls"])
with col2:
    ecobro_file = st.file_uploader("3. Cargar Base Ecobro", type=["xlsx", "xls"])
    proyecciones_file = st.file_uploader("4. Cargar Base Proyecciones", type=["xlsx", "xls"])

# Verificar si todos los archivos han sido cargados
all_files_uploaded = all([excel_1_file, excel_2_file, ecobro_file, proyecciones_file])

if all_files_uploaded:
    if st.button("🚀 Generar Reporte", type="primary"):
        try:
            with st.spinner('Procesando datos... Por favor, espere.'):
                # Leer los archivos cargados en DataFrames
                df_pabs = pd.read_excel(excel_1_file)
                df_siggo = pd.read_excel(excel_2_file)
                df_proyecciones = pd.read_excel(proyecciones_file)
                df_ecobro = pd.read_excel(ecobro_file)

                # Generar el reporte
                reporte_final = generate_report(df_pabs, df_siggo, df_proyecciones, df_ecobro)
                
                # Guardar en el estado de la sesión
                st.session_state.reporte_generado = reporte_final
                
                st.success("¡Reporte generado con éxito!")
                st.dataframe(reporte_final.head())

        except Exception as e:
            st.error(f"Ocurrió un error durante el procesamiento: {e}")
            st.session_state.reporte_generado = None # Limpiar en caso de error

else:
    st.info("Por favor, cargue los cuatro archivos para habilitar la generación del reporte.")


# Botón de descarga
if st.session_state.reporte_generado is not None:
    df_xlsx = to_excel(st.session_state.reporte_generado)
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=df_xlsx,
        file_name="reporte_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

