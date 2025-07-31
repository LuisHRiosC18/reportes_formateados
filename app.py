import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- Funciones insanas ---

def quitar_numeros(texto):
    """Elimina los dígitos de una cadena de texto."""
    if isinstance(texto, str):
        return re.sub(r'\d+', '', texto)
    return texto

def generate_report(cartera, excel_1, excel_2, proyecciones, ecobro):
    """
    Procesa los cinco DataFrames de entrada para generar el reporte final formateado.
    Esta función contiene la lógica de negocio principal y el formato de Excel.
    """
    # --- Inicio del código de procesamiento del usuario ---

    merge_bases = pd.merge(cartera, excel_1, how='left', left_on='contrato', right_on='contrato')

    # Asegurarse de que las columnas existen antes de seleccionarlas
    required_cols = ['contrato', 'Fecha contrato', 'forma_pago', 'estatus', 'cliente', 'domicilio', 'colonia', 'localidad', 'telefono', 'promotor', 'cobrador', 'Dia de visita semanal']
    for col in required_cols:
        if col not in cartera.columns:
            st.error(f"La columna '{col}' no se encuentra en el archivo 'cartera'. Por favor, verifique el archivo.")
            return None
            
    reporte = cartera[required_cols].copy()

    reporte = pd.merge(reporte, merge_bases[['contrato', 'monto_pago_actual']], on='contrato', how='left')
    reporte['Num Cobrador'] = 1

    days_of_week = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    for day in days_of_week:
        reporte[day] = 0

    orden_dias = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    reporte['Dia de visita semanal'] = pd.Categorical(reporte['Dia de visita semanal'], categories=orden_dias, ordered=True)
    reporte = reporte.sort_values(['Dia de visita semanal', 'contrato'])

    reporte['COBRADOR SIN SEGMENTO'] = reporte['cobrador'].apply(quitar_numeros)
    reporte['proyeccion'] = np.where(reporte['contrato'].isin(proyecciones['Contrato']), 'Proyeccion', 'Sin proyeccion')

    sala_dict = {
        'A0': 'ESPARTANOS', 'B0': 'AGUILAS', 'C0': '0', 'D0': 'SOLES', 'DD': 'SOLES', 'E0': '0',
        'F0': 'ESPARTANOS SLRC', 'FA': 'ESPARTANOS SLRC', 'G0': 'TORRE FUERTE', 'H0': 'LOBOS',
        'I0': 'VICTORIA', 'IN': 'VICTORIA', 'J0': 'DIAMANTE', 'JJ': 'DIAMANTE', 'K0': 'AGUILAS SLRC',
        'L0': 'INNOVA', 'LA': 'INNOVA', 'LB': 'INNOVA', 'LL': 'INNOVA', 'M0': 'LEGIONARIOS SLRC',
        'MA': '0', 'MF': '0', 'N0': 'HALCONES SLRC', 'O0': 'ALFAS', 'OO': 'ALFAS', 'OP': 'ALFAS',
        'P0': 'EMPLEADO', 'Q0': 'EAR OLIMPO', 'R0': 'ELITE', 'S0': 'ALFAS MXL', 'k0': 'AGUILAS SLRC'
    }
    reporte['contrato_prefix_reporte'] = reporte['contrato'].astype(str).str[:2].str.upper()
    reporte['SALA'] = reporte['contrato_prefix_reporte'].map(sala_dict)
    reporte = reporte.drop(columns=['contrato_prefix_reporte'])

    ecobro['Fecha'] = pd.to_datetime(ecobro['Fecha'])
    ecobro['Dia de visita semanal'] = ecobro['Fecha'].dt.day_name()
    day_mapping = {
        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles', 'Thursday': 'Jueves',
        'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
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

    reporte_merged = pd.merge(reporte, ecobro_processed, how='left', left_on='contrato', right_on='Contrato')
    reporte['Resultado'] = reporte_merged['Detalle']
    reporte['Aportacion actual'] = reporte_merged['Monto'].fillna(0.0)
    reporte['Dia Cobro'] = reporte_merged['Dia de visita semanal_y']
    reporte['Aportacion'] = (reporte['Aportacion actual'] > 0).astype(int)
    reporte['Dia Visita Correcto'] = (reporte['Dia de visita semanal'] == reporte['Dia Cobro']).map({True: 'Correcto', False: 'Incorrecto'})

    for index, row in reporte.iterrows():
        if pd.notna(row['Dia Cobro']):
            day = row['Dia Cobro']
            if day in reporte.columns:
                reporte.loc[index, day] = row['Resultado']

    columnas_finales = [
        'contrato', 'Fecha contrato', 'forma_pago', 'estatus', 'SALA', 'cliente', 'domicilio', 'colonia',
        'localidad', 'telefono', 'promotor', 'cobrador', 'Dia de visita semanal', 'monto_pago_actual',
        'Num Cobrador', 'COBRADOR SIN SEGMENTO', 'Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes',
        'Martes', 'Miércoles', 'Dia Visita Correcto', 'proyeccion', 'Resultado', 'Aportacion actual',
        'Aportacion', 'Dia Cobro'
    ]
    reporte = reporte[columnas_finales]

    rename_dict = {
        'contrato': 'CONTRATO', 'Fecha contrato': 'FECHA CONTRATO', 'forma_pago': 'FORMA PAGO',
        'estatus': 'ESTATUS', 'cliente': 'CLIENTE', 'domicilio': 'DOMICILIO', 'colonia': 'COLONIA',
        'localidad': 'LOCALIDAD', 'telefono': 'TELEFONO', 'promotor': 'PROMOTOR', 'cobrador': 'COBRADOR',
        'Aportacion': 'Aporto', 'Aportacion actual': 'Aportacion Actual'
    }
    reporte = reporte.rename(columns=rename_dict)
    
    # --- Fin del código de procesamiento y inicio de la creación del Excel ---

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        reporte.to_excel(writer, sheet_name='Reporte', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Reporte']
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#1F497D',
            'font_color': '#FFFFFF',
            'border': 1
        })
        for col_num, value in enumerate(reporte.columns.values):
            worksheet.write(0, col_num, value, header_format)
            # Auto-ajustar ancho de columna
            column_len = max(reporte[value].astype(str).map(len).max(), len(value))
            worksheet.set_column(col_num, col_num, column_len)

    processed_data = output.getvalue()
    return processed_data

# --- Interfaz de Usuario de Streamlit ---

st.set_page_config(page_title="Generador de Reportes", layout="wide")

st.title("📊 Creación de reportes para la semana")
st.write("Cargue los 5 archivos de Excel necesarios para generar el reporte consolidado.")

if 'reporte_generado' not in st.session_state:
    st.session_state.reporte_generado = None

# Contenedores para la carga de archivos
st.header("Carga de Archivos")
col1, col2 = st.columns(2)

with col1:
    cartera_file = st.file_uploader("1. Cargar Base Cartera (Principal)", type=["xlsx", "xls"])
    excel_1_file = st.file_uploader("2. Cargar Base PABS", type=["xlsx", "xls"])
    excel_2_file = st.file_uploader("3. Cargar Base SIGGO", type=["xlsx", "xls"])

with col2:
    ecobro_file = st.file_uploader("4. Cargar Base Ecobro", type=["xlsx", "xls"])
    proyecciones_file = st.file_uploader("5. Cargar Base Proyecciones", type=["xlsx", "xls"])

# Verificar si todos los archivos han sido cargados
all_files_uploaded = all([cartera_file, excel_1_file, excel_2_file, ecobro_file, proyecciones_file])

if all_files_uploaded:
    if st.button("🚀 Generar Reporte", type="primary"):
        try:
            with st.spinner('Procesando datos y generando reporte... Por favor, espere.'):
                # Leer los archivos cargados en DataFrames
                df_cartera = pd.read_excel(cartera_file, sheet_name=3)
                df_pabs = pd.read_excel(excel_1_file)
                df_siggo = pd.read_excel(excel_2_file)
                df_proyecciones = pd.read_excel(proyecciones_file)
                df_ecobro = pd.read_excel(ecobro_file)

                # Generar el reporte
                reporte_bytes = generate_report(df_cartera, df_pabs, df_siggo, df_proyecciones, df_ecobro)
                
                if reporte_bytes:
                    st.session_state.reporte_generado = reporte_bytes
                    st.success("¡Reporte generado con éxito! Ya puede descargarlo.")
                else:
                    # El error ya se mostró dentro de la función
                    st.session_state.reporte_generado = None

        except Exception as e:
            st.error(f"Ocurrió un error inesperado durante el procesamiento: {e}")
            st.session_state.reporte_generado = None

else:
    st.info("Por favor, cargue los cinco archivos para habilitar la generación del reporte.")

# Botón de descarga
if st.session_state.reporte_generado:
    st.download_button(
        label="📥 Descargar Reporte Formateado",
        data=st.session_state.reporte_generado,
        file_name="reporte_formateado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

