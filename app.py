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

def generate_report(cartera, excel_1, excel_2, proyecciones, ecobro_reporte):
    """
    Procesa los cinco DataFrames de entrada para generar el reporte final formateado.
    Esta función contiene la lógica de negocio principal y el formato de Excel.
    """
    # --- Empieza el skibidi procesamiento---

    # Procesamiento inicial de Ecobro
    ecobro = ecobro_reporte.copy()
    ecobro.columns = ecobro.iloc[0]
    # FIX: Remove duplicate columns that can cause reindexing errors
    ecobro = ecobro.loc[:,~ecobro.columns.duplicated()]
    ecobro = ecobro.iloc[1:].reset_index(drop=True)
    # Renombrar la columna ANTES de intentar usarla
    ecobro = ecobro.rename(columns={'No. de Contrato': 'Contrato'})


    merge_bases = pd.merge(cartera, excel_1, how='left', left_on='contrato', right_on='contrato')

    # Preparación inicial del DataFrame 'reporte'
    required_cols = ['contrato', 'cliente', 'domicilio', 'colonia', 'localidad', 'telefono', 'promotor']
    reporte = cartera[required_cols].copy()
    siggo_cols = ['contrato','forma_pago','estatus','cobrador', 'Dia de visita semanal']
    reporte = pd.merge(reporte, excel_2[siggo_cols], on='contrato', how='left')
    reporte = reporte.rename(columns={'contrato': 'Contrato'})

    orden_dias = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    reporte['Dia de visita semanal'] = pd.Categorical(reporte['Dia de visita semanal'], categories=orden_dias, ordered=True)
    
    # Lógica condicional para la columna 'sala'
    columns_to_merge = ['contrato', 'fecha_contrato', 'monto_pago_actual']
    if 'sala' in merge_bases.columns:
        columns_to_merge.append('sala')
        reporte = pd.merge(reporte, merge_bases[columns_to_merge], how='left', left_on='Contrato', right_on='contrato')
    else:
        reporte = pd.merge(reporte, merge_bases[columns_to_merge], how='left', left_on='Contrato', right_on='contrato')
        st.warning("La columna 'sala' no se encontró en la base de PABS. Se generará manualmente.")
        sala_dict = {
            'A0': 'ESPARTANOS', 'B0': 'AGUILAS', 'C0': '0', 'D0': 'SOLES', 'DD': 'SOLES', 'E0': '0',
            'F0': 'ESPARTANOS SLRC', 'FA': 'ESPARTANOS SLRC', 'G0': 'TORRE FUERTE', 'H0': 'LOBOS',
            'I0': 'VICTORIA', 'IN': 'VICTORIA', 'J0': 'DIAMANTE', 'JJ': 'DIAMANTE', 'K0': 'AGUILAS SLRC',
            'L0': 'INNOVA', 'LA': 'INNOVA', 'LB': 'INNOVA', 'LL': 'INNOVA', 'M0': 'LEGIONARIOS SLRC',
            'MA': '0', 'MF': '0', 'N0': 'HALCONES SLRC', 'O0': 'ALFAS', 'OO': 'ALFAS', 'OP': 'ALFAS',
            'P0': 'EMPLEADO', 'Q0': 'EAR OLIMPO', 'R0': 'ELITE', 'S0': 'ALFAS MXL', 'k0': 'AGUILAS SLRC'
        }
        reporte['contrato_prefix_reporte'] = reporte['Contrato'].astype(str).str[:2].str.upper()
        reporte['sala'] = reporte['contrato_prefix_reporte'].map(sala_dict)
        reporte = reporte.drop(columns=['contrato_prefix_reporte'])

    if 'contrato' in reporte.columns:
        reporte = reporte.drop(columns=['contrato'])

    # Estandarización y preparación de datos
    reporte['Contrato'] = reporte['Contrato'].astype(str).str.upper()
    ecobro['Contrato'] = ecobro['Contrato'].astype(str).str.upper()
    proyecciones['Contrato'] = proyecciones['Contrato'].astype(str).str.upper()

    reporte['Num Cobrador'] = 1
    reporte['COBRADOR SIN SEGMENTO'] = reporte['cobrador'].apply(quitar_numeros)
    reporte['proyeccion'] = np.where(reporte['Contrato'].isin(proyecciones['Contrato']), 'Proyeccion', 'Sin proyeccion')

    # Procesamiento avanzado de Ecobro
    ecobro['Fecha'] = pd.to_datetime(ecobro['Fecha'], errors='coerce')
    ecobro.dropna(subset=['Fecha'], inplace=True)
    # MODIFICACIÓN: Crear columna de Hora
    ecobro['Hora'] = ecobro['Fecha'].dt.time
    ecobro['Dia de visita semanal'] = ecobro['Fecha'].dt.day_name()
    day_mapping = {
        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
        'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
    }
    ecobro['Dia de visita semanal'] = ecobro['Dia de visita semanal'].map(day_mapping)
    ecobro['Monto'] = pd.to_numeric(ecobro['Monto'].astype(str).str.replace('[$,]', '', regex=True), errors='coerce').fillna(0)
    # Ordenar por fecha y hora para obtener el último detalle correctamente
    ecobro = ecobro.sort_values(by='Fecha') 

    # --- INICIO DEL NUEVO ALGORITMO EFICIENTE ---

    # 1. Inicializar columnas
    dias_semana = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    for dia in dias_semana:
        reporte[dia] = ''
    
    # 2. Iterar por día y prioridad para llenar las columnas
    for dia in dias_semana:
        df_aux = ecobro[ecobro['Dia de visita semanal'] == dia].copy()
        # Reset index on the auxiliary dataframe to prevent reindexing errors
        df_aux.reset_index(drop=True, inplace=True)
        
        # Procesar por prioridad
        for prioridad in ['Cobro']:
            contratos_prioritarios = df_aux[df_aux['Detalle'] == prioridad]['Contrato'].unique()
            if len(contratos_prioritarios) > 0:
                # Actualizar el reporte principal
                reporte.loc[reporte['Contrato'].isin(contratos_prioritarios), dia] = prioridad
                # Eliminar los contratos ya procesados del dataframe auxiliar
                df_aux = df_aux[~df_aux['Contrato'].isin(contratos_prioritarios)]
        
        # Procesar los detalles restantes (los que no tienen prioridad)
        if not df_aux.empty:
            # Quedarse solo con la última visita del día para los contratos restantes
            # Como el df_aux hereda el orden de 'ecobro' (ordenado por Fecha), 'keep=last' toma la última hora.
            ultimas_visitas = df_aux.drop_duplicates(subset='Contrato', keep='last')
            # Crear un diccionario para mapear Contrato -> Detalle
            mapa_ultimos_detalles = pd.Series(ultimas_visitas.Detalle.values, index=ultimas_visitas.Contrato).to_dict()
            # Actualizar el reporte usando el mapa
            reporte[dia] = reporte['Contrato'].map(mapa_ultimos_detalles).fillna(reporte[dia])

    # 4. Calcular el Resultado final y las columnas de aportación
    def calcular_resultados_finales(row):
        lista_de_prioridad = ['Cobro', 'No tenía dinero', 'Difirió el pago']
        resultado_final = ''
        ultimo_detalle_encontrado = ''
        
        detalles_de_la_semana = {dia: row.get(dia, '') for dia in dias_semana if row.get(dia, '')}

        for prioridad in lista_de_prioridad:
            if prioridad in detalles_de_la_semana.values():
                resultado_final = prioridad
                break
        
        if not resultado_final and detalles_de_la_semana:
            # Si no hay prioridad, encontrar el último detalle cronológico
            for dia in reversed(dias_semana): # Miércoles -> Jueves
                if dia in detalles_de_la_semana:
                    ultimo_detalle_encontrado = detalles_de_la_semana[dia]
                    break
            resultado_final = ultimo_detalle_encontrado

        aportacion_actual = 0.0
        aporto = 0
        dia_cobro = ''
        if resultado_final == 'Cobro':
            for dia, detalle in detalles_de_la_semana.items():
                if detalle == 'Cobro':
                    dia_cobro = dia
                    cobro_entry = ecobro[(ecobro['Contrato'] == row['Contrato']) & (ecobro['Dia de visita semanal'] == dia_cobro) & (ecobro['Detalle'] == 'Cobro')]
                    if not cobro_entry.empty:
                        aportacion_actual = cobro_entry['Monto'].iloc[-1]
                    break
            aporto = 1 if aportacion_actual > 0 else 0
            
        return pd.Series([resultado_final, aportacion_actual, aporto, dia_cobro])

    reporte[['Resultado', 'Aportacion Actual', 'Aporto', 'Dia Cobro']] = reporte.apply(calcular_resultados_finales, axis=1)
    
    # --- FIN DEL NUEVO ALGORITMO ---

    def verificar_dia_visita(fila):
        dia_programado = fila['Dia de visita semanal']
        if pd.isna(dia_programado): return 'Incorrecto'
        return 'Correcto' if fila.get(str(dia_programado), '') != '' else 'Incorrecto'

    reporte['Dia Visita Correcto'] = reporte.apply(verificar_dia_visita, axis=1)

    df_final = reporte
    columnas_finales = [
        'Contrato', 'fecha_contrato', 'forma_pago', 'estatus', 'sala', 'cliente',
        'domicilio', 'colonia', 'localidad', 'telefono', 'promotor', 'cobrador',
        'Dia de visita semanal', 'monto_pago_actual', 'Num Cobrador','COBRADOR SIN SEGMENTO', 'Jueves',
        'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles', 'Dia Visita Correcto',
        'proyeccion', 'Resultado', 'Aportacion Actual', 'Aporto', 'Dia Cobro'
    ]
    for col in columnas_finales:
        if col not in df_final.columns:
            df_final[col] = ''
            
    reporte_final = df_final[columnas_finales]
    rename_final = {
        'Contrato': 'CONTRATO', 'fecha_contrato': 'FECHA CONTRATO', 'forma_pago': 'FORMA PAGO',
        'estatus': 'ESTATUS', 'sala': 'SALA', 'cliente': 'CLIENTE', 'domicilio': 'DOMICILIO',
        'colonia': 'COLONIA', 'localidad': 'LOCALIDAD', 'telefono': 'TELEFONO',
        'promotor': 'PROMOTOR', 'cobrador': 'COBRADOR'
    }
    reporte_final = reporte_final.rename(columns=rename_final)
    
    reporte_final['FECHA CONTRATO'] = pd.to_datetime(reporte_final['FECHA CONTRATO'], errors='coerce').dt.date
    reporte_final = reporte_final.sort_values(['Dia de visita semanal', 'CONTRATO'], na_position='last')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        reporte_final.to_excel(writer, sheet_name='Reporte', index=False, na_rep='')
        workbook = writer.book
        worksheet = writer.sheets['Reporte']
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#1F497D', 'font_color': '#FFFFFF', 'border': 1
        })
        for col_num, value in enumerate(reporte_final.columns.values):
            worksheet.write(0, col_num, value, header_format)
            try:
                column_len = max(reporte_final[value].astype(str).map(len).max(), len(str(value))) + 2
                worksheet.set_column(col_num, col_num, column_len)
            except (AttributeError, TypeError):
                worksheet.set_column(col_num, col_num, 15)
            
    processed_data = output.getvalue()
    # Devolver también el DataFrame para mostrarlo en la UI
    return processed_data, reporte_final

# --- Interfaz de Usuario de Streamlit ---

st.set_page_config(page_title="Generador de Reportes", layout="wide")
st.title("📊 Creación de reportes para la semana")
st.write("Cargue los 5 archivos de Excel necesarios para generar el reporte consolidado.")

if 'reporte_generado' not in st.session_state:
    st.session_state.reporte_generado = None
if 'df_para_mostrar' not in st.session_state:
    st.session_state.df_para_mostrar = pd.DataFrame()

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

all_files_uploaded = all([cartera_file, excel_1_file, excel_2_file, ecobro_file, proyecciones_file])

if all_files_uploaded:
    if st.button("🚀 Generar Reporte", type="primary"):
        try:
            with st.spinner('Procesando datos y generando reporte... Por favor, espere.'):
                def safe_read_excel(file_uploader):
                    engine = 'openpyxl' if file_uploader.name.endswith('.xlsx') else 'xlrd'
                    return pd.read_excel(file_uploader, engine=engine)

                engine_cartera = 'openpyxl' if cartera_file.name.endswith('.xlsx') else 'xlrd'
                df_cartera = pd.read_excel(cartera_file, sheet_name=3, engine=engine_cartera)
                df_pabs = safe_read_excel(excel_1_file)
                df_siggo = safe_read_excel(excel_2_file)
                df_proyecciones = safe_read_excel(proyecciones_file)
                df_ecobro = safe_read_excel(ecobro_file)

                # Generar el reporte y guardar el DataFrame para mostrar
                reporte_bytes, df_display = generate_report(df_cartera, df_pabs, df_siggo, df_proyecciones, df_ecobro)
                
                if reporte_bytes is not None:
                    st.session_state.reporte_generado = reporte_bytes
                    st.session_state.df_para_mostrar = df_display
                    st.success("¡Reporte generado con éxito! Ya puede descargarlo y ver los resultados a continuación.")
                else:
                    st.session_state.reporte_generado = None
                    st.session_state.df_para_mostrar = pd.DataFrame()
        except Exception as e:
            st.error(f"Ocurrió un error inesperado durante el procesamiento: {e}")
            st.session_state.reporte_generado = None
            st.session_state.df_para_mostrar = pd.DataFrame()
            st.exception(e)
else:
    st.info("Por favor, cargue los cinco archivos para habilitar la generación del reporte.")

# --- Sección de Vista Previa y Conteo ---
if not st.session_state.df_para_mostrar.empty:
    st.markdown("---")
    st.header("Resultados de la Semana")
    
    st.subheader("Conteo de Resultados")
    # Limpiar la columna 'Resultado' para un conteo más limpio
    resultado_counts = st.session_state.df_para_mostrar['Resultado'].replace('', 'Sin Visita').value_counts()
    st.table(resultado_counts)

    st.subheader("Vista Previa del Reporte")
    st.dataframe(st.session_state.df_para_mostrar)

    st.download_button(
        label="📥 Descargar Reporte Formateado",
        data=st.session_state.reporte_generado,
        file_name="reporte_formateado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
