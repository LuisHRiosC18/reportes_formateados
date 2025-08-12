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
    # Hacer que la primera fila del ecobro sean los nuevos nombres de columnas
    ecobro.columns = ecobro.iloc[0]
    # Eliminar la primera fila (que ya no es necesaria)
    ecobro = ecobro.iloc[1:].reset_index(drop=True)

    merge_bases = pd.merge(cartera, excel_1, how='left', left_on='contrato', right_on='contrato')

    # Asegurarse de que las columnas existen antes de seleccionarlas
    required_cols = ['contrato', 'cliente', 'domicilio', 'colonia', 'localidad', 'telefono', 'promotor']
    for col in required_cols:
        if col not in cartera.columns:
            st.error(f"La columna '{col}' no se encuentra en el archivo 'cartera'. Por favor, verifique el archivo.")
            return None
            
    reporte = cartera[required_cols].copy()
    # Añadir columnas de SIGGO
    siggo_cols = ['contrato','forma_pago','estatus','cobrador', 'Dia de visita semanal']
    
    #Añadimos las columnas de SIGGO a Reporte
    reporte = pd.merge(reporte, excel_2[siggo_cols], on='contrato', how='left')

    reporte = reporte.rename(columns={'contrato': 'Contrato'}) # Renombrar para consistencia

    # Ordenar el DataFrame por el día de visita semanal
    orden_dias = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    reporte['Dia de visita semanal'] = pd.Categorical(reporte['Dia de visita semanal'], categories=orden_dias, ordered=True)
    
    # --- Lógica condicional para la columna 'sala' ---
    columns_to_merge = ['contrato', 'fecha_contrato', 'monto_pago_actual']
    
    # Verificar si 'sala' existe en la base de PABS (excel_1)
    if 'sala' in merge_bases.columns:
        columns_to_merge.append('sala')
        reporte = pd.merge(reporte, merge_bases[columns_to_merge], how='left', left_on='Contrato', right_on='contrato')
    else:
        # Si no existe, se fusiona sin 'sala' y se crea manualmente
        reporte = pd.merge(reporte, merge_bases[columns_to_merge], how='left', left_on='Contrato', right_on='contrato')
        
        st.warning("La columna 'sala' no se encontró en la base de PABS. Se generará manualmente a partir del prefijo del contrato.")
        
        sala_dict = {
            'A0': 'ESPARTANOS', 'B0': 'AGUILAS', 'C0': '0', 'D0': 'SOLES', 'DD': 'SOLES', 'E0': '0',
            'F0': 'ESPARTANOS SLRC', 'FA': 'ESPARTANOS SLRC', 'G0': 'TORRE FUERTE', 'H0': 'LOBOS',
            'I0': 'VICTORIA', 'IN': 'VICTORIA', 'J0': 'DIAMANTE', 'JJ': 'DIAMANTE', 'K0': 'AGUILAS SLRC',
            'L0': 'INNOVA', 'LA': 'INNOVA', 'LB': 'INNOVA', 'LL': 'INNOVA', 'M0': 'LEGIONARIOS SLRC',
            'MA': '0', 'MF': '0', 'N0': 'HALCONES SLRC', 'O0': 'ALFAS', 'OO': 'ALFAS', 'OP': 'ALFAS',
            'P0': 'EMPLEADO', 'Q0': 'EAR OLIMPO', 'R0': 'ELITE', 'S0': 'ALFAS MXL', 'k0': 'AGUILAS SLRC'
        }
        
        # Crear la columna 'sala' usando el mapeo
        reporte['contrato_prefix_reporte'] = reporte['Contrato'].astype(str).str[:2].str.upper()
        reporte['sala'] = reporte['contrato_prefix_reporte'].map(sala_dict)
        reporte = reporte.drop(columns=['contrato_prefix_reporte'])

    if 'contrato' in reporte.columns:
        reporte = reporte.drop(columns=['contrato']) # Limpiar columna de contrato duplicada

    reporte['Num Cobrador'] = 1
    reporte['COBRADOR SIN SEGMENTO'] = reporte['cobrador'].apply(quitar_numeros)
    reporte['proyeccion'] = np.where(reporte['Contrato'].isin(proyecciones['Contrato']), 'Proyeccion', 'Sin proyeccion')

    # Procesamiento avanzado de Ecobro
    ecobro['Fecha'] = pd.to_datetime(ecobro['Fecha'], errors='coerce')
    ecobro.dropna(subset=['Fecha'], inplace=True) # Eliminar filas donde la fecha no es válida
    ecobro['Dia de visita semanal'] = ecobro['Fecha'].dt.day_name()
    day_mapping = {
        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
        'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
    }
    ecobro['Dia de visita semanal'] = ecobro['Dia de visita semanal'].map(day_mapping)
    ecobro = ecobro.rename(columns={'No. de Contrato': 'Contrato'})
    ecobro = ecobro.loc[:,~ecobro.columns.duplicated()]
    ecobro['Monto'] = ecobro['Monto'].astype(str).str.replace('[$,]', '', regex=True)
    ecobro['Monto'] = pd.to_numeric(ecobro['Monto'], errors='coerce').fillna(0)
    
    # --- INICIO DEL NUEVO ALGORITMO ITERATIVO ---
    
    # MODIFICACIÓN: Convertir todos los contratos a mayúsculas para estandarizar
    reporte['Contrato'] = reporte['Contrato'].astype(str).str.upper()
    ecobro['Contrato'] = ecobro['Contrato'].astype(str).str.upper()
    proyecciones['Contrato'] = proyecciones['Contrato'].astype(str).str.upper()


    # 1. Inicializar columnas en el reporte final
    dias_semana = ['Jueves', 'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles']
    for dia in dias_semana:
        reporte[dia] = '' # Llenar con strings vacíos

    # Ordenar ecobro por fecha para obtener fácilmente la última entrada del día
    ecobro = ecobro.sort_values(by='Fecha')

    # 2. Llenar las columnas de los días iterando fila por fila
    for index, row in reporte.iterrows():
        contrato_actual = row['Contrato']
        for dia in dias_semana:
            # Filtrar ecobro para el contrato y día actual
            visitas_del_dia = ecobro[(ecobro['Contrato'] == contrato_actual) & (ecobro['Dia de visita semanal'] == dia)]

            if not visitas_del_dia.empty:
                # Si hubo visitas ese día
                if 'Cobro' in visitas_del_dia['Detalle'].values:
                    # Si hubo un 'Cobro', se le da prioridad
                    reporte.loc[index, dia] = 'Cobro'
                else:
                    # Si no, se toma el último detalle registrado para ese día
                    reporte.loc[index, dia] = visitas_del_dia['Detalle'].iloc[-1]

    # 3. Calcular las columnas de resultado y aportación basadas en los días ya llenos
    def calcular_resultados_finales(row):
        lista_de_prioridad = ['Cobro', 'No tenía dinero', 'Difirió pago']
        resultado_final = ''
        ultimo_detalle_encontrado = ''
        detalles_de_la_semana = {}

        # a. Recolectar todos los detalles de la semana
        for dia in dias_semana:
            detalle_dia = row.get(dia, '')
            if detalle_dia:
                detalles_de_la_semana[dia] = detalle_dia
                ultimo_detalle_encontrado = detalle_dia
        
        # b. Buscar el resultado con la mayor prioridad
        for prioridad in lista_de_prioridad:
            if prioridad in detalles_de_la_semana.values():
                resultado_final = prioridad
                break
        
        if not resultado_final:
            resultado_final = ultimo_detalle_encontrado

        # c. Calcular aportación solo si el resultado es 'Cobro'
        aportacion_actual = 0.0
        aporto = 0
        dia_cobro = ''

        if resultado_final == 'Cobro':
            for dia, detalle in detalles_de_la_semana.items():
                if detalle == 'Cobro':
                    dia_cobro = dia
                    # Buscar el monto correspondiente en el dataframe original de ecobro
                    cobro_entry = ecobro[
                        (ecobro['Contrato'] == row['Contrato']) &
                        (ecobro['Dia de visita semanal'] == dia_cobro) &
                        (ecobro['Detalle'] == 'Cobro')
                    ]
                    if not cobro_entry.empty:
                        aportacion_actual = cobro_entry['Monto'].iloc[-1] # Tomar el último monto si hay varios
                    break
            aporto = 1 if aportacion_actual > 0 else 0
            
        return pd.Series([resultado_final, aportacion_actual, aporto, dia_cobro])

    # Aplicar la función para generar las columnas de resultado
    reporte[['Resultado', 'Aportacion Actual', 'Aporto', 'Dia Cobro']] = reporte.apply(calcular_resultados_finales, axis=1)
    
    # --- FIN DEL NUEVO ALGORITMO ---

    def verificar_dia_visita(fila):
        dia_programado = fila['Dia de visita semanal']
        if pd.isna(dia_programado): 
            return 'Incorrecto'
        # Verificar si la columna del día programado tiene algún contenido
        detalle_ese_dia = fila.get(str(dia_programado), '')
        return 'Correcto' if detalle_ese_dia != '' else 'Incorrecto'

    reporte['Dia Visita Correcto'] = reporte.apply(verificar_dia_visita, axis=1)

    df_final = reporte # Renombrar para consistencia con el resto del script

    columnas_finales = [
        'Contrato', 'fecha_contrato', 'forma_pago', 'estatus', 'sala', 'cliente',
        'domicilio', 'colonia', 'localidad', 'telefono', 'promotor', 'cobrador',
        'Dia de visita semanal', 'monto_pago_actual', 'Num Cobrador','COBRADOR SIN SEGMENTO', 'Jueves',
        'Viernes', 'Sábado', 'Domingo', 'Lunes', 'Martes', 'Miércoles', 'Dia Visita Correcto',
        'proyeccion', 'Resultado', 'Aportacion Actual', 'Aporto', 'Dia Cobro'
    ]
    # Asegurarse de que todas las columnas existan antes de seleccionarlas
    for col in columnas_finales:
        if col not in df_final.columns:
            df_final[col] = '' # o 0 si es numérico
            
    reporte_final = df_final[columnas_finales]

    rename_final = {
        'Contrato': 'CONTRATO', 'fecha_contrato': 'FECHA CONTRATO', 'forma_pago': 'FORMA PAGO',
        'estatus': 'ESTATUS', 'sala': 'SALA', 'cliente': 'CLIENTE', 'domicilio': 'DOMICILIO',
        'colonia': 'COLONIA', 'localidad': 'LOCALIDAD', 'telefono': 'TELEFONO',
        'promotor': 'PROMOTOR', 'cobrador': 'COBRADOR'
    }
    reporte_final = reporte_final.rename(columns=rename_final)
    
    reporte_final['FECHA CONTRATO'] = pd.to_datetime(reporte_final['FECHA CONTRATO'], errors='coerce').dt.date
    # Ordenar al final, manejando los posibles valores nulos
    reporte_final = reporte_final.sort_values(['Dia de visita semanal', 'CONTRATO'], na_position='last')

    # --- Fin del código de procesamiento y inicio de la creación del Excel ---
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
                # Calcular el ancho de la columna
                column_len = max(reporte_final[value].astype(str).map(len).max(), len(str(value))) + 2
                worksheet.set_column(col_num, col_num, column_len)
            except (AttributeError, TypeError):
                worksheet.set_column(col_num, col_num, 15) # Ancho por defecto
            
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
                # Función para leer archivos de Excel de forma segura
                def safe_read_excel(file_uploader):
                    engine = 'openpyxl' if file_uploader.name.endswith('.xlsx') else 'xlrd'
                    return pd.read_excel(file_uploader, engine=engine)

                # Leer los archivos cargados en DataFrames
                # La hoja de Cartera es específica
                engine_cartera = 'openpyxl' if cartera_file.name.endswith('.xlsx') else 'xlrd'
                df_cartera = pd.read_excel(cartera_file, sheet_name=3, engine=engine_cartera)
                
                df_pabs = safe_read_excel(excel_1_file)
                df_siggo = safe_read_excel(excel_2_file)
                df_proyecciones = safe_read_excel(proyecciones_file)
                df_ecobro = safe_read_excel(ecobro_file)

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
            # Para depuración, muestra más detalles del error en la consola
            st.exception(e)

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
