import streamlit as st
import pandas as pd
from datetime import datetime

# Lista de nombres de columnas restantes en el orden deseado
columnas_restantes = [
    'CÓDIGO SECTOR GENERAL',
    'CÓDIGO SECTOR ESPECÍFICO',
    'ENSAYO',
    'TÉCNICA',
    'SUSTANCIA, MATERIAL, ELEMENTO O PRODUCTO A ENSAYAR',
    'INTERVALO DE MEDICIÓN',
    'DOCUMENTO NORMATIVO',
    'ACTIVIDAD DE ASEGURAMIENTO',
    'ACTIVIDAD DE ASEGURAMIENTO',
    'ACTIVIDAD DE ASEGURAMIENTO',
    'ACEPTACIÓN DE LA JUSTIFICACIÓN'
]

def procesar_archivos(archivos):
    consolidado_df = pd.DataFrame()
    for archivo in archivos:
        try:
            df = df = pd.read_excel(archivo, sheet_name='2019-05-09')
            version = 0
        except:
            df = pd.read_excel('file.xlsx', sheet_name='JNP')
            version = 1

        # Encuentra la fila que contiene el texto "CÓDIGO SECTOR GENERAL" en la primera columna y empieza a extraer la información a partir de la siguiente fila
        start_row = df[df.iloc[:, 1] == 'CÓDIGO SECTOR GENERAL'].index[0] + 1
        # Encuentra la última fila donde la tercera columna no tiene datos
        end_row = 21  # Inicialmente, establece end_row en 20 filas después del start_row
        # Encuentra la última fila del archivo para extraer la fecha de respuesta
        fecha_respuesta = df[df.iloc[:, 1] == 'Revisión Profesional'].index[0]

        
        # Itera desde la fila 24 hasta el final del DataFrame
        for row in range(21, len(df)):
            if pd.isnull(df.iloc[row, 6]):
                end_row = row  # Actualiza end_row cuando se encuentra una fila con la primera columna vacía
                break
        filtered_df = df.iloc[start_row:end_row, 1:]
        filtered_df.reset_index(drop=True, inplace=True)

        # Extrae la fecha de la columna 3, fila 7
        date_str = str(df.iloc[5, 2])
        date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')  # Convierte la cadena de fecha en un objeto datetime

        # Agrega la fecha extraída como una nueva columna en el DataFrame filtrado
        filtered_df['Fecha'] = date

        # Extrae el año de la fecha y crea una nueva columna en el DataFrame filtrado
        filtered_df['Año'] = date.year

        # Extrae el Codigo de Acreditación y crea una nueva columna en el DataFrame filtrado
        filtered_df['Codigo de Acreditacion']=df.iloc[13, 5]

        # Extrae el nombre del OEC y crea una nueva columna en el DataFrame filtrado
        filtered_df['OEC']=df.iloc[7, 2]

        # Extrae el la fecha de respuesta y crea una nueva columna en el DataFrame filtrado
        if version==1:
            date_str = str(df.iloc[fecha_respuesta, 11])
            filtered_df['Fecha de respuesta']=datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        elif version==0:
            date_str = str(df.iloc[fecha_respuesta, 8])
            date_str = date_str[6:]
            filtered_df['Fecha de respuesta']=datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')

        # Se origaniza el DataFrame para que tenga el orden adecuado
        filtered_df.insert(0, 'Fecha de respuesta', filtered_df.pop('Fecha de respuesta'))
        filtered_df.insert(0, 'Año', filtered_df.pop('Año'))
        filtered_df.insert(0, 'Fecha', filtered_df.pop('Fecha'))
        filtered_df.insert(0, 'Codigo de Acreditacion', filtered_df.pop('Codigo de Acreditacion'))
        filtered_df.insert(0, 'OEC', filtered_df.pop('OEC'))
        

        # Se le da el nombre a las columnas
        filtered_df.columns = ['OEC', 'Codigo de Acreditacion', 'Fecha', 'Año', 'Fecha de respuesta'] + columnas_restantes

        # Agrega los datos filtrados al DataFrame principal
        consolidado_df = pd.concat([consolidado_df, filtered_df], ignore_index=True)
    consolidado_df.to_excel('ConsolidadoStreamlit.xlsx', index=False)        
    return consolidado_df

st.title('Aplicación para consolidar JNP en ONAC')

uploaded_files = st.file_uploader("Selecciona los archivos a cargar: ", accept_multiple_files=True)
#st.write(uploaded_files)

# Botón para procesar los archivos
if st.button('Procesar archivos'):
    st.info('Procesando archivos... Esto puede llevar un tiempo.')
    archivo_resultado = procesar_archivos(uploaded_files)
    st.info('Proceso terminado con exito')

    with open('ConsolidadoStreamlit.xlsx', "rb") as template_file:
        template_byte = template_file.read()
        st.download_button(label='📥 Descargar el Resultado',
                        data=template_byte,
                        file_name="ArchivoFormateado.xlsx",
                        mime='application/octet-stream')
