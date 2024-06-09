import streamlit as st
import pandas as pd
from openai import OpenAI
from io import BytesIO
import xlsxwriter
from dotenv import load_dotenv
api_key=['OPENAI_API_KEY']

def convert_df_to_excel(df_resumido_):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })
        
        marcas_unicas = df_resumido_['Marca'].unique()
        
        for marca in marcas_unicas:
            df_marca = df_resumido_[df_resumido_['Marca'] == marca]
            df_marca_filtrado = df_marca[['Fecha de emisión', 'Nombre', 'Total', 'Comision', 'IGV']]
            total = df_marca['Total'].sum() - df_marca['ALQUILER'].iloc[0] - df_marca['Comision'].sum()
            
            resumen_data = {
                'Total Venta': [df_marca['Total'].sum()],
                'Alquiler': [df_marca['ALQUILER'].iloc[0]],
                'Comisión': [df_marca['Comision'].sum()],
                'Monto Total a Depositar': [total]
            }
            df_resumen = pd.DataFrame(resumen_data)
            
            sheet_name = marca if len(marca) <= 31 else marca[:31]
            df_marca_filtrado.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(df_marca_filtrado.columns.values):
                worksheet.write(0, col_num, value, header_format)
                column_len = max(df_marca_filtrado[value].astype(str).str.len().max(), len(value)) + 3
                worksheet.set_column(col_num, col_num, column_len)
            
            df_resumen.to_excel(writer, sheet_name=sheet_name, startrow=len(df_marca_filtrado)+3, index=False)
            for col_num, value in enumerate(df_resumen.columns.values):
                worksheet.write(len(df_marca_filtrado) + 3, col_num, value, header_format)
        
    output.seek(0)
    return output

def match_brand_with_openai_streaming(product_name, brands, api_key=''):
    client = OpenAI()
    brands_list = brands['MARCA'].tolist()
    prompt = f"Identifica la marca a partir de la siguiente descripción del producto: '{product_name}'. Las marcas posibles son: {', '.join(brands_list)}. Basado únicamente en la lista proporcionada, ¿a qué marca pertenece probablemente este producto? Por favor, devuelve solo el nombre de la marca en mayúsculas y sin ningún signo de puntuación. Si no existe la marca retorna OTROS"
    
    try:
        stream = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            stream=False
        )
        return stream.choices[0].message.content
        
    except Exception as e:
        print("Error with OpenAI API:", e)
        return "OTROS"
    
def load_excel_ventas(file):
    if file:
        return pd.read_excel(file,skiprows=2)
    return None

def load_excel_marcas(file):
    if file:
        return pd.read_excel(file)
    return None

st.title('Aplicación para cargar y procesar archivos de Excel')

# Subida de archivos de Excel
file1 = st.file_uploader("Sube el archivo de ventas.", type=['xlsx'])
file2 = st.file_uploader("Sube el listado de marcas.", type=['xlsx'])

# Botón para generar el reporte solo si ambos archivos están presentes
if file1 and file2:
    ventas = load_excel_ventas(file1)
    marcas = load_excel_marcas(file2)
    if st.button('Generar Reporte'):
        ventas_filtrado = ventas[\
                (ventas['Tipo de comprobante'].isin(['Boleta','Factura'])) &\
                (ventas['Estado del documento'].isin(['Emitido']))         &\
                (ventas['Estado'].isin(['Aceptado']))]
        
        ventas_resumido = ventas_filtrado[['Fecha de emisión','Nombre','Total']]

        ventas_resumido['Marca'] = ventas_resumido['Nombre'].apply(lambda x: match_brand_with_openai_streaming(x, marcas))

        df_resumido_ = ventas_resumido.merge(marcas, how='left', left_on='Marca', right_on='MARCA')

        # Rellenando los valores NaN con 0 para comisión y alquiler
        df_resumido_['COMISION'] = df_resumido_['COMISION'].fillna(0)
        df_resumido_['ALQUILER'] = df_resumido_['ALQUILER'].fillna(0)

        # Eliminar la columna 'MARCA' extra si está presente después del merge
        if 'MARCA' in df_resumido_.columns:
            df_resumido_.drop('MARCA', axis=1, inplace=True)
        
        df_resumido_['Comision'] = df_resumido_['COMISION']*df_resumido_['Total']

        df_resumido_['IGV'] = df_resumido_['Total']*0.18

        excel_file = convert_df_to_excel(df_resumido_)
        st.download_button(
            label="Descargar reporte de marcas en Excel",
            data=excel_file,
            file_name="reporte_marcas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.write("Por favor, carga ambos archivos para generar el reporte.")
