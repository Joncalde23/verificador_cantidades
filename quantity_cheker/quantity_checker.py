import pdfplumber
import pandas as pd
import re
from unidecode import unidecode
import os
from openpyxl import load_workbook

pdf_path = "PROFORMA.pdf"
df_inventario= pd.read_excel('INFORME DE SALDOS TOTALES.xlsx', header=6)

# Extraer todo el texto
with pdfplumber.open(pdf_path) as pdf:
    texto = ""
    for page in pdf.pages:
        texto += page.extract_text() + "\n"

# Filtramos las líneas de ítems (empiezan con un número)
lineas = [l.strip() for l in texto.splitlines() if re.match(r"^\d+\s", l.strip())]

productos = []
for linea in lineas:
    # Expresión regular que busca los campos típicos de tus proformas
    match = re.match(
        r"^(\d+)\s+(\S+)\s+(\d+)\s+(.*)\s+([\d,.]+)\s+([\d,.]+)$", linea
    )
    if match:
        productos.append(match.groups())
    else:
        print("⚠️ No coincidió:", linea)  # para depurar si alguna línea no encaja

# Crear DataFrame y exportar
columnas = ["Item", "Referencia", "Cantidad", "Descripción", "Vr Unitario", "Vr Total"]
df_proforma = pd.DataFrame(productos, columns=columnas)

# Hago tratamiento a la columnas, las convierto wen minúsculas, sin tildes y sin espacios
# Función de tratamiento de las columnas inicial
def columns_treatment (data):
    new_columns = []

    for column in data.columns:
        name_lowered = column.lower()
        name_striped = name_lowered.strip()
        name_separated = name_striped.replace(' ', '_' )
        #Elimine las tildes para manipular mejor los nombres de las columnas
        name_unicode = unidecode(name_separated)
        new_columns.append(name_unicode)
    data.columns = new_columns 

columns_treatment (df_proforma)
columns_treatment (df_inventario)


#Elimino en la columna referencia del excel los vacios que puedan exisitir

df_inventario = df_inventario.dropna(subset=['referencia'])
# Hago el merge entre los dos dataframes
df_merged = pd.merge(df_proforma, df_inventario[['referencia', 'saldo']], on='referencia', how='left')    



# convierto la columna cantidad en entero
df_merged['cantidad'] = df_merged['cantidad'].astype(int)



df_merged['agotados']=df_merged['saldo'] - df_merged['cantidad']
df_filtrado = df_merged.query('agotados < 0')
df_filtrado = df_filtrado[['referencia', 'descripcion', 'agotados']]

# Convierto este dataframe a excel
nombre_archivo = "productos_agotados.xlsx"
df_filtrado.to_excel(nombre_archivo, index=False)

# --- Ajustar anchos de columna con openpyxl ---
wb = load_workbook(nombre_archivo)
ws = wb.active
    
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter  # Letra de la columna (A, B, C...)
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    ajuste = max_length + 2  # un poquito de espacio extra
    ws.column_dimensions[col_letter].width = ajuste
    
wb.save(nombre_archivo)


if os.name == 'nt':  # Windows
    os.startfile(nombre_archivo)
elif os.name == 'posix':  # macOS o Linux
    os.system(f'open "{nombre_archivo}"') 
