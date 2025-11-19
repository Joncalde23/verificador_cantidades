import pdfplumber
import pandas as pd
import re
from unidecode import unidecode
import os
from openpyxl import load_workbook

pdf_path = "fichas_tecnicas/PROFORMA.pdf"
df_fichas= pd.read_excel('fichas_tecnicas/fichas_tecnicas.xlsx')

# Extraemos texto del PDF
with pdfplumber.open(pdf_path) as pdf:
    texto = ""
    for page in pdf.pages:
        texto += page.extract_text() + "\n"

# Filtramos líneas de ítems
lineas = [l.strip() for l in texto.splitlines() if re.match(r"^\d+\s", l.strip())]

# REGEX tolerante final
regex_tolerante = re.compile(
    r"""
    ^
    (\d+)\s+                                  # 1 ITEM
    ^[A-Z]+-\d+-\d+-\d+$                         # 2 REFERENCIA (toma todo hasta ANTES de la cantidad)
    (\d+)\s+                                  # 3 CANTIDAD
    (?:(\d+)\s+)?                              # 4 NUM EXTRA OPCIONAL
    (.+?)\s+                                   # 5 DESCRIPCION
    ([\d.,]+)\s+                               # 6 VR UNITARIO
    ([\d.,]+)$                                 # 7 VR TOTAL
    """,
    re.VERBOSE
)

# Procesamos cada línea
productos = []
for linea in lineas:
    match = regex_tolerante.match(linea)
    if match:
        productos.append(match.groups())
    else:
        print("⚠️ No coincidió:", linea)

# Convertimos a DataFrame
df_proforma = pd.DataFrame(productos, columns=[
    "item",
    "referencia",
    "cantidad",
    "num_extra",
    "descripcion",
    "vr_unitario",
    "vr_total"
])

# Convertimos valores numéricos
df_proforma["item"] = df_proforma["item"].astype(int)
df_proforma["cantidad"] = df_proforma["cantidad"].astype(int)

df_proforma["vr_unitario"] = df_proforma["vr_unitario"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
df_proforma["vr_total"] = df_proforma["vr_total"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)

print(df_proforma)



# Hago tratamiento a la columnas, las convierto en minúsculas, sin tildes y sin espacios
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
columns_treatment (df_fichas)

# renombro la columna 'referencia' del df_fichas por 'ref'
df_fichas = df_fichas.rename(columns={'referencia': 'ref'}) 
df_fichas['referencia'] = df_fichas['ref'].str[:15]  # Asegurarse de que las referencias coincidan en formato

# Muevo la columna de la posicion 4 a la 2
col = df_fichas.pop('referencia')
df_fichas.insert(0, 'referencia', col)  

'''print(df_proforma)'''

# Hago el merge entre los dos dataframes
'''df_merged = pd.merge(df_proforma[['referencia', 'descripcion']], df_fichas[['referencia', 'link']], on='referencia', how='left')  


# Convierto este dataframe a excel
nombre_archivo = "productos_agotados.xlsx"
df_merged.to_excel(nombre_archivo, index=False)

# --- Ajustar anchos de columna con openpyxl ---
wb = load_workbook(nombre_archivo)
ws = wb.active

# ===============================
# 🚀 Convertir la columna LINK en hipervínculo
# ===============================
for row in range(2, ws.max_row + 1):   # desde la fila 2 (omitir encabezados)
    cell = ws[f"C{row}"]  # Asumiendo que 'link' está en la columna C
    if cell.value and isinstance(cell.value, str):
        url = cell.value.strip()
        cell.hyperlink = url          # asignar hipervínculo
        cell.style = "Hyperlink"      # estilo azul y subrayado
    
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
'''