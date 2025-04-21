from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple

import pandas as pd
from dotenv import load_dotenv
import os

list_components_bch = ['IGB1 1', 'IGB1 2', 'IGB1', 'IGB2', 'DB1', 'DB2']
list_components_pwu = ['IGD5 U', 'IGD5 V', 'IGD5 W', 'IGU', 'IGV', 'IGW', 'IGX', 'IGY', 'IGZ']
def extract_xlsx():
    # Extraccion de URL del archivo xlsx desde un .env
    load_dotenv('urls_sarmiento.env')
    url_file_mit = os.getenv("URL_MITSUBISHI")
    URL = os.path.join(url_file_mit, "FALLAS VVVF MITSUBISHI 20240129_V1.0.xlsx")

    # Carga de archivo xlsx
    wb = load_workbook(URL)
    ws = wb.active

    try:
        df_test = pd.read_excel(URL, sheet_name=ws.title, header=None)
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")

    start_table = "A1"
    end_table = "X1"

    start_row, start_col = coordinate_to_tuple(start_table)
    _, end_col = coordinate_to_tuple(end_table)

    # Ubico tabla en la hoja activa y obtengo un objeto generator
    rows = ws.iter_rows(
        min_row=start_row,
        max_row=ws.max_row, 
        min_col=start_col, 
        max_col=end_col,
    )

    # Itero sobre generator para cargar valores en una lista de listas
    data = []

    for row in rows:
        data.append([cell.value for cell in row])

    # Descrimino entre encabezado y datos
    headers = data[0]
    rows_data = data[1:]

    # Creo dataframe
    df = pd.DataFrame(rows_data, columns=headers)    

    # Identifico y filtro solo filas validas - en base a num de registro
    df["Num"] = pd.to_numeric(df["N"], errors="coerce")
    df_clean = df.dropna(subset=["Num"])
    df_clean = df_clean.drop("Num", axis=1)

    df_clean = df_clean.reset_index(drop=True)

    return df_clean

def search_serie(df, num_serie):
    df_filter = df.copy()

    df_filter = df_filter[df_filter["Número de serie"] == num_serie]

    return df_filter

def filter_by_type(df, type):
    df_filter = df.copy()

    df_filter = df_filter[df_filter["Unidad en falla"] == type]
    #df_filter = df_filter.drop("Unidad en falla", axis=1)

    if type=='BCH':
        df_filter = df_filter.drop(list_components_pwu, axis=1)
    elif type=='PWU':
        df_filter = df_filter.drop(list_components_bch, axis=1)

    return df_filter

def segment_components(register):
    valores_validos = ['x', 'xp', 'p']
    list_components = []
    components = ""

    if register["Unidad en falla"] == "BCH":
        list_components = list_components_bch
    elif register["Unidad en falla"] == "PWU":
        list_components = list_components_pwu 

    for col in df.columns:
        row = str(register[col]).strip().lower()

        if row in valores_validos and col in list_components:
           components = components + ", "+ f"{col}" 

    components = components.strip(", ")

    if components:
        return components
    else:
        return "Sin registro"

if __name__ == "__main__":
    df = extract_xlsx()
    #df = filter_by_type(df, "BCH")
    #df = search_serie(df, "DA30765")
    df['componentes_reemplazados'] = df.apply(segment_components, axis=1)
    df = df.drop(list_components_pwu + list_components_bch, axis=1)
    print(df)
    print(df.info())