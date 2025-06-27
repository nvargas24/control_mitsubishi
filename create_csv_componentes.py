import csv
import os
from dotenv import load_dotenv
import pandas as pd

load_dotenv('urls_sarmiento.env')
url_desktop = os.getenv("URL_DESKTOP")

def export_to_csv(df, name="output_data"):
    """
    Crea archivo csv en escritorio segun df recibido
    """
    output_file = os.path.join(url_desktop, f"{name}.csv")
    try:
        df.to_csv(output_file, index=False, encoding="utf-8-sig", sep=";")
        print(f"Archivo CSV creado: ./{name}.csv")
    except PermissionError:
        print(f"Error: No se pudo guardar el archivo '{name}.csv'. Verifica que no esté abierto en otro programa.")
    except Exception as e:
        print(f"Error inesperado al guardar el archivo: {e}")


# Datos de ejemplo
componentes = [
    {"id_componente": 'IGB1 1', "tipo": "Placa de control", "Detalle": "Sin detalle", "Modulo": "BCH"},
    {"id_componente": 'IGB1 2', "tipo": "Placa de control", "Detalle": "Sin detalle", "Modulo": "BCH"},
    {"id_componente": 'IGB1', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "BCH"},
    {"id_componente": 'IGB2', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "BCH"},
    {"id_componente": 'DB1', "tipo": "Diodo", "Detalle": "CM2400HCB34N", "Modulo": "BCH"},
    {"id_componente": 'DB2', "tipo": "Diodo", "Detalle": "CM2400HCB34N", "Modulo": "BCH"},
    {"id_componente": 'IGD5 U', "tipo": "Placa de control", "Detalle": "Sin detalle", "Modulo": "PW"},
    {"id_componente": 'IGD5 V', "tipo": "Placa de control", "Detalle": "Sin detalle", "Modulo": "PW"},
    {"id_componente": 'IGD5 W', "tipo": "Placa de control", "Detalle": "Sin detalle", "Modulo": "PW"},
    {"id_componente": 'IGU', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
    {"id_componente": 'IGV', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
    {"id_componente": 'IGW', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
    {"id_componente": 'IGX', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
    {"id_componente": 'IGY', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
    {"id_componente": 'IGZ', "tipo": "IGBT", "Detalle": "RM600HS-34S", "Modulo": "PW"},
]


if __name__ == "__main__":
    # Creo df
    df = pd.DataFrame(componentes)

    # Exporta en formato csv
    export_to_csv(df, "componentes_mitsubishi")
 