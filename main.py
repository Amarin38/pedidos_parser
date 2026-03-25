import numpy as np
import pandas as pd
import regex as re
from pathlib import Path

OFICINA_PATH = Path("C:\\Users\\repuestos01\\Desktop\\autoparser\\archivos\\")
OFICINA_CSV_PATH = Path("C:\\Users\\repuestos01\\Desktop\\autoparser\\archivos_csv\\")
COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}

def pasar_a_csv() -> None:
    for archivo in OFICINA_PATH.rglob("Ped *.txt"):
        try:
            nombre_salida = OFICINA_CSV_PATH / f"{archivo.stem}.csv"
            
            df = pd.read_fwf(
                    archivo, 
                    encoding='latin1', 
                    skiprows=5
                )
            
            file_size = len(df)
            df = df.drop([0, file_size-1, file_size-2]) # elimino todo, dejo colo la tabla principal

            df = df.rename(columns=COLUMNAS)
            df["Articulo"] = df["Columnas"].str.extract(r'^.{0}(.*?) \s{2,}') # extraigo los articulos
            df["Codigo Proveedor"] = df["Columnas"].str.extract(r'^.{12}(.*?) \s{2,}') # extraigo los codigos proveedor
            df["Descripcion"] = df.apply(extraer_dinamico, axis=1, args=(df, "Codigo Proveedor", 12, 40)) #extraigo la descripcion
            df["Cantidad"] = df.apply(extraer_dinamico, axis=1, args=(df, "Descripcion", 40, 85)) # extraigo la cantidad

            df["Cantidad"] = df["Cantidad"].astype(float)

            df.to_csv(nombre_salida, index=False)
            df.to_csv(OFICINA_CSV_PATH / f"{archivo.stem}.txt", index=False)
            
            print(f"Convertido: {archivo.name} -> {nombre_salida.name}")

        except Exception as e:
            print(f"No se puede procesar el archivo {archivo.name} -> {e}")


def extraer_dinamico(fila, df, nombre_col_base, val_vacio, val_texto):
    texto_linea = str(fila["Columnas"])
    valor_celda = str(fila[nombre_col_base]) if pd.notna(fila[nombre_col_base]) else ""
    largo_celda = len(valor_celda)
    
    if largo_celda == 0:
        texto_final = texto_linea[val_vacio:].strip()
    else:
        texto_final = texto_linea[val_texto:].strip()

    match = re.search(r'^.*?(?=\s{2,})', texto_final)

    if match:
        return match.group(0).strip()
    else:
        return texto_final



def sacar_lista_codigos() -> None:
    df = pd.read_excel("codigos.xlsx")

    df["Codigo"] = df["Codigo"].astype(str)

    df[["Familia", "Articulo"]] = df["Codigo"].str.strip().str.split(".", expand=True) # separo codigos
    df["Familia"] = df["Familia"].astype('Int32')
    df["Articulo"] = df["Articulo"].astype('Int32')
    df = df.drop(columns="Codigo")


    df_flavio = df.loc[df["Deposito"] == "FLAVIO", :]
    df_flavio = df_flavio.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_flavio = df_flavio.reset_index(drop=True)
    df_flavio = df_flavio.loc[:, ~df_flavio.columns.str.contains("^Unnamed")]
    

    df_oficina = df.loc[df["Deposito"] == "OFICINA", :]
    df_oficina = df_oficina.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_oficina = df_oficina.reset_index(drop=True)
    df_oficina = df_oficina.loc[:, ~df_oficina.columns.str.contains("^Unnamed")]
    

    df_ropa = df.loc[df["Familia"] == 130, :]
    df_ropa = df_ropa.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_ropa = df_ropa.reset_index(drop=True)
    df_ropa = df_ropa.loc[:, ~df_ropa.columns.str.contains("^Unnamed")]


    pd.DataFrame(df_flavio).to_excel("codigos_flavio.xlsx")
    pd.DataFrame(df_oficina).to_excel("codigos_oficina.xlsx")
    pd.DataFrame(df_ropa).to_excel("codigos_ropa.xlsx")

if __name__ == '__main__':
    pasar_a_csv()