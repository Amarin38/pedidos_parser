import pandas as pd
import regex as re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OFICINA_PATH = BASE_DIR / "pedidos"
OFICINA_CSV_PATH = BASE_DIR / "pedidos_separados"
CODIGOS_PATH = BASE_DIR / "codigos"
COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}

def limpiar_pedido() -> None:
    archivos = list(OFICINA_PATH.rglob("Ped *.txt"))

    for archivo in archivos:
        try:
            nombre_salida = OFICINA_CSV_PATH / f"TODOS {archivo.stem}.xlsx"
            
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
            df = df.drop(columns="Columnas")

            df.to_excel(nombre_salida, index=False)
            
            print(f"Convertido: {archivo.name} -> {nombre_salida.name}")

        except Exception as e:
            print(f"No se puede procesar el archivo {archivo.name} -> {e}")


def filtrar_pedidos(tipo_filtro: str) -> None:
    df_lista_pertenece = []
    df_lista_no_pertenece = []
    archivos = list(OFICINA_CSV_PATH.rglob("TODOS *.xlsx"))

    try:
        codigos = pd.read_excel(CODIGOS_PATH / f"{tipo_filtro}.xlsx", dtype={"Codigos": str})["Codigos"]
        
        for archivo in archivos:
            df_todos = pd.read_excel(archivo, dtype={"Articulo": str})
            df_todos["Articulo"] = df_todos["Articulo"].astype(str).str.strip()
            
            pertenece = df_todos["Articulo"].isin(codigos)

            df_lista_pertenece.append(df_todos[pertenece])
            df_lista_no_pertenece.append(df_todos[~pertenece])

        df_final = pd.concat(df_lista_pertenece)
        df_resto = pd.concat(df_lista_no_pertenece)

        df_final.to_excel(OFICINA_CSV_PATH / f"PEDIDO {tipo_filtro}.xlsx", index=False)
        df_resto.to_excel(OFICINA_CSV_PATH / "RESTO PEDIDO.xlsx", index=False)

    except FileNotFoundError as e:
        print("No existe el archivo -> ", e.with_traceback)


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


# CODIGOS
def limpiar_codigos() -> None:
    df = pd.read_excel(CODIGOS_PATH / "codigos.xlsx")

    df["Codigo"] = df["Codigo"].astype(str)
    df["Articulo"] = df["Articulo"].astype(str)
    df = df.rename(columns={"Articulo":"Descripcion"})

    df[["Familia", "Articulo"]] = df["Codigo"].str.strip().str.split(".", expand=True)
    df["Familia"] = df["Familia"].astype('Int32')
    df["Articulo"] = df["Articulo"].astype('Int32')
    df = df.drop(columns="Codigo")

    familia = df["Familia"].astype(str).str.zfill(3)
    articulo = df["Articulo"].astype(str).str.zfill(5) 
    df["Codigos"] = familia + "." + articulo

    df.to_excel(CODIGOS_PATH / "codigos.xlsx")


def sacar_lista_codigos() -> None:
    df = pd.read_excel(CODIGOS_PATH / "codigos.xlsx")

    df["Familia"] = df["Familia"].astype('Int32')
    df["Articulo"] = df["Articulo"].astype('Int32')


    df_flavio = df.loc[df["Deposito"] == "FLAVIO", :]
    df_flavio = df_flavio.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_flavio = df_flavio.reset_index(drop=True)
    df_flavio = df_flavio.loc[:, ~df_flavio.columns.str.contains("^Unnamed")]
    
    fam_flavio = df_flavio["Familia"].astype(str).str.zfill(3)
    art_flavio = df_flavio["Articulo"].astype(str).str.zfill(5) 
    df_flavio["Codigos"] = fam_flavio + "." + art_flavio

    df_oficina = df.loc[df["Deposito"] == "OFICINA", :]
    df_oficina = df_oficina.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_oficina = df_oficina.reset_index(drop=True)
    df_oficina = df_oficina.loc[:, ~df_oficina.columns.str.contains("^Unnamed")]
    
    fam_ofi = df_oficina["Familia"].astype(str).str.zfill(3)
    art_ofi = df_oficina["Articulo"].astype(str).str.zfill(5) 
    df_oficina["Codigos"] = fam_ofi + "." + art_ofi

    df_ropa = df.loc[df["Familia"] == 130, :]
    df_ropa = df_ropa.sort_values(by=["Familia", "Articulo"], ascending=[True, True])
    df_ropa = df_ropa.reset_index(drop=True)
    df_ropa = df_ropa.loc[:, ~df_ropa.columns.str.contains("^Unnamed")]

    fam_ropa = df_ropa["Familia"].astype(str).str.zfill(3)
    art_ropa = df_ropa["Articulo"].astype(str).str.zfill(5) 
    df_ropa["Codigos"] = fam_ropa + "." + art_ropa

    pd.DataFrame(df_flavio).to_excel(CODIGOS_PATH / "codigos_flavio.xlsx")
    pd.DataFrame(df_oficina).to_excel(CODIGOS_PATH / "codigos_oficina.xlsx")
    pd.DataFrame(df_ropa).to_excel(CODIGOS_PATH / "codigos_ropa.xlsx")


if __name__ == '__main__':
    # limpiar_codigos()
    # sacar_lista_codigos()
    limpiar_pedido()

    print("""
          Separar por:
          0- Todos
          1- Oficina
          2- Flavio
          3- Ropa
          """)
    

    while True:
        try:
            tipo_filtrado = int(input(">> "))

            match tipo_filtrado:
                case 0:
                    filtrar_pedidos("codigos_oficina")
                    filtrar_pedidos("codigos_flavio")
                    filtrar_pedidos("codigos_ropa")
                    break
                case 1: 
                    filtrar_pedidos("codigos_oficina")
                    break
                case 2: 
                    filtrar_pedidos("codigos_flavio")
                    break
                case 3:
                    filtrar_pedidos("codigos_ropa")
                    break
        except ValueError as e:
            print("Introduce una de las opciones. -> ", e)

  

    

