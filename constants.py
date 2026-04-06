from enum import auto, StrEnum
from pathlib import Path
from strenum import UppercaseStrEnum
from typing import Tuple

BASE_DIR = Path(__file__).resolve().parent

DB_PATH: str = f"sqlite:///{BASE_DIR}/db/database.db"
CODIGOS_PATH = BASE_DIR / "assets" / "codigos.xlsx"
COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}

NO_DATA         : str = "Sin datos"
UNNAMED         : str = "^Unnamed"
SEPARATOR       : str = "."
ROPA_FAM        : int = 130


# Page
PAGE_TITLE          : str = "Separar Pedidos"
UPLOAD_TITLE        : str = "Inserta los archivos"
SELECT_BOX_LABEL    : str = "Separar por:"
RECARGAR_LABEL      : str = "Recargar códigos ⟳"
SEPARAR_LABEL       : str = "Separar pedidos"
SEPARAR_SPINNER     : str = "Separando pedidos..."
SUCCESS_FILE        : str = "Archivo generado."
DOWNLOAD_BTTN_LABEL : str = "Descargar pedidos."
PLACEHOLDER         : str = "-----"
CONTAINER_WIDTH     : int = 500
MAIN_COLS           : Tuple[float, ...] = (0.85,1,0.5)


# DF Cols
COL_COLUMNAS        : str = "Columnas"
COL_COD_PROVEEDOR   : str = "Codigo Proveedor"
COL_ARTICULO        : str = "Articulo"
COL_DESCRIPCION     : str = "Descripcion"
COL_CANTIDAD        : str = "Cantidad"
COL_FAMILIA         : str = "Familia"
COL_CODIGO          : str = "Codigo"
COL_CODIGOS         : str = "Codigos"
COL_DEPOSITO        : str ="Deposito"


# REGEX
PED_PEN             = r'Ped Pen'
REGEX_ARTICULO      = r'^.{0}(.*?) \s{2,}'
REGEX_COD_PROV      = r'^.{12}(.*?) \s{2,}'
REGEX_CANTIDAD      = r'(\d+\.?\d*)'
REGEX_PED_PEN       = r'(Ped Pen.*?\d{2}-\d{2}-\d{4})'
REGEX_PED           = r'(Ped.*?\d{2}-\d{2}-\d{4})'
REGEX_EXTRAER_DIN   = r'^.*?(?=\s{2,})'
REGEX_NOTA_PED      = r'NOTA DE PEDIDO\s*:\s*(\d+)'
REGEX_NOTA_PED_PEN  = r'PEDIDO PENDIENTE\s*:\s*(\d+)'
REGEX_FECHA         = r'FECHA\s*:\s*([\d/]+)'
REGEX_PARA          = r'Para:\s*(.*?)(?:\s{2,}|\n|$)'
REGEX_PROVEEDOR     = r'PROVEEDOR\s*:\s*(.*?)(?:\s{2,}|\n|$)'
REGEX_RAZON_SOCIAL  = r'R\.\s*SOCIAL\s*:\s*(.*?)(?:\s{2,}|\n|$)'
REGEX_PED_FECHAS    = r'Pedido de Fecha\s*(.*?)(?:\s{2,}|\n|$)'


# WORK SHEET
WS_BORDER_COLOR : str = "000000"
WS_BORDER_THICK : str = "thick"
WS_BORDER_THIN  : str = "thin"

WS_ENCODING: str = 'latin1' 

WS_TITLE        : str = "PEDIDO"
WS_NOTA_PED     : str = "NOTA DE PEDIDO:"
WS_R_SOCIAL     : str = "R.SOCIAL:"
WS_PED_PEN      : str = "PEDIDO PENDIENTE:"
WS_PED_FECHA    : str = "PEDIDO DE FECHA:"
WS_PROVEEDOR    : str = "PROVEEDOR:"
WS_FECHA        : str = "FECHA:"
WS_PARA         : str = "PARA:"
WS_PEDIDO_PARA  : str = "PEDIDO PARA:"
WS_TIPO_PEDIDO  : str = "TIPO DE PEDIDO:"

WS_A_WIDTH: int = 25
WS_B_WIDTH: int = 45
WS_C_WIDTH: int = 50
WS_D_WIDTH: int = 15


# Enums
class TipoPedidoEnum(UppercaseStrEnum):
    OFICINA = auto()
    FLAVIO = auto()
    ROPA = auto()
    RESTO = auto()


class SepararPorEnum(StrEnum):
    def _generate_next_value_(name, start, count, last_values): # type: ignore
        return name.replace("_", " ").capitalize()
    
    TODO = auto() # type: ignore
    SOLO_OFICINA = auto() # type: ignore
    SOLO_FLAVIO = auto() # type: ignore
    SOLO_ROPA = auto() # type: ignore


class FormatoPedidoEnum(UppercaseStrEnum):
    NORMAL = auto()
    PENDIENTE = auto()