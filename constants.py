import sys
from enum import auto, StrEnum
from pathlib import Path
from strenum import UppercaseStrEnum

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        # Si es un EXE, sacamos la ruta de donde está el archivo .exe físico
        return Path(sys.executable).resolve().parent
    else:
        # Si es el script de Python normal
        return Path(__file__).resolve().parent
    
BASE_DIR = obtener_ruta_base()

OFICINA_PATH = BASE_DIR / "pedidos"
OFICINA_XLSX_PATH = BASE_DIR / "pedidos_separados"
CODIGOS_PATH = BASE_DIR / "codigos"
# print(f"DEBUG: Buscando archivos en -> {BASE_DIR}")

COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}

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