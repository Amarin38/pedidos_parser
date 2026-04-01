from enum import auto, StrEnum
from pathlib import Path
from strenum import UppercaseStrEnum
    
BASE_DIR = Path(__file__).resolve().parent

CODIGOS_PATH = BASE_DIR / "codigos"
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