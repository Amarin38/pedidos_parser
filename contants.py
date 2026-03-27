from enum import auto
from pathlib import Path
from strenum import UppercaseStrEnum


BASE_DIR = Path(__file__).resolve().parent

OFICINA_PATH = BASE_DIR / "pedidos"
OFICINA_XLSX_PATH = BASE_DIR / "pedidos_separados"
CODIGOS_PATH = BASE_DIR / "codigos"

COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}

class TipoPedido(UppercaseStrEnum):
    OFICINA = auto()
    FLAVIO = auto()
    ROPA = auto()
    RESTO = auto()

