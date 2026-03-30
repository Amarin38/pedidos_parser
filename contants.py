import sys
from enum import auto
from pathlib import Path
from strenum import UppercaseStrEnum

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        # Ruta del .exe
        ruta_exe = Path(sys.executable).resolve()
        
        # SI EL EXE ESTÁ EN 'dist', SUBIMOS UN NIVEL
        if ruta_exe.parent.name == 'dist':
            return ruta_exe.parent.parent
        
        return ruta_exe.parent
    
    return Path(__file__).resolve().parent

BASE_DIR = obtener_ruta_base()

OFICINA_PATH = BASE_DIR / "pedidos"
OFICINA_XLSX_PATH = BASE_DIR / "pedidos_separados"
CODIGOS_PATH = BASE_DIR / "codigos"
# print(f"DEBUG: Buscando archivos en -> {BASE_DIR}")

COLUMNAS = {"Articulo    Cod Prov                     Descripcion                                 Cantidad   Prec Unit     S-Total":"Columnas"}
CLI = """
Separar por:
    0- Salir
    1- Todos
    2- Oficina
    3- Flavio
    4- Ropa
    """
VALUE_ERR_CLI = "Introduce una de las opciones. -> "


class TipoPedido(UppercaseStrEnum):
    OFICINA = auto()
    FLAVIO = auto()
    ROPA = auto()
    RESTO = auto()

class FormatoPedido(UppercaseStrEnum):
    NORMAL = auto()
    PENDIENTE = auto()