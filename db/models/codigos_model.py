from sqlalchemy import String
from sqlalchemy.orm import Mapped
from sqlalchemy.orm import mapped_column
from .. import DBBase


class CodigosModel(DBBase):
    __tablename__ = "CODIGOS"

    id: Mapped[int] = mapped_column(primary_key=True)
    Descripcion:    Mapped[str]
    Deposito:       Mapped[str] = mapped_column(String(7))
    Familia:        Mapped[int]
    Articulo:       Mapped[int]
    Codigos:        Mapped[str] = mapped_column(String(9))
    


