from typing import List
from sqlalchemy.orm import sessionmaker
from sqlalchemy import select
import pandas as pd

from .. import db_engine
from constants import TipoPedidoEnum
from ..models.codigos_model import CodigosModel
from ..vm.codigos_vm import CodigosVM


class CodigosRepository():
    def __init__(self):
        SessionLocal = sessionmaker(bind=db_engine)
        self.session = SessionLocal()

    # Create -------------------------------------------
    def insert_many(self, models: List[CodigosModel]) -> None:
        try:
            self.session.query(CodigosModel).delete()
            self.session.add_all(models)
            self.session.commit()  # <--- ESTO ES LO QUE FALTA
        except Exception as e:
            self.session.rollback() # Si algo falla, vuelve atrás para no romper la DB
            print(f"Error al insertar: {e}")
            raise e


    # Read -------------------------------------------
    def get_df(self):
        model = self.session.scalars(
            select(CodigosModel)
        ).all()

        return CodigosVM().to_df(model) if model else pd.DataFrame()


    def get_by_id(self, _id: int) -> CodigosModel:
        return self.session.scalars(
            select(CodigosModel).where(CodigosModel.id == _id)
        ).first()


    def get_by_deposito(self, deposito: TipoPedidoEnum):
        return self.session.scalars(
            select(CodigosModel).where(CodigosModel.Deposito == deposito)
        ).all()


    def get_by_codigos(self, codigo: str):
        return self.session.scalars(
            select(CodigosModel).where(CodigosModel.Codigos == codigo)
        ).first()


    # Delete -------------------------------------------
    def delete_by_id(self, _id: int) -> None:
        row = self.session.get(CodigosModel, _id)
        if row:
            self.session.delete(row)
    

