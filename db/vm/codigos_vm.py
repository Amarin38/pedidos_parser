import pandas as pd
from typing import Sequence, List
from ..models import CodigosModel


class CodigosVM:

    def to_model(self, df: pd.DataFrame) -> List[CodigosModel]:
        return [
            CodigosModel(
                id              = None,
                Descripcion     = row["Descripcion"],
                Deposito        = row["Deposito"],
                Familia         = row["Familia"],
                Articulo        = row["Articulo"],
                Codigos         = row["Codigos"],
            ) for index, row in df.iterrows()
        ]

    def to_df(self, model: Sequence[CodigosModel]) -> pd.DataFrame:
        df = pd.DataFrame(
                [
                    {
                        "id"            : e.id,
                        "Descripcion"   : e.Descripcion,
                        "Deposito"      : e.Deposito,
                        "Familia"       : e.Familia,
                        "Articulo"      : e.Articulo,
                        "Codigos"       : e.Codigos
                    }
                    for e in model
                ] 
            )
        
        if not df.empty:
            return df.astype({"Codigos": str})
        return df