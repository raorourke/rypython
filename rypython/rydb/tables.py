import json
from typing import List

import pandas as pd
import numpy as np

from rypython.rex import capture_all


class RyDBColumn:
    def __init__(self, column_name: str, column: pd.Series):
        self.name = column_name
        self.column = column


class RyDBTable:
    def __init__(self, table_name: str, table: pd.DataFrame):
        self.name = table_name
        self.df = table
        self._refresh()

    def __repr__(self):
        return repr(self.df)

    @property
    def shape(self):
        return self.df.shape

    def _refresh(self):
        self.__dict__.update(
            {
                column_name: self.df[column_name]
                for column_name in self.df.columns
            }
        )

    @staticmethod
    def regextract(pattern: str, value: str):
        return json.dumps(
            capture_all(
                pattern,
                value,
                flatten=True
            )
        )

    def add_columns_by_regex(
            self,
            pattern: str,
            source_column_name: str,
            column_order: List[str] = None,
            column_mappings: dict = None,
            boolean_columns: List[str] = None
    ) -> None:
        source_column = getattr(self, source_column_name)
        json_column_values = source_column.apply(
            lambda x: self.regextract(pattern, x)
        ).apply(json.loads).tolist()
        json_columns = pd.DataFrame(json_column_values)
        for json_column in json_columns:
            if column_mapping := column_mappings.get(json_column):
                json_columns[json_column] = json_columns[json_column].apply(
                    lambda x: column_mapping.get(x)
                )
            if json_column in boolean_columns:
                json_columns[json_column] = np.where(
                    json_columns[json_column].notna(),
                    True,
                    False
                )
        if column_order is not None:
            json_columns = json_columns[column_order]
        for json_column in json_columns:
            self.df[json_column] = json_columns[json_column]
        self._refresh()

    def append_right(self, right_df: pd.DataFrame):
        self.df = self.df.join(right_df)
        self._refresh()


