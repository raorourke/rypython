import json
import logging
from dataclasses import dataclass
from typing import List, Union, Dict, Any

import numpy as np
import pandas as pd
from pydantic import BaseModel
from rich.table import Table

from rypython.rex import capture_all
from rypython.ryagram import erDiagram

logging.basicConfig(level=logging.DEBUG)


@dataclass
class DType:
    pandas: str
    sql: str = None


class DataType:
    MAPPING = {
        "string": {
            "tiny": "TINYTEXT",
            "text": "TEXT",
            "med": "MEDIUMTEXT",
            "long": "LONGTEXT"
        },
        "int": {
            "int8": "TINYINT",
            "int16": "SMALLINT",
            "int32": "INT",
            "int64": "BIGINT"
        },
        "float": "FLOAT",
        "date": "DATE",
        "time": "TIME",
        "datetime": "DATETIME"
    }

    def __init__(self, column: pd.Series):
        self.pandas = column.dtype
        self.sql = self.to_sql_dtype(column)

    @staticmethod
    def category_to_sql_dtype(column: pd.Series):
        categories = column.astype("category").dtype.categories.to_list()
        return f"ENUM({', '.join(category for category in categories)})"

    @staticmethod
    def str_to_sql_char(char_lens: np.ndarray):
        if char_lens.shape[0] == 1:
            return f"CHAR({char_lens[0]})"

    @staticmethod
    def str_to_sql_dtype(column: pd.Series):
        string_mapping = DataType.MAPPING.get("string")
        max_len = column.map(len).max()

        if (
                (char_lens := column.map(len).unique())
        ):
            if sql_dtype := DataType.str_to_sql_char(char_lens):
                return sql_dtype

        if max_len < pow(2, 8):
            return string_mapping.get("tiny")
        if max_len < pow(2, 16):
            return string_mapping.get("text")
        if max_len < pow(2, 24):
            return string_mapping.get("med")
        return string_mapping.get("long")

    @staticmethod
    def int_to_sql_dtype(column: pd.Series):
        int_mapping = DataType.MAPPING.get("int")
        dtype = column.dtype.lower()

        if sql_dtype := int_mapping.get(dtype):
            return sql_dtype

        max_value = column.max()
        if max_value < pow(2, 16):
            return int_mapping.get("int16")
        if max_value < pow(2, 32):
            return int_mapping.get("int32")
        return int_mapping.get("int64")

    @staticmethod
    def to_sql_dtype(
            column: pd.Series = Any
    ):
        dtype = column.dtype.lower()
        sql_mapping = DataType.MAPPING.get(dtype)
        if isinstance(sql_mapping, str):
            return sql_mapping
        if dtype == "category":
            return DataType.category_to_sql_dtype(column)
        if "int" in dtype:
            if (
                    (len_counts := column.astype("string").map(len).unique().shape[0]) == 1
            ):
                return DataType.to_sql_type(column.astype("string"))
            return DataType.int_to_sql_dtype(column)
        if dtype == "string":
            return DataType.int_to_sql_dtype(column)


class RyDBColumn:
    def __init__(self, column_name: str, column: pd.Series):
        self.name = column_name
        self.column = column
        self.dtype = DataType(column)


class CalculatedColumn(RyDBColumn):
    def __init__(
            self,
            column_name: str,
            source: Union[pd.Series, pd.DataFrame],
            source_label: str = None,
            **config
    ):
        column = self._transform(source, config)
        super().__init__(column_name, column)
        self.source_label = source_label

    def _transform(self, source: Union[pd.Series, pd.DataFrame], config: Dict[str, Any]):
        pass


class RyDBTable:
    def __init__(self, table_name: str, table: pd.DataFrame):
        self.name = table_name
        self.df = table
        self._refresh()

    def to_mermaid(self):
        diagram = erDiagram()
        diagram.add_entity_by_dtypes(
            self.name,
            self.df.dtypes
        )

    @staticmethod
    def _extract_attribute_from_base_model(base_model: BaseModel, attr_name: str):
        if not attr_name.startswith("$"):
            return getattr(base_model, attr_name)
        attr_path = attr_name.replace("$", "").split(".")
        attr_value = base_model
        for attr_level in attr_path:
            attr_value = getattr(attr_value, attr_level)
        return attr_value

    @staticmethod
    def _transform_value(value: Any, transform_type: str):
        pass

    @classmethod
    def from_pydantic(
            cls,
            table_name: str,
            table_def: Dict[str, str],
            column_mappings: List[Dict[str, str]],
            *base_models: BaseModel,
            ignore_by_any: Any = None,
            ignore_by_path: Any = None
    ):
        table_values = []
        for base_model in base_models:
            row_values = []
            for column_name, dtype in table_def.items():
                logging.debug(f"Calculating {column_name}:{dtype}")
                column_mapping = column_mappings.get(column_name)

                attr_name = column_mapping.get("attr_name")
                attr_value = cls._extract_attribute_from_base_model(
                    base_model,
                    attr_name
                )
                if transform_config := column_mapping.get("transform", {}):
                    attr_value = cls._transform_value(
                        attr_value,
                        **transform_config
                    )
                row_values.append(attr_value)
        table = pd.DataFrame(
            table_values,
            columns=list(table_def.keys())
        ).astype(table_def)
        return cls(table_name, table)

    def __repr__(self):
        return repr(self.df)

    @property
    def shape(self):
        return self.df.shape

    @property
    def columns(self):
        return self.df.columns.tolist()

    def loc(
            self,
            query: Union[str, pd.Series],
            column_names: List[str] = None,
            squeeze: bool = False
    ):
        mask = self.df.eval(query) if isinstance(query, str) else query
        target = (mask,) if column_names is None else (mask, column_names)
        result = self.df.loc[target]
        return result.squeeze() if squeeze else result

    @staticmethod
    def format_where_query(where_dict: dict):
        where_conds = set()
        for where_column, where_value in where_dict.items():
            where_cond = f"{where_column} == {where_value}"
            where_conds.add(where_cond)
        if len(where_conds) == 1:
            return list(where_conds)[0]
        where_conds = [
            f"({where_cond})"
            for where_cond in where_conds
        ]
        return " & ".join(where_conds)

    def stringify(self, *args):
        return [
            str(arg)
            for arg in args
        ]

    def update_row(
            self,
            set_dict: dict,
            where_dict: dict
    ):
        set_columns = list(set_dict.keys())
        set_values = list(set_dict.values())
        where_query = self.format_where_query(where_dict)
        self.df.loc[self.df.eval(where_query), set_columns] = set_values
        logging.info(f"({','.join(self.stringify(set_columns))}) updated to ({','.join(self.stringify(set_values))}) for {where_query}")

    def lookup(self, *args):
        return self.loc(*args)

    def _refresh(self):
        self.__dict__.update(
            {
                column_name: self.df[column_name]
                for column_name in self.columns
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

    def to_rich_console_table(
            self,
            query: str,
            name: str = None,
            column_formats: dict = None,
            count: int = 10,
            fillna: str = ""
    ):
        table = Table(title=name or self.name)
        results = self.df.fillna("").query(query).head(count)

        column_formats = column_formats or {
            column: {} for column in self.columns
        }

        for column, column_format in column_formats.items():
            # column_format = column_formats.get(column, {})
            table.add_column(column, **column_format)

        for row in results.astype("string").itertuples():
            row_values = [
                getattr(row, column, fillna)
                for column in column_formats
            ]
            table.add_row(*row_values)

        return table
