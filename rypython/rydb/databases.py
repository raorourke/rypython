import os
from pathlib import Path
from typing import List, Union, Dict

import pandas as pd

from rypython.randas import DataFrame
from rypython.ry365 import O365Account
from rypython.rydb.tables import RyDBTable

DEFAULT_DOWNLOAD_DIR = os.environ.get(
    'RYPYTHON_DEFAULT_DOWNLOAD_DIR',
    Path.home() / 'Downloads'
)


class RyDBSource:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)

    def collect(self):
        ...

    def replace(self, tables: Dict[str, pd.DataFrame]) -> None:
        ...

    def update(self):
        ...


class O365DB(RyDBSource):
    def replace(self, tables: Dict[str, pd.DataFrame]) -> None:
        folder = self.folder
        existing_db_file = folder.get_item(self.db_filename)
        new_local_db_file = DEFAULT_DOWNLOAD_DIR / self.db_filename
        with pd.ExcelWriter(
                new_local_db_file,
                engine="xlsxwriter"
        ) as writer:
            for table_name, table in tables.items():
                if table_name == "source":
                    continue
                table.to_excel(writer, sheet_name=table_name, index=False)
        if new_local_db_file.exists() and existing_db_file is not None:
            existing_db_file.delete()
            self.folder.upload_file(new_local_db_file)


class RyDB:
    def __init__(self, source: RyDBSource, **data):
        self.source = source
        self.__dict__.update(
            {
                table_name: RyDBTable(table_name, table)
                for table_name, table in data.items()
            }
        )

    def __setattr__(self, table_name: str, table: pd.DataFrame):
        super().__setattr__(table_name, table)

    @classmethod
    def read_o365(
            cls,
            site: str,
            filepath: Union[str, List[str]],
            db_filename: str
    ):
        account = O365Account(site=site)
        filepath = filepath.split("/") if isinstance(filepath, str) else filepath
        folder = account.get_folder(*filepath)
        remote_db_file = folder.get_item(db_filename)
        source_config = {
            "account": account,
            "folder": folder,
            "db_filename": db_filename
        }
        with DataFrame(remote_db_file) as wb:
            return cls(
                source=O365DB(**source_config),
                **wb.data
            )

    def replace(self):
        self.source.replace(
            {
                table_name: table.df
                for table_name, table in self.__dict__.items()
                if isinstance(table, RyDBTable)
            }
        )
