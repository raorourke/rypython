import logging
import os
from pathlib import Path
from typing import List, Union, Dict, Any

import pandas as pd
import pyodbc

from rypython.randas import DataFrame
from rypython.ry365 import O365Account
from rypython.rydb.tables import RyDBTable

logging.basicConfig(level=logging.DEBUG)

DEFAULT_DOWNLOAD_DIR = os.environ.get(
    'RYPYTHON_DEFAULT_DOWNLOAD_DIR',
    Path.home() / 'Downloads'
)


class RyDBSource:
    def __init__(self, type: str, **kwargs):
        self.type = type
        self.updated = set()

    @property
    def has_changed(self):
        return bool(self.updated)

    def connect(self):
        ...

    def commit(self):
        ...

    def collect_table(self, table_name: str):
        return self.conn.get(table_name)

    def collect_all(self):
        ...

    def replace(self, tables: Dict[str, pd.DataFrame]) -> None:
        ...

    def update_row(
            self,
            table_name: RyDBTable,
            set_dict: dict,
            where_dict: dict
    ):
        # table = self.collect_table(table_name)
        self.updated.add(table_name)
        self.conn[table_name].update_row(set_dict=set_dict, where_dict=where_dict)


class O365DB(RyDBSource):
    def __init__(
            self,
            site: str,
            filepath: str,
            db_filename: str
    ):
        super().__init__(type="o365")
        self.site = site
        self.filepath = "/".split(filepath)
        self.db_filename = db_filename
        self.db_file = self.connect()
        self.conn = self.collect_all()

    def connect(self):
        account = O365Account(site=self.site)
        folder = account.get_folder(*self.filepath)
        return folder.get_item(self.db_filename)

    def collect_table(self, table_name):
        with DataFrame(self.db_file, sheet_name=table_name) as db:
            return RyDBTable(table_name, db.data)


    def collect_all(self):
        with DataFrame(self.db_file) as db:
            return {
                table_name: RyDBTable(table_name, table)
                for table_name, table in db.data.items()
            }

    def commit(self):
        if self.has_changed:
            local_db_file = DEFAULT_DOWNLOAD_DIR / self.db_file.name
            with pd.ExcelWriter(local_db_file, engine="xlsxwriter") as writer:
                for table_name, table in self.conn.items():
                    table.to_excel(writer, sheet_name=table_name, index=False)
                if not local_db_file.exists():
                    logging.error(f"Could not find {local_db_file}!")
                    return
                folder = self.db_file.get_parent()
                self.db_file.delete()
                folder.upload_file(local_db_file)
                if new_db_file := folder.get_item(local_db_file.name):
                    logging.info(f"Database file updated at {new_db_file}")

    def list_tables(self):
        return list(self.data.keys())
    def update_row(
            self,
            table_name: str,
            set_dict: dict,
            where_dict: dict
    ):
        super().update_row(table_name, set_dict=set_dict, where_dict=where_dict)

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


class SQLDB(RyDBSource):
    def __init__(
            self,
            username: str,
            password: str,
            server: str,
            database: str,
            port: int = 1433,
            driver: str = "{ODBC Driver 17 for SQL Server}",
            engine: str = "pyodbc"
    ):
        super().__init__(type="sql")
        self.server = server
        self.host = server.split(".", 1)[0]
        self.database = database
        self.port = port
        self.driver = driver
        self.engine = engine
        self._conn = self.connect(username, password)
        self.conn = self.collect_all()
        self.curr = self._conn.cursor()

    def connect(self, username: str, password: str):
        if self.engine != "pyodbc":
            return
        conn_string = ";".join(
            [
                f"DRIVER={self.driver}",
                f"PORT={self.port}",
                f"SERVER={self.server}",
                f"DATABASE={self.database}",
                f"UID={username}",
                f"PWD={password}"
            ]
        )
        return pyodbc.connect(conn_string)

    def collect_table(self, table_name: str = None):
        query = f"""
        SELECT * FROM {table_name} 
        """
        return RyDBTable(table_name, pd.read_sql(query, self._conn))

    def collect_all(self):
        table_names = self.list_tables()
        return {
            table_name: self.collect_table(table_name)
            for table_name in table_names
        }

    def list_tables(self):
        query = """
        SELECT schema_name(t.schema_id) as schema_name,
            t.name as table_name,
            t.create_date,
            t.modify_date
        FROM sys.tables t
        ORDER BY schema_name, table_name;
        """
        df = pd.read_sql(query, self._conn)
        return df.table_name.tolist()

    def commit(self):
        if self.has_changed:
            self._conn.commit()

    def execute(
            self,
            query: str,
            *values: Any
    ):
        logging.info(f"Sending {query} with {values}")
        self.curr.execute(
            query,
            *values
        )

    def update_row(
            self,
            table_name: str,
            set_dict: dict,
            where_dict: dict,
            _execute: bool = True,
            _replace: bool = False
    ):
        super().update_row(table_name, set_dict=set_dict, where_dict=where_dict)
        if _execute:
            logging.info("Updating server table")
            set_clause = ", ".join(f"{set_key} = ?" for set_key in set_dict)
            where_clause = ", ".join(f"{where_key} = ?" for where_key in where_dict)
            values = [
                *set_dict.values(),
                *where_dict.values()
            ]
            query = f"""
                    UPDATE
                        {table_name}
                        SET {set_clause}
                        WHERE {where_clause}
                    """
            self.execute(query, *values)
        if _replace:
            logging.info("Print replacing server table")
            table = self.conn[table_name]
            table.df.to_sql(table_name, con=self._conn, if_exists="replace")







class RyDB:
    def __init__(
            self,
            source: RyDBSource,
            read_only: bool = False,
            collect_all: bool = True
    ):
        self.source = source
        self.read_only = read_only
        if collect_all:
            self.source.collect_all()
        self.conn = self.source.conn

    def __setattr__(self, table_name: str, table: pd.DataFrame):
        super().__setattr__(table_name, table)

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        if not self.read_only:
            self.commit()

    def commit(self):
        self.source.commit()

    def collect_table(self, table_name: str):
        return self.source.collect_table(table_name)

    def collect_all(self):
        return self.source.collect_all()

    @classmethod
    def read_o365(
            cls,
            site: str,
            filepath: Union[str, List[str]],
            db_filename: str
    ):
        SOURCE = O365DB(
            site=site,
            filepath=filepath,
            db_filename=db_filename
        )
        return cls(source=SOURCE)

    @classmethod
    def read_sql(
            cls,
            server: str,
            username: str,
            password: str,
            database: str,
            port: str = 1433,
            driver: str = "{ODBC Driver 17 for SQL Server}",
            engine: str = "pyodbc"
    ):
        SOURCE = SQLDB(
            username,
            password,
            server=server,
            database=database,
            port=port,
            driver=driver,
            engine=engine
        )
        return cls(source=SOURCE)

    def replace(self):
        self.source.replace(
            {
                table_name: table.df
                for table_name, table in self.__dict__.items()
                if isinstance(table, RyDBTable)
            }
        )

    def count(self, table_name: str, column_name: str = None):
        return

    def update_row(self, table_name: str, set_dict: dict, where_dict: dict):
        self.source.update_row(table_name, set_dict, where_dict)
