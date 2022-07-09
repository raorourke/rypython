import pandas as pd
import os
from pathlib import Path
from O365.drive import File
from typing import Union

from rypython.ry365 import WorkBook


DEFAULT_DOWNLOAD_DIR = os.environ.get('RYPYTHON_DEFAULT_DOWNLOAD_DIR', Path.home() / 'Downloads')


def read_excel365(xl_file: Union[File, WorkBook], sheet_name: str = False, skip_rows: int = 0):
    wb = WorkBook(xl_file) if isinstance(xl_file, File) else xl_file
    if sheet_name is None:
        wss = wb.get_worksheets()
        dfs = {}
        for ws in wss:
            dfs[ws.name] = read_excel365(wb, sheet_name=ws.name, skip_rows=skip_rows)
        return dfs
    if sheet_name is False:
        wss = wb.get_worksheets()
        ws = wss[0]
    elif sheet_name:
        ws = wb.get_worksheet(sheet_name)
    _range = ws.get_used_range()
    cols, *values = _range.values[skip_rows:]
    df = pd.DataFrame(values, columns=cols)
    df.range = _range
    return df


class DataFrame:
    def __init__(
            self,
            source_file: File,
            download_dir: Path = DEFAULT_DOWNLOAD_DIR,
            sheet_name: str = None
    ):
        self.source_file = source_file
        self.file_type = source_file.name.rsplit('.', 1)[1].lower()
        self.download_dir = download_dir
        self.sheet_name = sheet_name

    def __enter__(self):
        self.source_file.download(self.download_dir)
        self.local_file = self.download_dir / self.source_file.name
        self.data = self._parse_source_file()
        return self

    def _parse_source_file(self):
        if self.file_type == 'csv':
            return pd.read_csv(self.local_file)
        return pd.read_excel(
            self.local_file,
            sheet_name=self.sheet_name
        )

    def __exit__(self, type, value, traceback):
        os.remove(self.local_file)