import pandas as pd
from O365.drive import File
from typing import Union

from rypython.ry365 import WorkBook


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



