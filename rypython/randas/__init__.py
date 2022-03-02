import pandas as pd
from O365.drive import File

from rypython.ry365 import WorkBook


def read_excel365(xl_file: File, sheet_name: str = None, skip_rows: int = 0):
    wb = WorkBook(xl_file)
    if sheet_name:
        ws = wb.get_worksheet(sheet_name)
    else:
        wss = wb.get_worksheets()
        ws = wss[0]
    _range = ws.get_used_range()
    cols, *values = _range.values[skip_rows:]
    df = pd.DataFrame(values, columns=cols)
    df.range = _range
    return df
