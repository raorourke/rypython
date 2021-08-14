from pathlib import Path
import pandas as pd
from rypython.ry365 import WorkBook
from O365.drive import File

def read_excel365(xl_file: File):
    wb = WorkBook(xl_file)
    wss = wb.get_worksheet()
    ws = wss[0]
    _range = ws.get_used_range()
    cols, *values = _range.values
    df = pd.DataFrame(values, columns=cols)
    df.range = _range
    return df
