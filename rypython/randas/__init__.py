import pandas as pd
import welo365

def read_excel365(file_url: str, sheet_name: str):
    xl = welo365.account.O365Account().search(file_url)
    wb = welo365.excel.WorkBook(xl)
    ws = wb.get_worksheet(sheet_name)
    _range = ws.get_used_range()
    cols, *values = _range.values
    df = pd.DataFrame(values, columns=cols)
    df.range = _range
    return df