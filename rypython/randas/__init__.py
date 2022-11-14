import logging
import os
import re
from pathlib import Path
from typing import Union
from urllib.error import HTTPError
from urllib.parse import urlparse

import pandas as pd
import requests
from O365.drive import File
from bs4 import BeautifulSoup as bs

from rypython.ry365 import WorkBook

logging.basicConfig(level=logging.INFO)

HEADER = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.75 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest"
}

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


class HTMLDataFrame:
    def __init__(self, url: str):
        self.url = urlparse(url)
        self.tables = self._parse_url() or []

    def _parse_url(self, text: str = None):
        try:
            return pd.read_html(text or self.url.geturl())
        except AttributeError as e:
            logging.error(f"Error parsing {self.url.geturl()}: {e}!")
            r = requests.get(self.url.geturl(), headers=HEADER)
            text = r.text
            rowspan_pattern = r'(?P<rowspan>rowspan="\d+)(&([a-z])+;)+"'
            text = re.sub(rowspan_pattern, r'\g<rowspan>"', text)
            colspan_pattern = r'(?P<colspan>colspan="\d+)(&([a-z])+;)+"'
            text = re.sub(colspan_pattern, r'\g<colspan>"', text)
            return pd.read_html(text)
        except HTTPError as e:
            logging.error(f"Error parsing {self.url.geturl()}: {e}!")
            r = requests.get(self.url.geturl(), headers=HEADER)
            self._parse_url(text=r.text)
        except ValueError as e:
            logging.error(f"Error parsing tables for {self.url.geturl()}: {e}!")
            return []

    def get_list_items(self):
        items = {}
        r = requests.get(self.url.geturl(), headers=HEADER)
        soup = bs(r.text, 'lxml')
        uls = soup.find_all('ul')
        for ul in uls:
            for li in ul.find_all('li'):
                if item := li.text:
                    items.setdefault('Bulleted Lists', []).append(item)
        ols = soup.find_all('ol')
        for ol in ols:
            for li in ol.find_all('li'):
                if item := li.text:
                    items.setdefault('Ordered Lists', []).append(item)
        if items:
            return {
                list_type: pd.DataFrame(list_items, columns=['List Items'])
                for list_type, list_items in items.items()
            }
        return items

    def export(self, output_dir: Path = None, filename: str = None):
        output_dir = output_dir or DEFAULT_DOWNLOAD_DIR
        filename = filename or self.url.path.replace('/', '_')
        if not filename.endswith('.xlsx'):
            filename = f"{filename}.xlsx"
        outfile = output_dir / filename
        table_count = 1
        with pd.ExcelWriter(outfile, engine='xlsxwriter') as writer:
            source_df = pd.DataFrame(
                [[self.url.geturl()]],
                columns=['Source']
            )
            source_df.to_excel(
                writer,
                sheet_name='Source',
                index=False
            )
            for table in self.tables:
                first_column = table.columns[0]
                if (not isinstance(first_column, str)) or all(
                        l.isdigit()
                        for l in first_column
                ):
                    table.columns = table.iloc[0]
                    table = table.drop(table.index[0])
                if table.empty:
                    continue
                table.to_excel(
                    writer,
                    sheet_name=f"Table {table_count}",
                    index=False
                )
                table_count += 1
            for list_type, list_df in self.get_list_items().items():
                list_df.to_excel(
                    writer,
                    sheet_name=list_type,
                    index=False
                )


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
