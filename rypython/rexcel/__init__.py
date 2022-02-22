from pathlib import Path
from typing import Tuple, Callable
from xlsxwriter.format import Format

import pandas as pd
import xlsxwriter


class RexcelWorkbook:
    COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    def __init__(
            self,
            output_file: Path
    ) -> None:
        self.worksheets = []
        self.output_file = output_file

    def __enter__(self):
        self.workbook = xlsxwriter.Workbook(self.output_file)
        return self

    def __exit__(self, type, value, traceback):
        self.workbook.close()

    def add_format(self, config: dict):
        return self.workbook.add_format(config)

    def add_worksheet_by_dataframe(
            self,
            df: pd.DataFrame,
            worksheet_name: str,
            column_widths: list,
            include_index: bool = False,
            format_rows: Tuple[Callable, Format] = None,
            formula_columns: dict = None,
            conditional_formatting: dict = None
    ):
        df = df.where(pd.notnull(df), '')
        format_test, row_format = format_rows if format_rows else (None, None)
        if include_index:
            df = df.reset_index()
        wks = self.workbook.add_worksheet(name=worksheet_name)
        bold = self.workbook.add_format({'bold': 1})
        if column_widths:
            for first, last, width in column_widths:
                wks.set_column(first, last, width)
        columns = df.columns.tolist()
        if formula_columns is not None:
            for formula_column in formula_columns:
                columns.append(formula_column)
        for j, header in enumerate(columns):
            if not header or header == 'index':
                continue
            wks.write(
                f"{self.COLS[j]}1",
                header,
                bold
            )
        rows = df.values.tolist()
        row_number = 1
        col_number = 0
        write_funcs = {
            str: wks.write_string,
            bool: wks.write_boolean,
            int: wks.write_number
        }
        for row in rows:
            offset = 0
            for j, cell in enumerate(row):
                if type(cell) not in write_funcs:
                    cell = str(cell)
                cell_info = [
                    row_number,
                    col_number + offset,
                    cell
                ]
                if format_rows is not None and format_test(row_number):
                    cell_info.append(row_format)
                write_func = write_funcs.get(type(cell))
                write_func(
                    *cell_info
                )
                offset += 1
            if formula_columns is not None:
                for row_test, formula_format in formula_columns.values():
                    if row_test(row):
                        cell_info = [
                            row_number,
                            col_number + offset,
                            formula_format(row, row_number)
                        ]
                        wks.write_formula(*cell_info)
                    offset += 1
            row_number += 1
        if conditional_formatting:
            for format_range, config in conditional_formatting.items():
                if '{end}' in format_range:
                    format_range = format_range.format(end=len(rows) + 1)
                wks.conditional_format(
                    format_range,
                    config
                )
        self.worksheets.append(wks)
