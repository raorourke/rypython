from pathlib import Path
from typing import Tuple, Callable

import pandas as pd
import xlsxwriter
from xlsxwriter.format import Format


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

    @staticmethod
    def get_column_letter(column_index: int):
        COLS = RexcelWorkbook.COLS
        if column_index < 26:
            return COLS[column_index]
        first_letter = COLS[(column_index // 26) - 1]
        second_letter = COLS[column_index % 26]
        return f"{first_letter}{second_letter}"

    def add_worksheet_by_dataframe(
            self,
            df: pd.DataFrame,
            worksheet_name: str = 'Master',
            column_widths: list = None,
            include_index: bool = False,
            format_rows: Tuple[Callable, Format] = None,
            format_columns: dict = None,
            formula_columns: dict = None,
            conditional_formatting: dict = None,
            hidden_rows: list = None,
            hidden_columns: list = None,
            header_calculations: list = None,
            skip_rows: int = 0,
            freeze_panes: Tuple[int, int] = None,
            data_validation_columns: dict = None,
            comment_column: str = None
    ):
        hidden_rows = hidden_rows or []
        hidden_columns = hidden_columns or []
        header_calculations = header_calculations or []
        df = df.where(pd.notnull(df), '')
        row_number = skip_rows
        format_test, row_format = format_rows if format_rows else (None, None)
        if include_index:
            df = df.reset_index()
        wks = self.workbook.add_worksheet(name=worksheet_name)
        FORMATS = {
            'bold': self.workbook.add_format({'bold': 1}),
            'text': self.workbook.add_format({'num_format': '@'}),
            'percent': self.workbook.add_format({'num_format': 9}),
            'integer': self.workbook.add_format({'num_format': 1}),
            'decimal': self.workbook.add_format({'num_format': 2}),
            'locked': self.workbook.add_format({'locked': True})
        }
        if column_widths:
            for first, last, width in column_widths:
                wks.set_column(first, last, width)

        if header_calculations:
            for cell, write_func, cell_text, cell_format in header_calculations:
                cell_format = FORMATS.get(cell_format)
                write_func = getattr(wks, write_func)
                write_func(cell, cell_text, cell_format)

        columns = df.columns.tolist()
        if formula_columns is not None:
            for formula_column in formula_columns:
                columns.append(formula_column)
        if data_validation_columns is not None:
            for data_validation_column in data_validation_columns:
                columns.append(data_validation_column)
        if comment_column is not None:
            columns.append(comment_column)
        for j, header in enumerate(columns):
            if not header or header == 'index':
                continue
            wks.write(
                f"{self.get_column_letter(j)}{row_number + 1}",
                header,
                FORMATS.get('bold')
            )
        rows = df.values.tolist()
        row_number += 1
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
                if row_format is not None and format_test(row_number):
                    cell_info.append(row_format)
                if format_columns and (column_format := format_columns.get(columns[offset])):
                    column_format = FORMATS.get(column_format)
                    cell_info.append(column_format)
                write_func = write_funcs.get(type(cell))
                if cell and isinstance(cell, str) and cell[0] == '=':
                    write_func = wks.write_formula
                if cell == "":
                    write_func = wks.write_blank
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
            if data_validation_columns is not None:
                for row_test, data_validation in data_validation_columns.values():
                    if row_test(row):
                        cell = f"{self.get_column_letter(offset)}{row_number}"
                        wks.data_validation(
                            cell,
                            data_validation
                        )
                    offset += 1
            if comment_column is not None:
                wks.write_blank(
                    row_number,
                    col_number + offset,
                    ""
                )
            row_number += 1
        if conditional_formatting:
            for format_range, config in conditional_formatting.items():
                if '{end}' in format_range:
                    format_range = format_range.format(end=len(rows) + 1)
                wks.conditional_format(
                    format_range,
                    config
                )
        for hidden_row in hidden_rows:
            wks.set_row(hidden_row, None, None, {'hidden': True})
        for hidden_column in hidden_columns:
            wks.set_column(f"{hidden_column}:{hidden_column}", None, None, {'hidden': True})
        if freeze_panes:
            freeze_row, freeze_column = freeze_panes
            wks.freeze_panes(freeze_row, freeze_column)
        self.worksheets.append(wks)
