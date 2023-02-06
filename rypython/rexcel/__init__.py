import string
from dataclasses import dataclass
from pathlib import Path
from typing import Tuple, Callable, List, Any, Union, Generator, Dict

import numpy as np
import pandas as pd
from xlsxwriter import Workbook
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

DataFrameRow = Any  # TODO: Figure out how to type hint Pandas rows
DataFrameIndex = Any


@dataclass
class RexcelFormat:
    test_func: Callable
    config: dict
    format_type: str  # TODO: Set up as ``Enum``

    def get_format(
            self,
            check_value: Any
    ):
        ...


class RexcelWorksheet:
    """
    Worksheet object for adding formatted DataFrames to Excel

    Parameters
    ----------
    workbook: Workbook
        ``Workbook`` object used to add worksheets
    worksheet_name: str
        Name of new Excel worksheet

    Attributes
    ----------
    wb: Workbook
        ``Workbook`` object used to add worksheets
    wks: Worksheet
        New worksheet as ``Worksheet`` object
    column_widths: List[Tuple[int, int, int]]
        List of (start, stop, width) tuples for setting column widths

    BUFFER_FACTOR: float
        Factor by which to increase production work file count as buffer
    UTTERANCE_TOTALS: Dict[str,int]
        Total utterances required for each locale (default
        is 50000 unless specified in ``UTTERANCE_TOTALS``)

    Examples
    --------
    >>> from lex_engineering import Scenarios
    >>> output_dir = Path.home() / 'Downloads'
    >>> scenarios = Scenarios('MIT', 'fr-FR')
    >>> scenarios.get_work_files('standard', output_dir)
    """

    def __init__(
            self,
            workbook: Workbook,
            worksheet_name: str,
            column_widths: List[Tuple[int, int, int]] = None
    ) -> None:
        self.wb = workbook
        self.wks = self.wb.add_worksheet(
            name=worksheet_name
        )

        if column_widths:
            for start, stop, width in column_widths:
                self.wks.set_column(
                    start,
                    stop,
                    width
                )

        self.current_row = 1
        self.current_column = 0

        self.data_validation_columns = []

        self.FORMATS = {
            'bold': self.workbook.add_format({'bold': 1}),
            'text': self.workbook.add_format({'num_format': '@'}),
            'percent': self.workbook.add_format({'num_format': 9}),
            'integer': self.workbook.add_format({'num_format': 1}),
            'decimal': self.workbook.add_format({'num_format': 2}),
            'locked': self.workbook.add_format({'locked': True})
        }

    @staticmethod
    def get_column_generator(
            limit: int = 26,
            row_number: int = None
    ) -> Generator[Tuple[str, int], None, None]:
        """
        Creates generator object for identifying columns

        Parameters
        ----------
        limit: int

        Returns
        -------
        Generator[Tuple[str, int], None, None]
            Returns generator that yields (column_letter, column_number) tuple
        """
        BASE_COLUMNS = list(string.ascii_uppercase)
        row_number = str(row_number) if row_number else ""
        prefix_offset = -1
        column_offset = 0
        column_count = 0
        while column_count < limit:
            if column_offset == 26:
                column_offset = 0
                prefix_offset += 1
            letter_prefix = BASE_COLUMNS[prefix_offset] if prefix_offset >= 0 else ""
            base_letter = BASE_COLUMNS[column_offset]
            column_letter = f"{letter_prefix}{base_letter}{row_number}"
            yield column_letter, column_count
            column_offset += 1
            column_count += 1

    def write_cell(
            self,
            cell: str,
            write_func: str,
            cell_text: str,
            cell_format: str
    ) -> None:
        """
        Writes given text to specified cell according to write function and cell format

        Parameters
        ----------
        cell: str
            Cell address in "A1" format
        write_func: str
            Write function to call ("write_string", "write_formula", "write_boolean", "write_number")
        cell_text: str
            Text to write
        cell_format: str
            Format for cell ("bold", "text", etc.)
        """
        cell_format = self.FORMATS.get(cell_format)
        write_func = getattr(self.wks, write_func)
        write_func(cell, cell_text, cell_format)

    def add_formula_column(
            self,
            df: pd.DataFrame,
            new_column_name: str,
            row_test: Callable[[DataFrameRow], bool],
            formula_func: Callable[[DataFrameRow, DataFrameIndex], str]
    ) -> pd.DataFrame:
        """
        Adds formula column at right side of given DataFrame according to row and index values

        Parameters
        ----------
        df: pd.DataFrame
        new_column_name: str
        row_test: Callable[[DataFrameRow], bool]
        formula_func: Callable[[DataFrameRow, DataFrameIndex], str]

        Returns
        -------
        df: pd.DataFrame
            Updated DataFrame object with new formula row added at right
        """

        def get_formula(row):
            """
            Passes row to ``row_test`` returning formatted Excel formula if True
            """
            if not row_test(row):
                return ""
            return formula_func(row, row.Index)

        df[new_column_name] = df.apply(
            lambda row: get_formula(row),
            axis=1
        )

        return df

    def add_list_data_validation_column(
            self,
            df: pd.DataFrame,
            new_column_name: str,
            row_test: Callable[[DataFrameRow], bool],
            *list_values: str
    ) -> pd.DataFrame:
        """

        Parameters
        ----------
        df: pd.DataFrame
        new_column_name: str
        row_test: Callable[[DataFrameRow], bool]
        *list_values: str

        Returns
        -------
        df: pd.DataFrame
        """

        # Adds data validation column to ``format_only_columns`` list to avoid writing it later
        # TODO: Figure out how to write default values to data validation list cells
        self.data_validation_columns.append(new_column_name)

        # Use ``row_test`` function to create mask for vectorized mapping of data validation list values
        mask = df.apply(
            lambda row: row_test(row),
            axis=1
        )

        # Map data validation values to cells where row test is True
        df[new_column_name] = np.where(
            mask,
            ",".join(list_values),
            ""
        )

        return df

    def add_right_df(
            self,
            df: pd.DataFrame,
            right_df: pd.DataFrame,
            use_index: bool = True
    ) -> pd.DataFrame:
        """
        Adds full ``right_df`` DataFrame to original ``df``

        Parameters
        ----------
        df: pd.DataFrame
        right_df: pd.DataFrame
        use_index: bool

        Returns
        -------
        df: pd.DataFrame
            Updated DataFrame object with columns from ``right_df`` added at right
        """

        if use_index:
            return df.join(right_df)

        for column_name in right_df.columns:
            # Not using index directly so first must check that columns have same length
            if right_df[column_name].shape[0] == df.shape[0]:
                df[column_name] = right_df[column_name]

        return df

    def _write_column_headers(
            self,
            df: pd.DataFrame,
            header_format: str = "bold",
            ignore_strings: List[str] = None
    ) -> None:
        """

        Parameters
        ----------
        df: pd.DataFrame
        row_number: int
        header_format: str
        ignore_strings: List[str]
            List of case-insensitive strings to skip writing as column headers (default is "index" and "unnamed")
        """
        column_headers = df.columns.tolist()
        ignore_strings = ignore_strings or ("index", "unnamed")
        columns = self.get_column_generator(
            limit=len(column_headers),
            row_number=self.current_row
        )
        column_format = self.FORMATS.get(header_format)
        for (cell, _), column_header in zip(columns, column_headers):
            if any(
                    ignore_string in column_header.lower()
                    for ignore_string in ignore_strings
            )
                self.write_cell(
                    cell,
                    'write_string',
                    column_header,
                    column_format
                )
        self.current_row += 1

    def _set_skip_rows(
            self,
            df: pd.DataFrame,
            skip_rows: int,
            include_index: bool = False
    ) -> pd.DataFrame:
        """
        Syncs df and skip_rows so that index matches for easier manipulation

        Parameters
        ----------
        df: pd.DataFrame
        skip_rows: int

        Returns
        -------
        df: pd.DataFrame
        """
        self.current_row += skip_rows
        df = df.reset_index(drop=True)
        df.index += self.current_row
        if include_index:
            df = df.reset_index().rename(columns={"index": "_index"})
            df.index += self.current_row
        return df

    def _get_format(
            self,
            formats: List[RexcelFormat],
            row: np.ndarray,
            column_header: str,
            cell_value: Any
    ) -> Union[Any, None]:
        # First check for cell-level formats
        for cell_formatter in filter(
                lambda format: format.check_type == "cell",
                formats
        ):
            if cell_format := cell_formatter.get_format(
                    (column_header, cell_value)
            ):
                return cell_format
        # If no cell-level formats found, try column-level
        else:
            for column_formatter in filter(
                    lambda format: format.check_type == "column",
                    formats
            ):
                if column_format := column_formatter.get_format(
                        column_header
                ):
                    return column_format
            # Finally, look for row-level formats
            else:
                for row_formatter in filter(
                        lambda format: format.check_type == "row",
                        formats
                ):
                    if row_format := row_formatter.get_format(
                            row
                    ):
                        return row_format

    @staticmethod
    def _get_row_edges(df: pd.DataFrame):
        return df.index[0], df.index[-1]

    def _apply_conditional_format(
            self,
            format_range: str,
            config: dict,
            start: int,
            end: int
    ):
        format_range = format_range.format_map(
            {
                "start": start,
                "end": end
            }
        )
        self.wks.conditional_format(
            format_range,
            config
        )

    def write(
            self,
            df: pd.DataFrame,
            skip_rows: int = 0,
            include_index: bool = False,
            header_format: str = "bold",
            ignore_strings: List[str] = None,
            formats: List[RexcelFormat] = None,
            conditional_formats: Dict[str, Any] = None,
            hidden_rows: List[int] = None,
            hidden_columns: List[int] = None,
            freeze_panes: Tuple[Any, Any] = None,
            hide_right_columns: str = None
    ):

        # Sync skip rows if needed
        if skip_rows:
            df = self._set_skip_rows(
                df,
                skip_rows,
                include_index=include_index
            )

        # Write column headers from ``df``
        self._write_column_headers(
            df,
            header_format=header_format,
            ignore_strings=ignore_strings
        )

        # TODO: Create better mapping to data types for rydb module
        write_funcs = {
            str: self.wks.write_string,
            bool: self.wks.write_boolean,
            int: self.wks.write_number
        }

        columns = df.columns.tolist()
        start_row, end_row = self._get_row_edges(df)

        # Iterate over rows as generator without initializing values list object
        for row in df.fillna("").values:

            for column_offset, cell_value in enumerate(row):
                cell_value = str(cell_value) if type(cell_value) not in write_funcs else cell_value
                row_number = self.current_row
                column_header = columns[column_offset]

                cell_format = self._get_format(
                    formats,
                    row,
                    column_header,
                    cell_value
                )

                cell_info = [
                    row_number,
                    column_offset,
                    cell_value,
                    cell_format
                ]

                write_func = write_funcs.get(type(cell_value))
                if cell_value and isinstance(cell_value, str) and cell_value[0] == '=':
                    write_func = self.wks.write_formula
                if cell_value == "":
                    write_func = self.wks.write_blank
                write_func(
                    *cell_info
                )

            self.current_row += 1

        # Apply conditional formats
        conditional_formats = conditional_formats or {}
        for format_range, config in conditional_formats.items():
            self._apply_conditional_format(
                format_range,
                config,
                start_row,
                end_row
            )

        for hidden_row in hidden_rows:
            self.wks.set_row(
                hidden_row,
                None,
                None,
                {'hidden': True}
            )
        for hidden_column in hidden_columns:
            self.wks.set_column(
                f"{hidden_column}:{hidden_column}",
                None,
                None,
                {'hidden': True}
            )
        if freeze_panes:
            freeze_row, freeze_column = freeze_panes
            self.wks.freeze_panes(freeze_row, freeze_column)
        if hide_right_columns is not None:
            self.wks.set_column(
                hide_right_columns,
                None,
                None,
                {
                    'hidden': True
                }
            )


class RexcelWorkbook:
    """
    The Scenarios object is the primary mechanism for viewing and manipulating existing scenarios

    Parameters
    ----------
    output_dir: Path
        Local output directory where resources should be saved (default is None)

    Attributes
    ----------
    scenarios: List[Scenario]
        Parsed list of Scenario objects representing full context of each scenario for given
        domain-locale pair
    BUFFER_FACTOR: float
        Factor by which to increase production work file count as buffer
    UTTERANCE_TOTALS: Dict[str,int]
        Total utterances required for each locale (default
        is 50000 unless specified in ``UTTERANCE_TOTALS``)

    Examples
    --------
    >>> from lex_engineering import Scenarios
    >>> output_dir = Path.home() / 'Downloads'
    >>> scenarios = Scenarios('MIT', 'fr-FR')
    >>> scenarios.get_work_files('standard', output_dir)
    """

    COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    def __init__(
            self,
            output_file: Path
    ) -> None:
        self.worksheets = []
        self.new_worksheets = {}
        self.output_file = output_file

    def __enter__(self):
        self.wb = Workbook(self.output_file)
        return self

    def __exit__(self, type, value, traceback):
        self.wb.close()

    def add_format(self, config: dict):
        return self.wb.add_format(config)

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
            right_df: pd.DataFrame = None,
            comment_column: str = None,
            header_format: Format = None,
            hide_right_columns: str = None
    ):
        hidden_rows = hidden_rows or []
        hidden_columns = hidden_columns or []
        header_calculations = header_calculations or []
        df = df.where(
            pd.notnull(df), ''
        )
        row_number = skip_rows
        format_test, row_format = format_rows if format_rows else (None, None)

        """
        wks = RexcelWorksheet(
            self.wb, 
            worksheet_name,
            column_widths=column_widths
        )
        """

        wks = self.wb.add_worksheet(name=worksheet_name)

        FORMATS = {
            'bold': self.wb.add_format({'bold': 1}),
            'text': self.wb.add_format({'num_format': '@'}),
            'percent': self.wb.add_format({'num_format': 9}),
            'integer': self.wb.add_format({'num_format': 1}),
            'decimal': self.wb.add_format({'num_format': 2}),
            'locked': self.wb.add_format({'locked': True})
        }

        if column_widths:
            for first, last, width in column_widths:
                wks.set_column(first, last, width)

        if header_calculations:
            """
            for cell_config in header_calculations:
                wks.write_cell(**cell_config)
            """
            for cell, write_func, cell_text, cell_format in header_calculations:
                cell_format = FORMATS.get(cell_format)
                write_func = getattr(wks, write_func)
                write_func(cell, cell_text, cell_format)

        # Create list of column headers
        columns = df.columns.tolist()

        # Add formula column headers, if present
        """        
        for new_column_name, (row_test, formula_func) in formula_columns.items():
            df = wks.add_formula_column(
                df,
                new_column_name,
                row_test,
                formula_func
            )
        """
        if formula_columns is not None:
            for formula_column in formula_columns:
                columns.append(formula_column)

        # Add data validation column headers, if present
        if data_validation_columns is not None:
            for data_validation_column in data_validation_columns:
                columns.append(data_validation_column)

        # Add right DataFrame column headers, if present
        if right_df is not None:
            columns.extend(right_df.columns.tolist())

        # Add comment column header, if present
        if comment_column is not None:
            columns.append(comment_column)

        # Write column headers
        for j, header in enumerate(columns):
            if not header or header.lower() in ('index',) or 'unnamed:' in header.lower():
                continue
            wks.write(
                f"{self.get_column_letter(j)}{row_number + 1}",
                header,
                header_format or FORMATS.get('bold')
            )

        # Create list of DataFrame values
        # Create index column, as needed
        if include_index:
            df = df.reset_index()
        rows = df.values.tolist()

        # Create list of right DataFrame values, if present
        right_rows = right_df.values.tolist() if right_df is not None else right_df

        # Increase starting row number to account for header row
        row_number += 1

        write_funcs = {
            str: wks.write_string,
            bool: wks.write_boolean,
            int: wks.write_number
        }
        for i, row in enumerate(rows):
            col_number = 0
            for j, cell in enumerate(row):
                if type(cell) not in write_funcs:
                    cell = str(cell)
                cell_info = [
                    row_number,
                    col_number,
                    cell
                ]
                if row_format is not None and format_test(row_number):
                    cell_info.append(row_format)
                if format_columns and (column_format := format_columns.get(columns[col_number])):
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
                col_number += 1

            # Adds formula column info rendered redundant above
            if formula_columns is not None:
                for row_test, formula_format in formula_columns.values():
                    if row_test(row):
                        cell_info = [
                            row_number,
                            col_number,
                            formula_format(row, row_number)
                        ]
                        wks.write_formula(*cell_info)
                    col_number += 1
            if data_validation_columns is not None:
                for row_test, data_validation in data_validation_columns.values():
                    if row_test(row):
                        cell = f"{self.get_column_letter(col_number)}{row_number}"
                        wks.data_validation(
                            cell,
                            data_validation
                        )
                    col_number += 1
            if right_rows is not None:
                for k, cell in enumerate(right_rows[i]):
                    if type(cell) not in write_funcs:
                        cell = str(cell)
                    cell_info = [
                        row_number,
                        col_number,
                        cell
                    ]
                    write_func = write_funcs.get(type(cell))
                    if cell and isinstance(cell, str) and cell[0] == '=':
                        write_func = wks.write_formula
                    if cell == "":
                        write_func = wks.write_blank
                    write_func(
                        *cell_info
                    )
                    col_number += 1
            if comment_column is not None:
                wks.write_blank(
                    row_number,
                    col_number,
                    ""
                )
                col_number += 1
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
        if hide_right_columns is not None:
            wks.set_column(
                hide_right_columns,
                None,
                None,
                {
                    'hidden': True
                }
            )
        self.worksheets.append(wks)

    def add_worksheet(
            self,
            worksheet_name: str,
            column_widths: List[Tuple[int, int, int]]
    ) -> RexcelWorksheet:
        wks = RexcelWorksheet(
            self.wb,
            worksheet_name,
            column_widths=column_widths
        )
        self.new_worksheets[worksheet_name] = wks
        return wks



    def new_add_worksheet_by_dataframe(
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
            ad_hoc_cells: list = None,
            skip_rows: int = 0,
            freeze_panes: Tuple[int, int] = None,
            data_validation_columns: dict = None,
            right_df: pd.DataFrame = None,
            comment_column: str = None,
            header_format: Format = None,
            hide_right_columns: str = None
    ):
        hidden_rows = hidden_rows or []
        hidden_columns = hidden_columns or []
        ad_hoc_cells = ad_hoc_cells or []
        df = df.where(
            pd.notnull(df), ''
        )
        row_number = skip_rows
        format_test, row_format = format_rows if format_rows else (None, None)

        # Register new worksheet
        wks = self.add_worksheet(
            sheet_name=worksheet_name,
            column_widths=column_widths
        )

        # Write ad hoc cells
        for cell_config in ad_hoc_cells:
            wks.write_cell(**cell_config)




        """        
        for new_column_name, (row_test, formula_func) in formula_columns.items():
            df = wks.add_formula_column(
                df,
                new_column_name,
                row_test,
                formula_func
            )
        """
        if formula_columns is not None:
            for formula_column in formula_columns:
                columns.append(formula_column)

        # Add data validation column headers, if present
        if data_validation_columns is not None:
            for data_validation_column in data_validation_columns:
                columns.append(data_validation_column)

        # Add right DataFrame column headers, if present
        if right_df is not None:
            columns.extend(right_df.columns.tolist())

        # Add comment column header, if present
        if comment_column is not None:
            columns.append(comment_column)

        # Write column headers
        for j, header in enumerate(columns):
            if not header or header.lower() in ('index',) or 'unnamed:' in header.lower():
                continue
            wks.write(
                f"{self.get_column_letter(j)}{row_number + 1}",
                header,
                header_format or FORMATS.get('bold')
            )

        # Create list of DataFrame values
        # Create index column, as needed
        if include_index:
            df = df.reset_index()
        rows = df.values.tolist()

        # Create list of right DataFrame values, if present
        right_rows = right_df.values.tolist() if right_df is not None else right_df

        # Increase starting row number to account for header row
        row_number += 1

        write_funcs = {
            str: wks.write_string,
            bool: wks.write_boolean,
            int: wks.write_number
        }
        for i, row in enumerate(rows):
            col_number = 0
            for j, cell in enumerate(row):
                if type(cell) not in write_funcs:
                    cell = str(cell)
                cell_info = [
                    row_number,
                    col_number,
                    cell
                ]
                if row_format is not None and format_test(row_number):
                    cell_info.append(row_format)
                if format_columns and (column_format := format_columns.get(columns[col_number])):
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
                col_number += 1

            # Adds formula column info rendered redundant above
            if formula_columns is not None:
                for row_test, formula_format in formula_columns.values():
                    if row_test(row):
                        cell_info = [
                            row_number,
                            col_number,
                            formula_format(row, row_number)
                        ]
                        wks.write_formula(*cell_info)
                    col_number += 1
            if data_validation_columns is not None:
                for row_test, data_validation in data_validation_columns.values():
                    if row_test(row):
                        cell = f"{self.get_column_letter(col_number)}{row_number}"
                        wks.data_validation(
                            cell,
                            data_validation
                        )
                    col_number += 1
            if right_rows is not None:
                for k, cell in enumerate(right_rows[i]):
                    if type(cell) not in write_funcs:
                        cell = str(cell)
                    cell_info = [
                        row_number,
                        col_number,
                        cell
                    ]
                    write_func = write_funcs.get(type(cell))
                    if cell and isinstance(cell, str) and cell[0] == '=':
                        write_func = wks.write_formula
                    if cell == "":
                        write_func = wks.write_blank
                    write_func(
                        *cell_info
                    )
                    col_number += 1
            if comment_column is not None:
                wks.write_blank(
                    row_number,
                    col_number,
                    ""
                )
                col_number += 1
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
        if hide_right_columns is not None:
            wks.set_column(
                hide_right_columns,
                None,
                None,
                {
                    'hidden': True
                }
            )
        self.worksheets.append(wks)