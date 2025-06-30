from pathlib import Path

import pandas as pd

from uvmlib.export.sheets import Sheet, DocSheet


class ExcelExporter:
    """
    Excel exporter with automatic formatting and multi-sheet support.

    This class handles the export of pandas DataFrames to Excel with consistent
    formatting, including date formats, autofit columns, and autofilter.

    Parameters
    ----------
    date_format : str, optional
        Format string for dates, by default 'DD-MM-YYYY'
    datetime_format : str, optional
        Format string for datetime values, by default 'DD-MM-YYYY'
    """

    def __init__(
        self,
        date_format: str = 'DD-MM-YYYY',
        datetime_format: str = 'DD-MM-YYYY',
    ) -> None:
        self.date_format = date_format
        self.datetime_format = datetime_format

    def _write_sheet(
        self,
        writer: pd.ExcelWriter,
        sheet: Sheet,
    ) -> None:
        """Schrijf sheet naar excelbestand."""
        if hasattr(sheet, 'write_to_excel'):
            sheet.write_to_excel(writer)
            return

        df = sheet.data
        n_rows, n_cols = df.shape
        col_nlevels = df.columns.nlevels + (df.columns.nlevels > 1)

        # Bereken welke rijen/kolommen bevroren moeten worden
        freeze_row = col_nlevels
        freeze_col = (df.index.nlevels if sheet.toon_index else 0)

        df.to_excel(
            writer,
            sheet_name = sheet.naam,
            index = sheet.toon_index,
            freeze_panes = (freeze_row, freeze_col),
        )

        # Formatteer sheet
        excel_sheet = writer.sheets[sheet.naam]
        self._format_sheet(
            sheet = excel_sheet,
            col_nlevels = col_nlevels,
            idx_offset = df.index.nlevels if sheet.toon_index else 0,
            n_rows = n_rows,
            n_cols = n_cols,
        )

        if sheet.tab_kleur:
            excel_sheet.set_tab_color(sheet.tab_kleur)

    def _format_sheet(
        self,
        sheet: object,
        col_nlevels: int,
        idx_offset: int,
        n_rows: int,
        n_cols: int,
    ) -> None:
        """Apply formatting to a worksheet.

        Parameters
        ----------
        sheet : Any
            Excel worksheet object
        col_nlevels : int
            Number of column levels
        idx_offset : int
            Number of columns to offset for index (if shown)
        n_rows : int
            Number of rows
        n_cols : int
            Number of columns
        """
        sheet.autofit()
        sheet.autofilter(
            first_row = col_nlevels - 1,
            first_col = 0,
            last_row  = n_rows,
            last_col  = idx_offset + n_cols - 1
        )

    def export_workbook(
        self,
        sheets: Sheet | pd.DataFrame | list[Sheet | pd.DataFrame],
        filepath: Path|str,
    ) -> None:
        """
        Exporteer één pf meerdere DataFrames naar een Excelbestand.

        Parameters
        ----------
        sheets : Sheet | pd.DataFrame | list[Sheet | pd.DataFrame]
            Een enkele sheet/DataFrame of lijst van sheets/DataFrames
        filepath : Path|str
            Bestandslocatie om Excel file op te slaan
        """
        filepath = Path(filepath)

        if not isinstance(sheets, list):
            sheets = [sheets]

        sheets = [
            Sheet(naam='data', data=sheet, omschrijving="")
            if isinstance(sheet, pd.DataFrame)
            else sheet
            for sheet in sheets
        ]

        with pd.ExcelWriter(
            filepath,
            date_format = self.date_format,
            datetime_format = self.datetime_format,
        ) as writer:
            for sheet in sheets:
                self._write_sheet(writer, sheet)
