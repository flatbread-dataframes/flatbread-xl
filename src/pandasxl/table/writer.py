from pathlib import Path
from typing import Any, cast

import pandas as pd
import numpy as np
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from uvmlib.export.formatter import spans
from pandasxl.layout import TableLayout, CellPosition
from pandasxl.pattern import PatternMatcher
from pandasxl.merge import MergeManager
from pandasxl.borders import BorderManager
from pandasxl.style import (
    StyleManager,
    HorizontalAlignment,
    VerticalAlignment,
)


class ExcelTableWriter:
    """
    Writes pandas DataFrame to Excel with proper formatting and layout.
    Uses the TableLayout system to manage cell positioning.
    """
    # Default styles
    INDEX_NAMES_STYLE = {
        "bold": True,
        "horizontal_alignment": HorizontalAlignment.LEFT,
        "vertical_alignment": VerticalAlignment.CENTER,
    }

    COLUMN_NAMES_STYLE = {
        "bold": True,
        "horizontal_alignment": HorizontalAlignment.RIGHT,
        "vertical_alignment": VerticalAlignment.CENTER,
    }

    INDEX_STYLE = {
        "bold": True,
        "horizontal_alignment": HorizontalAlignment.LEFT,
        "vertical_alignment": VerticalAlignment.TOP,
    }

    COLUMN_STYLE = {
        "bold": True,
        "horizontal_alignment": HorizontalAlignment.CENTER,
        "vertical_alignment": VerticalAlignment.CENTER,
    }

    # Default value for NA/NaN values
    NA_REPRESENTATION = ""

    def __init__(
        self,
        df: pd.DataFrame,
        worksheet: Worksheet,
        x_offset: int = 0,
        y_offset: int = 0,
        default_number_format: str | None = None,
        number_formats: dict | None = None,
        border_specs: Any | None = None,
    ):
        """
        Initialize the Excel writer with DataFrame and layout information.

        Parameters
        ----------
        df : pd.DataFrame
            The DataFrame to write to Excel
        worksheet : Worksheet
            The openpyxl worksheet to write to
        x_offset : int, default 0
            Horizontal offset in cells (0-based)
        y_offset : int, default 0
            Vertical offset in cells (0-based)
        default_number_format : str, optional
            Default Excel number format string for numeric cells
        number_formats : dict, optional
            Specific number formats for rows/columns
        border_specs : dict or list or str, optional
            Labels where borders should be drawn
        """
        self.df = df
        self.worksheet = worksheet
        self.layout = TableLayout.from_df(df, x_offset, y_offset)
        self.default_number_format = default_number_format

        self.style_manager = StyleManager(default_styles={
            'index': self.INDEX_STYLE,
            'column': self.COLUMN_STYLE,
            'index_names': self.INDEX_NAMES_STYLE,
            'column_names': self.COLUMN_NAMES_STYLE,
        })
        self.merge_manager = MergeManager(worksheet)
        self.border_manager = BorderManager(worksheet)

        # Store formatting and border configs for later processing
        self.number_formats = number_formats or {}
        self.border_specs = border_specs

        # Initialize format specs (will be processed in a separate method)
        self.row_formats: list[str | None] = [None] * self.layout.index.height
        self.column_formats: list[str | None] = [None] * self.layout.columns.width
        self.row_borders: list[bool] = [False] * self.layout.index.height
        self.column_borders: list[bool] = [False] * self.layout.columns.width

        # Process number formats
        self._process_number_formats()
        self._process_border_specs()

    def _process_number_formats(self) -> None:
        """Pre-process number formats into position-specific formats."""
        # Skip if no formats provided
        if not self.number_formats:
            return

        # Handle different specification types
        if isinstance(self.number_formats, dict):
            # Check if it has rows/columns keys
            if 'rows' in self.number_formats or 'columns' in self.number_formats:
                # Process rows
                if 'rows' in self.number_formats:
                    row_specs = self.number_formats['rows']
                    if isinstance(row_specs, dict):
                        row_patterns = list(row_specs.items())
                    else:
                        row_patterns = [row_specs] if not isinstance(row_specs, list) else row_specs

                    # Apply to rows
                    for i, row_idx in enumerate(self.df.index):
                        match = PatternMatcher.find_match(row_idx, row_patterns)
                        if match is not None:
                            self.row_formats[i] = match

                # Process columns
                if 'columns' in self.number_formats:
                    col_specs = self.number_formats['columns']
                    if isinstance(col_specs, dict):
                        column_patterns = list(col_specs.items())
                    else:
                        column_patterns = [col_specs] if not isinstance(col_specs, list) else col_specs

                    # Apply to columns
                    for j, col_name in enumerate(self.df.columns):
                        match = PatternMatcher.find_match(col_name, column_patterns)
                        if match is not None:
                            self.column_formats[j] = match
            else:
                # Apply to both axes
                patterns = list(self.number_formats.items())

                # Apply to rows
                for i, row_idx in enumerate(self.df.index):
                    match = PatternMatcher.find_match(row_idx, patterns)
                    if match is not None:
                        self.row_formats[i] = match

                # Apply to columns
                for j, col_name in enumerate(self.df.columns):
                    match = PatternMatcher.find_match(col_name, patterns)
                    if match is not None:
                        self.column_formats[j] = match
        else:
            # Not a dict - interpret as pattern with default format
            pattern = self.number_formats
            patterns = [(pattern, self.default_number_format)]

            # Apply to rows
            for i, row_idx in enumerate(self.df.index):
                if PatternMatcher.is_match(row_idx, pattern):
                    self.row_formats[i] = self.default_number_format

            # Apply to columns
            for j, col_name in enumerate(self.df.columns):
                if PatternMatcher.is_match(col_name, pattern):
                    self.column_formats[j] = self.default_number_format

    def _process_border_specs(self) -> None:
        """Pre-process border specifications into position-specific flags."""
        # Skip if no border specs provided
        if not self.border_specs:
            return

        # Handle different specification types
        if isinstance(self.border_specs, dict):
            # Check if it has rows/columns keys
            if 'rows' in self.border_specs or 'columns' in self.border_specs:
                # Process rows
                if 'rows' in self.border_specs:
                    row_specs = self.border_specs['rows']
                    row_patterns = [row_specs] if not isinstance(row_specs, list) else row_specs

                    # Apply to rows
                    for i, row_idx in enumerate(self.df.index):
                        for pattern in row_patterns:
                            if PatternMatcher.is_match(row_idx, pattern):
                                self.row_borders[i] = True
                                break

                # Process columns
                if 'columns' in self.border_specs:
                    col_specs = self.border_specs['columns']
                    column_patterns = [col_specs] if not isinstance(col_specs, list) else col_specs

                    # Apply to columns
                    for j, col_name in enumerate(self.df.columns):
                        for pattern in column_patterns:
                            if PatternMatcher.is_match(col_name, pattern):
                                self.column_borders[j] = True
                                break
            else:
                # Apply to both axes - keys are the patterns
                patterns = list(self.border_specs.keys())

                # Apply to rows
                for i, row_idx in enumerate(self.df.index):
                    for pattern in patterns:
                        if PatternMatcher.is_match(row_idx, pattern):
                            self.row_borders[i] = True
                            break

                # Apply to columns
                for j, col_name in enumerate(self.df.columns):
                    for pattern in patterns:
                        if PatternMatcher.is_match(col_name, pattern):
                            self.column_borders[j] = True
                            break
        else:
            # Not a dict - interpret as a pattern
            pattern = self.border_specs
            patterns = [pattern] if not isinstance(pattern, list) else pattern

            # Apply to rows
            for i, row_idx in enumerate(self.df.index):
                for p in patterns:
                    if PatternMatcher.is_match(row_idx, p):
                        self.row_borders[i] = True
                        break

            # Apply to columns
            for j, col_name in enumerate(self.df.columns):
                for p in patterns:
                    if PatternMatcher.is_match(col_name, p):
                        self.column_borders[j] = True
                        break

    def write_all(self) -> None:
        """Write all DataFrame components to the worksheet."""
        self.write_column_headers()
        self.write_index_values()

        if self.layout.has_column_names:
            self.write_column_names()

        if self.layout.has_index_names:
            self.write_index_names()

        self.write_data()

        # Calculate merge ranges and apply them
        merge_ranges = spans.get_all_merge_ranges(self.df, self.layout)
        self.merge_manager.apply_merges(merge_ranges)

        # Calculate spans for level borders
        row_spans = spans.get_level_spans(self.df.index)
        column_spans = spans.get_level_spans(self.df.columns)

        # Add various types of borders separately
        self.border_manager.add_vertical_index_border(self.layout)
        self.border_manager.add_horizontal_header_border(self.layout)
        self.border_manager.add_level_borders(
            self.layout,
            row_spans,
            column_spans,
            min_border_level=1,
        )
        self.border_manager.add_custom_borders(
            self.layout,
            self.row_borders,
            self.column_borders
        )

    def _write_cell(
        self,
        position: CellPosition,
        value: Any,
        style: dict | None = None,
        number_format: str | None = None
    ) -> Cell:
        """
        Write a value to a cell with proper formatting.

        Parameters
        ----------
        position : CellPosition
            The position where to write the cell
        value : Any
            The value to write
        style : dict, optional
            Style dictionary to apply
        number_format : str, optional
            Excel number format string
        """
        # Handle NA values
        value = self._handle_na_value(value)

        # Get Excel coordinates
        row, col = position.excel_position

        # Write to worksheet
        # Cast type to Cell because we can assume that this is not a merged cell
        # (merging happens after writing the data)
        cell = cast(Cell, self.worksheet.cell(row=row, column=col, value=value))

        # Apply style if provided
        if style:
            self.style_manager.apply_style(cell, style)

        # Apply number format if it's a number and format is provided
        if number_format and isinstance(value, (int, float, np.number)):
            cell.number_format = number_format

        return cell

    def _handle_na_value(self, value) -> Any:
        """Convert NA/NaN values to a representation Excel can handle."""
        if pd.isna(value):
            return self.NA_REPRESENTATION
        return value

    # Placeholder methods that will be implemented next
    def write_data(self) -> None:
        """Write the actual data values to the worksheet."""
        for i, j, position in self.layout.iter_data_positions():
            value = self.df.iloc[i, j]

            # Determine the number format to use
            col_format = self.column_formats[j]
            row_format = self.row_formats[i]
            # Priority: Column format > Row format > Default format
            number_format = row_format or col_format or self.default_number_format

            self._write_cell(
                position = position,
                value = value,
                number_format = number_format
            )

    def write_column_headers(self) -> None:
        """Write column headers with proper formatting."""
        for j, level, position in self.layout.iter_column_positions():
            if isinstance(self.df.columns[j], tuple):
                value = self.df.columns[j][level]
            else:
                # For single-level columns, only write at level 0
                if level > 0:
                    continue
                value = self.df.columns[j]

            self._write_cell(
                position=position,
                value=value,
                style=self.COLUMN_STYLE
            )

    def write_index_values(self) -> None:
        """Write index values with proper formatting."""
        for i, level, position in self.layout.iter_index_positions():
            if isinstance(self.df.index[i], tuple):
                value = self.df.index[i][level]
            else:
                # For single-level indices, only write at level 0
                if level > 0:
                    continue
                value = self.df.index[i]

            self._write_cell(
                position=position,
                value=value,
                style=self.INDEX_STYLE
            )

    def write_column_names(self) -> None:
        """Write column names with proper formatting."""
        if not self.layout.has_column_names:
            return

        for level, position in self.layout.iter_column_name_positions():
            name = self.df.columns.names[level]
            if name is not None:
                self._write_cell(
                    position=position,
                    value=name,
                    style=self.COLUMN_NAMES_STYLE
                )

    def write_index_names(self) -> None:
        """Write index names with proper formatting."""
        if not self.layout.has_index_names:
            return

        for level, position in self.layout.iter_index_name_positions():
            name = self.df.index.names[level]
            if name is not None:
                self._write_cell(
                    position=position,
                    value=name,
                    style=self.INDEX_NAMES_STYLE
                )
