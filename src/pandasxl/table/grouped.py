# uvmlib/export/formatter/grouped.py
from typing import Any

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from pandasxl.table.writer import ExcelTableWriter
from pandasxl.style import HorizontalAlignment, VerticalAlignment


class GroupHeaderValue:
    """A wrapper for group header values that displays the same but compares differently."""

    def __init__(self, value):
        self.value = value

    def __str__(self):
        return str(self.value)

    def __repr__(self):
        return f"GroupHeaderValue({repr(self.value)})"

    # This ensures it's never equal to the original value
    def __eq__(self, other):
        if isinstance(other, GroupHeaderValue):
            return self.value == other.value
        return False

    # Make the class hashable for pandas operations
    def __hash__(self):
        return hash(self.value)

    def __lt__(self, other):
        if isinstance(other, GroupHeaderValue):
            return self.value < other.value
        return self.value < other

    def __le__(self, other):
        if isinstance(other, GroupHeaderValue):
            return self.value <= other.value
        return self.value <= other

    def __gt__(self, other):
        if isinstance(other, GroupHeaderValue):
            return self.value > other.value
        return self.value > other

    def __ge__(self, other):
        if isinstance(other, GroupHeaderValue):
            return self.value >= other.value
        return self.value >= other


def add_group_headers(
    df: pd.DataFrame,
    group_levels: int | list[int],
) -> tuple[pd.DataFrame, pd.Series]:
    """
    Add group header rows to a DataFrame.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with MultiIndex to transform
    group_levels : int | list[int]
        List of index levels to use for grouping

    Returns
    -------
    tuple[pd.DataFrame, pd.Series]
        Transformed DataFrame and marker series indicating header rows
    """
    if not isinstance(df.index, pd.MultiIndex):
        raise ValueError("DataFrame must have a MultiIndex for grouping")

    group_levels = [group_levels] if isinstance(group_levels, int) else group_levels

    # Sort levels to process hierarchically
    group_levels = sorted(group_levels)

    # Make sure we're not trying to group by all levels (need at least one data level)
    if len(group_levels) >= df.index.nlevels:
        group_levels = group_levels[:-1]

    # Add marker column
    result_df = df.copy()
    result_df['__group_header__'] = False

    # Apply grouping recursively starting from highest level
    for level in group_levels:
        result_df = result_df.groupby(
            level=level,
            group_keys=False,
            sort=False
        ).apply(_stack_group)

    # Extract marker column
    marker_column = result_df['__group_header__']

    # Remove marker column from result
    result_df = result_df.drop(columns=['__group_header__'])

    return result_df, marker_column


def _stack_group(group: pd.DataFrame) -> pd.DataFrame:
    """
    Add a header row to a group.

    Parameters
    ----------
    group : pd.DataFrame
        Group to add header to

    Returns
    -------
    pd.DataFrame
        Group with header row added
    """
    if '__group_header__' in group.columns and group['__group_header__'].any():
        return group.droplevel(-1)

    # Create header row data
    data: dict[str | tuple[str, ...], Any] = {col: pd.NA for col in group.columns}

    # Set marker column to True
    if isinstance(group.columns, pd.MultiIndex):
        nlev = group.columns.nlevels - 1
        col_key = tuple(['__group_header__'] + [''] * nlev)
        data[col_key] = True
    else:
        data['__group_header__'] = True

    # Create appropriate index for header row
    if isinstance(group.index, pd.MultiIndex):
        nlev = group.index.nlevels - 2
        if nlev <= 0:
            # Wrap the group name in GroupHeaderValue
            idx = [GroupHeaderValue(group.name)]
        else:
            # Wrap the first element in the tuple with GroupHeaderValue
            key = tuple([GroupHeaderValue(group.name)] + [pd.NA] * nlev)
            names = group.index.names[1:]
            idx = pd.MultiIndex.from_tuples([key], names=names)
    else:
        # Wrap the group name in GroupHeaderValue
        idx = [GroupHeaderValue(group.name)]

    header_row = pd.DataFrame(data, index=idx)

    # Combine header row with group data
    return pd.concat([header_row, group.droplevel(0)])


class GroupedExcelTableWriter(ExcelTableWriter):
    """
    Excel writer that supports grouped data with header rows.
    Extends the ExcelTableWriter with special handling for group headers.
    """
    # Default style for group headers
    GROUP_HEADER_STYLE = {
        "bold": True,
        "horizontal_alignment": HorizontalAlignment.LEFT,
        "vertical_alignment": VerticalAlignment.BOTTOM,
        "font_size": 13,
    }

    # Default height for group header rows
    DEFAULT_HEADER_ROW_HEIGHT = 32

    def __init__(
        self,
        df: pd.DataFrame,
        worksheet: Worksheet,
        x_offset: int = 0,
        y_offset: int = 0,
        default_number_format: str | None = None,
        number_formats: dict | None = None,
        border_specs: Any | None = None,
        group_levels: int | list[int] = 0,
    ):
        """
        Initialize the grouped Excel writer.

        Parameters
        ----------
        df : pd.DataFrame
            DataFrame to write
        worksheet : Worksheet
            Worksheet to write to
        x_offset : int, default 0
            X offset to start writing at
        y_offset : int, default 0
            Y offset to start writing at
        default_number_format : str, optional
            Default number format for numeric cells
        number_formats : dict, optional
            Number formats for specific patterns
        border_specs : Any, optional
            Border specifications
        group_levels : int or list[int], default [0]
            Index levels to group by
        """
        # Store original DataFrame
        self.original_df = df

        # Add group headers to DataFrame
        grouped_df, self.marker_column = add_group_headers(df, group_levels)

        # Initialize parent class with transformed DataFrame
        super().__init__(
            df                    = grouped_df,
            worksheet             = worksheet,
            x_offset              = x_offset,
            y_offset              = y_offset,
            default_number_format = default_number_format,
            number_formats        = number_formats,
            border_specs          = border_specs,
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

            # Extract original value if it's a GroupHeaderValue
            if isinstance(value, GroupHeaderValue):
                value = value.value
                is_header = True
            else:
                is_header = False

            # Skip NA values
            if pd.isna(value):
                continue

            # Use header style for header rows
            style = self.GROUP_HEADER_STYLE if is_header else self.INDEX_STYLE

            self._write_cell(
                position=position,
                value=value,
                style=style
            )

    def write_all(self) -> None:
        """Override to apply row heights after writing all data."""
        # Call the parent method to write all data
        super().write_all()

        # Use marker_column to identify header rows and apply heights
        for i, is_header in enumerate(self.marker_column):
            if is_header:
                # Convert DataFrame index position to Excel row
                h = self.DEFAULT_HEADER_ROW_HEIGHT
                excel_row = self.layout.index.cell_at(0, i).excel_y
                self.worksheet.row_dimensions[excel_row].height = h
