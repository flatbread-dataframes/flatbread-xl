from typing import Any, Tuple, Iterator

import pandas as pd


# region Cell
class CellPosition:
    """
    Represents a cell position with conversions between 0-based (Python/pandas)
    and 1-based (Excel/OpenPyXL) indexing.
    """
    def __init__(self, x: int, y: int, excel_based: bool = False) -> None:
        """
        Initialize a cell position.

        Parameters
        ----------
        x : int
            X-coordinate (column)
        y : int
            Y-coordinate (row)
        excel_based : bool, default False
            If True, x and y are treated as 1-based Excel coordinates.
            If False, x and y are treated as 0-based Python coordinates.
        """
        if excel_based:
            self.x = x - 1  # Convert to 0-based
            self.y = y - 1  # Convert to 0-based
        else:
            self.x = x
            self.y = y

    @property
    def excel_x(self) -> int:
        """Get the Excel/OpenPyXL column coordinate (1-based)."""
        return self.x + 1

    @property
    def excel_y(self) -> int:
        """Get the Excel/OpenPyXL row coordinate (1-based)."""
        return self.y + 1

    @property
    def excel_position(self) -> tuple[int, int]:
        """Get the Excel/OpenPyXL position as (row, column) tuple."""
        return (
            self.excel_y,
            self.excel_x,
        )

    def __add__(self, other: 'CellPosition') -> 'CellPosition':
        """Add two positions together."""
        return CellPosition(
            self.x + other.x,
            self.y + other.y,
        )

    def __sub__(self, other: 'CellPosition') -> 'CellPosition':
        """Subtract one position from another."""
        return CellPosition(
            self.x - other.x,
            self.y - other.y,
        )

    def offset(self, dx: int = 0, dy: int = 0) -> 'CellPosition':
        """Create a new position offset from this one."""
        return CellPosition(
            self.x + dx,
            self.y + dy,
        )

    def __eq__(self, other) -> bool:
        if not isinstance(other, CellPosition):
            return False
        return self.x == other.x and self.y == other.y

    def __repr__(self) -> str:
        return f"CellPosition(x={self.x}, y={self.y}, excel=({self.excel_y}, {self.excel_x}))"


# region Base
class BaseLayout:
    """Base class for all layout components."""

    def __init__(
        self,
        width: int,
        height: int,
        x_offset: int,
        y_offset: int,
    ) -> None:
        self.width = width
        self.height = height
        self.position = CellPosition(x_offset, y_offset)

    @property
    def x_start(self) -> int:
        return self.position.x

    @property
    def x_end(self) -> int:
        return self.x_start + self.width - 1

    @property
    def y_start(self) -> int:
        return self.position.y

    @property
    def y_end(self) -> int:
        return self.y_start + self.height - 1

    @property
    def excel_x_start(self) -> int:
        """Get the starting Excel column (1-based)."""
        return self.position.excel_x

    @property
    def excel_y_start(self) -> int:
        """Get the starting Excel row (1-based)."""
        return self.position.excel_y

    @property
    def excel_x_end(self) -> int:
        """Get the ending Excel column (1-based)."""
        return self.x_end + 1

    @property
    def excel_y_end(self) -> int:
        """Get the ending Excel row (1-based)."""
        return self.y_end + 1

    def cell_at(self, x_offset: int = 0, y_offset: int = 0) -> CellPosition:
        """Get the position of a cell relative to the layout's start position."""
        return CellPosition(
            self.x_start + x_offset,
            self.y_start + y_offset,
        )

    def excel_cell_at(self, x_offset: int = 0, y_offset: int = 0) -> tuple[int, int]:
        """Get the Excel position (row, col) of a cell relative to the layout's start."""
        pos = self.cell_at(x_offset, y_offset)
        return pos.excel_position

    def iter_positions(self) -> Iterator[CellPosition]:
        """Iterate through all cell positions in this layout."""
        for y in range(self.height):
            for x in range(self.width):
                yield self.cell_at(x, y)

    def iter_rows(self) -> Iterator[list[CellPosition]]:
        """Iterate through rows of positions."""
        for y in range(self.height):
            row = [self.cell_at(x, y) for x in range(self.width)]
            yield row

    def iter_columns(self) -> Iterator[list[CellPosition]]:
        """Iterate through columns of positions."""
        for x in range(self.width):
            column = [self.cell_at(x, y) for y in range(self.height)]
            yield column


# region Axes
class AxisLayout(BaseLayout):
    def __init__(
        self,
        width: int,
        height: int,
        has_names: bool,
        x_offset: int,
        y_offset: int,
    ) -> None:
        super().__init__(width, height, x_offset, y_offset)
        self.has_names = has_names


class NamesLayout(BaseLayout):
    pass


class DataLayout(BaseLayout):
    pass


# region Table
class TableLayout:
    def __init__(
        self,
        index_width: int,
        index_height: int,
        column_width: int,
        column_height: int,
        has_index_names: bool = False,
        has_column_names: bool = False,
        x_offset: int = 0,
        y_offset: int = 0,
    ) -> None:
        self.position         = CellPosition(x_offset, y_offset)
        self.index_width      = index_width
        self.index_height     = index_height
        self.column_width     = column_width
        self.column_height    = column_height
        self.has_index_names  = has_index_names
        self.has_column_names = has_column_names
        self._init_layouts()

    def _init_layouts(self) -> None:
        # Initialize layouts in dependency order to avoid circular references:
        # 1. First the axis layouts which depend only on dimensions
        self.columns = self._create_column_layout()
        self.index = self._create_index_layout()

        # 2. Then names layouts which depend on axis layouts
        self.index_names = self._create_index_names_layout()
        self.column_names = self._create_column_names_layout()

        # 3. Finally data layout which depends on all other layouts
        self.data_layout = self._create_data_layout()

    def _create_column_layout(self) -> AxisLayout:
        return AxisLayout(
            width     = self.column_width,
            height    = self.column_height,
            has_names = self.has_column_names,
            x_offset  = self.position.x + self.index_width,
            y_offset  = self.position.y,
        )

    def _create_index_layout(self) -> AxisLayout:
        # Index layout starts one row down if there are index names
        additional_y_offset = 1 if self.has_index_names else 0

        return AxisLayout(
            width     = self.index_width,
            height    = self.index_height,
            has_names = self.has_index_names,
            x_offset  = self.position.x,
            y_offset  = self.position.y + self.column_height + additional_y_offset,
        )

    def _create_data_layout(self) -> DataLayout:
        return DataLayout(
            width    = self.columns.width,
            height   = self.index.height,
            x_offset = self.columns.x_start,
            y_offset = self.index.y_start,
        )

    def _create_index_names_layout(self) -> NamesLayout:
        # Index names are positioned above the index
        return NamesLayout(
            width    = self.index.width,
            height   = 1 if self.has_index_names else 0,
            x_offset = self.index.x_start,
            y_offset = self.index.y_start - (1 if self.has_index_names else 0),
        )

    def _create_column_names_layout(self) -> NamesLayout:
        # Column names are positioned to the left of the columns
        return NamesLayout(
            width    = 1 if self.has_column_names else 0,
            height   = self.columns.height,
            x_offset = self.columns.x_start - (1 if self.has_column_names else 0),
            y_offset = self.columns.y_start,
        )

    @property
    def total_width(self) -> int:
        return self.index.width + self.data_layout.width

    @property
    def total_height(self) -> int:
        return (
            + self.columns.height
            + self.index_names.height
            + self.index.height
        )

    @property
    def x_start(self) -> int:
        # Leftmost position of the table
        return self.position.x

    @property
    def y_start(self) -> int:
        # Topmost position of the table
        return self.position.y

    @property
    def x_end(self) -> int:
        # Rightmost position of the table
        return self.data_layout.x_end

    @property
    def y_end(self) -> int:
        # Bottommost position of the table
        return self.data_layout.y_end

    @property
    def excel_x_start(self) -> int:
        """Get the starting Excel column (1-based)."""
        return self.position.excel_x

    @property
    def excel_y_start(self) -> int:
        """Get the starting Excel row (1-based)."""
        return self.position.excel_y

    @property
    def excel_x_end(self) -> int:
        """Get the ending Excel column (1-based)."""
        return self.x_end + 1

    @property
    def excel_y_end(self) -> int:
        """Get the ending Excel row (1-based)."""
        return self.y_end + 1

    def cell_at(self, x_offset: int = 0, y_offset: int = 0) -> CellPosition:
        """Get the position of a cell relative to the table's start position."""
        return CellPosition(
            self.x_start + x_offset,
            self.y_start + y_offset,
        )

    def excel_cell_at(self, x_offset: int = 0, y_offset: int = 0) -> tuple[int, int]:
        """Get the Excel position (row, col) of a cell relative to the table's start."""
        pos = self.cell_at(x_offset, y_offset)
        return pos.excel_position

    def iter_data_positions(self) -> Iterator[tuple[int, int, CellPosition]]:
        """
        Iterate through data cell positions, yielding (row_idx, col_idx, position) tuples.
        row_idx and col_idx are 0-based indices.
        """
        for i in range(self.index.height):
            for j in range(self.columns.width):
                pos = CellPosition(
                    self.data_layout.x_start + j,
                    self.data_layout.y_start + i,
                )
                yield i, j, pos

    def iter_index_positions(self) -> Iterator[tuple[int, int, CellPosition]]:
        """
        Iterate through index cell positions, yielding (row_idx, level, position) tuples.
        """
        for i in range(self.index.height):
            for level in range(self.index.width):
                pos = CellPosition(
                    self.index.x_start + level,
                    self.index.y_start + i,
                )
                yield i, level, pos

    def iter_column_positions(self) -> Iterator[tuple[int, int, CellPosition]]:
        """
        Iterate through column header positions, yielding (col_idx, level, position) tuples.
        """
        for j in range(self.columns.width):
            for level in range(self.columns.height):
                pos = CellPosition(
                    self.columns.x_start + j,
                    self.columns.y_start + level,
                )
                yield j, level, pos

    def iter_index_name_positions(self) -> Iterator[tuple[int, CellPosition]]:
        """
        Iterate through index name positions if they exist.
        """
        if not self.has_index_names:
            return

        for level in range(self.index_names.width):
            pos = CellPosition(
                self.index_names.x_start + level,
                self.index_names.y_start,
            )
            yield level, pos

    def iter_column_name_positions(self) -> Iterator[tuple[int, CellPosition]]:
        """
        Iterate through column name positions if they exist.
        """
        if not self.has_column_names:
            return

        for level in range(self.column_names.height):
            pos = CellPosition(
                self.column_names.x_start,
                self.column_names.y_start + level,
            )
            yield level, pos

    def get_data_range(self) -> tuple[CellPosition, CellPosition]:
        """Get the top-left and bottom-right positions of the data area."""
        top_left = CellPosition(self.data_layout.x_start, self.data_layout.y_start)
        bottom_right = CellPosition(self.data_layout.x_end, self.data_layout.y_end)
        return top_left, bottom_right

    def get_index_range(self) -> tuple[CellPosition, CellPosition]:
        """Get the top-left and bottom-right positions of the index area."""
        top_left = CellPosition(self.index.x_start, self.index.y_start)
        bottom_right = CellPosition(self.index.x_end, self.index.y_end)
        return top_left, bottom_right

    def get_columns_range(self) -> tuple[CellPosition, CellPosition]:
        """Get the top-left and bottom-right positions of the columns area."""
        top_left = CellPosition(self.columns.x_start, self.columns.y_start)
        bottom_right = CellPosition(self.columns.x_end, self.columns.y_end)
        return top_left, bottom_right

    @classmethod
    def from_df(
        cls,
        df: pd.DataFrame,
        x_offset: int = 0,
        y_offset: int = 0,
    ) -> 'TableLayout':
        """Create a TableLayout from a DataFrame."""
        has_index_names = any(name is not None for name in df.index.names)
        has_column_names = any(name is not None for name in df.columns.names)

        return cls(
            index_width      = df.index.nlevels,
            index_height     = len(df),
            column_width     = len(df.columns),
            column_height    = df.columns.nlevels,
            has_index_names  = has_index_names,
            has_column_names = has_column_names,
            x_offset         = x_offset,
            y_offset         = y_offset,
        )
