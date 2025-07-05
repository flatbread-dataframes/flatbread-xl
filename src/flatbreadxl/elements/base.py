# uvmlib/export/elements/base.py
from typing import Any, cast

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell

from flatbreadxl.layout import CellPosition, BaseLayout
from flatbreadxl.style import StyleManager


class WorksheetElement:
    """
    Base class for all worksheet elements.
    Uses layout classes for positioning and adds rendering capabilities.
    """
    def __init__(
        self,
        x_offset: int = 0,
        y_offset: int = 0,
    ) -> None:
        """
        Initialize element with position information.

        Parameters
        ----------
        x_offset : int, default 0
            X coordinate (column) where element starts
        y_offset : int, default 0
            Y coordinate (row) where element starts
        """
        self.position = CellPosition(x_offset, y_offset)
        self._style_manager = StyleManager()

    @property
    def x_start(self) -> int:
        """Starting X coordinate (column)."""
        return self.position.x

    @property
    def y_start(self) -> int:
        """Starting Y coordinate (row)."""
        return self.position.y

    @property
    def width(self) -> int:
        """Element width in cells."""
        raise NotImplementedError("Subclasses must implement width property")

    @property
    def height(self) -> int:
        """Element height in cells."""
        raise NotImplementedError("Subclasses must implement height property")

    @property
    def x_end(self) -> int:
        """Ending X coordinate (column)."""
        return self.x_start + self.width - 1

    @property
    def y_end(self) -> int:
        """Ending Y coordinate (row)."""
        return self.y_start + self.height - 1

    def render(self, worksheet: Worksheet) -> None:
        """
        Render the element to the worksheet.
        Must be implemented by subclasses.
        """
        raise NotImplementedError("Subclasses must implement render()")

    def get_position_below(self, spacing: int = 1) -> CellPosition:
        """Get position for an element directly below this one."""
        return CellPosition(self.x_start, self.y_end + spacing)

    def get_position_right(self, spacing: int = 1) -> CellPosition:
        """Get position for an element directly to the right of this one."""
        return CellPosition(self.x_end + spacing, self.y_start)

    def _write_cell(
        self,
        worksheet: Worksheet,
        position: CellPosition,
        value: Any,
        style: dict | None = None
    ) -> Cell:
        """Write a value to a cell with styling."""
        cell = cast(
            Cell,
            worksheet.cell(
                row = position.excel_y,
                column = position.excel_x,
                value = value
            )
        )

        if style:
            self._style_manager.apply_style(cell, style)

        return cell
