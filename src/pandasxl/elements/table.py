import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from pandasxl.layout import TableLayout
from pandasxl.table.writer import ExcelTableWriter
from pandasxl.table.grouped import GroupedExcelTableWriter
from pandasxl.elements.base import WorksheetElement
from pandasxl.elements.text import TextElement


class TableElement(WorksheetElement):
    """
    Element that renders a DataFrame as a table in a worksheet.
    Can include an optional title and caption.
    """
    def __init__(
        self,
        df: pd.DataFrame,
        x_offset: int = 0,
        y_offset: int = 0,
        title: str | None = None,
        caption: str | None = None,
        **table_kwargs
    ) -> None:
        """
        Initialize a table element.

        Parameters
        ----------
        df : pd.DataFrame
            The data to display
        x_offset : int, default 0
            X coordinate (column) where element starts
        y_offset : int, default 0
            Y coordinate (row) where element starts
        title : str, optional
            Optional title to display above the table
        caption : str, optional
            Optional caption to display below the table
        **table_kwargs
            Additional keyword arguments for ExcelTableWriter
        """
        super().__init__(x_offset, y_offset)
        self.df = df
        self.title = title
        self.caption = caption
        self.table_kwargs = table_kwargs

        # Calculate layout
        self._calculate_layout()

    def _calculate_layout(self) -> None:
        """Calculate the layout of the table and any title/caption."""
        current_y = self.y_start

        # Account for title if present
        if self.title:
            self.title_element = TextElement(
                text         = self.title,
                x_offset     = self.x_start,
                y_offset     = current_y,
                style_preset = "subtitle"
            )
            current_y += self.title_element.height
        else:
            self.title_element = None

        # Create table layout
        self.table_layout = TableLayout.from_df(
            self.df,
            x_offset=self.x_start,
            y_offset=current_y
        )
        current_y += self.table_layout.total_height

        # Account for caption if present
        if self.caption:
            self.caption_element = TextElement(
                text         = self.caption,
                x_offset     = self.x_start,
                y_offset     = current_y,
                style_preset = "caption"
            )
            current_y += self.caption_element.height
        else:
            self.caption_element = None

        # Store total dimensions
        self._calculated_height = current_y - self.y_start
        self._calculated_width = self.table_layout.total_width

    @property
    def width(self) -> int:
        """Total width of the table element."""
        return self._calculated_width

    @property
    def height(self) -> int:
        """Total height of the table element."""
        return self._calculated_height

    def render(self, worksheet: Worksheet) -> None:
        """Render the table element to the worksheet."""
        # Render title if present
        if self.title_element:
            self.title_element.render(worksheet)

        # Render table
        writer = ExcelTableWriter(
            df        = self.df,
            worksheet = worksheet,
            x_offset  = self.table_layout.position.x,
            y_offset  = self.table_layout.position.y,
            **self.table_kwargs
        )
        writer.write_all()

        # Render caption if present
        if self.caption_element:
            self.caption_element.render(worksheet)


class GroupedTableElement(TableElement):
    """
    Element that renders a DataFrame as a grouped table in a worksheet.
    Can include an optional title and caption.
    """
    def __init__(
        self,
        df: pd.DataFrame,
        x_offset: int = 0,
        y_offset: int = 0,
        title: str | None = None,
        caption: str | None = None,
        group_levels: int | list[int] = 0,
        **table_kwargs
    ) -> None:
        """
        Initialize a grouped table element.

        Parameters
        ----------
        df : pd.DataFrame
            The data to display
        x_offset : int, default 0
            X coordinate (column) where element starts
        y_offset : int, default 0
            Y coordinate (row) where element starts
        title : str, optional
            Optional title to display above the table
        caption : str, optional
            Optional caption to display below the table
        group_levels : int | list[int], default 0
            Index levels to group by
        **table_kwargs
            Additional keyword arguments for ExcelTableWriter
        """
        super().__init__(df, x_offset, y_offset, title, caption, **table_kwargs)
        self.group_levels = group_levels

    def render(self, worksheet: Worksheet) -> None:
        """Render the grouped table element to the worksheet."""
        # Render title if present
        if self.title_element:
            self.title_element.render(worksheet)

        # Render grouped table
        writer = GroupedExcelTableWriter(
            df           = self.df,
            worksheet    = worksheet,
            x_offset     = self.table_layout.position.x,
            y_offset     = self.table_layout.position.y,
            group_levels = self.group_levels,
            **self.table_kwargs
        )
        writer.write_all()

        # Render caption if present
        if self.caption_element:
            self.caption_element.render(worksheet)
