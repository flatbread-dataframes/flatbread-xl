# uvmlib/export/elements/text.py
from typing import Optional, Dict, Any

from openpyxl.worksheet.worksheet import Worksheet

from pandasxl.layout import CellPosition
from pandasxl.style import HorizontalAlignment, VerticalAlignment
from pandasxl.elements.base import WorksheetElement


class TextElement(WorksheetElement):
    """
    Element for rendering text (titles, captions, etc.) in a worksheet.
    """
    # Default styles
    DEFAULT_STYLE = {
        "horizontal_alignment": HorizontalAlignment.LEFT,
        "vertical_alignment": VerticalAlignment.CENTER,
    }

    TITLE_STYLE = {
        "bold": True,
        "font_size": 16,
        "horizontal_alignment": HorizontalAlignment.LEFT,
    }

    SUBTITLE_STYLE = {
        "bold": True,
        "font_size": 14,
        "horizontal_alignment": HorizontalAlignment.LEFT,
    }

    CAPTION_STYLE = {
        "italic": True,
        "font_size": 11,
        "horizontal_alignment": HorizontalAlignment.LEFT,
    }

    def __init__(
        self,
        text: str,
        x_offset: int = 0,
        y_offset: int = 0,
        style: dict[str, Any] | None = None,
        style_preset: str = "default",
    ) -> None:
        """
        Initialize a text element.

        Parameters
        ----------
        text : str
            The text content to display
        x_offset : int, default 0
            X coordinate (column) where element starts
        y_offset : int, default 0
            Y coordinate (row) where element starts
        style : Dict[str, Any], optional
            Custom style to apply
        style_preset : str, default "default"
            Style preset to use ("default", "title", "subtitle", "caption")
        """
        super().__init__(x_offset, y_offset)
        self.text = text

        # Set style based on preset or custom style
        if style is not None:
            self.style = style
        else:
            self.style = self._get_preset_style(style_preset)

    def _get_preset_style(self, preset: str) -> Dict[str, Any]:
        """Get a style preset by name."""
        presets = {
            "default": self.DEFAULT_STYLE,
            "title": self.TITLE_STYLE,
            "subtitle": self.SUBTITLE_STYLE,
            "caption": self.CAPTION_STYLE,
        }
        return presets.get(preset.lower(), self.DEFAULT_STYLE)

    @property
    def width(self) -> int:
        """Text elements are assumed to span one column."""
        return 1

    @property
    def height(self) -> int:
        """Text elements are assumed to span one row."""
        return 1

    def render(self, worksheet: Worksheet) -> None:
        """Render the text element to the worksheet."""
        self._write_cell(
            worksheet,
            self.position,
            self.text,
            self.style
        )


class MultiColumnTextElement(TextElement):
    """
    Element for rendering text that spans multiple columns with text wrapping.
    """
    def __init__(
        self,
        text: str,
        width: int,
        x_offset: int = 0,
        y_offset: int = 0,
        style: dict[str, Any] | None = None,
        style_preset: str = "default",
        row_height: float | None = None,  # Optional explicit row height
    ) -> None:
        """
        Initialize a multi-column text element.

        Parameters
        ----------
        text : str
            The text content to display
        width : int
            Number of columns this text should span
        x_offset : int, default 0
            X coordinate (column) where element starts
        y_offset : int, default 0
            Y coordinate (row) where element starts
        style : dict[str, Any], optional
            Custom style to apply
        style_preset : str, default "default"
            Style preset to use ("default", "title", "subtitle", "caption")
        row_height : float, optional
            Explicit row height in points (if None, Excel will auto-adjust)
        """
        super().__init__(
            text=text,
            x_offset=x_offset,
            y_offset=y_offset,
            style=style,
            style_preset=style_preset
        )
        self._width = width
        self.row_height = row_height

        # Ensure wrap_text is enabled in the style
        if self.style is None:
            self.style = {}
        self.style["wrap_text"] = True

    @property
    def width(self) -> int:
        """Multi-column text elements span the specified width."""
        return self._width

    def render(self, worksheet: Worksheet) -> None:
        """Render the text element to the worksheet and merge cells if needed."""
        # Write the text to the first cell
        cell = self._write_cell(
            worksheet,
            self.position,
            self.text,
            self.style
        )

        # If spanning multiple columns, merge the cells
        if self.width > 1:
            start_cell = f"{cell.column_letter}{cell.row}"
            end_col_letter = worksheet.cell(
                row=cell.row,
                column=cell.column + self.width - 1
            ).column_letter
            end_cell = f"{end_col_letter}{cell.row}"

            worksheet.merge_cells(f"{start_cell}:{end_cell}")

        # Set custom row height if specified
        if self.row_height is not None:
            worksheet.row_dimensions[cell.row].height = self.row_height
