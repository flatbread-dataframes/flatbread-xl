from enum import Enum
from typing import Any

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.cell import Cell


class HorizontalAlignment(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class VerticalAlignment(Enum):
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"


class BorderStyle(Enum):
    THIN = "thin"
    MEDIUM = "medium"
    THICK = "thick"
    DASHED = "dashed"
    DOTTED = "dotted"
    DOUBLE = "double"
    NONE = None


class StyleManager:
    """
    Manages the application of styles to Excel cells.
    Separates styling logic from the table writing logic.
    """

    def __init__(self, default_styles: dict[str, dict[str, Any]] | None = None):
        """
        Initialize with optional default styles for different cell types.

        Parameters
        ----------
        default_styles : dict[str, dict[str, Any]], optional
            Default styles for different cell types (e.g., 'header', 'data', 'index')
        """
        self.default_styles = default_styles or {}

    def apply_style(self, cell: Cell, style: dict[str, Any]) -> None:
        """
        Apply a style dictionary to a cell, preserving existing properties
        when not explicitly overridden.

        Parameters
        ----------
        cell : Cell
            The openpyxl cell to style
        style : Dict[str, Any]
            Dictionary of style properties to apply
        """
        self._apply_font(cell, style)
        self._apply_alignment(cell, style)
        self._apply_fill(cell, style)
        self._apply_border(cell, style)
        self._apply_number_format(cell, style)

    def _apply_font(self, cell: Cell, style: dict[str, Any]) -> None:
        """Apply font styling to a cell."""
        # Get current font properties
        current_font = cell.font
        font_kwargs = {
            "name": current_font.name,
            "size": current_font.size,
            "bold": current_font.bold,
            "italic": current_font.italic,
            "color": current_font.color,
            "underline": current_font.underline,
            "strike": current_font.strike,
            "vertAlign": current_font.vertAlign,
        }

        # Update with new properties from style
        if style.get("bold") is not None:
            font_kwargs["bold"] = style["bold"]
        if style.get("italic") is not None:
            font_kwargs["italic"] = style["italic"]
        if "font_size" in style:
            font_kwargs["size"] = style["font_size"]
        if "font_color" in style:
            font_kwargs["color"] = style["font_color"]
        if "underline" in style:
            font_kwargs["underline"] = style["underline"]

        # Apply updated font
        cell.font = Font(**font_kwargs)

    def _apply_alignment(self, cell: Cell, style: dict[str, Any]) -> None:
        """Apply alignment styling to a cell."""
        # Get current alignment properties
        current_alignment = cell.alignment
        alignment_kwargs = {
            "horizontal": current_alignment.horizontal,
            "vertical": current_alignment.vertical,
            "textRotation": current_alignment.textRotation,
            "wrapText": current_alignment.wrapText,
            "shrinkToFit": current_alignment.shrinkToFit,
            "indent": current_alignment.indent,
        }

        # Update with new properties
        horiz = style.get("horizontal_alignment")
        if horiz is not None:
            alignment_kwargs["horizontal"] = horiz.value

        vert = style.get("vertical_alignment")
        if vert is not None:
            alignment_kwargs["vertical"] = vert.value

        if "text_rotation" in style:
            alignment_kwargs["textRotation"] = style["text_rotation"]

        if "wrap_text" in style:
            alignment_kwargs["wrapText"] = style["wrap_text"]

        # Apply updated alignment
        cell.alignment = Alignment(**alignment_kwargs)

    def _apply_fill(self, cell: Cell, style: dict[str, Any]) -> None:
        """Apply fill styling to a cell."""
        if "fill_color" in style:
            cell.fill = PatternFill(
                start_color=style["fill_color"],
                end_color=style["fill_color"],
                fill_type="solid"
            )

    def _apply_border(self, cell: Cell, style: dict[str, Any]) -> None:
        """Apply border styling to a cell."""
        # Get current border properties
        current_border = cell.border
        border_kwargs = {
            "left": current_border.left,
            "right": current_border.right,
            "top": current_border.top,
            "bottom": current_border.bottom,
        }

        # Update with new properties
        for side in ["left", "right", "top", "bottom"]:
            border_key = f"{side}_border"
            if border_key in style:
                border_style = style[border_key]
                if isinstance(border_style, BorderStyle):
                    border_kwargs[side] = Side(style=border_style.value)
                elif isinstance(border_style, dict):
                    # Handle more complex border specifications
                    border_kwargs[side] = Side(**border_style)

        # Apply updated border
        cell.border = Border(**border_kwargs)

    def _apply_number_format(self, cell: Cell, style: dict[str, Any]) -> None:
        """Apply number format to a cell."""
        if "number_format" in style:
            cell.number_format = style["number_format"]
