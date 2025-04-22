from typing import TypeAlias

from openpyxl.worksheet.worksheet import Worksheet

MergeRange: TypeAlias = tuple[int, int, int, int]


class MergeManager:
    """Handles applying cell merges in Excel worksheets."""

    def __init__(self, worksheet: Worksheet):
        """
        Initialize with the worksheet and optional pre-calculated merge ranges.

        Parameters
        ----------
        worksheet : Worksheet
            The openpyxl worksheet where merges will be applied
        merge_ranges : List[Tuple[int, int, int, int]], optional
            Pre-calculated merge ranges to apply
        """
        self.worksheet = worksheet
        self.applied_merges = []

    def apply_merges(self, merge_ranges: list[MergeRange]) -> None:
        """
        Apply the specified merge ranges to the worksheet.

        Parameters
        ----------
        merge_ranges : list[MergeRange]
            Merge ranges to apply as (start_row, start_col, end_row, end_col) tuples
        """
        for start_row, start_col, end_row, end_col in merge_ranges:
            try:
                # Apply merge
                self.worksheet.merge_cells(
                    start_row    = start_row,
                    start_column = start_col,
                    end_row      = end_row,
                    end_column   = end_col,
                )

                # Track the applied merge
                self.applied_merges.append((start_row, start_col, end_row, end_col))
            except ValueError:
                # Skip already merged cells
                pass

    def is_merged_cell(self, row: int, col: int) -> bool:
        """
        Check if a cell is part of a merged range.

        Parameters
        ----------
        row : int
            Excel row (1-based)
        col : int
            Excel column (1-based)

        Returns
        -------
        bool
            True if the cell is in a merged range, False otherwise
        """
        for start_row, start_col, end_row, end_col in self.applied_merges:
            if (start_row <= row <= end_row and
                start_col <= col <= end_col):
                return True
        return False
