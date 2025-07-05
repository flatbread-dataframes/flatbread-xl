from typing import Any

import pandas as pd
from flatbreadxl.layout import TableLayout, CellPosition


def get_contiguous_spans(values: list[Any]) -> list[dict[str, Any]]:
    """
    Identify contiguous spans of values in hierarchical indices.

    Parameters
    ----------
    values : list[Any]
        list of index values (tuples for MultiIndex)

    Returns
    -------
    list[dict[str, Any]]
        list of span objects with start position, value, and count
    """
    spans = []
    prev = None
    start_idx = 0

    for i, val in enumerate(values):
        if val != prev:
            if i > 0:
                spans.append({
                    'start': start_idx,
                    'value': prev,
                    'count': i - start_idx
                })
            start_idx = i
            prev = val

    # Add the last span
    if len(values) > 0:
        spans.append({
            'start': start_idx,
            'value': prev,
            'count': len(values) - start_idx
        })

    return spans


def get_level_spans(
    index: pd.Index,
    max_level: int | None = None,
) -> list[list[dict[str, Any]]]:
    """
    Get span information for each level of a MultiIndex.

    Parameters
    ----------
    index : pd.Index
        Index to analyze
    max_level : Optional[int], default None
        Maximum level to consider (defaults to all levels)

    Returns
    -------
    list[list[dict[str, Any]]]
        list of spans for each level
    """
    if not isinstance(index, pd.MultiIndex):
        # Handle single-level index
        return [get_contiguous_spans(index)]

    levels = []
    n_levels = max_level if max_level is not None else index.nlevels

    for level in range(n_levels):
        # Get values up to this level for each index entry
        level_values = [idx[:level+1] for idx in index]
        spans = get_contiguous_spans(level_values)
        levels.append(spans)

    return levels


def get_merge_ranges_from_spans(
    spans: list[list[dict[str, Any]]],
    layout,
    is_row_index: bool
) -> list[Any, int, int]:
    """
    Convert spans into Excel merge ranges using layout information.

    Parameters
    ----------
    spans : list[list[dict[str, Any]]]
        list of spans from get_level_spans()
    layout : AxisLayout
        The layout for the axis (index or columns)
    is_row_index : bool
        If True, spans are for row index; otherwise column index

    Returns
    -------
    list[Any, int, int]
        list of merge ranges as tuples (start_row, start_col, end_row, end_col)
    """
    merge_ranges = []

    for level, level_spans in enumerate(spans):
        for span in level_spans:
            # Only merge if span is longer than 1
            if span['count'] > 1:
                if is_row_index:
                    # For row index: merge vertically
                    cell_start = layout.cell_at(level, span['start'])
                    cell_end = layout.cell_at(level, span['start'] + span['count'] - 1)

                    merge_ranges.append((
                        cell_start.excel_y,
                        cell_start.excel_x,
                        cell_end.excel_y,
                        cell_start.excel_x  # Same column
                    ))
                else:
                    # For column index: merge horizontally
                    cell_start = layout.cell_at(span['start'], level)
                    cell_end = layout.cell_at(span['start'] + span['count'] - 1, level)

                    merge_ranges.append((
                        cell_start.excel_y,
                        cell_start.excel_x,
                        cell_start.excel_y,  # Same row
                        cell_end.excel_x
                    ))

    return merge_ranges


def get_empty_spans(
    df: pd.DataFrame,
    layout: TableLayout
) -> list[Any, int, int]:
    """
    Identify spans of empty cells in MultiIndex row indices that should be merged.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with MultiIndex to analyze
    layout : TableLayout
        Table layout for coordinate translation

    Returns
    -------
    list[Any, int, int]
        list of merge ranges as (start_row, start_col, end_row, end_col) tuples
    """
    if not isinstance(df.index, pd.MultiIndex):
        return []

    merge_ranges = []

    # Check each row's index values for empty spans
    for i, idx in enumerate(df.index):
        if not isinstance(idx, tuple):
            continue

        # Find spans of empty values
        span_start = None
        last_non_empty = 0

        for level in range(1, len(idx)):
            value = idx[level]

            if pd.isna(value) or value == '':
                # Start a new span if we're not already in one
                if span_start is None:
                    span_start = level
            else:
                # End any current span
                if span_start is not None:
                    # Convert to Excel coordinates
                    start_cell = layout.index.cell_at(last_non_empty, i)
                    end_cell = layout.index.cell_at(level - 1, i)

                    merge_ranges.append((
                        start_cell.excel_y,
                        start_cell.excel_x,
                        end_cell.excel_y,
                        end_cell.excel_x
                    ))
                    span_start = None

                # Update the last non-empty position
                last_non_empty = level

        # Handle any span that extends to the end
        if span_start is not None:
            # Convert to Excel coordinates
            start_cell = layout.index.cell_at(last_non_empty, i)
            end_cell = layout.index.cell_at(df.index.nlevels - 1, i)

            merge_ranges.append((
                start_cell.excel_y,
                start_cell.excel_x,
                end_cell.excel_y,
                end_cell.excel_x
            ))

    return merge_ranges


def get_all_merge_ranges(
    df: pd.DataFrame,
    layout: TableLayout
) -> list[Any, int, int]:
    """
    Calculate all merge ranges for a DataFrame including spans and empty cells.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame to analyze
    layout : TableLayout
        Table layout for coordinate translation

    Returns
    -------
    list[Any, int, int]
        Combined list of all merge ranges
    """
    # Calculate spans
    row_spans = get_level_spans(df.index)
    column_spans = get_level_spans(df.columns)

    # Convert spans to merge ranges
    row_merges = get_merge_ranges_from_spans(
        row_spans,
        layout.index,
        is_row_index=True
    )

    column_merges = get_merge_ranges_from_spans(
        column_spans,
        layout.columns,
        is_row_index=False
    )

    # Get empty spans
    empty_merges = get_empty_spans(df, layout)

    # Combine all merge ranges
    return row_merges + column_merges + empty_merges
