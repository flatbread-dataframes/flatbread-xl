# File: uvmlib/export/formatter/pattern.py

from typing import Any, List, Dict, Tuple, Union, Optional, TypeVar, Generic

import pandas as pd

T = TypeVar('T')  # Type for the format/style/etc. that matches a pattern
PatternSpec = str | int | float | tuple | list | dict


class PatternMatcher:
    """
    Utility class for matching pandas index/column labels against patterns.
    Used for applying formats, styles, borders, etc. based on label patterns.
    """

    @staticmethod
    def is_match(label: Any, pattern: PatternSpec) -> bool:
        """
        Check if a label matches a pattern specification.

        Parameters
        ----------
        label : Any
            The index or column label to check
        pattern : PatternSpec
            Pattern to match against (string, tuple, etc.)

        Returns
        -------
        bool
            True if the label matches the pattern, False otherwise
        """
        # Check for tuple pattern exact match
        if isinstance(pattern, tuple) and isinstance(label, tuple) and label == pattern:
            return True

        # For non-tuple patterns
        if not isinstance(pattern, tuple):
            # Direct equality match
            if label == pattern:
                return True

            # Check parts of tuple label
            if isinstance(label, tuple):
                return PatternMatcher._tuple_contains_match(label, pattern)

        return False

    @staticmethod
    def _tuple_contains_match(label_tuple: Tuple, pattern: Any) -> bool:
        """
        Check if any part of a tuple label matches the pattern.

        Parameters
        ----------
        label_tuple : Tuple
            Tuple containing label components (e.g., MultiIndex levels)
        pattern : Any
            Pattern to match against any part of the tuple

        Returns
        -------
        bool
            True if any part of the label tuple matches the pattern
        """
        for part in label_tuple:
            # Exact match with tuple part
            if part == pattern:
                return True

            # String prefix match for string patterns
            if isinstance(part, str) and isinstance(pattern, str):
                if part.startswith(pattern):
                    return True

        return False

    @staticmethod
    def find_match(label: Any, patterns: List[Tuple[PatternSpec, T]]) -> Optional[T]:
        """
        Find the first matching value for a label from a list of pattern-value pairs.

        Parameters
        ----------
        label : Any
            The index or column label to check
        patterns : List[Tuple[PatternSpec, T]]
            List of (pattern, value) tuples to check against

        Returns
        -------
        Optional[T]
            The first matching value, or None if no match is found
        """
        for pattern, value in patterns:
            if PatternMatcher.is_match(label, pattern):
                return value
        return None

    @staticmethod
    def create_position_map(
        labels: Union[pd.Index, List],
        patterns: List[Tuple[PatternSpec, T]]
    ) -> List[Optional[T]]:
        """
        Create a position-based map of values for each label.

        Parameters
        ----------
        labels : Union[pd.Index, List]
            List of labels to process
        patterns : List[Tuple[PatternSpec, T]]
            List of (pattern, value) pairs to match against

        Returns
        -------
        List[Optional[T]]
            List where index i contains the value for the i-th label, or None if no match
        """
        result = [None] * len(labels)

        for i, label in enumerate(labels):
            match = PatternMatcher.find_match(label, patterns)
            if match is not None:
                result[i] = match

        return result

    @staticmethod
    def process_spec_dict(
        spec_dict: Dict[str, Any]
    ) -> Tuple[List[Tuple[PatternSpec, T]], List[Tuple[PatternSpec, T]]]:
        """
        Process a specification dictionary with 'rows' and 'columns' keys into usable patterns.

        Parameters
        ----------
        spec_dict : Dict[str, Any]
            Dictionary with 'rows' and/or 'columns' keys containing pattern specifications

        Returns
        -------
        Tuple[List[Tuple[PatternSpec, T]], List[Tuple[PatternSpec, T]]]
            Tuple of (row_patterns, column_patterns)
        """
        row_patterns = []
        column_patterns = []

        if 'rows' in spec_dict:
            raw_row_specs = spec_dict['rows']
            if isinstance(raw_row_specs, dict):
                row_patterns = list(raw_row_specs.items())
            elif isinstance(raw_row_specs, list):
                row_patterns = raw_row_specs

        if 'columns' in spec_dict:
            raw_col_specs = spec_dict['columns']
            if isinstance(raw_col_specs, dict):
                column_patterns = list(raw_col_specs.items())
            elif isinstance(raw_col_specs, list):
                column_patterns = raw_col_specs

        # If no 'rows' or 'columns' keys, treat entire dict as column patterns
        if not row_patterns and not column_patterns and spec_dict:
            if all(isinstance(item, tuple) and len(item) == 2 for item in spec_dict.items()):
                column_patterns = list(spec_dict.items())

        return row_patterns, column_patterns
