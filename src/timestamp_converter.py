"""
Unix timestamp detection and conversion utilities.
"""

import pandas as pd

from .constants import (
    UNIX_TIMESTAMP_MIN,
    UNIX_TIMESTAMP_MAX,
    TIMESTAMP_KEYWORDS,
    TIMESTAMP_READABLE_SUFFIX
)


def is_unix_timestamp_column(series: pd.Series) -> bool:
    """
    Check if a column contains Unix timestamps.
    Returns True if column appears to contain Unix timestamps.
    """
    # Check if column name suggests timestamp
    name_lower = str(series.name).lower()
    if not any(keyword in name_lower for keyword in TIMESTAMP_KEYWORDS):
        return False
    
    # Check if values are numeric and in reasonable Unix timestamp range
    if not pd.api.types.is_numeric_dtype(series):
        return False
    
    # Unix timestamp range: 1/1/2000 to 1/1/2100
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    
    return non_null.between(UNIX_TIMESTAMP_MIN, UNIX_TIMESTAMP_MAX).all()


def convert_timestamps_to_readable(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detect and convert Unix timestamp columns to readable datetime format.
    Adds new columns with '_readable' suffix next to original timestamp columns.
    """
    new_columns = {}
    insert_positions = {}
    
    for col in df.columns:
        if is_unix_timestamp_column(df[col]):
            # Create readable column name
            readable_col = f"{col}{TIMESTAMP_READABLE_SUFFIX}"
            
            # Convert Unix timestamp to datetime
            new_columns[readable_col] = pd.to_datetime(df[col], unit='s', errors='coerce')
            
            # Store position to insert after original column
            insert_positions[readable_col] = list(df.columns).index(col) + 1
    
    # Insert new columns at appropriate positions
    for col_name in sorted(insert_positions.keys(), key=lambda x: insert_positions[x], reverse=True):
        position = insert_positions[col_name]
        df.insert(position, col_name, new_columns[col_name])
    
    return df
