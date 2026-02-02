"""
Database operations for SQLite to Excel converter.
"""

import sqlite3
import re
import pandas as pd


def _validate_sql_identifier(identifier: str) -> str:
    """
    Validate and sanitize SQL identifier (table/column name).
    
    Ensures identifier contains only safe characters to prevent SQL injection.
    Raises ValueError if identifier contains unsafe characters.
    
    Args:
        identifier: Table or column name to validate
        
    Returns:
        Validated identifier
        
    Raises:
        ValueError: If identifier contains unsafe characters
    """
    if not identifier:
        raise ValueError("SQL identifier cannot be empty")
    
    # Allow only alphanumeric characters, underscores, and spaces
    # SQLite allows many characters in identifiers when quoted
    if not re.match(r'^[\w\s]+$', identifier):
        raise ValueError(
            f"Invalid SQL identifier '{identifier}': "
            f"Only alphanumeric characters, underscores, and spaces are allowed"
        )
    
    return identifier


def _quote_identifier(identifier: str) -> str:
    """
    Safely quote SQL identifier for SQLite.
    
    Uses SQLite's double-quote escaping by doubling any quotes in the identifier.
    
    Args:
        identifier: Validated SQL identifier
        
    Returns:
        Properly quoted identifier safe for SQL query
    """
    # Escape any double quotes by doubling them (SQLite standard)
    escaped = identifier.replace('"', '""')
    return f'"{escaped}"'


def get_all_tables(db_path: str) -> list[str]:
    """
    Get a list of all tables from the SQLite database.
    
    Uses parameterized query to safely filter system tables.
    """
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        
        # Get all tables (excluding SQLite system tables)
        # Using parameterized query for the LIKE pattern
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type=? AND name NOT LIKE ?",
            ('table', 'sqlite_%')
        )
        tables = [row[0] for row in cursor.fetchall()]
    
    return tables


def rename_data_format_columns(df: pd.DataFrame, db_path: str) -> pd.DataFrame:
    """
    Rename data_format_X columns to readable names from data_format table.
    Maps data_format_0, data_format_1, etc. to their comment descriptions.
    
    Uses safe SQL queries with validation and proper quoting.
    """
    try:
        with sqlite3.connect(db_path) as conn:
            # Validate table name exists
            cursor = conn.cursor()
            cursor.execute(
                "SELECT 1 FROM sqlite_master WHERE type=? AND name=?",
                ('table', 'data_format')
            )
            if not cursor.fetchone():
                return df
            
            # Read the data_format table to get column descriptions
            # Table name is validated above, columns are literals
            format_df = pd.read_sql_query(
                "SELECT data_format_index, comment FROM data_format",
                conn
            )
            
            # Create mapping from data_format_X to comment
            rename_map = {}
            for _, row in format_df.iterrows():
                col_name = f"data_format_{row['data_format_index']}"
                if col_name in df.columns:
                    rename_map[col_name] = row['comment']
            
            # Rename columns
            if rename_map:
                df = df.rename(columns=rename_map)
        
    except Exception:
        # If data_format table doesn't exist or has different structure, skip renaming
        pass
    
    return df


def read_table(db_path: str, table_name: str) -> pd.DataFrame:
    """
    Read a table from SQLite database into a DataFrame.
    
    Validates table name and uses proper SQL quoting to prevent SQL injection.
    
    Args:
        db_path: Path to SQLite database file
        table_name: Name of table to read (validated for safety)
        
    Returns:
        DataFrame containing table data
        
    Raises:
        ValueError: If table name contains unsafe characters
    """
    # Validate table name to prevent SQL injection
    validated_name = _validate_sql_identifier(table_name)
    
    with sqlite3.connect(db_path) as conn:
        # Verify table exists first (using parameterized query)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT 1 FROM sqlite_master WHERE type=? AND name=?",
            ('table', validated_name)
        )
        if not cursor.fetchone():
            raise ValueError(f"Table '{validated_name}' does not exist in database")
        
        # Use properly quoted identifier for the query
        # Table name is validated, so this is safe
        quoted_table = _quote_identifier(validated_name)
        query = f'SELECT * FROM {quoted_table}'
        df = pd.read_sql_query(query, conn)
    
    return df
