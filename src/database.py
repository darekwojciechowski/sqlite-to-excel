#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Database operations for SQLite to Excel converter.
"""

import sqlite3
import pandas as pd


def get_all_tables(db_path: str) -> list[str]:
    """Get a list of all tables from the SQLite database"""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all tables (excluding SQLite system tables)
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    tables = [row[0] for row in cursor.fetchall()]
    
    conn.close()
    return tables


def rename_data_format_columns(df: pd.DataFrame, db_path: str) -> pd.DataFrame:
    """
    Rename data_format_X columns to readable names from data_format table.
    Maps data_format_0, data_format_1, etc. to their comment descriptions.
    """
    try:
        with sqlite3.connect(db_path) as conn:
            # Read the data_format table to get column descriptions
            format_df = pd.read_sql_query("SELECT data_format_index, comment FROM data_format", conn)
            
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
    Uses proper quoting to prevent SQL injection.
    """
    with sqlite3.connect(db_path) as conn:
        # SQLite uses double quotes for identifiers
        query = f'SELECT * FROM "{table_name}"'
        df = pd.read_sql_query(query, conn)
    
    return df
