"""
Excel file writer for SQLite database tables.
"""

from pathlib import Path
import pandas as pd

from .protocols import LoggerProtocol
from .constants import DB_FILE_EXTENSION, EXCEL_FILE_EXTENSION
from .database import get_all_tables, rename_data_format_columns, read_table
from .timestamp_converter import convert_timestamps_to_readable
from .formatters import add_row_numbers, format_worksheet
from .utils import validate_non_empty_string, sanitize_excel_sheet_name


def convert_db_to_excel(
    db_path: str,
    output_path: str,
    logger: LoggerProtocol
) -> None:
    """Convert all tables from SQLite database to Excel file"""
    # Validate and sanitize inputs
    db_path = validate_non_empty_string(db_path, "Database path")
    output_path = validate_non_empty_string(output_path, "Output path")
    db_file = Path(db_path)
    
    if not db_file.exists():
        raise FileNotFoundError(f"Database file not found: {db_path}")
    
    if not db_file.is_file():
        raise ValueError(f"Path is not a file: {db_path}")
    
    # Validate it's a SQLite database
    if db_file.suffix.lower() != DB_FILE_EXTENSION:
        logger.warning(f"File does not have {DB_FILE_EXTENSION} extension: {db_path}")
    
    # Validate output file extension
    output_file = Path(output_path)
    if output_file.suffix.lower() != EXCEL_FILE_EXTENSION:
        raise ValueError(f"Output file must have {EXCEL_FILE_EXTENSION} extension: {output_path}")
    
    # Create output directory if it doesn't exist
    output_dir = output_file.parent
    if output_dir != Path('.'):
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            raise PermissionError(f"No permission to create output directory: {output_dir}")
        except OSError as e:
            raise OSError(f"Failed to create output directory {output_dir}: {e}")
    
    # Get list of tables
    tables = get_all_tables(db_path)
    
    if not tables:
        raise ValueError("Database does not contain any tables")
    
    logger.info(f"\nFound {len(tables)} table(s) in database:")
    for table in tables:
        logger.info(f"  - {table}")
    
    # Create Excel writer
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for table in tables:
            # Read table into DataFrame
            df = read_table(db_path, table)
            
            # Rename data_format_X columns to readable names
            df = rename_data_format_columns(df, db_path)
            
            # Convert Unix timestamps to readable format
            df = convert_timestamps_to_readable(df)
            
            # Add row numbers for better readability
            df = add_row_numbers(df)
            
            # Save to sheet (sheet name is table name)
            sheet_name = sanitize_excel_sheet_name(table)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Format the worksheet
            worksheet = writer.sheets[sheet_name]
            format_worksheet(worksheet, df)
            
            logger.info(f"  Table '{table}': {len(df)} rows, {len(df.columns)} columns")
    
    logger.info(f"\nSuccess! Data saved to: {output_path}")
