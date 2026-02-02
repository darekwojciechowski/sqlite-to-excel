"""
Utility functions for SQLite to Excel converter.
"""

import logging
import logging.config
from pathlib import Path

from .constants import (
    DEFAULT_INPUT_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_CONFIG_DIR,
    DEFAULT_LOGS_DIR,
    DB_FILE_EXTENSION,
    EXCEL_FILE_EXTENSION,
    EXCEL_MAX_SHEET_NAME_LENGTH,
    INVALID_FILENAME_CHARS,
    LOGGING_CONFIG_FILE
)


def validate_non_empty_string(value: str, name: str) -> str:
    """Validate that string is not empty or whitespace only."""
    if not value or not value.strip():
        raise ValueError(f"{name} cannot be empty")
    return value.strip()


def sanitize_excel_sheet_name(name: str) -> str:
    """Sanitize and truncate name for Excel sheet (max 31 chars)."""
    return name[:EXCEL_MAX_SHEET_NAME_LENGTH] if len(name) > EXCEL_MAX_SHEET_NAME_LENGTH else name


def setup_logging() -> logging.Logger:
    """Configure logging from configuration file"""
    log_dir = Path(DEFAULT_LOGS_DIR)
    log_dir.mkdir(exist_ok=True)
    
    config_path = Path(DEFAULT_CONFIG_DIR) / LOGGING_CONFIG_FILE
    if config_path.exists():
        logging.config.fileConfig(config_path)
    else:
        # Fallback to basic config if file not found
        logging.basicConfig(
            level=logging.INFO,
            format='%(message)s'
        )
    return logging.getLogger(__name__)


def find_all_db_files(input_dir: str = DEFAULT_INPUT_DIR) -> list[str]:
    """Find all .db files in the input/ folder"""
    input_dir = validate_non_empty_string(input_dir, "Input directory")
    input_path = Path(input_dir)
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    
    if not input_path.is_dir():
        raise ValueError(f"Path is not a directory: {input_dir}")
    
    db_files = [str(f) for f in input_path.glob(f'*{DB_FILE_EXTENSION}')]
    
    if not db_files:
        raise FileNotFoundError(f"No .db files found in '{input_dir}/' folder")
    
    return db_files


def get_output_path(db_path: str, output_dir: str = DEFAULT_OUTPUT_DIR) -> str:
    """Generate output Excel file path based on input database filename"""
    db_path = validate_non_empty_string(db_path, "Database path")
    output_dir = validate_non_empty_string(output_dir, "Output directory")
    
    # Get the base filename without extension
    db_file = Path(db_path)
    base_name = db_file.stem
    
    if not base_name:
        raise ValueError(f"Could not extract filename from path: {db_path}")
    
    # Sanitize filename (remove invalid characters for Windows/Unix)
    for char in INVALID_FILENAME_CHARS:
        base_name = base_name.replace(char, '_')
    
    # Create output path with .xlsx extension
    output_path = Path(output_dir) / f"{base_name}{EXCEL_FILE_EXTENSION}"
    return str(output_path)
