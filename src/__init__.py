"""
SQLite to Excel converter package.
"""

__version__ = "1.0.0"

from .excel_writer import convert_db_to_excel
from .utils import (
    setup_logging,
    find_all_db_files,
    get_output_path,
    validate_non_empty_string,
    sanitize_excel_sheet_name
)

__all__ = [
    'convert_db_to_excel',
    'setup_logging',
    'find_all_db_files',
    'get_output_path',
    'validate_non_empty_string',
    'sanitize_excel_sheet_name',
]
