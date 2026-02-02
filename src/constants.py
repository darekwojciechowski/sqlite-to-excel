#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Constants for SQLite to Excel converter.
"""

# Directory paths
DEFAULT_INPUT_DIR = 'input'
DEFAULT_OUTPUT_DIR = 'output'
DEFAULT_CONFIG_DIR = 'config'
DEFAULT_LOGS_DIR = 'logs'

# File extensions
DB_FILE_EXTENSION = '.db'
EXCEL_FILE_EXTENSION = '.xlsx'

# Excel formatting
EXCEL_MAX_SHEET_NAME_LENGTH = 31
EXCEL_MAX_COLUMN_WIDTH = 50
EXCEL_COLUMN_WIDTH_PADDING = 2
EXCEL_HEADER_ROW_HEIGHT = 20

# Excel colors
EXCEL_HEADER_BG_COLOR = "4472C4"
EXCEL_HEADER_FONT_COLOR = "FFFFFF"
EXCEL_BORDER_COLOR = "D3D3D3"

# Excel font sizes
EXCEL_HEADER_FONT_SIZE = 11

# Unix timestamp validation
# Range: 1/1/2000 00:00:00 to 1/1/2100 00:00:00
UNIX_TIMESTAMP_MIN = 946684800
UNIX_TIMESTAMP_MAX = 4102444800

# Timestamp column keywords
TIMESTAMP_KEYWORDS = ['time', 'timestamp', 'date']

# Timestamp conversion
TIMESTAMP_READABLE_SUFFIX = '_readable'

# Row numbering
ROW_NUMBER_COLUMN_NAME = 'Row'

# Invalid filename characters (Windows/Unix)
INVALID_FILENAME_CHARS = '<>:"|?*'

# Excel datetime format
EXCEL_DATETIME_FORMAT = 'YYYY-MM-DD HH:MM:SS'

# Logging config
LOGGING_CONFIG_FILE = 'logging.ini'
