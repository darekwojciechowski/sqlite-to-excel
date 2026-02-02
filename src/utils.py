#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utility functions for SQLite to Excel converter.
"""

import os
import glob
import logging
import logging.config
from pathlib import Path


def setup_logging() -> logging.Logger:
    """Configure logging from configuration file"""
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    config_path = Path('config/logging.ini')
    if config_path.exists():
        logging.config.fileConfig(config_path)
    else:
        # Fallback to basic config if file not found
        logging.basicConfig(
            level=logging.INFO,
            format='%(message)s'
        )
    return logging.getLogger(__name__)


def find_all_db_files(input_dir: str = 'input') -> list[str]:
    """Find all .db files in the input/ folder"""
    if not input_dir or not input_dir.strip():
        raise ValueError("Input directory cannot be empty")
    
    input_dir = input_dir.strip()
    
    if not os.path.exists(input_dir):
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    
    if not os.path.isdir(input_dir):
        raise ValueError(f"Path is not a directory: {input_dir}")
    
    db_files = glob.glob(os.path.join(input_dir, '*.db'))
    
    if not db_files:
        raise FileNotFoundError(f"No .db files found in '{input_dir}/' folder")
    
    return db_files


def get_output_path(db_path: str, output_dir: str = 'output') -> str:
    """Generate output Excel file path based on input database filename"""
    if not db_path or not db_path.strip():
        raise ValueError("Database path cannot be empty")
    
    if not output_dir or not output_dir.strip():
        raise ValueError("Output directory cannot be empty")
    
    db_path = db_path.strip()
    output_dir = output_dir.strip()
    
    # Get the base filename without extension
    base_name = os.path.splitext(os.path.basename(db_path))[0]
    
    if not base_name:
        raise ValueError(f"Could not extract filename from path: {db_path}")
    
    # Sanitize filename (remove invalid characters for Windows/Unix)
    invalid_chars = '<>:"|?*'
    for char in invalid_chars:
        base_name = base_name.replace(char, '_')
    
    # Create output path with .xlsx extension
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")
    return output_path
