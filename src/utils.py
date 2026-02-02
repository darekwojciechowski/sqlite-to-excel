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
    db_files = glob.glob(os.path.join(input_dir, '*.db'))
    
    if not db_files:
        raise FileNotFoundError(f"No .db files found in '{input_dir}/' folder")
    
    return db_files


def get_output_path(db_path: str, output_dir: str = 'output') -> str:
    """Generate output Excel file path based on input database filename"""
    # Get the base filename without extension
    base_name = os.path.splitext(os.path.basename(db_path))[0]
    # Create output path with .xlsx extension
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")
    return output_path
