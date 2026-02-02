#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main entry point for SQLite to Excel converter.
"""

import sqlite3
from convert_db_to_excel import setup_logging, find_all_db_files, get_output_path, convert_db_to_excel


def main() -> int:
    """Main program function"""
    logger = setup_logging()
    
    try:
        logger.info("=" * 60)
        logger.info("SQLite to Excel Conversion")
        logger.info("=" * 60)
        
        # Find all .db files
        db_files = find_all_db_files()
        logger.info(f"\nFound {len(db_files)} database file(s):")
        for db_file in db_files:
            logger.info(f"  - {db_file}")
        
        logger.info("\n" + "-" * 60)
        
        # Convert each database file
        success_count = 0
        error_count = 0
        
        for db_path in db_files:
            try:
                # Generate output path based on input filename
                output_path = get_output_path(db_path)
                
                logger.info(f"\nConverting: {db_path}")
                logger.info(f"Output: {output_path}")
                
                # Convert to Excel
                convert_db_to_excel(db_path, output_path=output_path, logger=logger)
                success_count += 1
                
            except Exception as e:
                logger.error(f"\nError converting {db_path}: {e}")
                error_count += 1
                continue
        
        logger.info("\n" + "=" * 60)
        logger.info(f"Conversion completed!")
        logger.info(f"Successfully converted: {success_count} file(s)")
        if error_count > 0:
            logger.info(f"Failed: {error_count} file(s)")
        logger.info("=" * 60)
        
        return 1 if error_count > 0 else 0
        
    except FileNotFoundError as e:
        logger.error(f"\nError: {e}")
        logger.error("Make sure .db files are located in the 'input/' folder")
        return 1
    
    except Exception as e:
        logger.exception(f"\nUnexpected error: {e}")
        return 1


if __name__ == "__main__":
    exit(main())
