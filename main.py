"""
Main entry point for SQLite to Excel converter.
"""

import uuid
import time
from src import setup_logging, find_all_db_files, get_output_path, convert_db_to_excel
from src.observability import log_structured, log_error


def main() -> int:
    """Main program function"""
    logger = setup_logging()
    batch_trace_id = str(uuid.uuid4())
    batch_start_time = time.time()
    
    try:
        log_structured(
            logger,
            "info",
            "Starting batch conversion",
            batch_trace_id=batch_trace_id,
            operation="batch_conversion"
        )
        
        # Find all .db files
        db_files = find_all_db_files()
        log_structured(
            logger,
            "info",
            f"Found {len(db_files)} database file(s)",
            batch_trace_id=batch_trace_id,
            files_count=len(db_files),
            files=db_files
        )
        
        # Convert each database file
        success_count = 0
        error_count = 0
        
        for db_path in db_files:
            file_start_time = time.time()
            try:
                # Generate output path based on input filename
                output_path = get_output_path(db_path)
                
                log_structured(
                    logger,
                    "info",
                    f"Converting database",
                    batch_trace_id=batch_trace_id,
                    db_file=db_path,
                    output_file=output_path
                )
                
                # Convert to Excel
                convert_db_to_excel(db_path, output_path, logger)
                
                file_duration = (time.time() - file_start_time) * 1000
                log_structured(
                    logger,
                    "info",
                    f"Successfully converted database",
                    batch_trace_id=batch_trace_id,
                    db_file=db_path,
                    duration_ms=f"{file_duration:.2f}",
                    success="true"
                )
                success_count += 1
                
            except Exception as e:
                log_error(logger, batch_trace_id, e, context=f"Converting {db_path}")
                error_count += 1
                continue
        
        batch_duration = (time.time() - batch_start_time) * 1000
        log_structured(
            logger,
            "info",
            "Batch conversion completed",
            batch_trace_id=batch_trace_id,
            total_files=len(db_files),
            successful=success_count,
            failed=error_count,
            duration_ms=f"{batch_duration:.2f}",
            success_rate=f"{(success_count / len(db_files) * 100):.1f}%" if db_files else "0%"
        )
        
        return 1 if error_count > 0 else 0
        
    except FileNotFoundError as e:
        log_error(logger, batch_trace_id, e, context="Finding database files")
        logger.error("Make sure .db files are located in the 'input/' folder")
        return 1
    
    except Exception as e:
        log_error(logger, batch_trace_id, e, context="Batch conversion")
        logger.exception("Unexpected error occurred")
        return 1


if __name__ == "__main__":
    exit(main())
