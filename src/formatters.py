#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel formatting utilities for worksheets.
"""

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


def add_row_numbers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add sequential row numbers starting from 1 at the beginning of DataFrame.
    """
    df.insert(0, 'Row', range(1, len(df) + 1))
    return df


def format_worksheet(worksheet: Worksheet, df: pd.DataFrame) -> None:
    """
    Format worksheet for better readability and aesthetics.
    - Auto-adjust column widths
    - Format headers (bold, background color)
    - Add borders
    - Freeze top row
    """
    # Define styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    cell_alignment = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    
    # Format header row
    for col_num, column in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Format data cells and calculate column widths
    for col_num, column in enumerate(df.columns, 1):
        column_letter = get_column_letter(col_num)
        
        # Calculate maximum width for this column
        max_length = len(str(column))  # Start with header length
        
        # Check if this is a datetime column
        is_datetime_col = pd.api.types.is_datetime64_any_dtype(df[column])
        
        for row_num, value in enumerate(df[column], 2):  # Start from row 2 (data rows)
            cell = worksheet.cell(row=row_num, column=col_num)
            cell.alignment = cell_alignment
            cell.border = thin_border
            
            # Format datetime cells
            if is_datetime_col and pd.notna(value):
                cell.number_format = 'YYYY-MM-DD HH:MM:SS'
            
            # Update max length
            if value is not None:
                cell_length = len(str(value))
                max_length = max(max_length, cell_length)
        
        # Set column width (add some padding)
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Freeze top row (header)
    worksheet.freeze_panes = worksheet['A2']
    
    # Set row height for header
    worksheet.row_dimensions[1].height = 20
