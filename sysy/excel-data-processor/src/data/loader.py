#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Data Loader

This module contains the DataLoader class which handles loading data from various file formats.
"""

import os
import pandas as pd
import logging

logger = logging.getLogger(__name__)

class DataLoader:
    """Handles loading data from various file formats."""
    
    def __init__(self):
        """Initialize the DataLoader."""
        self.supported_extensions = {
            'excel': ['.xlsx', '.xls', '.xlsm'],
            'csv': ['.csv'],
            'text': ['.txt']
        }
    
    def load_file(self, file_path, **kwargs):
        """Load data from a file based on its extension.
        
        Args:
            file_path (str): Path to the file to load.
            **kwargs: Additional arguments to pass to the loader function.
            
        Returns:
            pandas.DataFrame or dict: Loaded data, either as a single DataFrame 
                                      or a dict of sheet_name -> DataFrame mappings.
        """
        if not os.path.exists(file_path):
            logger.error(f"File does not exist: {file_path}")
            raise FileNotFoundError(f"File not found: {file_path}")
            
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        # Determine file type and use appropriate loader
        if ext in self.supported_extensions['excel']:
            return self.load_excel(file_path, **kwargs)
        elif ext in self.supported_extensions['csv']:
            return self.load_csv(file_path, **kwargs)
        elif ext in self.supported_extensions['text']:
            return self.load_text(file_path, **kwargs)
        else:
            logger.error(f"Unsupported file extension: {ext}")
            raise ValueError(f"Unsupported file extension: {ext}")
    
    def get_excel_sheet_names(self, file_path):
        """Get the sheet names from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file.
            
        Returns:
            list: List of sheet names.
        """
        try:
            xls = pd.ExcelFile(file_path)
            return xls.sheet_names
        except Exception as e:
            logger.error(f"Error getting Excel sheet names: {str(e)}")
            raise
    
    def load_excel(self, file_path, sheet_name=0, **kwargs):
        """Load data from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file.
            sheet_name (str or int or list or None): Sheet(s) to load. Default is 0 (first sheet).
                - If None, all sheets are loaded.
                - If str, the sheet with that name is loaded.
                - If int, the sheet at that position is loaded (0-based).
                - If list, all sheets with those names are loaded.
            **kwargs: Additional arguments to pass to pandas.read_excel().
            
        Returns:
            pandas.DataFrame or dict: If sheet_name is str or int, returns a DataFrame.
                                    If sheet_name is list or None, returns a dict of DataFrames.
        """
        try:
            logger.info(f"Loading Excel file: {file_path}")
            return pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise
    
    def load_excel_sheet(self, file_path, sheet_name, **kwargs):
        """Load a specific sheet from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file.
            sheet_name (str or int): Name or index of the sheet to load.
            **kwargs: Additional arguments to pass to pandas.read_excel().
            
        Returns:
            pandas.DataFrame: The loaded DataFrame.
        """
        try:
            logger.info(f"Loading sheet '{sheet_name}' from Excel file: {file_path}")
            return pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
        except Exception as e:
            logger.error(f"Error loading Excel sheet: {str(e)}")
            raise
    
    def load_csv(self, file_path, **kwargs):
        """Load data from a CSV file.
        
        Args:
            file_path (str): Path to the CSV file.
            **kwargs: Additional arguments to pass to pandas.read_csv().
            
        Returns:
            pandas.DataFrame: The loaded DataFrame.
        """
        try:
            logger.info(f"Loading CSV file: {file_path}")
            # Attempt to auto-detect delimiter if not specified
            if 'sep' not in kwargs and 'delimiter' not in kwargs:
                with open(file_path, 'r', encoding=kwargs.get('encoding', 'utf-8')) as f:
                    sample = f.readline() + f.readline()
                
                # Check common delimiters
                for delimiter in [',', ';', '\t', '|']:
                    if delimiter in sample:
                        kwargs['sep'] = delimiter
                        break
            
            return pd.read_csv(file_path, **kwargs)
        except Exception as e:
            logger.error(f"Error loading CSV file: {str(e)}")
            raise
    
    def load_text(self, file_path, **kwargs):
        """Load data from a text file.
        
        Args:
            file_path (str): Path to the text file.
            **kwargs: Additional arguments to pass to pandas.read_fwf() or pandas.read_csv().
            
        Returns:
            pandas.DataFrame: The loaded DataFrame.
        """
        try:
            logger.info(f"Loading text file: {file_path}")
            
            # Try to determine if it's a fixed-width or delimited file
            with open(file_path, 'r', encoding=kwargs.get('encoding', 'utf-8')) as f:
                sample_lines = [f.readline() for _ in range(min(5, sum(1 for _ in open(file_path))))]
                sample = ''.join(sample_lines)
            
            # Check for common delimiters
            delimiters = [',', ';', '\t', '|']
            for delimiter in delimiters:
                if delimiter in sample:
                    # If delimiter found, treat as delimited file
                    return pd.read_csv(file_path, sep=delimiter, **kwargs)
            
            # If no common delimiter found, try fixed-width
            return pd.read_fwf(file_path, **kwargs)
        except Exception as e:
            logger.error(f"Error loading text file: {str(e)}")
            raise