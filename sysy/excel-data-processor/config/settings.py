#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Application Settings

This file contains configuration settings for the Excel Data Processor application.
"""

# Application information
APP_NAME = "Excel Data Processor"
APP_VERSION = "1.0.0"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"
APP_SIZE = "1024x768"

# File settings
SUPPORTED_FILE_TYPES = [
    ("Excel files", "*.xlsx *.xls *.xlsm"),
    ("CSV files", "*.csv"),
    ("All files", "*.*")
]

# Data processing settings
MAX_PREVIEW_ROWS = 100
DEFAULT_ENCODING = "utf-8"
DATE_FORMAT = "%Y-%m-%d"

# Analysis settings
DEFAULT_SIGNIFICANCE_LEVEL = 0.05
OUTLIER_THRESHOLD = 3.0  # Z-score threshold for outlier detection
CORRELATION_THRESHOLD = 0.7  # Threshold for significant correlation

# Data quality thresholds
COMPLETENESS_THRESHOLD = 0.9  # 90% of values should be non-null
DUPLICATES_THRESHOLD = 0.05  # No more than 5% of rows should be duplicates

# GUI settings
PADDING = 10
STEP_TITLES = [
    "1. Load Data",
    "2. Assess Data Quality",
    "3. Merge and Clean Columns",
    "4. Analyze Data",
    "5. Export Results"
]

# Default save locations
DEFAULT_REPORT_DIR = "reports"
DEFAULT_EXPORT_DIR = "exports"

# Feature flags
ENABLE_ADVANCED_ANALYTICS = True
ENABLE_DATA_VISUALIZATION = True
ENABLE_AUTO_RECOMMENDATIONS = True

# Performance settings
CHUNK_SIZE = 10000  # For processing large files in chunks