#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Data Processor - Main Application

This is the entry point for the Excel Data Processor application.
It initializes the GUI and starts the application.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

# Add the project directory to the path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# Import application modules
try:
    from src.gui.app import ExcelDataProcessorApp
    from config.settings import APP_TITLE, APP_SIZE
except ImportError as e:
    logger.critical(f"Failed to import required modules: {e}")
    print(f"Error: Failed to import required modules: {e}")
    print("Please make sure you have installed all required dependencies.")
    sys.exit(1)

def main():
    """Main function to start the application."""
    try:
        # Create the main Tkinter window
        root = tk.Tk()
        root.title(APP_TITLE)
        root.geometry(APP_SIZE)
        
        # Set theme and style
        style = ttk.Style()
        style.theme_use("clam")  # Use a modern theme
        
        # Initialize the application
        app = ExcelDataProcessorApp(root)
        
        # Center the window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'+{x}+{y}')
        
        # Start the application
        logger.info("Starting Excel Data Processor application")
        root.mainloop()
    except Exception as e:
        logger.critical(f"Application crashed: {e}", exc_info=True)
        print(f"Error: Application crashed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()