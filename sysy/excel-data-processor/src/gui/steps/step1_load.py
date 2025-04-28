#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Load Data Step

This module contains the LoadDataStep class which handles data loading functionality.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import pandas as pd

from src.gui.steps import BaseStep
from src.data.loader import DataLoader
from src.gui.widgets.data_preview import DataPreviewFrame
from config.settings import SUPPORTED_FILE_TYPES, MAX_PREVIEW_ROWS

logger = logging.getLogger(__name__)

class LoadDataStep(BaseStep):
    """Step for loading data from files."""
    
    def _get_title(self):
        return "Step 1: Load Data"
    
    def _get_description(self):
        return "Select an Excel or CSV file to load. You can preview the data before proceeding."
    
    def _init_ui(self):
        """Initialize the step UI."""
        # File selection area
        file_frame = ttk.LabelFrame(self.content_frame, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # File path entry and browse button
        self.file_path_var = tk.StringVar()
        ttk.Label(file_frame, text="File Path:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse...", command=self.select_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Sheet selection
        self.sheet_frame = ttk.Frame(file_frame)
        self.sheet_frame.grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(self.sheet_frame, text="Sheet:").pack(side=tk.LEFT, padx=5)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.sheet_frame, textvariable=self.sheet_var, state="readonly", width=30)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # Hide sheet selection initially
        self.sheet_frame.pack_forget()
        
        # Options frame
        options_frame = ttk.LabelFrame(self.content_frame, text="Loading Options", padding=10)
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Header option
        self.has_header_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="First row contains headers",
            variable=self.has_header_var,
            command=self.refresh_preview
        ).pack(anchor="w", padx=5, pady=2)
        
        # Skip rows option
        skip_frame = ttk.Frame(options_frame)
        skip_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(skip_frame, text="Skip rows:").pack(side=tk.LEFT, padx=5)
        self.skip_rows_var = tk.StringVar(value="0")
        ttk.Spinbox(
            skip_frame,
            from_=0,
            to=100,
            width=5,
            textvariable=self.skip_rows_var,
            command=self.refresh_preview
        ).pack(side=tk.LEFT, padx=5)
        
        # Preview area
        preview_frame = ttk.LabelFrame(self.content_frame, text="Data Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Preview widget
        self.preview = DataPreviewFrame(preview_frame)
        self.preview.pack(fill=tk.BOTH, expand=True)
        
        # Status area
        status_frame = ttk.Frame(self.content_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="Ready to load file")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        
        # Reload button
        ttk.Button(
            status_frame,
            text="Reload",
            command=self.reload_file
        ).pack(side=tk.RIGHT, padx=5)
        
        # Create DataLoader instance
        self.loader = DataLoader()
    
    def select_file(self):
        """Open file dialog to select a file."""
        file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=SUPPORTED_FILE_TYPES
        )
        
        if not file_path:
            return
        
        self.file_path_var.set(file_path)
        self.session_data["loaded_file"] = file_path
        
        # Try to load the file
        self.load_file(file_path)
    
    def load_file(self, file_path):
        """Load the selected file."""
        self.update_status(f"Loading file: {os.path.basename(file_path)}...")
        
        try:
            # Get file extension to determine file type
            _, ext = os.path.splitext(file_path)
            ext = ext.lower()
            
            # Excel file - show sheet selector
            if ext in ['.xlsx', '.xls', '.xlsm']:
                # Get sheet names
                sheet_names = self.loader.get_excel_sheet_names(file_path)
                
                if not sheet_names:
                    messagebox.showerror("Error", "No sheets found in the Excel file.")
                    return
                
                # Update sheet combo
                self.sheet_combo['values'] = sheet_names
                self.sheet_combo.current(0)  # Select first sheet
                
                # Show sheet selection
                self.sheet_frame.grid()
                
                # Load the first sheet
                self.load_sheet(file_path, sheet_names[0])
            
            # CSV file - load directly
            elif ext in ['.csv']:
                # Hide sheet selection
                self.sheet_frame.grid_remove()
                
                # Load CSV data
                self.load_csv(file_path)
            
            else:
                messagebox.showerror("Error", f"Unsupported file type: {ext}")
                return
            
        except Exception as e:
            logger.error(f"Error loading file: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            self.update_status("Error loading file")
    
    def load_sheet(self, file_path, sheet_name):
        """Load a specific sheet from an Excel file."""
        try:
            # Get loading options
            has_header = self.has_header_var.get()
            skip_rows = int(self.skip_rows_var.get())
            
            # Load the sheet
            df = self.loader.load_excel_sheet(
                file_path,
                sheet_name,
                header=0 if has_header else None,
                skiprows=skip_rows
            )
            
            if df is None or df.empty:
                messagebox.showwarning("Warning", f"Sheet '{sheet_name}' is empty.")
                self.update_status(f"Sheet '{sheet_name}' is empty")
                return
            
            # Store in session data
            self.session_data["dataframes"][sheet_name] = df
            self.session_data["current_sheet"] = sheet_name
            
            # Update preview
            self.preview.set_dataframe(df.head(MAX_PREVIEW_ROWS))
            
            # Update status
            self.update_status(
                f"Loaded sheet '{sheet_name}' ({len(df)} rows, {len(df.columns)} columns)"
            )
            
        except Exception as e:
            logger.error(f"Error loading sheet: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load sheet: {str(e)}")
            self.update_status("Error loading sheet")
    
    def load_csv(self, file_path):
        """Load a CSV file."""
        try:
            # Get loading options
            has_header = self.has_header_var.get()
            skip_rows = int(self.skip_rows_var.get())
            
            # Load the CSV
            df = self.loader.load_csv(
                file_path,
                header=0 if has_header else None,
                skiprows=skip_rows
            )
            
            if df is None or df.empty:
                messagebox.showwarning("Warning", "CSV file is empty.")
                self.update_status("CSV file is empty")
                return
            
            # Store in session data
            sheet_name = os.path.basename(file_path)
            self.session_data["dataframes"][sheet_name] = df
            self.session_data["current_sheet"] = sheet_name
            
            # Update preview
            self.preview.set_dataframe(df.head(MAX_PREVIEW_ROWS))
            
            # Update status
            self.update_status(
                f"Loaded CSV file ({len(df)} rows, {len(df.columns)} columns)"
            )
            
        except Exception as e:
            logger.error(f"Error loading CSV: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
            self.update_status("Error loading CSV file")
    
    def on_sheet_selected(self, event=None):
        """Handle sheet selection."""
        sheet_name = self.sheet_var.get()
        if not sheet_name:
            return
        
        file_path = self.file_path_var.get()
        if not file_path:
            return
        
        # Load the selected sheet
        self.load_sheet(file_path, sheet_name)
    
    def refresh_preview(self, event=None):
        """Refresh the data preview based on current options."""
        file_path = self.file_path_var.get()
        if not file_path:
            return
        
        # Reload with current options
        if self.sheet_frame.winfo_ismapped():
            # Excel file
            sheet_name = self.sheet_var.get()
            if sheet_name:
                self.load_sheet(file_path, sheet_name)
        else:
            # CSV file
            self.load_csv(file_path)
    
    def reload_file(self):
        """Reload the current file."""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showinfo("Info", "No file selected.")
            return
        
        self.load_file(file_path)
    
    def on_show(self):
        """Called when the step is shown."""
        # Check if we already have a file loaded
        if self.session_data.get("loaded_file"):
            file_path = self.session_data["loaded_file"]
            self.file_path_var.set(file_path)
            
            # Update preview if we have data
            current_sheet = self.session_data.get("current_sheet")
            if current_sheet and current_sheet in self.session_data.get("dataframes", {}):
                df = self.session_data["dataframes"][current_sheet]
                self.preview.set_dataframe(df.head(MAX_PREVIEW_ROWS))
                
                self.update_status(
                    f"Showing {current_sheet} ({len(df)} rows, {len(df.columns)} columns)"
                )
    
    def validate(self):
        """Validate step data before proceeding."""
        if not self.session_data.get("loaded_file"):
            messagebox.showerror("Error", "Please load a file before proceeding.")
            return False
        
        if not self.session_data.get("dataframes"):
            messagebox.showerror("Error", "No data loaded. Please load a valid file.")
            return False
        
        return True
    
    def save_state(self):
        """Save step state to session data."""
        # State is already saved in session_data during loading
        pass
    
    def show_preview(self):
        """Show a larger data preview dialog."""
        current_sheet = self.session_data.get("current_sheet")
        if not current_sheet or current_sheet not in self.session_data.get("dataframes", {}):
            messagebox.showinfo("Info", "No data loaded to preview.")
            return
        
        df = self.session_data["dataframes"][current_sheet]
        
        # Create dialog
        preview_dialog = tk.Toplevel(self.frame)
        preview_dialog.title(f"Data Preview: {current_sheet}")
        preview_dialog.geometry("800x600")
        
        # Create preview widget
        preview = DataPreviewFrame(preview_dialog)
        preview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Show more rows in the dialog
        preview.set_dataframe(df.head(100))
        
        # Add close button
        ttk.Button(
            preview_dialog,
            text="Close",
            command=preview_dialog.destroy
        ).pack(pady=10)