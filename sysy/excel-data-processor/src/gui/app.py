#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Main Application Window

This module contains the main application window class for the Excel Data Processor.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import logging

from src.gui.step_manager import StepManager
from src.gui.steps.step1_load import LoadDataStep
from src.gui.steps.step2_quality import DataQualityStep
from src.gui.steps.step3_merge import MergeColumnsStep
from src.gui.steps.step4_analyze import AnalyzeDataStep
from src.gui.steps.step5_save import SaveResultsStep
from config.settings import STEP_TITLES, PADDING

logger = logging.getLogger(__name__)

class ExcelDataProcessorApp:
    """Main application class for Excel Data Processor."""
    
    def __init__(self, root):
        """Initialize the application.
        
        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Session data - shared across steps
        self.session_data = {
            "loaded_file": None,
            "dataframes": {},  # Sheet name -> DataFrame mapping
            "current_sheet": None,
            "analysis_results": {},
            "quality_results": {},
            "merged_columns": [],
            "removed_columns": []
        }
        
        self._create_widgets()
        self._setup_menu()
        
        # Log application start
        logger.info("Excel Data Processor initialized")
    
    def _create_widgets(self):
        """Create the main application widgets."""
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding=PADDING)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, PADDING))
        
        # App title
        title_font = ("Helvetica", 16, "bold")
        self.title_label = ttk.Label(
            self.header_frame, 
            text="Excel Data Processor",
            font=title_font
        )
        self.title_label.pack(side=tk.LEFT)
        
        # Progress indicator - step numbers
        self.progress_frame = ttk.Frame(self.header_frame)
        self.progress_frame.pack(side=tk.RIGHT)
        
        self.step_indicators = []
        for i, title in enumerate(STEP_TITLES, 1):
            step_indicator = ttk.Label(
                self.progress_frame,
                text=f"Step {i}",
                padding=5,
                relief="raised"
            )
            step_indicator.pack(side=tk.LEFT, padx=2)
            self.step_indicators.append(step_indicator)
        
        # Separator
        ttk.Separator(self.main_frame, orient='horizontal').pack(fill=tk.X, pady=PADDING)
        
        # Content area - will contain the steps
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Navigation buttons at the bottom
        self.nav_frame = ttk.Frame(self.main_frame)
        self.nav_frame.pack(fill=tk.X, pady=PADDING)
        
        self.back_button = ttk.Button(
            self.nav_frame, 
            text="← Back",
            command=self._on_back_clicked,
            state=tk.DISABLED
        )
        self.back_button.pack(side=tk.LEFT)
        
        self.next_button = ttk.Button(
            self.nav_frame, 
            text="Next →",
            command=self._on_next_clicked
        )
        self.next_button.pack(side=tk.RIGHT)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(
            self.root, 
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize step manager and steps
        self._init_steps()
    
    def _init_steps(self):
        """Initialize the step manager and step screens."""
        # Create steps
        steps = [
            LoadDataStep(self.content_frame, self.session_data, self.update_status),
            DataQualityStep(self.content_frame, self.session_data, self.update_status),
            MergeColumnsStep(self.content_frame, self.session_data, self.update_status),
            AnalyzeDataStep(self.content_frame, self.session_data, self.update_status),
            SaveResultsStep(self.content_frame, self.session_data, self.update_status)
        ]
        
        # Create step manager
        self.step_manager = StepManager(steps, self.update_nav_buttons, self.update_step_indicators)
        
        # Show the first step
        self.step_manager.show_step(0)
    
    def _setup_menu(self):
        """Setup the application menu."""
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Session", command=self._new_session)
        file_menu.add_command(label="Open File...", command=self._load_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Settings", command=self._show_settings)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Data Preview", command=self._show_data_preview)
        view_menu.add_command(label="Column Summary", command=self._show_column_summary)
        menubar.add_cascade(label="View", menu=view_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="User Guide", command=self._show_help)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def update_status(self, message):
        """Update the status bar message.
        
        Args:
            message (str): Message to display in the status bar.
        """
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def update_nav_buttons(self, can_go_back, can_go_next):
        """Update navigation button states.
        
        Args:
            can_go_back (bool): Whether the back button should be enabled.
            can_go_next (bool): Whether the next button should be enabled.
        """
        self.back_button.config(state=tk.NORMAL if can_go_back else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if can_go_next else tk.DISABLED)
    
    def update_step_indicators(self, current_step):
        """Update step indicators to highlight the current step.
        
        Args:
            current_step (int): Index of the current step.
        """
        for i, indicator in enumerate(self.step_indicators):
            if i == current_step:
                indicator.configure(relief="sunken", background="#ccccff")
            elif i < current_step:
                indicator.configure(relief="flat", background="#e6e6ff")
            else:
                indicator.configure(relief="raised", background="")
        
        # Update window title with step
        self.root.title(f"Excel Data Processor - {STEP_TITLES[current_step]}")
    
    def _on_back_clicked(self):
        """Handle back button click."""
        self.step_manager.previous_step()
    
    def _on_next_clicked(self):
        """Handle next button click."""
        # Call validation on current step before proceeding
        current_step = self.step_manager.get_current_step()
        if current_step.validate():
            # Save current step data before moving to next step
            current_step.save_state()
            self.step_manager.next_step()
    
    def _new_session(self):
        """Start a new session, clearing all data."""
        if messagebox.askyesno("New Session", "Start a new session? This will clear all current data."):
            # Reset session data
            self.session_data = {
                "loaded_file": None,
                "dataframes": {},
                "current_sheet": None,
                "analysis_results": {},
                "quality_results": {},
                "merged_columns": [],
                "removed_columns": []
            }
            
            # Reset steps
            self._init_steps()
            
            # Show first step
            self.step_manager.show_step(0)
            
            self.update_status("New session started")
    
    def _load_file(self):
        """Load a file through the File menu."""
        # Delegate to the load step's file loading function
        load_step = self.step_manager.steps[0]
        load_step.select_file()
        
        # If a file was loaded, show the load step
        if self.session_data["loaded_file"]:
            self.step_manager.show_step(0)
    
    def _show_settings(self):
        """Show the settings dialog."""
        # Placeholder for settings dialog
        messagebox.showinfo("Settings", "Settings dialog not implemented yet.")
    
    def _show_data_preview(self):
        """Show a data preview dialog."""
        # Delegate to the current step if it has a preview method
        current_step = self.step_manager.get_current_step()
        if hasattr(current_step, "show_preview") and callable(current_step.show_preview):
            current_step.show_preview()
        else:
            messagebox.showinfo("Data Preview", "No data preview available at this step.")
    
    def _show_column_summary(self):
        """Show a column summary dialog."""
        # Check if we have data loaded
        if not self.session_data["dataframes"]:
            messagebox.showinfo("Column Summary", "No data loaded yet.")
            return
        
        # Placeholder for column summary dialog
        messagebox.showinfo("Column Summary", "Column summary not implemented yet.")
    
    def _show_help(self):
        """Show help documentation."""
        # Placeholder for help window
        messagebox.showinfo("Help", "User guide not implemented yet.")
    
    def _show_about(self):
        """Show about dialog."""
        messagebox.showinfo("About", "Excel Data Processor v1.0.0\n\nA comprehensive tool for processing, analyzing, and visualizing Excel data.")
    
    def on_close(self):
        """Handle application close."""
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            # Clean up resources
            logger.info("Excel Data Processor shutting down")
            self.root.destroy()