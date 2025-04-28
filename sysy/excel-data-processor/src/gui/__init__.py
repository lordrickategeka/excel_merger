#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Base Step Class

This module contains the BaseStep class which all step screens should inherit from.
"""

import tkinter as tk
from tkinter import ttk
import logging

logger = logging.getLogger(__name__)

class BaseStep:
    """Base class for all step screens."""
    
    def __init__(self, parent, session_data, update_status_callback):
        """Initialize the base step.
        
        Args:
            parent (tk.Frame): Parent frame to place this step in.
            session_data (dict): Shared session data.
            update_status_callback (callable): Callback to update status bar.
        """
        self.parent = parent
        self.session_data = session_data
        self.update_status = update_status_callback
        self.dependency = None
        
        # Create the step frame which will contain all the step's widgets
        self.frame = ttk.Frame(parent)
        
        # Add a title label
        self.title = self._get_title()
        title_font = ("Helvetica", 14, "bold")
        self.title_label = ttk.Label(
            self.frame,
            text=self.title,
            font=title_font
        )
        self.title_label.pack(anchor="w", pady=(0, 10))
        
        # Add a description label
        self.description = self._get_description()
        self.description_label = ttk.Label(
            self.frame,
            text=self.description,
            wraplength=600
        )
        self.description_label.pack(anchor="w", pady=(0, 15))
        
        # Add a separator
        ttk.Separator(self.frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Content area - to be filled by subclasses
        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Initialize the step's UI
        self._init_ui()
        
        # Log step creation
        logger.debug(f"Created step: {self.__class__.__name__}")
    
    def _get_title(self):
        """Get the step title. Should be overridden by subclasses.
        
        Returns:
            str: The step title.
        """
        return "Step Title"
    
    def _get_description(self):
        """Get the step description. Should be overridden by subclasses.
        
        Returns:
            str: The step description.
        """
        return "Step description goes here."
    
    def _init_ui(self):
        """Initialize the step UI. Must be implemented by subclasses."""
        raise NotImplementedError("Subclasses must implement _init_ui method")
    
    def show(self):
        """Show this step."""
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.on_show()
    
    def hide(self):
        """Hide this step."""
        self.frame.pack_forget()
        self.on_hide()
    
    def on_show(self):
        """Called when the step is shown. Can be overridden by subclasses."""
        pass
    
    def on_hide(self):
        """Called when the step is hidden. Can be overridden by subclasses."""
        pass
    
    def validate(self):
        """Validate the step data before proceeding to the next step.
        
        Returns:
            bool: True if validation passes, False otherwise.
        """
        return True
    
    def save_state(self):
        """Save the step state to the session data. Can be overridden by subclasses."""
        pass
    
    def set_dependency(self, step):
        """Set the dependency for this step.
        
        Args:
            step (BaseStep): The step this step depends on.
        """
        self.dependency = step
    
    def is_dependency_met(self):
        """Check if this step's dependency is met.
        
        Returns:
            bool: True if dependency is met or there is no dependency.
        """
        if self.dependency is None:
            return True
        
        # Custom dependency logic can be implemented by subclasses
        return True