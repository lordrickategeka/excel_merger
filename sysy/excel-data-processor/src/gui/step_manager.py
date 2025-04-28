#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Step Manager

This module contains the StepManager class which handles navigation between step screens.
"""

import logging

logger = logging.getLogger(__name__)

class StepManager:
    """Manages step-based workflow navigation."""
    
    def __init__(self, steps, update_nav_callback, update_indicators_callback):
        """Initialize the StepManager.
        
        Args:
            steps (list): List of Step objects.
            update_nav_callback (callable): Callback for updating navigation buttons.
            update_indicators_callback (callable): Callback for updating step indicators.
        """
        self.steps = steps
        self.current_step_index = 0
        self.update_nav_callback = update_nav_callback
        self.update_indicators_callback = update_indicators_callback
        
        # Initialize dependencies between steps
        self._setup_dependencies()
    
    def _setup_dependencies(self):
        """Setup dependencies between steps."""
        # Example: Step 2 depends on Step 1, Step 3 depends on Step 2, etc.
        for i in range(1, len(self.steps)):
            self.steps[i].set_dependency(self.steps[i-1])
    
    def show_step(self, step_index):
        """Show the specified step.
        
        Args:
            step_index (int): Index of the step to show.
        """
        if step_index < 0 or step_index >= len(self.steps):
            logger.error(f"Invalid step index: {step_index}")
            return
        
        # Hide current step
        self.steps[self.current_step_index].hide()
        
        # Show new step
        self.current_step_index = step_index
        self.steps[self.current_step_index].show()
        
        # Update UI elements
        self._update_ui()
        
        logger.info(f"Showing step {step_index + 1}: {self.steps[step_index].__class__.__name__}")
    
    def next_step(self):
        """Navigate to the next step if possible."""
        if self.current_step_index < len(self.steps) - 1:
            self.show_step(self.current_step_index + 1)
    
    def previous_step(self):
        """Navigate to the previous step if possible."""
        if self.current_step_index > 0:
            self.show_step(self.current_step_index - 1)
    
    def get_current_step(self):
        """Get the current step object.
        
        Returns:
            object: The current step object.
        """
        return self.steps[self.current_step_index]
    
    def _update_ui(self):
        """Update UI elements based on current step."""
        # Update navigation buttons
        can_go_back = self.current_step_index > 0
        can_go_next = self.current_step_index < len(self.steps) - 1
        
        # Check if next step dependencies are met
        if can_go_next:
            next_step = self.steps[self.current_step_index + 1]
            if hasattr(next_step, "is_dependency_met") and callable(next_step.is_dependency_met):
                can_go_next = next_step.is_dependency_met()
        
        self.update_nav_callback(can_go_back, can_go_next)
        
        # Update step indicators
        self.update_indicators_callback(self.current_step_index)