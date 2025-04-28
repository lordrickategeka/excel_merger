import tkinter as tk
import sys
import os

# Add parent directory to path for imports to work
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.merger import ExcelColumnMerger
from ui.main_window import MainWindow

def main():
    """
    Main application entry point
    """
    # Create the main window
    root = tk.Tk()
    
    # Initialize merger
    merger = ExcelColumnMerger()
    
    # Initialize the main window
    app = MainWindow(root, merger)
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()