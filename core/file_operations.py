import os
import pandas as pd
import datetime
from tkinter import filedialog

class FileOperations:
    """
    Handles file operations like opening and saving Excel files
    """
    @staticmethod
    def select_file():
        """
        Opens a file dialog to select an Excel file
            
        Returns:
            Selected file path or None
        """
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        
        return file_path if file_path else None
    
    @staticmethod
    def save_merged_file(merger, file_dialog_func=None):
        """
        Save the merged data to a new Excel file
        
        Args:
            merger: ExcelColumnMerger instance
            file_dialog_func: Optional function to open a save file dialog
                
        Returns:
            Path to saved file or None if cancelled
        """
        if not merger.current_sheets:
            return None
            
        # Generate default filename with timestamp
        input_dir = os.path.dirname(merger.input_file)
        input_filename = os.path.basename(merger.input_file)
        basename, ext = os.path.splitext(input_filename)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{basename}_merged_{timestamp}{ext}"
        default_path = os.path.join(input_dir, default_filename)
        
        # Ask user where to save the file
        if file_dialog_func:
            output_file = file_dialog_func(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename,
                initialdir=input_dir,
                title="Save Merged File"
            )
        else:
            # Use the standard tkinter filedialog
            from tkinter import filedialog
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename,
                initialdir=input_dir,
                title="Save Merged File"
            )
        
        if not output_file:
            return None
            
        # Save the merged data
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in merger.current_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            return output_file
        except Exception as e:
            raise Exception(f"Failed to save file: {str(e)}")
        
    @staticmethod
    def open_file(file_path):
        """
        Open a file with the default application
        
        Args:
            file_path: Path to the file to open
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if os.name == 'nt':  # Windows
                os.startfile(file_path)
            else:  # Mac/Linux
                import subprocess
                subprocess.call(('open', file_path) if os.name == 'posix' else ('xdg-open', file_path))
            return True
        except Exception as e:
            raise Exception(f"Could not open file: {str(e)}")