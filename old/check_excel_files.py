import os
import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import threading

class ExcelCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Files Column Checker")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to check Excel files")
        
        # Columns to check for
        self.check_columns = [
            'Position', 'Vial', 'FreezerName', 'Personel', 'Box', 'SAMPLE ID', 
            'Barcode', 'Quality Label', 'Quality', 'Aliquot ID', 'Row', 'Column', 'GUID'
        ]
        
        # Minimum number of columns required
        self.min_columns_required = 2
        
        self.valid_files = []
        self.invalid_files = []
        self.all_data = []
        
        # Create UI elements
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="Input", padding="10")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Select folder containing Excel files:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(input_frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Output file (optional):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Required columns display
        columns_frame = ttk.LabelFrame(main_frame, text="Column Requirements", padding="10")
        columns_frame.pack(fill=tk.X, pady=5)
        
        columns_text = ", ".join(self.check_columns)
        columns_label = ttk.Label(columns_frame, text=columns_text, wraplength=750)
        columns_label.pack(fill=tk.X, pady=5)
        ttk.Label(columns_frame, 
                  text=f"Files must contain AT LEAST {self.min_columns_required} of these columns to be valid.", 
                  font=('', 9, 'bold')).pack(pady=2)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Check Files", command=self.start_checking).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export Sorted Data", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Results", command=self.clear_results).pack(side=tk.LEFT, padx=5)
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Log tab
        log_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(log_frame, text="Log")
        
        # Create a scrollable text area for logs
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, yscrollcommand=log_scroll.set)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        # Valid files tab
        valid_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(valid_frame, text="Valid Files")
        
        valid_scroll = ttk.Scrollbar(valid_frame)
        valid_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.valid_files_list = tk.Listbox(valid_frame, yscrollcommand=valid_scroll.set, font=('', 9))
        self.valid_files_list.pack(fill=tk.BOTH, expand=True)
        valid_scroll.config(command=self.valid_files_list.yview)
        
        # Invalid files tab
        invalid_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(invalid_frame, text="Invalid Files")
        
        invalid_scroll = ttk.Scrollbar(invalid_frame)
        invalid_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.invalid_files_list = tk.Listbox(invalid_frame, yscrollcommand=invalid_scroll.set, font=('', 9))
        self.invalid_files_list.pack(fill=tk.BOTH, expand=True)
        invalid_scroll.config(command=self.invalid_files_list.yview)
        
        # Status bar
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        if folder:
            self.folder_path.set(folder)
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Sorted Data As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
    
    def start_checking(self):
        folder_path = self.folder_path.get()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Error", "Please select a valid folder")
            return
        
        # Clear previous results
        self.clear_results()
        
        # Run the check in a separate thread to avoid freezing the UI
        self.status_var.set("Checking files...")
        threading.Thread(target=self.check_excel_files, daemon=True).start()
    
    def check_excel_files(self):
        folder_path = self.folder_path.get()
        
        # Find all Excel files in the folder
        excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))
        
        if not excel_files:
            self.log("No Excel files found in the selected folder")
            self.status_var.set("No Excel files found")
            return
        
        self.log(f"Found {len(excel_files)} Excel file(s) in the folder.")
        self.log(f"Checking for columns: {', '.join(self.check_columns)}")
        self.log(f"Files must have at least {self.min_columns_required} of these columns to be considered valid.")
        
        # Process each Excel file
        for file_path in excel_files:
            file_name = os.path.basename(file_path)
            try:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Check which of the required columns are present
                present_columns = [col for col in self.check_columns if col in df.columns]
                missing_columns = [col for col in self.check_columns if col not in df.columns]
                
                # Consider valid if at least min_columns_required are present
                if len(present_columns) >= self.min_columns_required:
                    message = f"✓ {file_name} - Valid ({len(present_columns)} columns found: {', '.join(present_columns)})"
                    self.log(message)
                    
                    self.valid_files.append(file_name)
                    self.all_data.append(df)
                    
                    # Update UI in the main thread
                    display_msg = f"{file_name} - {len(present_columns)} columns found"
                    self.root.after(0, lambda msg=display_msg: self.valid_files_list.insert(tk.END, msg))
                else:
                    message = f"✗ {file_name} - Invalid (only {len(present_columns)} columns found: {', '.join(present_columns)})"
                    self.log(message)
                    self.invalid_files.append((file_name, missing_columns))
                    
                    # Update UI in the main thread
                    display_msg = f"{file_name} - Only {len(present_columns)} columns found"
                    self.root.after(0, lambda msg=display_msg: self.invalid_files_list.insert(tk.END, msg))
            except Exception as e:
                error_msg = f"Error processing {file_name}: {str(e)}"
                self.log(error_msg)
                self.invalid_files.append((file_name, ["Error: " + str(e)]))
                
                # Update UI in the main thread
                self.root.after(0, lambda msg=error_msg: self.invalid_files_list.insert(tk.END, msg))
        
        # Print summary
        self.log("\n--- Summary ---")
        self.log(f"Total files checked: {len(excel_files)}")
        self.log(f"Valid files (with at least {self.min_columns_required} required columns): {len(self.valid_files)}")
        self.log(f"Invalid files: {len(self.invalid_files)}")
        
        # Update status
        summary = f"Checked {len(excel_files)} files - {len(self.valid_files)} valid, {len(self.invalid_files)} invalid"
        self.root.after(0, lambda: self.status_var.set(summary))
    
    def export_data(self):
        if not self.all_data:
            messagebox.showinfo("Info", "No valid data to export")
            return
        
        output_path = self.output_path.get()
        if not output_path:
            # Generate default filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"combined_data_{timestamp}.xlsx"
            self.output_path.set(output_path)
        
        try:
            # Combine all valid dataframes
            combined_df = pd.concat(self.all_data, ignore_index=True)
            
            # Sort by specified columns if they exist
            sort_columns = ['FreezerName', 'Box', 'Position']
            available_sort_columns = [col for col in sort_columns if col in combined_df.columns]
            
            if available_sort_columns:
                self.log(f"Sorting data by: {', '.join(available_sort_columns)}")
                combined_df = combined_df.sort_values(by=available_sort_columns)
            else:
                self.log("No sort columns available in the data")
            
            # Export to the specified output file
            if not output_path.lower().endswith(('.xlsx', '.xls')):
                output_path += '.xlsx'
            
            combined_df.to_excel(output_path, index=False)
            self.log(f"\nExported sorted data to {output_path}")
            messagebox.showinfo("Success", f"Exported sorted data to {output_path}")
        except Exception as e:
            self.log(f"Error exporting data: {str(e)}")
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def clear_results(self):
        self.log_text.delete(1.0, tk.END)
        self.valid_files_list.delete(0, tk.END)
        self.invalid_files_list.delete(0, tk.END)
        self.valid_files = []
        self.invalid_files = []
        self.all_data = []
        self.status_var.set("Ready to check Excel files")
    
    def log(self, message):
        self.root.after(0, lambda: self.log_text.insert(tk.END, message + "\n"))
        self.root.after(0, lambda: self.log_text.see(tk.END))

def main():
    root = tk.Tk()
    app = ExcelCheckerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()