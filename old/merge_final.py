import os
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import datetime
import re

class ExcelMerger:
    def __init__(self):
        self.input_folder = None
        self.output_file = None
        self.merged_data = None
        self.all_dataframes = {}  # Store individual dataframes for lookup
        self.skipped_files = []  # Store files that couldn't be processed
        self.processed_folders = {}  # Track which folders were processed
        
    def select_folder(self):
        """Let user select the folder containing Excel files"""
        self.input_folder = filedialog.askdirectory(title="Select Root Folder with Excel/CSV/Text Files")
        return self.input_folder
    
    def analyze_folder_recursive(self):
        """Analyze the folder and subfolders recursively for Excel, CSV, and text files"""
        if not self.input_folder:
            return None
        
        files = []
        self.processed_folders = {}  # Reset folder tracking
        
        # Walk through all directories recursively
        for root, dirs, filenames in os.walk(self.input_folder):
            folder_files = []
            
            for file in filenames:
                # Include Excel, CSV, and TXT files
                if file.endswith(('.xlsx', '.xls', '.xlsm', '.csv', '.txt')):
                    full_path = os.path.join(root, file)
                    files.append(full_path)
                    folder_files.append(file)
            
            # Record how many files found in this folder
            if folder_files:
                rel_path = os.path.relpath(root, self.input_folder)
                if rel_path == '.':
                    rel_path = 'Root Folder'
                self.processed_folders[rel_path] = folder_files
                
        return {
            'folder': self.input_folder,
            'file_count': len(files),
            'files': [os.path.basename(f) for f in files],
            'full_paths': files
        }
    
    def merge_files(self, analysis):
        """Merge all Excel, CSV, and text files in the folder"""
        if not analysis or analysis['file_count'] == 0:
            return False
            
        # Create an empty list to hold all dataframes
        all_data = []
        self.skipped_files = []  # Reset skipped files
        
        # Process each file
        for file_path in analysis['full_paths']:
            try:
                file_name = os.path.basename(file_path)
                rel_folder = os.path.relpath(os.path.dirname(file_path), self.input_folder)
                if rel_folder == '.':
                    rel_folder = 'Root'
                
                if file_path.lower().endswith('.csv'):
                    # Read CSV file
                    df = pd.read_csv(file_path)
                    df['Source_File'] = file_name
                    df['Source_Folder'] = rel_folder
                    df['Sheet_Name'] = 'CSV'  # CSV files don't have sheets
                    all_data.append(df)
                    
                    # Store for lookup
                    self.all_dataframes[f"{rel_folder}|{file_name}|CSV"] = df
                elif file_path.lower().endswith('.txt'):
                    # Read text file with various delimiters
                    # Try to detect delimiter by reading first few lines
                    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                        sample = ''.join([f.readline() for _ in range(5)])
                    
                    # Try to guess the delimiter
                    delimiter = None
                    for potential_delim in [',', '\t', '|', ';', ' ']:
                        if potential_delim in sample:
                            counts = sample.count(potential_delim)
                            if counts > 2:  # At least need a few occurrences
                                delimiter = potential_delim
                                break
                    
                    if delimiter:
                        # Try pandas read_csv with the detected delimiter
                        df = pd.read_csv(file_path, delimiter=delimiter, engine='python', error_bad_lines=False)
                    else:
                        # If delimiter detection fails, try to read it as a fixed-width or space-delimited file
                        df = pd.read_fwf(file_path)
                    
                    # If we have only one column, it might be unstructured text - convert to dataframe
                    if len(df.columns) == 1:
                        # Read raw text and create a proper dataframe
                        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                            lines = f.readlines()
                        
                        # Create dataframe with lines as rows
                        df = pd.DataFrame({'Text_Content': lines})
                    
                    df['Source_File'] = file_name
                    df['Source_Folder'] = rel_folder
                    df['Sheet_Name'] = 'TXT'  # Text files don't have sheets
                    all_data.append(df)
                    
                    # Store for lookup
                    self.all_dataframes[f"{rel_folder}|{file_name}|TXT"] = df
                else:
                    # Read Excel file with multiple sheets
                    xls = pd.ExcelFile(file_path)
                    
                    # Process each sheet in the Excel file
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        
                        # Add file and sheet info as columns
                        df['Source_File'] = file_name
                        df['Source_Folder'] = rel_folder
                        df['Sheet_Name'] = sheet_name
                        
                        # Add to our list
                        all_data.append(df)
                        
                        # Store for lookup
                        self.all_dataframes[f"{rel_folder}|{file_name}|{sheet_name}"] = df
            except Exception as e:
                print(f"Error processing {file_path}: {str(e)}")
                # Record the skipped file with error details
                self.skipped_files.append({
                    'file': os.path.basename(file_path),
                    'folder': os.path.relpath(os.path.dirname(file_path), self.input_folder),
                    'error': str(e)
                })
                
        # Merge all dataframes
        if all_data:
            self.merged_data = pd.concat(all_data, ignore_index=True)
            return True
        return False
    
    def save_merged_file(self):
        """Save the merged data to a new Excel file"""
        if self.merged_data is None:
            return None
            
        # Generate default filename with timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"merged_excel_{timestamp}.xlsx"
        
        # Ask user where to save the file
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=default_filename,
            title="Save Merged File"
        )
        
        if not self.output_file:
            return None
            
        # Save the merged data
        if self.output_file.lower().endswith('.csv'):
            self.merged_data.to_csv(self.output_file, index=False)
        else:
            writer = pd.ExcelWriter(self.output_file, engine='openpyxl')
            
            # Save main data
            self.merged_data.to_excel(writer, sheet_name='Merged_Data', index=False)
            
            # Save skipped files report if any
            if self.skipped_files:
                skipped_df = pd.DataFrame(self.skipped_files)
                skipped_df.to_excel(writer, sheet_name='Skipped_Files', index=False)
                
            # Save folder summary
            folders_data = []
            for folder, files in self.processed_folders.items():
                folders_data.append({
                    'Folder': folder,
                    'Files_Count': len(files),
                    'Files': ', '.join(files)
                })
            
            folders_df = pd.DataFrame(folders_data)
            folders_df.to_excel(writer, sheet_name='Folders_Summary', index=False)
                
            writer.close()
            
        return self.output_file
    
    def show_skipped_files(self):
        """Show a window with skipped files"""
        if not self.skipped_files:
            messagebox.showinfo("Skipped Files", "No files were skipped during processing.")
            return
            
        # Create window to show skipped files
        skip_window = tk.Toplevel()
        skip_window.title("Files Not Processed")
        skip_window.geometry("700x500")
        
        # Create treeview
        columns = ("file", "folder", "error")
        tree = ttk.Treeview(skip_window, columns=columns, show="headings")
        
        # Define headings
        tree.heading("file", text="File")
        tree.heading("folder", text="Folder")
        tree.heading("error", text="Error")
        
        # Set column widths
        tree.column("file", width=150)
        tree.column("folder", width=200)
        tree.column("error", width=300)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(skip_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(skip_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Pack widgets
        tree.pack(side="top", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        
        # Insert data
        for item in self.skipped_files:
            tree.insert("", "end", values=(item['file'], item['folder'], item['error']))
            
        # Export button
        def export_skipped():
            export_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                title="Export Skipped Files List"
            )
            
            if not export_file:
                return
                
            skipped_df = pd.DataFrame(self.skipped_files)
            
            if export_file.lower().endswith('.csv'):
                skipped_df.to_csv(export_file, index=False)
            else:
                skipped_df.to_excel(export_file, index=False)
                
            messagebox.showinfo("Export Complete", f"Skipped files list exported to {export_file}")
            
        ttk.Button(skip_window, text="Export List", command=export_skipped).pack(pady=10)
    
    def show_folder_summary(self):
        """Show a window with folder summary"""
        if not self.processed_folders:
            messagebox.showinfo("Folder Summary", "No folders were processed.")
            return
            
        # Create window
        folder_window = tk.Toplevel()
        folder_window.title("Processed Folders Summary")
        folder_window.geometry("700x500")
        
        # Create treeview
        tree = ttk.Treeview(folder_window)
        tree["columns"] = ("files", "count")
        tree["show"] = "tree headings"
        
        tree.heading("#0", text="Folder")
        tree.heading("files", text="Files")
        tree.heading("count", text="Count")
        
        tree.column("#0", width=250)
        tree.column("files", width=350)
        tree.column("count", width=70)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(folder_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(folder_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Pack widgets
        tree.pack(side="top", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        
        # Insert data
        for folder, files in self.processed_folders.items():
            tree.insert("", "end", text=folder, values=(", ".join(files[:3]) + ("..." if len(files) > 3 else ""), len(files)))
            
        # Export button
        def export_folders():
            export_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                title="Export Folder Summary"
            )
            
            if not export_file:
                return
                
            folders_data = []
            for folder, files in self.processed_folders.items():
                folders_data.append({
                    'Folder': folder,
                    'Files_Count': len(files),
                    'Files': ', '.join(files)
                })
            
            folders_df = pd.DataFrame(folders_data)
            
            if export_file.lower().endswith('.csv'):
                folders_df.to_csv(export_file, index=False)
            else:
                folders_df.to_excel(export_file, index=False)
                
            messagebox.showinfo("Export Complete", f"Folder summary exported to {export_file}")
            
        ttk.Button(folder_window, text="Export Summary", command=export_folders).pack(pady=10)
    
    def perform_lookup(self):
        """Open a window to perform lookups across all loaded files"""
        if not self.all_dataframes:
            messagebox.showerror("Error", "No data loaded to search in.")
            return
        
        # Create lookup window
        lookup_window = tk.Toplevel()
        lookup_window.title("Data Lookup Tool")
        lookup_window.geometry("900x600")
        
        # Create frames
        search_frame = ttk.Frame(lookup_window, padding="10")
        search_frame.pack(fill="x")
        
        results_frame = ttk.Frame(lookup_window, padding="10")
        results_frame.pack(fill="both", expand=True)
        
        # Create search controls
        ttk.Label(search_frame, text="Search Term:").grid(row=0, column=0, padx=5, pady=5)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Column selection
        ttk.Label(search_frame, text="Search In:").grid(row=0, column=2, padx=5, pady=5)
        search_option_var = tk.StringVar(value="All Columns")
        
        # Collect all possible columns from all dataframes
        all_columns = set()
        for df in self.all_dataframes.values():
            all_columns.update(df.columns)
        
        column_options = ["All Columns"] + sorted(list(all_columns))
        column_dropdown = ttk.Combobox(search_frame, textvariable=search_option_var, values=column_options, width=20)
        column_dropdown.grid(row=0, column=3, padx=5, pady=5)
        
        # Case sensitive option
        case_sensitive_var = tk.BooleanVar(value=False)
        case_option = ttk.Checkbutton(search_frame, text="Case Sensitive", variable=case_sensitive_var)
        case_option.grid(row=0, column=4, padx=5, pady=5)
        
        # Results treeview
        columns = ("folder", "file", "sheet", "row", "column", "value", "context")
        results_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        # Define headings
        results_tree.heading("folder", text="Folder")
        results_tree.heading("file", text="File")
        results_tree.heading("sheet", text="Sheet")
        results_tree.heading("row", text="Row")
        results_tree.heading("column", text="Column")
        results_tree.heading("value", text="Matching Value")
        results_tree.heading("context", text="Context")
        
        # Set column widths
        results_tree.column("folder", width=100)
        results_tree.column("file", width=100)
        results_tree.column("sheet", width=80)
        results_tree.column("row", width=50)
        results_tree.column("column", width=100)
        results_tree.column("value", width=150)
        results_tree.column("context", width=250)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
        results_tree.configure(yscroll=scrollbar.set)
        
        # Pack the treeview and scrollbar
        results_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Export results button
        export_button = ttk.Button(lookup_window, text="Export Results", state="disabled")
        export_button.pack(pady=10)
        
        # Status bar
        status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(lookup_window, textvariable=status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")
        
        # Search function
        def search():
            # Clear previous results
            for item in results_tree.get_children():
                results_tree.delete(item)
            
            search_term = search_var.get()
            if not search_term:
                status_var.set("Please enter a search term")
                return
            
            search_column = search_option_var.get()
            case_sensitive = case_sensitive_var.get()
            
            results_count = 0
            
            for key, df in self.all_dataframes.items():
                folder, file_name, sheet_name = key.split('|')
                
                # Select columns to search
                columns_to_search = [search_column] if search_column != "All Columns" and search_column in df.columns else df.columns
                
                # Convert df to string to allow string searching
                for col in columns_to_search:
                    # Skip non-string columns if they can't be converted to string
                    try:
                        # Case sensitivity handling
                        if case_sensitive:
                            mask = df[col].astype(str).str.contains(search_term, na=False)
                        else:
                            mask = df[col].astype(str).str.contains(search_term, case=False, na=False)
                            
                        # Get matching rows
                        matching_rows = df[mask]
                        
                        for idx, row in matching_rows.iterrows():
                            # Get original row index
                            original_idx = df.index.get_loc(idx) + 2  # +2 for Excel row number (header + 1-based index)
                            
                            # Get matching value
                            match_value = str(row[col])
                            
                            # Create context (show a few values from the row)
                            context_cols = [c for c in df.columns if c not in ['Source_File', 'Sheet_Name', 'Source_Folder']][:3]
                            context = " | ".join(f"{c}: {row[c]}" for c in context_cols if c in row)
                            
                            # Add to treeview
                            results_tree.insert("", "end", values=(
                                folder,
                                file_name, 
                                sheet_name, 
                                original_idx, 
                                col, 
                                match_value[:50] + ('...' if len(match_value) > 50 else ''),
                                context
                            ))
                            results_count += 1
                    except Exception as e:
                        print(f"Error searching column {col}: {str(e)}")
            
            # Update status
            status_var.set(f"Found {results_count} matches")
            
            # Enable export if we have results
            if results_count > 0:
                export_button.configure(state="normal")
            else:
                export_button.configure(state="disabled")
        
        # Export function
        def export_results():
            if not results_tree.get_children():
                return
                
            # Ask for save location
            export_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                title="Export Search Results"
            )
            
            if not export_file:
                return
                
            # Create dataframe from results
            results_data = []
            for item in results_tree.get_children():
                values = results_tree.item(item, "values")
                results_data.append({
                    'Folder': values[0],
                    'File': values[1],
                    'Sheet': values[2],
                    'Row': values[3],
                    'Column': values[4],
                    'Matching Value': values[5],
                    'Context': values[6]
                })
                
            results_df = pd.DataFrame(results_data)
            
            # Save based on extension
            if export_file.lower().endswith('.csv'):
                results_df.to_csv(export_file, index=False)
            else:
                results_df.to_excel(export_file, index=False)
                
            messagebox.showinfo("Export Complete", f"Results exported to {export_file}")
        
        # Connect button to export function
        export_button.configure(command=export_results)
        
        # Connect search button
        search_button = ttk.Button(search_frame, text="Search", command=search)
        search_button.grid(row=0, column=5, padx=5, pady=5)
        
        # Bind Enter key to search
        search_entry.bind("<Return>", lambda event: search())
        
        # Focus search entry
        search_entry.focus_set()
        
        # Double-click to view full row
        def on_double_click(event):
            if not results_tree.selection():
                return
                
            item = results_tree.selection()[0]
            values = results_tree.item(item, "values")
            folder, file_name, sheet_name, row_idx = values[0], values[1], values[2], int(values[3])
            
            # Get the dataframe
            key = f"{folder}|{file_name}|{sheet_name}"
            if key in self.all_dataframes:
                df = self.all_dataframes[key]
                
                # Get the row data (row_idx - 2 because Excel is 1-based and has header)
                row_idx_adjusted = row_idx - 2  # Adjust for Excel row number
                if 0 <= row_idx_adjusted < len(df):
                    row_data = df.iloc[row_idx_adjusted]
                    
                    # Create a temporary window to show full row
                    row_window = tk.Toplevel(lookup_window)
                    row_window.title(f"Row Details - {folder} - {file_name} - {sheet_name} - Row {row_idx}")
                    row_window.geometry("600x400")
                    
                    # Create text widget with scrollbar
                    text_frame = ttk.Frame(row_window)
                    text_frame.pack(fill="both", expand=True, padx=10, pady=10)
                    
                    row_text = tk.Text(text_frame, wrap="word")
                    row_scroll = ttk.Scrollbar(text_frame, command=row_text.yview)
                    row_text.configure(yscrollcommand=row_scroll.set)
                    
                    row_text.pack(side="left", fill="both", expand=True)
                    row_scroll.pack(side="right", fill="y")
                    
                    # Format and insert row data
                    row_text.insert("1.0", f"Folder: {folder}\nFile: {file_name}\nSheet: {sheet_name}\nRow: {row_idx}\n\n")
                    
                    # Insert each column and value
                    for col_name, value in row_data.items():
                        row_text.insert("end", f"{col_name}: {value}\n")
                        
                    row_text.configure(state="disabled")  # Make read-only
        
        # Bind double-click
        results_tree.bind("<Double-1>", on_double_click)

def main():
    # Create the main window but hide it
    root = tk.Tk()
    root.withdraw()
    
    merger = ExcelMerger()
    
    # Create a simple menu
    menu_window = tk.Toplevel(root)
    menu_window.title("Excel, CSV & Text File Merger")
    menu_window.geometry("320x350")
    menu_window.resizable(False, False)
    
    # Center the window
    window_width = 320
    window_height = 350
    screen_width = menu_window.winfo_screenwidth()
    screen_height = menu_window.winfo_screenheight()
    x = (screen_width / 2) - (window_width / 2)
    y = (screen_height / 2) - (window_height / 2)
    menu_window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
    
    # Frame for buttons
    button_frame = ttk.Frame(menu_window, padding="20")
    button_frame.pack(expand=True)
    
    # Style for buttons
    s = ttk.Style()
    s.configure('Big.TButton', font=('Helvetica', 12))
    
    # Status variable
    status_var = tk.StringVar(value="Ready")
    
    # Buttons
    ttk.Button(
        button_frame, 
        text="Merge Excel, CSV & Text Files", 
        command=lambda: process_files(merger, menu_window, status_var),
        style='Big.TButton',
        width=25
    ).pack(pady=10)
    
    ttk.Button(
        button_frame, 
        text="Lookup Data", 
        command=lambda: lookup_data(merger, menu_window, status_var),
        style='Big.TButton',
        width=25
    ).pack(pady=10)
    
    ttk.Button(
        button_frame, 
        text="Show Folder Summary", 
        command=lambda: show_folders(merger, status_var),
        style='Big.TButton',
        width=25
    ).pack(pady=10)
    
    ttk.Button(
        button_frame, 
        text="Show Skipped Files", 
        command=lambda: show_skipped(merger, status_var),
        style='Big.TButton',
        width=25
    ).pack(pady=10)
    
    ttk.Button(
        button_frame, 
        text="Exit", 
        command=root.destroy,
        style='Big.TButton',
        width=25
    ).pack(pady=10)
    
    # Status bar
    status_bar = ttk.Label(menu_window, textvariable=status_var, relief="sunken", anchor="w")
    status_bar.pack(side="bottom", fill="x")
    
    # Prevent closing the root window directly
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    menu_window.protocol("WM_DELETE_WINDOW", root.destroy)
    
    # Functions for menu commands
    def process_files(merger, window, status):
        status.set("Selecting folder...")
        
        # Select folder
        folder = merger.select_folder()
        if not folder:
            status.set("No folder selected")
            return
        
        status.set("Analyzing folders recursively...")
        
        # Analyze folder and subfolders
        analysis = merger.analyze_folder_recursive()
        if not analysis or analysis['file_count'] == 0:
            messagebox.showerror("Error", "No Excel, CSV, or text files found in the selected folders.")
            status.set("No files found")
            return
        
        # Show analysis
        messagebox.showinfo(
            "Folder Analysis", 
            f"Found {analysis['file_count']} Excel/CSV/text files in {len(merger.processed_folders)} folders/subfolders"
        )
        
        status.set(f"Processing {analysis['file_count']} files...")
        
        # Merge files
        if merger.merge_files(analysis):
            status.set("Files merged, saving...")
            
            # Report skipped files if any
            if merger.skipped_files:
                status.set(f"Files merged with {len(merger.skipped_files)} skipped files")
                messagebox.showwarning(
                    "Warning", 
                    f"{len(merger.skipped_files)} files could not be processed. Click 'Show Skipped Files' to see details."
                )
            
            # Save merged file
            output_file = merger.save_merged_file()
            if output_file:
                messagebox.showinfo("Success", f"Merged file saved as:\n{output_file}")
                
                # Ask if the user wants to open the file
                if messagebox.askyesno("Open File", "Do you want to open the merged file?"):
                    os.startfile(output_file) if os.name == 'nt' else os.system(f"xdg-open {output_file}")
                
                status.set(f"Saved to {os.path.basename(output_file)}")
                
                # Ask if they want to see folder summary
                if len(merger.processed_folders) > 1:
                    if messagebox.askyesno("Folder Summary", "Do you want to see the folder summary?"):
                        merger.show_folder_summary()
            else:
                status.set("Save cancelled")
        else:
            messagebox.showerror("Error", "Failed to merge files.")
            status.set("Merge failed")
    
    def lookup_data(merger, window, status):
        if not merger.all_dataframes:
            # No data loaded, need to select and process files first
            status.set("Loading data for lookup...")
            
            # Select folder
            folder = merger.select_folder()
            if not folder:
                status.set("No folder selected")
                return
            if not analysis or analysis:
            # Analyze folder recursively
                analysis = merger.analyze_folder_recursive()
            if not analysis or analysis['file_count'] == 0:
                messagebox.showerror("Error", "No Excel, CSV, or text files found in the selected folders.")
                status.set("No files found")
                return
                
            # Process files for lookup (don't merge yet)
            if not merger.merge_files(analysis):
                messagebox.showerror("Error", "Failed to process files for lookup.")
                status.set("Processing failed")
                return
                
            status.set(f"Loaded {analysis['file_count']} files for lookup")
        
        # Open lookup window
        merger.perform_lookup()
    
    def show_folders(merger, status):
        if not merger.processed_folders:
            status.set("No folders processed yet")
            messagebox.showinfo("Folder Summary", "No folders have been processed yet. Please merge files first.")
            return
            
        status.set("Showing folder summary")
        merger.show_folder_summary()
    
    def show_skipped(merger, status):
        
        def merge_final_merged_files(merger, status):
            status.set("Selecting folder with final merged files...")
            folder = merger.select_folder()
            if not folder:
                status.set("No folder selected")
                return

            status.set("Merging 'Merged_Data' sheets from Excel files...")
            if merger.merge_merged_excel_files(folder):
                status.set("Files merged, saving...")
                output_file = merger.save_merged_file()
                if output_file:
                    messagebox.showinfo("Success", f"Final merged file saved as:\n{output_file}")
                    if messagebox.askyesno("Open File", "Do you want to open the merged file?"):
                        os.startfile(output_file) if os.name == 'nt' else os.system(f"xdg-open {output_file}")
                    status.set(f"Saved to {os.path.basename(output_file)}")
                else:
                    status.set("Save cancelled")
            else:
                messagebox.showerror("Error", "No valid 'Merged_Data' sheets found.")
                status.set("No valid data found")

                if not merger.skipped_files:
                    status.set("No skipped files")
                    messagebox.showinfo("Skipped Files", "No files were skipped during processing.")
                    return
                
            status.set(f"Showing {len(merger.skipped_files)} skipped files")
            merger.show_skipped_files()
    
    # Start the application
    root.mainloop()

if __name__ == "__main__":
    main()