import os
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import datetime
import re
from core.header_similarity import HeaderSimilarityAnalyzer

class ExcelMerger:
    def __init__(self):
        self.input_folder = None
        self.output_file = None
        self.merged_data = None
        self.all_dataframes = {}  # Store individual dataframes for lookup
        self.skipped_files = []  # Store files that couldn't be processed
        self.processed_folders = {}  # Track which folders were processed
        self.current_sheets = {}  # Store sheets for column merging
        self.header_analyzer = HeaderSimilarityAnalyzer()  # For analyzing similar headers
        
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

    # New method for loading Excel files in a specific folder for column merging
    def load_excel_files_for_merging(self, folder=None):
        """Load Excel files from a folder for column merging operations"""
        if folder is None:
            folder = filedialog.askdirectory(title="Select Folder with Excel Files to Merge")
            if not folder:
                return False
        
        # Reset current sheets
        self.current_sheets = {}
        
        # Find all Excel files in the folder (non-recursive)
        excel_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                     if f.endswith(('.xlsx', '.xls', '.xlsm')) and os.path.isfile(os.path.join(folder, f))]
        
        if not excel_files:
            messagebox.showinfo("No Files Found", "No Excel files found in the selected folder.")
            return False
            
        # Load each Excel file
        for file_path in excel_files:
            file_name = os.path.basename(file_path)
            try:
                # Read all sheets from the Excel file
                xls = pd.ExcelFile(file_path)
                
                # Process each sheet
                for sheet_name in xls.sheet_names:
                    sheet_key = f"{file_name}_{sheet_name}"
                    self.current_sheets[sheet_key] = pd.read_excel(file_path, sheet_name=sheet_name)
                    print(f"Loaded sheet: {sheet_key}")
            except Exception as e:
                print(f"Error loading {file_path}: {str(e)}")
                messagebox.showwarning("Warning", f"Could not load {file_name}: {str(e)}")
        
        if not self.current_sheets:
            messagebox.showinfo("No Data", "No valid Excel sheets were found or could be loaded.")
            return False
            
        return True
    
    # Method to analyze similar columns across loaded sheets
    def analyze_similar_columns(self):
        """Analyze columns across all loaded sheets to find similar field names"""
        if not self.current_sheets:
            messagebox.showinfo("No Data", "No Excel sheets are currently loaded.")
            return None
            
        # Collect all column names from all sheets
        all_columns = {}
        
        for sheet_name, df in self.current_sheets.items():
            all_columns[sheet_name] = list(df.columns)
        
        # Flatten the list of all columns for similarity analysis
        flat_columns = []
        for cols in all_columns.values():
            flat_columns.extend(cols)
            
        # Use HeaderSimilarityAnalyzer to find similar columns
        self.header_analyzer.set_similarity_threshold(0.7)  # Default threshold
        results, suggestion_text = self.header_analyzer.analyze_and_suggest_merges(flat_columns)
        
        return {
            'sheet_columns': all_columns,
            'similarity_results': results,
            'suggestion_text': suggestion_text
        }
    
    # Method to merge columns across sheets
    def merge_columns_across_sheets(self, column_mapping, output_column_name=None):
        """
        Merge similar columns from different sheets into a single dataset
        
        Args:
            column_mapping: Dictionary mapping target column name to list of source columns
            output_column_name: Name for the merged output column (defaults to first key in mapping)
        
        Returns:
            DataFrame with merged data
        """
        if not self.current_sheets or not column_mapping:
            return None
            
        if not output_column_name:
            # Use the first key as the output column name
            output_column_name = next(iter(column_mapping))
            
        # Create list to store dataframes with standardized column names
        standardized_dfs = []
        
        for sheet_name, df in self.current_sheets.items():
            # Create a copy to avoid modifying the original
            new_df = df.copy()
            
            # Flag to track if this sheet has any of the target columns
            has_target_column = False
            
            # Rename columns according to the mapping
            for target_col, source_cols in column_mapping.items():
                for source_col in source_cols:
                    if source_col in df.columns:
                        # Rename the source column to the target name
                        new_df = new_df.rename(columns={source_col: target_col})
                        has_target_column = True
                        # Add sheet identifier column
                        new_df['Source_Sheet'] = sheet_name
                        break  # Only rename the first matching column in this sheet
            
            # Only add sheets that have at least one of the target columns
            if has_target_column:
                standardized_dfs.append(new_df)
        
        # Merge all dataframes
        if standardized_dfs:
            merged_df = pd.concat(standardized_dfs, ignore_index=True)
            return merged_df
        else:
            return None
    
    # Method to show the column merge UI
    def show_column_merge_ui(self):
        """Show a UI for merging similar columns across sheets"""
        # First load Excel files
        if not self.load_excel_files_for_merging():
            return
            
        # Analyze columns
        analysis = self.analyze_similar_columns()
        if not analysis:
            return
            
        # Create UI for column merging
        merge_window = tk.Toplevel()
        merge_window.title("Merge Similar Columns Across Sheets")
        merge_window.geometry("800x600")
        
        # Create frames
        top_frame = ttk.Frame(merge_window, padding="10")
        top_frame.pack(fill="x")
        
        # Instructions
        ttk.Label(
            top_frame,
            text="This tool identifies similar columns across sheets and merges them into a single dataset.",
            wraplength=750
        ).pack