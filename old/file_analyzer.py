import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
from collections import defaultdict

class ExcelColumnMerger:
    def __init__(self):
        self.input_file = None
        self.output_file = None
        self.df = None
        self.duplicate_columns = None
        self.merge_strategy = tk.StringVar(value="first_non_empty")
        self.current_sheets = {}  # Store current dataframes for each sheet
        
    def select_file(self):
        """Let user select the Excel file to process"""
        self.input_file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if self.input_file:
            try:
                # Read all sheets
                xl = pd.ExcelFile(self.input_file)
                self.current_sheets = {
                    sheet: pd.read_excel(self.input_file, sheet_name=sheet)
                    for sheet in xl.sheet_names
                }
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file: {str(e)}")
                self.input_file = None
                return None
        return self.input_file
    
    def analyze_file(self):
        """Analyze the Excel file for duplicate column names (case-insensitive)"""
        if not self.input_file or not self.current_sheets:
            return None
            
        try:
            sheets_data = {}
            
            for sheet_name, df in self.current_sheets.items():
                # Check for duplicate columns (case-insensitive)
                column_groups = defaultdict(list)
                
                # Group columns by lowercase name
                for col in df.columns:
                    column_groups[str(col).lower()].append(col)
                
                # Keep only groups with more than one column
                duplicate_columns = {k: v for k, v in column_groups.items() if len(v) > 1}
                
                if duplicate_columns:
                    sheets_data[sheet_name] = {
                        'dataframe': df,
                        'duplicate_columns': duplicate_columns
                    }
            
            if not sheets_data:
                return {"status": "no_duplicates"}
            
            return {
                "status": "ok",
                "sheets_data": sheets_data,
                "sheet_names": list(sheets_data.keys())
            }
            
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }
    
    def merge_columns(self, analysis, strategy="first_non_empty"):
        """Merge duplicate columns based on the selected strategy"""
        if not analysis or analysis["status"] != "ok":
            return False
            
        try:
            sheets_data = analysis["sheets_data"]
            
            for sheet_name, sheet_data in sheets_data.items():
                df = sheet_data['dataframe']
                duplicate_columns = sheet_data['duplicate_columns']
                
                # Process each group of duplicate columns
                for base_name, columns in duplicate_columns.items():
                    # Choose merge strategy
                    if strategy == "first_non_empty":
                        # Create a new column combining non-empty values
                        new_values = df[columns].apply(
                            lambda row: next((x for x in row if pd.notna(x)), None), 
                            axis=1
                        )
                    elif strategy == "sum":
                        # Sum numeric values, ignoring non-numeric
                        numeric_columns = []
                        for col in columns:
                            if pd.api.types.is_numeric_dtype(df[col]):
                                numeric_columns.append(col)
                        
                        if numeric_columns:
                            new_values = df[numeric_columns].sum(axis=1)
                        else:
                            # If no numeric columns, use first_non_empty
                            new_values = df[columns].apply(
                                lambda row: next((x for x in row if pd.notna(x)), None), 
                                axis=1
                            )
                    elif strategy == "concatenate":
                        # Concatenate all non-empty string values
                        new_values = df[columns].apply(
                            lambda row: " ".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]), 
                            axis=1
                        )
                    
                    # Create a new dataframe without the duplicate columns
                    new_df = df.drop(columns=columns)
                    
                    # Add the merged column (use the first duplicate name)
                    new_df[columns[0]] = new_values
                    
                    # Update the dataframe
                    df = new_df
                
                # Update the current sheet
                self.current_sheets[sheet_name] = df
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge columns: {str(e)}")
            return False
    
    def manual_merge_columns(self, sheet_name, columns_to_merge, new_column_name, strategy="first_non_empty", delete_source=True, remove_empty=True):
        """Manually merge selected columns into a new column"""
        if not sheet_name or not columns_to_merge or sheet_name not in self.current_sheets:
            return False
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure all columns exist
            for col in columns_to_merge:
                if col not in df.columns:
                    messagebox.showerror("Error", f"Column '{col}' not found in sheet '{sheet_name}'")
                    return False
            
            # Apply merge strategy
            if strategy == "first_non_empty":
                new_values = df[columns_to_merge].apply(
                    lambda row: next((x for x in row if pd.notna(x)), None), 
                    axis=1
                )
            elif strategy == "sum":
                numeric_columns = []
                for col in columns_to_merge:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        numeric_columns.append(col)
                
                if numeric_columns:
                    new_values = df[numeric_columns].sum(axis=1)
                else:
                    new_values = df[columns_to_merge].apply(
                        lambda row: next((x for x in row if pd.notna(x)), None), 
                        axis=1
                    )
            elif strategy == "concatenate":
                new_values = df[columns_to_merge].apply(
                    lambda row: " ".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]), 
                    axis=1
                )
            
            # Add the new column
            df[new_column_name] = new_values
            
            # Delete source columns if requested
            if delete_source:
                df = df.drop(columns=columns_to_merge)
            
            # Remove empty columns if requested
            if remove_empty:
                # Find columns that are all NaN or empty strings
                empty_cols = []
                for col in df.columns:
                    # Check if column is all NaN or empty strings
                    is_empty = df[col].isna().all() or (
                        (df[col].astype(str).str.strip() == '').all()
                    )
                    if is_empty:
                        empty_cols.append(col)
                
                # Drop empty columns
                if empty_cols:
                    df = df.drop(columns=empty_cols)
                    print(f"Removed {len(empty_cols)} empty columns")
            
            # Update the current sheet
            self.current_sheets[sheet_name] = df
            
            return True, empty_cols if remove_empty else []
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge columns: {str(e)}")
            return False, []
    
    def compare_columns_for_duplicates(self, sheet_name, columns_to_compare):
        """Compare values across selected columns and identify duplicate values"""
        if not sheet_name or not columns_to_compare or sheet_name not in self.current_sheets:
            return None
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure all columns exist
            for col in columns_to_compare:
                if col not in df.columns:
                    messagebox.showerror("Error", f"Column '{col}' not found in sheet '{sheet_name}'")
                    return None
            
            # Create a result dataframe
            result_df = pd.DataFrame(index=df.index)
            
            # Add the columns we're comparing
            for col in columns_to_compare:
                result_df[col] = df[col]
            
            # Find rows with duplicate values across columns
            duplicates_mask = pd.DataFrame(index=df.index)
            
            # Compare each pair of columns
            for i, col1 in enumerate(columns_to_compare):
                for col2 in columns_to_compare[i+1:]:
                    # Check for exact matches (excluding NaN)
                    mask = (df[col1] == df[col2]) & df[col1].notna()
                    duplicates_mask[f"{col1}_vs_{col2}"] = mask
            
            # Combine all masks to find any row with at least one duplicate
            has_duplicates = duplicates_mask.any(axis=1)
            result_df['Has_Duplicates'] = has_duplicates
            
            # Add information about which columns have duplicates
            for i, col1 in enumerate(columns_to_compare):
                for col2 in columns_to_compare[i+1:]:
                    dup_col_name = f"{col1}_eq_{col2}"
                    result_df[dup_col_name] = (df[col1] == df[col2]) & df[col1].notna()
            
            return {
                'result_df': result_df,
                'duplicate_rows': result_df[has_duplicates],
                'duplicate_count': has_duplicates.sum(),
                'total_rows': len(df)
            }
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare columns: {str(e)}")
            return None
    
    def create_common_column(self, sheet_name, columns_to_combine, new_column_name, strategy="first_non_empty", mark_duplicates=False):
        """Create a common column from selected columns and optionally mark duplicate values"""
        if not sheet_name or not columns_to_combine or sheet_name not in self.current_sheets:
            return False
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure all columns exist
            for col in columns_to_combine:
                if col not in df.columns:
                    messagebox.showerror("Error", f"Column '{col}' not found in sheet '{sheet_name}'")
                    return False
            
            # First identify duplicates
            comparison_result = self.compare_columns_for_duplicates(sheet_name, columns_to_combine)
            if not comparison_result:
                return False
            
            # Apply merge strategy
            if strategy == "first_non_empty":
                new_values = df[columns_to_combine].apply(
                    lambda row: next((x for x in row if pd.notna(x)), None), 
                    axis=1
                )
            elif strategy == "prioritize_duplicates":
                # For each row, prioritize values that appear in multiple columns
                def get_duplicate_value(row):
                    values = [row[col] for col in columns_to_combine if pd.notna(row[col])]
                    if not values:
                        return None
                    
                    # Count occurrences of each value
                    value_counts = {}
                    for val in values:
                        value_counts[val] = value_counts.get(val, 0) + 1
                    
                    # Get the value with the highest count
                    max_count = max(value_counts.values())
                    if max_count > 1:  # If there are duplicates
                        for val, count in value_counts.items():
                            if count == max_count:
                                return val
                    
                    # If no duplicates, use first non-empty
                    return values[0]
                
                new_values = pd.Series([get_duplicate_value(row) for _, row in df[columns_to_combine].iterrows()], index=df.index)
                
            elif strategy == "mark_duplicates":
                # Similar to first_non_empty but mark duplicate values
                def mark_if_duplicate(row):
                    values = [row[col] for col in columns_to_combine if pd.notna(row[col])]
                    if not values:
                        return None
                    
                    # Count occurrences of each value
                    value_counts = {}
                    for val in values:
                        value_counts[val] = value_counts.get(val, 0) + 1
                    
                    # Get the first non-empty value
                    value = values[0]
                    
                    # If it appears multiple times, mark it
                    if value_counts[value] > 1:
                        return f"{value} (duplicate)"
                    return value
                
                new_values = pd.Series([mark_if_duplicate(row) for _, row in df[columns_to_combine].iterrows()], index=df.index)
                
            elif strategy == "concatenate":
                new_values = df[columns_to_combine].apply(
                    lambda row: " | ".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]), 
                    axis=1
                )
            
            # Add the new column
            df[new_column_name] = new_values
            
            # If requested, add a column that marks which rows have duplicates
            if mark_duplicates:
                df[f"{new_column_name}_has_duplicate"] = comparison_result['result_df']['Has_Duplicates']
            
            # Update the current sheet
            self.current_sheets[sheet_name] = df
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create common column: {str(e)}")
            return False
    
    def save_merged_file(self):
        """Save the merged data to a new Excel file"""
        if not self.current_sheets:
            return None
            
        # Generate default filename with timestamp
        input_dir = os.path.dirname(self.input_file)
        input_filename = os.path.basename(self.input_file)
        basename, ext = os.path.splitext(input_filename)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{basename}_merged_{timestamp}{ext}"
        default_path = os.path.join(input_dir, default_filename)
        
        # Ask user where to save the file
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename,
            initialdir=input_dir,
            title="Save Merged File"
        )
        
        if not self.output_file:
            return None
            
        # Save the merged data
        try:
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                for sheet_name, df in self.current_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            return self.output_file
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            return None

class ManualMergeWindow:
    def __init__(self, parent, merger, sheet_name, column_list):
        self.parent = parent
        self.merger = merger
        self.sheet_name = sheet_name
        self.column_list = column_list
        
        # Create a new window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Manual Column Merge - {sheet_name}")
        self.window.geometry("650x650")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        window_width = 650
        window_height = 650
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width / 2) - (window_width / 2)
        y = (screen_height / 2) - (window_height / 2)
        self.window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
        
        # Create frames
        header_frame = ttk.Frame(self.window, padding="10")
        header_frame.pack(fill="x")
        
        columns_frame = ttk.Frame(self.window, padding="10")
        columns_frame.pack(fill="both", expand=True)
        
        options_frame = ttk.Frame(self.window, padding="10")
        options_frame.pack(fill="x")
        
        button_frame = ttk.Frame(self.window, padding="10")
        button_frame.pack(fill="x", side="bottom")
        
        # Header
        ttk.Label(
            header_frame, 
            text=f"Select Columns to Merge in Sheet: {sheet_name}", 
            font=("Helvetica", 12)
        ).pack(pady=5)
        
        ttk.Label(
            header_frame,
            text="Select columns from the list below and specify a new column name",
            wraplength=550
        ).pack(pady=5)
        
        # Columns listbox with scrollbar
        columns_frame_inner = ttk.Frame(columns_frame)
        columns_frame_inner.pack(fill="both", expand=True)
        
        self.columns_listbox = tk.Listbox(
            columns_frame_inner, 
            selectmode=tk.MULTIPLE, 
            height=15,
            exportselection=0
        )
        
        scrollbar = ttk.Scrollbar(columns_frame_inner, orient="vertical")
        self.columns_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.columns_listbox.yview)
        
        self.columns_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add columns to listbox
        for col in self.column_list:
            self.columns_listbox.insert(tk.END, col)
        
        # Add a preview frame
        preview_frame = ttk.LabelFrame(columns_frame, text="Preview Selected Columns", padding="10")
        preview_frame.pack(fill="x", pady=10)
        
        self.preview_text = tk.Text(preview_frame, height=5, wrap="word")
        self.preview_text.pack(fill="x")
        
        # New column name
        ttk.Label(options_frame, text="New Column Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.new_column_var = tk.StringVar()
        ttk.Entry(options_frame, textvariable=self.new_column_var, width=30).grid(row=0, column=1, padx=5, pady=5)
        
        # Merge strategy
        ttk.Label(options_frame, text="Merge Strategy:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.strategy_var = tk.StringVar(value="first_non_empty")
        
        strategies = [
            ("First Non-Empty Value", "first_non_empty"),
            ("Sum Numeric Values", "sum"),
            ("Concatenate Text Values", "concatenate")
        ]
        
        strategy_frame = ttk.Frame(options_frame)
        strategy_frame.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        for i, (text, value) in enumerate(strategies):
            ttk.Radiobutton(
                strategy_frame, 
                text=text, 
                variable=self.strategy_var, 
                value=value
            ).pack(anchor="w", padx=5, pady=2)
        
        # Additional options
        options_additional_frame = ttk.LabelFrame(options_frame, text="Additional Options", padding="5")
        options_additional_frame.grid(row=2, column=0, columnspan=2, sticky="we", padx=5, pady=10)
        
        # Delete source columns option
        self.delete_source_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_additional_frame,
            text="Delete source columns after merging",
            variable=self.delete_source_var
        ).pack(anchor="w", padx=5, pady=2)
        
        # Remove empty columns option
        self.remove_empty_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_additional_frame,
            text="Remove empty columns from sheet",
            variable=self.remove_empty_var
        ).pack(anchor="w", padx=5, pady=2)
        
        # Buttons
        ttk.Button(button_frame, text="Merge Selected Columns", command=self.merge_selected).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.window.destroy).pack(side="right", padx=5)
        
        # Bind listbox selection to update preview
        self.columns_listbox.bind('<<ListboxSelect>>', self.update_preview)
    
    def update_preview(self, event=None):
        """Update the preview of selected columns"""
        self.preview_text.delete(1.0, tk.END)
        
        selected_indices = self.columns_listbox.curselection()
        if not selected_indices:
            self.preview_text.insert(tk.END, "No columns selected")
            return
        
        selected_columns = [self.column_list[i] for i in selected_indices]
        self.preview_text.insert(tk.END, f"Selected columns: {', '.join(selected_columns)}\n\n")
        
        # Suggest a name for the new column based on selection
        if len(selected_columns) > 0 and not self.new_column_var.get():
            # Use the shortest column name as a base
            base_name = min(selected_columns, key=len)
            self.new_column_var.set(f"Merged_{base_name}")
        
        # Show sample data if available
        if self.sheet_name in self.merger.current_sheets:
            df = self.merger.current_sheets[self.sheet_name]
            if all(col in df.columns for col in selected_columns):
                sample_data = df[selected_columns].head(3)
                self.preview_text.insert(tk.END, "Sample data (first 3 rows):\n")
                for i, row in sample_data.iterrows():
                    row_display = ", ".join([f"{col}: {row[col]}" for col in selected_columns])
                    self.preview_text.insert(tk.END, f"Row {i+1}: {row_display}\n")
    
    def merge_selected(self):
        """Merge the selected columns"""
        selected_indices = self.columns_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least two columns to merge")
            return
        
        selected_columns = [self.column_list[i] for i in selected_indices]
        if len(selected_columns) < 2:
            messagebox.showwarning("Warning", "Please select at least two columns to merge")
            return
        
        new_column_name = self.new_column_var.get()
        if not new_column_name:
            messagebox.showwarning("Warning", "Please enter a name for the new column")
            return
        
        # Check if new column name already exists
        df = self.merger.current_sheets[self.sheet_name]
        if new_column_name in df.columns and new_column_name not in selected_columns:
            overwrite = messagebox.askyesno(
                "Column Exists", 
                f"Column '{new_column_name}' already exists. Do you want to overwrite it?"
            )
            if not overwrite:
                return
        
        # Get option values
        delete_source = self.delete_source_var.get()
        remove_empty = self.remove_empty_var.get()
        
        # Perform the merge
        strategy = self.strategy_var.get()
        result, empty_cols = self.merger.manual_merge_columns(
            self.sheet_name, 
            selected_columns, 
            new_column_name, 
            strategy,
            delete_source,
            remove_empty
        )
        
        if result:
            success_message = f"Columns merged into '{new_column_name}'"
            
            if delete_source:
                success_message += f"\nDeleted {len(selected_columns)} source columns"
            
            if remove_empty and empty_cols:
                success_message += f"\nRemoved {len(empty_cols)} empty columns"
                
            messagebox.showinfo("Success", success_message)
            self.window.destroy()
            
            # Signal to parent to refresh
            self.parent.event_generate("<<RefreshSheets>>")
        else:
            messagebox.showerror("Error", "Failed to merge columns")

class CompareColumnsWindow:
    def __init__(self, parent, merger, sheet_name, column_list):
        self.parent = parent
        self.merger = merger
        self.sheet_name = sheet_name
        self.column_list = column_list
        
        # Create a new window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Compare Columns for Duplicates - {sheet_name}")
        self.window.geometry("750x700")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        window_width = 750
        window_height = 700
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width / 2) - (window_width / 2)
        y = (screen_height / 2) - (window_height / 2)
        self.window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
        
        # Create frames
        header_frame = ttk.Frame(self.window, padding="10")
        header_frame.pack(fill="x")
        
        columns_frame = ttk.Frame(self.window, padding="10")
        columns_frame.pack(fill="both", expand=True)
        
        results_frame = ttk.Frame(self.window, padding="10")
        results_frame.pack(fill="both", expand=True)
        
        options_frame = ttk.Frame(self.window, padding="10")
        options_frame.pack(fill="x")
        
        button_frame = ttk.Frame(self.window, padding="10")
        button_frame.pack(fill="x", side="bottom")
        
        # Header
        ttk.Label(
            header_frame, 
            text=f"Compare Columns for Duplicates in Sheet: {sheet_name}", 
            font=("Helvetica", 12)
        ).pack(pady=5)
        
        ttk.Label(
            header_frame,
            text="Select columns to compare for duplicate values, then click 'Find Duplicates'",
            wraplength=650
        ).pack(pady=5)
        
        # Columns listbox with scrollbar
        columns_frame_inner = ttk.Frame(columns_frame)
        columns_frame_inner.pack(fill="both", expand=True)
        
        self.columns_listbox = tk.Listbox(
            columns_frame_inner, 
            selectmode=tk.MULTIPLE, 
            height=10,
            exportselection=0
        )
        
        scrollbar = ttk.Scrollbar(columns_frame_inner, orient="vertical")
        self.columns_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.columns_listbox.yview)
        
        self.columns_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add columns to listbox
        for col in self.column_list:
            self.columns_listbox.insert(tk.END, col)
        
        # Results frame
        self.results_text = tk.Text(results_frame, height=10, wrap="word")
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        self.results_text.config(yscrollcommand=results_scrollbar.set)
        
        self.results_text.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")
        
        # Button to find duplicates
        ttk.Button(
            columns_frame,
            text="Find Duplicates",
            command=self.find_duplicates
        ).pack(pady=10)
        
        # Options for creating a common column
        ttk.Label(options_frame, text="Create Common Column").grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        # New column name
        ttk.Label(options_frame, text="New Column Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.new_column_var = tk.StringVar()
        ttk.Entry(options_frame, textvariable=self.new_column_var, width=30).grid(row=1, column=1, padx=5, pady=5)
        
        # Strategy for common column
        ttk.Label(options_frame, text="Column Strategy:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.common_strategy_var = tk.StringVar(value="first_non_empty")
        
        common_strategy_frame = ttk.Frame(options_frame)
        common_strategy_frame.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        common_strategies = [
            ("First Non-Empty Value", "first_non_empty"),
            ("Prioritize Duplicate Values", "prioritize_duplicates"),
            ("Mark Duplicate Values", "mark_duplicates"),
            ("Concatenate All Values", "concatenate")
        ]
        
        for i, (text, value) in enumerate(common_strategies):
            ttk.Radiobutton(
                common_strategy_frame, 
                text=text, 
                variable=self.common_strategy_var, 
                value=value
            ).pack(anchor="w", padx=5, pady=2)
        
        # Mark duplicates option
        self.mark_duplicates_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Add column indicating which rows have duplicates",
            variable=self.mark_duplicates_var
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        # Buttons
        ttk.Button(button_frame, text="Create Common Column", command=self.create_common_column).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=self.window.destroy).pack(side="right", padx=5)
        
        # Initialize with empty results
        self.results_text.insert(tk.END, "Select columns and click 'Find Duplicates' to analyze them for duplicate values.")
        
    def find_duplicates(self):
        """Find duplicate values across selected columns"""
        selected_indices = self.columns_listbox.curselection()
        if not selected_indices or len(selected_indices) < 2:
            messagebox.showwarning("Warning", "Please select at least two columns to compare")
            return
        
        selected_columns = [self.column_list[i] for i in selected_indices]
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"Comparing {len(selected_columns)} columns for duplicate values...\n\n")
        
        # Perform comparison
        comparison_result = self.merger.compare_columns_for_duplicates(self.sheet_name, selected_columns)
        if not comparison_result:
            self.results_text.insert(tk.END, "Failed to compare columns.")
            return
        
        # Display results
        self.results_text.insert(tk.END, f"Found {comparison_result['duplicate_count']} rows with duplicate values out of {comparison_result['total_rows']} total rows.\n\n")
        
        # Show some of the duplicates
        if comparison_result['duplicate_count'] > 0:
            self.results_text.insert(tk.END, "Sample of rows with duplicate values:\n")
            dup_rows = comparison_result['duplicate_rows']
            
            # Show at most 5 duplicate rows
            sample_size = min(5, len(dup_rows))
            for i, (idx, row) in enumerate(dup_rows.iloc[:sample_size].iterrows()):
                self.results_text.insert(tk.END, f"Row {idx+1}:\n")
                
                # Show only the original columns (not the added metadata)
                for col in selected_columns:
                    self.results_text.insert(tk.END, f"  {col}: {row[col]}\n")
                
                # Add separator between rows
                if i < sample_size - 1:
                    self.results_text.insert(tk.END, "\n")
        
        # Suggest a name for the new column
        if not self.new_column_var.get():
            self.new_column_var.set(f"Common_{'_'.join(selected_columns[:2])}")
    
    def create_common_column(self):
        """Create a common column based on the selected columns"""
        selected_indices = self.columns_listbox.curselection()
        if not selected_indices or len(selected_indices) < 2:
            messagebox.showwarning("Warning", "Please select at least two columns to create a common column")
            return
        
        selected_columns = [self.column_list[i] for i in selected_indices]
        
        new_column_name = self.new_column_var.get()
        if not new_column_name:
            messagebox.showwarning("Warning", "Please enter a name for the new column")
            return
        
        # Check if new column name already exists
        df = self.merger.current_sheets[self.sheet_name]
        if new_column_name in df.columns and new_column_name not in selected_columns:
            overwrite = messagebox.askyesno(
                "Column Exists", 
                f"Column '{new_column_name}' already exists. Do you want to overwrite it?"
            )
            if not overwrite:
                return
        
        # Get strategy and mark_duplicates option
        strategy = self.common_strategy_var.get()
        mark_duplicates = self.mark_duplicates_var.get()
        
        # Create the common column
        result = self.merger.create_common_column(
            self.sheet_name,
            selected_columns,
            new_column_name,
            strategy,
            mark_duplicates
        )
        
        if result:
            success_message = f"Created common column '{new_column_name}'"
            if mark_duplicates:
                success_message += f" and added duplicate indicator column"
                
            messagebox.showinfo("Success", success_message)
            self.window.destroy()
            
            # Signal to parent to refresh
            self.parent.event_generate("<<RefreshSheets>>")
        else:
            messagebox.showerror("Error", "Failed to create common column")

def main():
    # Create the main window
    root = tk.Tk()
    root.title("Excel Column Merger")
    root.geometry("800x600")
    
    # Center the window
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width / 2) - (window_width / 2)
    y = (screen_height / 2) - (window_height / 2)
    root.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
    
    # Initialize merger
    merger = ExcelColumnMerger()
    
    # Create frames
    title_frame = ttk.Frame(root, padding="10")
    title_frame.pack(fill="x")
    
    file_frame = ttk.Frame(root, padding="10")
    file_frame.pack(fill="x")
    
    options_frame = ttk.Frame(root, padding="10")
    options_frame.pack(fill="x")
    
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)
    
    auto_frame = ttk.Frame(notebook, padding="10")
    notebook.add(auto_frame, text="Auto Merge")
    
    manual_frame = ttk.Frame(notebook, padding="10")
    notebook.add(manual_frame, text="Manual Merge")
    
    compare_frame = ttk.Frame(notebook, padding="10")
    notebook.add(compare_frame, text="Compare Columns")
    
    button_frame = ttk.Frame(root, padding="10")
    button_frame.pack(fill="x", side="bottom")
    
    # Title
    title_label = ttk.Label(
        title_frame, 
        text="Excel Column Merger", 
        font=("Helvetica", 16)
    )
    title_label.pack(pady=10)
    
    description = ttk.Label(
        title_frame,
        text="This tool identifies duplicate column names, allows manual merging of columns, and can compare values across columns to find duplicates.",
        wraplength=750,
        justify="center"
    )
    description.pack(pady=5)
    
    # File selection
    file_var = tk.StringVar()
    ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(file_frame, textvariable=file_var, width=70).grid(row=0, column=1, padx=5, pady=5)
    
    def select_file():
        file = merger.select_file()
        if file:
            file_var.set(file)
            # Clear previous results
            for item in auto_results_tree.get_children():
                auto_results_tree.delete(item)
            # Update sheet selector
            update_sheet_selector()
            status_var.set("File selected. Use tabs to perform different operations.")
            # Enable analyze button
            analyze_button.configure(state="normal")
            # Enable manual merge
            manual_merge_button.configure(state="normal")
            # Enable compare columns
            compare_columns_button.configure(state="normal")
    
    ttk.Button(file_frame, text="Browse", command=select_file).grid(row=0, column=2, padx=5, pady=5)
    
    # Auto Merge tab
    auto_options_frame = ttk.Frame(auto_frame)
    auto_options_frame.pack(fill="x", pady=5)
    
    ttk.Label(auto_options_frame, text="Merge Strategy:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    strategies = [
        ("First Non-Empty Value (default)", "first_non_empty"),
        ("Sum Numeric Values", "sum"),
        ("Concatenate Text Values", "concatenate")
    ]
    
    for i, (text, value) in enumerate(strategies):
        ttk.Radiobutton(
            auto_options_frame, 
            text=text, 
            variable=merger.merge_strategy, 
            value=value
        ).grid(row=0, column=i+1, sticky="w", padx=5, pady=5)
    
    # Results treeview for auto merge
    ttk.Label(auto_frame, text="Duplicate Columns (Case-Insensitive):").pack(anchor="w", padx=5, pady=5)
    
    auto_columns = ("sheet", "column", "duplicates")
    auto_results_tree = ttk.Treeview(auto_frame, columns=auto_columns, show="headings")
    
    auto_results_tree.heading("sheet", text="Sheet")
    auto_results_tree.heading("column", text="Base Column Name")
    auto_results_tree.heading("duplicates", text="Duplicate Columns")
    
    auto_results_tree.column("sheet", width=100)
    auto_results_tree.column("column", width=150)
    auto_results_tree.column("duplicates", width=500)
    
    auto_scrollbar = ttk.Scrollbar(auto_frame, orient="vertical", command=auto_results_tree.yview)
    auto_results_tree.configure(yscrollcommand=auto_scrollbar.set)
    
    auto_results_tree.pack(side="left", fill="both", expand=True)
    auto_scrollbar.pack(side="right", fill="y")
    
    # Manual Merge tab
    manual_top_frame = ttk.Frame(manual_frame)
    manual_top_frame.pack(fill="x", pady=5)
    
    ttk.Label(manual_top_frame, text="Select Sheet:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    sheet_var = tk.StringVar()
    sheet_selector = ttk.Combobox(manual_top_frame, textvariable=sheet_var, width=30, state="readonly")
    sheet_selector.grid(row=0, column=1, padx=5, pady=5)
    
    def update_sheet_selector():
        """Update the sheet selector dropdown"""
        if hasattr(merger, 'current_sheets') and merger.current_sheets:
            sheet_selector['values'] = list(merger.current_sheets.keys())
            compare_sheet_selector['values'] = list(merger.current_sheets.keys())
            if sheet_selector['values']:
                sheet_selector.current(0)
                compare_sheet_selector.current(0)
                refresh_columns_display()
                refresh_compare_columns_display()
    
    # Columns display
    ttk.Label(manual_frame, text="Available Columns:").pack(anchor="w", padx=5, pady=5)
    
    columns_frame = ttk.Frame(manual_frame)
    columns_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    columns_listbox = tk.Listbox(columns_frame, height=10, exportselection=0)
    columns_scrollbar = ttk.Scrollbar(columns_frame, orient="vertical", command=columns_listbox.yview)
    columns_listbox.configure(yscrollcommand=columns_scrollbar.set)
    
    columns_listbox.pack(side="left", fill="both", expand=True)
    columns_scrollbar.pack(side="right", fill="y")
    
    def refresh_columns_display():
        """Refresh the columns display based on selected sheet"""
        columns_listbox.delete(0, tk.END)
        selected_sheet = sheet_var.get()
        if selected_sheet and selected_sheet in merger.current_sheets:
            df = merger.current_sheets[selected_sheet]
            for col in df.columns:
                columns_listbox.insert(tk.END, col)
    
    # Compare Columns tab
    compare_top_frame = ttk.Frame(compare_frame)
    compare_top_frame.pack(fill="x", pady=5)
    
    ttk.Label(compare_top_frame, text="Select Sheet:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    compare_sheet_var = tk.StringVar()
    compare_sheet_selector = ttk.Combobox(compare_top_frame, textvariable=compare_sheet_var, width=30, state="readonly")
    compare_sheet_selector.grid(row=0, column=1, padx=5, pady=5)
    
    # Compare columns display
    ttk.Label(compare_frame, text="Available Columns:").pack(anchor="w", padx=5, pady=5)
    
    compare_columns_frame = ttk.Frame(compare_frame)
    compare_columns_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    compare_columns_listbox = tk.Listbox(compare_columns_frame, height=10, exportselection=0)
    compare_columns_scrollbar = ttk.Scrollbar(compare_columns_frame, orient="vertical", command=compare_columns_listbox.yview)
    compare_columns_listbox.configure(yscrollcommand=compare_columns_scrollbar.set)
    
    compare_columns_listbox.pack(side="left", fill="both", expand=True)
    compare_columns_scrollbar.pack(side="right", fill="y")
    
    def refresh_compare_columns_display():
        """Refresh the compare columns display based on selected sheet"""
        compare_columns_listbox.delete(0, tk.END)
        selected_sheet = compare_sheet_var.get()
        if selected_sheet and selected_sheet in merger.current_sheets:
            df = merger.current_sheets[selected_sheet]
            for col in df.columns:
                compare_columns_listbox.insert(tk.END, col)
    
    # Bind sheet selectors to refresh columns
    sheet_selector.bind('<<ComboboxSelected>>', lambda e: refresh_columns_display())
    compare_sheet_selector.bind('<<ComboboxSelected>>', lambda e: refresh_compare_columns_display())
    
    # Status bar
    status_var = tk.StringVar(value="Ready")
    status_bar = ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w")
    status_bar.pack(side="bottom", fill="x")
    
    # Analysis function
    def analyze_file():
        if not merger.input_file:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
            
        status_var.set("Analyzing file for duplicate columns...")
        
        # Clear previous results
        for item in auto_results_tree.get_children():
            auto_results_tree.delete(item)
            
        # Analyze the file
        analysis = merger.analyze_file()
        
        if not analysis:
            status_var.set("Analysis failed. Please try again.")
            return
            
        if analysis["status"] == "error":
            messagebox.showerror("Error", f"Failed to analyze file: {analysis['message']}")
            status_var.set("Analysis failed.")
            return
            
        if analysis["status"] == "no_duplicates":
            messagebox.showinfo("No Duplicates", "No duplicate column names found in the Excel file.")
            status_var.set("No duplicate columns found.")
            return
            
        # Populate the results tree
        sheets_data = analysis["sheets_data"]
        for sheet_name, sheet_data in sheets_data.items():
            duplicate_columns = sheet_data["duplicate_columns"]
            
            for base_name, columns in duplicate_columns.items():
                auto_results_tree.insert(
                    "", 
                    "end", 
                    values=(
                        sheet_name,
                        base_name, 
                        ", ".join(columns)
                    )
                )
                
        # Store analysis
        merger.file_analysis = analysis
        
        # Update status
        total_duplicates = sum(len(sheet["duplicate_columns"]) for sheet in sheets_data.values())
        status_var.set(f"Found {total_duplicates} duplicate column groups. Choose a merge strategy and click 'Auto Merge'.")
        
        # Select the Auto Merge tab
        notebook.select(0)
        
        # Enable merge button
        auto_merge_button.configure(state="normal")
    
    # Auto Merge function
    def auto_merge_columns():
        if not hasattr(merger, 'file_analysis'):
            messagebox.showerror("Error", "Please analyze the file first.")
            return
            
        strategy = merger.merge_strategy.get()
        status_var.set(f"Merging columns using '{strategy}' strategy...")
        
        # Perform the merge
        if merger.merge_columns(merger.file_analysis, strategy):
            messagebox.showinfo(
                "Success", 
                "Duplicate columns merged successfully!\n\nYou can now use the 'Manual Merge' tab to merge additional columns, or save the file."
            )
            status_var.set("Columns merged. Switch to 'Manual Merge' tab or save the file.")
            
            # Update sheet selector
            update_sheet_selector()
            
            # Enable save button
            save_button.configure(state="normal")
            
            # Select Manual Merge tab
            notebook.select(1)
        else:
            status_var.set("Failed to merge columns.")
    
    # Manual Merge function
    def open_manual_merge():
        selected_sheet = sheet_var.get()
        if not selected_sheet or selected_sheet not in merger.current_sheets:
            messagebox.showerror("Error", "Please select a valid sheet.")
            return
            
        df = merger.current_sheets[selected_sheet]
        column_list = list(df.columns)
        
        if len(column_list) < 2:
            messagebox.showwarning("Warning", "Sheet has less than 2 columns. No merging possible.")
            return
            
        # Open the manual merge window
        ManualMergeWindow(root, merger, selected_sheet, column_list)
    
    # Compare Columns function
    def open_compare_columns():
        selected_sheet = compare_sheet_var.get()
        if not selected_sheet or selected_sheet not in merger.current_sheets:
            messagebox.showerror("Error", "Please select a valid sheet.")
            return
            
        df = merger.current_sheets[selected_sheet]
        column_list = list(df.columns)
        
        if len(column_list) < 2:
            messagebox.showwarning("Warning", "Sheet has less than 2 columns. No comparison possible.")
            return
            
        # Open the compare columns window
        CompareColumnsWindow(root, merger, selected_sheet, column_list)
    
    # Save function
    def save_file():
        if not hasattr(merger, 'current_sheets') or not merger.current_sheets:
            messagebox.showerror("Error", "No data to save.")
            return
            
        status_var.set("Saving file...")
        
        # Save the file
        output_file = merger.save_merged_file()
        
        if output_file:
            messagebox.showinfo(
                "Success", 
                f"File saved successfully as:\n{output_file}"
            )
            
            # Ask if the user wants to open the file
            if messagebox.askyesno("Open File", "Do you want to open the merged file?"):
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(output_file)
                    else:  # Mac/Linux
                        import subprocess
                        subprocess.call(('open', output_file) if os.name == 'posix' else ('xdg-open', output_file))
                except Exception as e:
                    messagebox.showwarning("Warning", f"Could not open file: {str(e)}")
            
            status_var.set(f"File saved as {os.path.basename(output_file)}")
        else:
            status_var.set("Save cancelled or failed.")
    
    # Refresh handler
    def refresh_after_manual_merge(event):
        refresh_columns_display()
        refresh_compare_columns_display()
    
    # Bind the custom event
    root.bind("<<RefreshSheets>>", refresh_after_manual_merge)
    
    # Action buttons
    analyze_button = ttk.Button(button_frame, text="Analyze for Duplicates", command=analyze_file, state="disabled")
    analyze_button.pack(side="left", padx=5)
    
    auto_merge_button = ttk.Button(button_frame, text="Auto Merge", command=auto_merge_columns, state="disabled")
    auto_merge_button.pack(side="left", padx=5)
    
    manual_merge_button = ttk.Button(button_frame, text="Manual Merge Selected", command=open_manual_merge, state="disabled")
    manual_merge_button.pack(side="left", padx=5)
    
    compare_columns_button = ttk.Button(button_frame, text="Compare Columns", command=open_compare_columns, state="disabled")
    compare_columns_button.pack(side="left", padx=5)
    
    save_button = ttk.Button(button_frame, text="Save File", command=save_file, state="disabled")
    save_button.pack(side="left", padx=5)
    
    ttk.Button(button_frame, text="Exit", command=root.destroy).pack(side="right", padx=5)
    
    # Set focus
    root.focus_force()
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()