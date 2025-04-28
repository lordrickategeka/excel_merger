import tkinter as tk
from tkinter import ttk, messagebox
import difflib  # For finding similar text
import re  # For text cleaning
import pandas as pd  # Make sure this is imported in your main file
from core.header_similarity import HeaderSimilarityAnalyzer

class ManualMergeWindow:
    """
    Window for manually merging selected columns with enhanced similar column detection
    """
    def __init__(self, parent, merger, sheet_name, column_list):
        self.parent = parent
        self.merger = merger
        self.sheet_name = sheet_name
        self.column_list = column_list
        
        # Create a new window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Manual Column Merge - {sheet_name}")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create scrollable main content frame
        main_container = ttk.Frame(self.window)
        main_container.pack(fill="both", expand=True)
        
        # Create canvas with scrollbar for the main content
        self.canvas = tk.Canvas(main_container)
        self.scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configure canvas and scrollable frame
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Add mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Create frames within scrollable frame
        header_frame = ttk.Frame(self.scrollable_frame, padding="10")
        header_frame.pack(fill="x")
        
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
        
        # Add similar columns detection section
        similar_frame = ttk.LabelFrame(self.scrollable_frame, text="Similar Column Detection", padding="10")
        similar_frame.pack(fill="x", padx=10, pady=5)
        
        # Create detection controls
        detection_frame = ttk.Frame(similar_frame)
        detection_frame.pack(fill="x", pady=5)
        
        ttk.Label(detection_frame, text="Similarity threshold:").pack(side="left", padx=5)
        self.similarity_var = tk.DoubleVar(value=0.6)  # Default similarity threshold
        similarity_scale = ttk.Scale(
            detection_frame, 
            from_=0.1, 
            to=0.9, 
            length=200,
            orient="horizontal", 
            variable=self.similarity_var,
            command=self.update_similarity_label
        )
        similarity_scale.pack(side="left", padx=5)
        
        self.similarity_label = ttk.Label(detection_frame, text="0.6")
        self.similarity_label.pack(side="left", padx=5)
        
        ttk.Button(
            detection_frame, 
            text="Find Similar Columns", 
            command=self.find_similar_columns
        ).pack(side="left", padx=20)
        
        # Create similar columns results
        self.similar_results = ttk.Frame(similar_frame)
        self.similar_results.pack(fill="x", pady=5)
        
        self.similar_groups_frame = ttk.Frame(self.similar_results)
        self.similar_groups_frame.pack(fill="both", expand=True)
        
        # Add columns selection frame
        columns_frame = ttk.LabelFrame(self.scrollable_frame, text="Manual Column Selection", padding="10")
        columns_frame.pack(fill="x", padx=10, pady=5)
        
        # Columns listbox with scrollbar
        columns_frame_inner = ttk.Frame(columns_frame)
        columns_frame_inner.pack(fill="both")
        
        self.columns_listbox = tk.Listbox(
            columns_frame_inner, 
            selectmode=tk.MULTIPLE, 
            height=15,
            exportselection=0
        )
        
        listbox_scrollbar = ttk.Scrollbar(columns_frame_inner, orient="vertical", command=self.columns_listbox.yview)
        self.columns_listbox.config(yscrollcommand=listbox_scrollbar.set)
        
        self.columns_listbox.pack(side="left", fill="both", expand=True)
        listbox_scrollbar.pack(side="right", fill="y")
        
        # Add columns to listbox
        for col in self.column_list:
            self.columns_listbox.insert(tk.END, col)
        
        # Add a preview frame
        preview_frame = ttk.LabelFrame(columns_frame, text="Preview Selected Columns", padding="10")
        preview_frame.pack(fill="x", pady=10)
        
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical")
        self.preview_text = tk.Text(preview_frame, height=5, wrap="word", yscrollcommand=preview_scroll.set)
        preview_scroll.config(command=self.preview_text.yview)
        
        self.preview_text.pack(side="left", fill="x", expand=True)
        preview_scroll.pack(side="right", fill="y")
        
        # Options frame
        options_frame = ttk.LabelFrame(self.scrollable_frame, text="Merge Options", padding="10")
        options_frame.pack(fill="x", padx=10, pady=5)
        
        # New column name
        name_frame = ttk.Frame(options_frame)
        name_frame.pack(fill="x", pady=5)
        
        ttk.Label(name_frame, text="New Column Name:").pack(side="left", padx=5)
        self.new_column_var = tk.StringVar()
        ttk.Entry(name_frame, textvariable=self.new_column_var, width=30).pack(side="left", padx=5)
        
        # Merge strategy
        strategy_frame = ttk.LabelFrame(options_frame, text="Merge Strategy", padding="5")
        strategy_frame.pack(fill="x", pady=5)
        
        self.strategy_var = tk.StringVar(value="first_non_empty")
        
        strategies = [
            ("First Non-Empty Value", "first_non_empty"),
            ("Sum Numeric Values", "sum"),
            ("Concatenate Text Values", "concatenate"),
            ("Stack Values in Rows", "stack_values")  # New strategy
        ]
        
        for i, (text, value) in enumerate(strategies):
            ttk.Radiobutton(
                strategy_frame, 
                text=text, 
                variable=self.strategy_var, 
                value=value
            ).pack(anchor="w", padx=5, pady=2)
        
        # Additional options
        options_additional_frame = ttk.LabelFrame(options_frame, text="Additional Options", padding="5")
        options_additional_frame.pack(fill="x", pady=5)
        
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
        
        # Create a fixed button frame at the bottom (outside scrollable area)
        button_frame = ttk.Frame(self.window, padding="10")
        button_frame.pack(fill="x", side="bottom")
        
        # Buttons - always visible at the bottom
        merge_button = ttk.Button(button_frame, text="Merge Selected Columns", command=self.merge_selected)
        merge_button.pack(side="left", padx=5)
        
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.window.destroy)
        cancel_button.pack(side="right", padx=5)
        
        # Bind listbox selection to update preview
        self.columns_listbox.bind('<<ListboxSelect>>', self.update_preview)
        
        # ---- IMPROVED WINDOW SIZING CODE ----
        self.window.update_idletasks()
        
        # Set initial size (large enough to show important elements)
        width = 900
        height = 800
        
        # Center the window
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Ensure window isn't too large for the screen
        width = min(width, int(screen_width * 0.9))
        height = min(height, int(screen_height * 0.9))
        
        # Calculate position
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # Set window geometry
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        
        # Make window resizable
        self.window.resizable(True, True)
        
        # Set minimum size
        self.window.minsize(800, 650)
        
        # Update canvas scroll region after all widgets are placed
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Configure canvas width to match window width
        self.canvas.config(width=width-20)  # Account for scrollbar
    
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def update_similarity_label(self, *args):
        """Update the similarity threshold label"""
        value = round(self.similarity_var.get(), 2)
        self.similarity_label.config(text=f"{value}")
    
    def normalize_text(self, text):
        """Normalize text for better similarity matching"""
        # Convert to lowercase
        text = str(text).lower()
        # Remove non-alphanumeric characters except spaces
        text = re.sub(r'[^\w\s]', '', text)
        # Replace multiple spaces with single space
        text = re.sub(r'\s+', ' ', text)
        # Remove leading/trailing spaces
        text = text.strip()
        return text
    
    def find_similar_columns(self):
        """Find similar column names based on string similarity using HeaderSimilarityAnalyzer"""
        # Clear previous results
        for widget in self.similar_groups_frame.winfo_children():
            widget.destroy()
        
        # Get threshold from the slider
        threshold = self.similarity_var.get()
        
        # Create analyzer instance if it doesn't exist or update threshold
        if not hasattr(self, 'header_analyzer'):
            self.header_analyzer = HeaderSimilarityAnalyzer()
        
        self.header_analyzer.set_similarity_threshold(threshold)
        
        # Analyze column headers
        results, suggestion_text = self.header_analyzer.analyze_and_suggest_merges(self.column_list)
        
        # Get the similar groups from analysis results
        similar_groups = results['similar_groups']
        exact_duplicates = results['exact_duplicates']
        common_word_groups = results['common_word_groups']
        
        # Display exact duplicates (highest priority)
        if exact_duplicates:
            exact_frame = ttk.LabelFrame(self.similar_groups_frame, text="Exact Duplicates (Highest Priority)")
            exact_frame.pack(fill="x", pady=5)
            
            # Add explanation
            ttk.Label(
                exact_frame,
                text="These columns have identical names (ignoring case and whitespace)",
                wraplength=550
            ).pack(anchor="w", pady=5)
            
            # Display each duplicate group
            for norm_name, dupes in exact_duplicates.items():
                group_frame = ttk.Frame(exact_frame)
                group_frame.pack(fill="x", pady=2, padx=5)
                
                ttk.Label(
                    group_frame,
                    text=f"Group: {', '.join(dupes)}",
                    wraplength=550
                ).pack(side="left", padx=5)
                
                ttk.Button(
                    group_frame,
                    text="Use Group",
                    command=lambda grp=dupes: self.use_selected_group(grp)
                ).pack(side="right", padx=5)
        
        # Display similar groups
        if similar_groups:
            similar_heading = ttk.LabelFrame(self.similar_groups_frame, text="Similar Column Names")
            similar_heading.pack(fill="x", pady=5)
            
            # Add explanation
            ttk.Label(
                similar_heading,
                text="These columns have similar names that may indicate spelling variations or related data",
                wraplength=550
            ).pack(anchor="w", pady=5)
            
            # Create a frame for each group with checkboxes
            for i, group in enumerate(similar_groups):
                group_frame = ttk.LabelFrame(similar_heading, text=f"Similar Group #{i+1}")
                group_frame.pack(fill="x", pady=5)
                
                # Add select all checkbox
                select_var = tk.BooleanVar(value=False)
                select_all = ttk.Checkbutton(
                    group_frame, 
                    text="Select All", 
                    variable=select_var, 
                    command=lambda var=select_var, grp=group: self.toggle_group_selection(var, grp)
                )
                select_all.pack(anchor="w", pady=5)
                
                # Add column checkboxes
                column_frame = ttk.Frame(group_frame)
                column_frame.pack(fill="x", padx=10)
                
                # Arrange in columns (3 per row)
                for j, col in enumerate(group):
                    col_var = tk.BooleanVar(value=False)
                    cb = ttk.Checkbutton(
                        column_frame, 
                        text=col, 
                        variable=col_var,
                        command=lambda col=col, var=col_var: self.toggle_column_selection(col, var)
                    )
                    cb.grid(row=j//3, column=j%3, sticky="w", padx=5, pady=2)
                    
                # Add merge button for this group
                ttk.Button(
                    group_frame, 
                    text="Use Selected Items", 
                    command=lambda grp=group: self.use_selected_group(grp)
                ).pack(anchor="e", pady=5, padx=5)
        
        # Display common word groups
        if common_word_groups:
            common_heading = ttk.LabelFrame(self.similar_groups_frame, text="Common Word Groups")
            common_heading.pack(fill="x", pady=5)
            
            # Add explanation
            ttk.Label(
                common_heading,
                text="These columns share significant words that may indicate related data",
                wraplength=550
            ).pack(anchor="w", pady=5)
            
            # Create a frame for each group
            for i, group in enumerate(common_word_groups):
                group_frame = ttk.Frame(common_heading)
                group_frame.pack(fill="x", pady=2, padx=5)
                
                ttk.Label(
                    group_frame,
                    text=f"Group {i+1}: {', '.join(group)}",
                    wraplength=550
                ).pack(side="left", padx=5)
                
                ttk.Button(
                    group_frame,
                    text="Use Group",
                    command=lambda grp=group: self.use_selected_group(grp)
                ).pack(side="right", padx=5)
        
        # Display message if no similar columns found
        if not exact_duplicates and not similar_groups and not common_word_groups:
            no_results = ttk.Label(
                self.similar_groups_frame, 
                text="No similar columns found. Try lowering the threshold.",
                wraplength=550
            )
            no_results.pack(pady=10)
        
        # Update window size
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    
    def toggle_group_selection(self, var, group):
        """Select or deselect all columns in a group"""
        select_all = var.get()
        
        # Find all checkboxes for this group and update
        for widget in self.similar_groups_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame):
                # Process widgets in the group frame
                for frame in widget.winfo_children():
                    if isinstance(frame, ttk.Frame):
                        for child in frame.winfo_children():
                            if isinstance(child, ttk.Checkbutton):
                                # Get column name from checkbox text
                                col_name = child.cget("text")
                                if col_name in group:
                                    # Select or deselect in the main listbox
                                    self.select_column_in_listbox(col_name, select_all)
    
    def toggle_column_selection(self, column, var):
        """Select or deselect a column in the main listbox"""
        self.select_column_in_listbox(column, var.get())
        
    def select_column_in_listbox(self, column, select):
        """Select or deselect a column in the main listbox"""
        try:
            idx = self.column_list.index(column)
            if select:
                if idx not in self.columns_listbox.curselection():
                    self.columns_listbox.selection_set(idx)
            else:
                if idx in self.columns_listbox.curselection():
                    self.columns_listbox.selection_clear(idx)
            
            # Update preview
            self.update_preview()
        except ValueError:
            pass  # Column not found
    
    def use_selected_group(self, group):
        """Use the selected columns from this group"""
        # Clear current selection in listbox
        self.columns_listbox.selection_clear(0, tk.END)
        
        # Select all columns in the group
        for col in group:
            try:
                idx = self.column_list.index(col)
                self.columns_listbox.selection_set(idx)
            except ValueError:
                pass  # Column not found
        
        # Update preview
        self.update_preview()
        
        # Auto-create a name for the merged column
        if len(group) > 0:
            # Find the common parts of the column names
            base_name = self.find_common_text(group)
            if base_name:
                self.new_column_var.set(f"Merged_{base_name}")
            else:
                # Use the shortest name as base
                base_name = min(group, key=len)
                self.new_column_var.set(f"Merged_{base_name}")
    
    def find_common_text(self, strings):
        """Find common text elements across strings"""
        if not strings:
            return ""
        
        # Normalize all strings
        normalized = [self.normalize_text(s) for s in strings]
        
        # Find words that appear in all strings
        all_words = set(normalized[0].split())
        for s in normalized[1:]:
            words = set(s.split())
            all_words &= words
        
        if all_words:
            return " ".join(sorted(all_words, key=lambda w: normalized[0].find(w)))
        return ""
    
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
            # Find common elements or use shortest name
            base_name = self.find_common_text(selected_columns)
            if base_name:
                self.new_column_var.set(f"Merged_{base_name}")
            else:
                base_name = min(selected_columns, key=len)
                self.new_column_var.set(f"Merged_{base_name}")
        
        # Show sample data if available
        try:
            if self.sheet_name in self.merger.current_sheets:
                df = self.merger.current_sheets[self.sheet_name]
                if all(col in df.columns for col in selected_columns):
                    sample_data = df[selected_columns].head(3)
                    self.preview_text.insert(tk.END, "Sample data (first 3 rows):\n")
                    for i, row in sample_data.iterrows():
                        row_display = ", ".join([f"{col}: {row[col]}" for col in selected_columns])
                        self.preview_text.insert(tk.END, f"Row {i+1}: {row_display}\n")
                    
                    # If "Stack Values" strategy is selected, show a preview of how it would look
                    if self.strategy_var.get() == "stack_values":
                        self.preview_text.insert(tk.END, "\nStack Values Preview (will create new rows):\n")
                        stack_preview = []
                        for i, row in sample_data.iterrows():
                            for col in selected_columns:
                                if pd.notna(row[col]) and row[col] != "":
                                    stack_preview.append(f"{row[col]}")
                        
                        for i, val in enumerate(stack_preview[:5]):  # Show first 5 values
                            self.preview_text.insert(tk.END, f"New Row {i+1}: {val}\n")
                        
                        if len(stack_preview) > 5:
                            self.preview_text.insert(tk.END, "... (more rows will be created)\n")
        except Exception as e:
            self.preview_text.insert(tk.END, f"Error previewing data: {str(e)}")
    
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
        
        try:
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
            
            # Handle stack_values strategy separately
            if strategy == "stack_values":
                result, empty_cols, rows_added = self.stack_values_merge(
                    self.sheet_name,
                    selected_columns,
                    new_column_name,
                    delete_source,
                    remove_empty
                )
                
                if result:
                    success_message = f"Columns stacked into '{new_column_name}'"
                    if rows_added > 0:
                        success_message += f"\nAdded {rows_added} new rows to preserve all values"
                    
                    if delete_source:
                        success_message += f"\nDeleted {len(selected_columns)} source columns"
                    
                    if remove_empty and empty_cols:
                        success_message += f"\nRemoved {len(empty_cols)} empty columns"
                        
                    messagebox.showinfo("Success", success_message)
                    self.window.destroy()
                    
                    # Signal to parent to refresh
                    self.parent.event_generate("<<RefreshSheets>>")
                else:
                    messagebox.showerror("Error", "Failed to stack columns")
            else:
                # Use the original merge method for other strategies
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
                    
        except Exception as e:
            messagebox.showerror("Error", str(e))

    
    def is_column_empty(self, df, column_name):
        """
        Check if a column is empty (all values are NaN or empty strings)
        
        Args:
            df: DataFrame to check
            column_name: Name of the column to check
            
        Returns:
            bool: True if the column is empty, False otherwise
        """
        if column_name not in df.columns:
            return True
            
        # Check if all values are NaN or empty strings
        return df[column_name].isna().all() or (df[column_name].astype(str).str.strip() == "").all()
    
    
    def stack_values_merge(self, sheet_name, columns, new_column_name, delete_source=True, remove_empty=True):
        """
        Merge columns by stacking their values in separate rows.
        This preserves all data by creating new rows when needed.
        
        Args:
            sheet_name: Name of the sheet to modify
            columns: List of column names to merge
            new_column_name: Name for the merged column
            delete_source: Whether to delete source columns after merging
            remove_empty: Whether to remove empty columns from the sheet
            
        Returns:
            tuple: (success, empty_columns_removed, rows_added)
        """
        try:
            if sheet_name not in self.merger.current_sheets:
                return False, [], 0
                
            # Get the dataframe
            df = self.merger.current_sheets[sheet_name].copy()
            
            # Check if all columns exist
            if not all(col in df.columns for col in columns):
                messagebox.showerror("Error", "One or more selected columns don't exist in the sheet")
                return False, [], 0
            
            # First, let's create a map of indices where each column has data
            columns_data = {}
            for col in columns:
                # Get indices where the column has data (not empty or NA)
                has_data = df[col].notna() & (df[col].astype(str).str.strip() != "")
                columns_data[col] = set(df.index[has_data])
            
            # Get all indices where at least one column has data
            all_data_indices = set()
            for indices in columns_data.values():
                all_data_indices.update(indices)
            
            # Create a new dataframe with the same schema as the original, but empty
            non_merge_cols = [c for c in df.columns if c not in columns]
            
            # Build the result dataframe as a list of rows and then convert to DataFrame at the end
            # This avoids the FutureWarning from pandas
            result_rows = []
            
            # For each row in the original dataframe
            for idx in df.index:
                row_data = df.loc[idx]
                
                # Create a base row with all non-merge columns
                base_row = {c: row_data[c] for c in non_merge_cols}
                
                # Count how many values we added for this row
                values_added = 0
                
                # For each merge column, add its value if non-empty
                for col in columns:
                    if idx in columns_data[col]:  # Column has data at this index
                        new_row = base_row.copy()
                        new_row[new_column_name] = row_data[col]
                        # Add the new row to our results list
                        result_rows.append(new_row)
                        values_added += 1
                
                # If no values were added and we want to preserve row structure
                if values_added == 0:
                    # Add a row with empty value in the merged column
                    base_row[new_column_name] = None
                    result_rows.append(base_row)
            
            # Convert the list of rows to a DataFrame
            result_df = pd.DataFrame(result_rows)
            
            # If the result is empty, create an empty DataFrame with the correct columns
            if result_df.empty:
                result_df = pd.DataFrame(columns=non_merge_cols + [new_column_name])
            
            # Calculate how many new rows were added
            rows_added = len(result_df) - len(df)
            
            # Apply the changes to the sheet
            self.merger.current_sheets[sheet_name] = result_df
            
            # Track empty columns if needed
            empty_cols = []
            if remove_empty:
                for col in self.merger.current_sheets[sheet_name].columns:
                    if self.merger.is_column_empty(self.merger.current_sheets[sheet_name], col):
                        empty_cols.append(col)
                
                # Remove empty columns
                if empty_cols:
                    self.merger.current_sheets[sheet_name] = self.merger.current_sheets[sheet_name].drop(columns=empty_cols)
            
            # Mark sheet as modified in the merger
            self.merger.modified_sheets.add(sheet_name)
            
            return True, empty_cols, rows_added
        except Exception as e:
            messagebox.showerror("Error", f"Failed to stack columns: {str(e)}")
            return False, [], 0
        
    