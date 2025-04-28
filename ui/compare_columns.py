import tkinter as tk
from tkinter import ttk, messagebox

class CompareColumnsWindow:
    """
    Window for comparing columns for duplicate values
    """
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
        try:
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
                
        except Exception as e:
            messagebox.showerror("Error", f"Error comparing columns: {str(e)}")
            self.results_text.insert(tk.END, f"Error: {str(e)}")
    
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
        except Exception as e:
            messagebox.showerror("Error", f"Error creating common column: {str(e)}")

    def __init__(self, parent, merger, sheet_name, column_list):
        self.parent = parent
        self.merger = merger
        self.sheet_name = sheet_name
        self.column_list = column_list
        
        # Create a new window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Compare Columns for Duplicates - {sheet_name}")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create frames and add all content
        # [existing code for creating frames and content]
        
        # After adding all content, update window size to fit contents
        self.window.update_idletasks()  # Make sure all widgets are created
        width = self.window.winfo_reqwidth()
        height = self.window.winfo_reqheight()
        
        # Add some padding
        width += 20
        height += 20
        
        # Set minimum dimensions
        width = max(width, 750)
        height = max(height, 700)
        
        # Center the window with the calculated dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.window.geometry(f"{width}x{height}+{int(x)}+{int(y)}")