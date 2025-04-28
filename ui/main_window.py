import tkinter as tk
from tkinter import ttk, messagebox
import os

from core.file_operations import FileOperations
from ui.manual_merge import ManualMergeWindow
from ui.compare_columns import CompareColumnsWindow
from ui.column_preview import ColumnPreviewWindow
from ui.common import (
    center_window, 
    create_header, 
    create_file_selector, 
    create_sheet_selector,
    create_status_bar,
    create_columns_listbox,
    populate_columns_listbox,
    create_button_frame,
    add_button
)

class MainWindow:
    """
    Main application window
    """
    def __init__(self, root, merger):
        self.root = root
        self.merger = merger
        
        # Window setup
        root.title("Excel Column Merger")
        root.geometry("800x600")
        center_window(root, 800, 600)
        
        # Variables
        self.file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.compare_sheet_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        
        # Create frames
        title_frame = create_header(
            root, 
            "Excel Column Merger", 
            "This tool identifies duplicate column names, allows manual merging of columns, and can compare values across columns to find duplicates."
        )
        title_frame.pack(fill="x")
        
        file_frame = create_file_selector(root, self.file_var, self.select_file)
        file_frame.pack(fill="x")
        
        # Create notebook with tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Auto Merge tab
        self.auto_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.auto_frame, text="Auto Merge")
        self.setup_auto_merge_tab()
        
        # Manual Merge tab
        self.manual_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.manual_frame, text="Manual Merge")
        self.setup_manual_merge_tab()
        
        # Compare Columns tab
        self.compare_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.compare_frame, text="Compare Columns")
        self.setup_compare_columns_tab()
        
        # Column Preview tab
        self.preview_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.preview_frame, text="Column Preview")
        self.setup_preview_tab()
        
        # Button frame
        self.button_frame = create_button_frame(root)
        self.button_frame.pack(fill="x", side="bottom")
        
        # Action buttons
        self.analyze_button = add_button(self.button_frame, "Analyze for Duplicates", self.analyze_file, "left")
        self.analyze_button.configure(state="disabled")
        
        self.auto_merge_button = add_button(self.button_frame, "Auto Merge", self.auto_merge_columns, "left")
        self.auto_merge_button.configure(state="disabled")
        
        self.manual_merge_button = add_button(self.button_frame, "Manual Merge Selected", self.open_manual_merge, "left")
        self.manual_merge_button.configure(state="disabled")
        
        self.compare_columns_button = add_button(self.button_frame, "Compare Columns", self.open_compare_columns, "left")
        self.compare_columns_button.configure(state="disabled")
        
        self.preview_column_button = add_button(self.button_frame, "Preview Column", self.open_column_preview, "left")
        self.preview_column_button.configure(state="disabled")
        
        self.save_button = add_button(self.button_frame, "Save File", self.save_file, "left")
        self.save_button.configure(state="disabled")
        
        add_button(self.button_frame, "Exit", root.destroy, "right")
        
        # Status bar
        self.status_bar = create_status_bar(root, self.status_var)
        
        # Refresh handler
        root.bind("<<RefreshSheets>>", self.refresh_after_update)
        
        # Set focus
        root.focus_force()
    
    def setup_auto_merge_tab(self):
        """Setup the Auto Merge tab"""
        # Auto merge options
        auto_options_frame = ttk.Frame(self.auto_frame)
        auto_options_frame.pack(fill="x", pady=5)
        
        ttk.Label(auto_options_frame, text="Merge Strategy:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.merge_strategy = tk.StringVar(value="first_non_empty")
        strategies = [
            ("First Non-Empty Value (default)", "first_non_empty"),
            ("Sum Numeric Values", "sum"),
            ("Concatenate Text Values", "concatenate")
        ]
        
        for i, (text, value) in enumerate(strategies):
            ttk.Radiobutton(
                auto_options_frame, 
                text=text, 
                variable=self.merge_strategy, 
                value=value
            ).grid(row=0, column=i+1, sticky="w", padx=5, pady=5)
        
        # Results treeview for auto merge
        ttk.Label(self.auto_frame, text="Duplicate Columns (Case-Insensitive):").pack(anchor="w", padx=5, pady=5)
        
        auto_columns = ("sheet", "column", "duplicates")
        self.auto_results_tree = ttk.Treeview(self.auto_frame, columns=auto_columns, show="headings")
        
        self.auto_results_tree.heading("sheet", text="Sheet")
        self.auto_results_tree.heading("column", text="Base Column Name")
        self.auto_results_tree.heading("duplicates", text="Duplicate Columns")
        
        self.auto_results_tree.column("sheet", width=100)
        self.auto_results_tree.column("column", width=150)
        self.auto_results_tree.column("duplicates", width=500)
        
        auto_scrollbar = ttk.Scrollbar(self.auto_frame, orient="vertical", command=self.auto_results_tree.yview)
        self.auto_results_tree.configure(yscrollcommand=auto_scrollbar.set)
        
        self.auto_results_tree.pack(side="left", fill="both", expand=True)
        auto_scrollbar.pack(side="right", fill="y")
    
    def setup_manual_merge_tab(self):
        """Setup the Manual Merge tab"""
        # Sheet selector
        manual_top_frame, self.sheet_selector = create_sheet_selector(self.manual_frame, self.sheet_var)
        manual_top_frame.pack(fill="x", pady=5)
        
        # Bind sheet selector to refresh columns
        self.sheet_selector.bind('<<ComboboxSelected>>', lambda e: self.refresh_columns_display())
        
        # Columns display
        ttk.Label(self.manual_frame, text="Available Columns:").pack(anchor="w", padx=5, pady=5)
        
        columns_frame, self.columns_listbox = create_columns_listbox(self.manual_frame, height=15)
        columns_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    def setup_compare_columns_tab(self):
        """Setup the Compare Columns tab"""
        # Sheet selector
        compare_top_frame, self.compare_sheet_selector = create_sheet_selector(self.compare_frame, self.compare_sheet_var)
        compare_top_frame.pack(fill="x", pady=5)
        
        # Bind sheet selector to refresh columns
        self.compare_sheet_selector.bind('<<ComboboxSelected>>', lambda e: self.refresh_compare_columns_display())
        
        # Columns display
        ttk.Label(self.compare_frame, text="Available Columns:").pack(anchor="w", padx=5, pady=5)
        
        compare_columns_frame, self.compare_columns_listbox = create_columns_listbox(self.compare_frame, height=15)
        compare_columns_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    def setup_preview_tab(self):
        """Setup the Column Preview tab with filter for non-empty columns"""
        # Sheet selector
        preview_top_frame, self.preview_sheet_selector = create_sheet_selector(self.preview_frame, tk.StringVar())
        preview_top_frame.pack(fill="x", pady=5)
        
        # Bind sheet selector to refresh columns
        self.preview_sheet_selector.bind('<<ComboboxSelected>>', lambda e: self.refresh_preview_columns_display())
        
        # Add filter options
        filter_frame = ttk.Frame(self.preview_frame)
        filter_frame.pack(fill="x", pady=5)
        
        # Show empty columns checkbox
        self.show_empty_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            filter_frame,
            text="Show empty columns",
            variable=self.show_empty_var,
            command=self.refresh_preview_columns_display
        ).pack(side="left", padx=5)
        
        # Columns display
        ttk.Label(self.preview_frame, text="Select a column to preview:").pack(anchor="w", padx=5, pady=5)
        
        preview_columns_frame, self.preview_columns_listbox = create_columns_listbox(self.preview_frame, height=15)
        preview_columns_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Single selection mode
        self.preview_columns_listbox.config(selectmode=tk.SINGLE)
        
        # Add a filter/search box
        search_frame = ttk.Frame(self.preview_frame)
        search_frame.pack(fill="x", pady=5)
        
        ttk.Label(search_frame, text="Filter columns:").pack(side="left", padx=5)
        self.column_filter_var = tk.StringVar()
        column_filter_entry = ttk.Entry(search_frame, textvariable=self.column_filter_var, width=30)
        column_filter_entry.pack(side="left", padx=5)
        
        # Bind to refresh on entry change
        column_filter_entry.bind("<KeyRelease>", lambda e: self.refresh_preview_columns_display())
    
    def select_file(self):
        """Let user select the Excel file to process"""
        try:
            file_path = FileOperations.select_file()
            if file_path:
                self.file_var.set(file_path)
                
                # Set the file in the merger
                self.merger.set_input_file(file_path)
                
                # Clear previous results
                for item in self.auto_results_tree.get_children():
                    self.auto_results_tree.delete(item)
                
                # Update sheet selectors
                self.update_sheet_selectors()
                
                self.status_var.set("File selected. Use tabs to perform different operations.")
                
                # Enable buttons
                self.analyze_button.configure(state="normal")
                self.manual_merge_button.configure(state="normal")
                self.compare_columns_button.configure(state="normal")
                self.preview_column_button.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def update_sheet_selectors(self):
        """Update all sheet selectors with current sheet names"""
        if hasattr(self.merger, 'current_sheets') and self.merger.current_sheets:
            sheet_names = list(self.merger.current_sheets.keys())
            
            self.sheet_selector['values'] = sheet_names
            self.compare_sheet_selector['values'] = sheet_names
            self.preview_sheet_selector['values'] = sheet_names
            
            if sheet_names:
                self.sheet_selector.current(0)
                self.compare_sheet_selector.current(0)
                self.preview_sheet_selector.current(0)
                
                # Refresh column displays
                self.refresh_columns_display()
                self.refresh_compare_columns_display()
                self.refresh_preview_columns_display()
    
    def refresh_columns_display(self):
        """Refresh the columns display in the Manual Merge tab"""
        selected_sheet = self.sheet_var.get()
        if selected_sheet and selected_sheet in self.merger.current_sheets:
            df = self.merger.current_sheets[selected_sheet]
            populate_columns_listbox(self.columns_listbox, df.columns)
    
    def refresh_compare_columns_display(self):
        """Refresh the columns display in the Compare Columns tab"""
        selected_sheet = self.compare_sheet_var.get()
        if selected_sheet and selected_sheet in self.merger.current_sheets:
            df = self.merger.current_sheets[selected_sheet]
            populate_columns_listbox(self.compare_columns_listbox, df.columns)
    
    def refresh_preview_columns_display(self):
        """Refresh the columns display in the Preview tab with filtering options"""
        selected_sheet = self.preview_sheet_selector.get()
        if selected_sheet and selected_sheet in self.merger.current_sheets:
            df = self.merger.current_sheets[selected_sheet]
            
            # Get columns based on empty filter
            if self.show_empty_var.get():
                # Show all columns
                columns = list(df.columns)
            else:
                # Show only non-empty columns
                columns = self.merger.get_non_empty_columns(selected_sheet)
            
            # Apply name filter if any
            filter_text = self.column_filter_var.get().lower()
            if filter_text:
                columns = [col for col in columns if filter_text in col.lower()]
            
            # Clear existing items
            self.preview_columns_listbox.delete(0, tk.END)
            
            # Add filtered columns
            for col in columns:
                self.preview_columns_listbox.insert(tk.END, col)
            
            # Update status message with count
            total_columns = len(df.columns)
            if not self.show_empty_var.get():
                non_empty_count = len(self.merger.get_non_empty_columns(selected_sheet))
                empty_count = total_columns - non_empty_count
                shown_count = len(columns)
                
                if filter_text:
                    self.status_var.set(f"Showing {shown_count} columns matching '{filter_text}'. {empty_count} empty columns are hidden.")
                else:
                    self.status_var.set(f"Showing {non_empty_count} columns with data. {empty_count} empty columns are hidden.")
            else:
                shown_count = len(columns)
                if filter_text:
                    self.status_var.set(f"Showing {shown_count} columns matching '{filter_text}' out of {total_columns} total columns.")
                else:
                    self.status_var.set(f"Showing all {total_columns} columns (including empty ones).")
    
    def analyze_file(self):
        """Analyze the file for duplicate columns"""
        if not self.merger.input_file:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
            
        self.status_var.set("Analyzing file for duplicate columns...")
        
        # Clear previous results
        for item in self.auto_results_tree.get_children():
            self.auto_results_tree.delete(item)
            
        try:
            # Analyze the file
            analysis = self.merger.analyze_file()
            
            if not analysis:
                self.status_var.set("Analysis failed. Please try again.")
                return
                
            if analysis["status"] == "error":
                messagebox.showerror("Error", f"Failed to analyze file: {analysis['message']}")
                self.status_var.set("Analysis failed.")
                return
                
            if analysis["status"] == "no_duplicates":
                messagebox.showinfo("No Duplicates", "No duplicate column names found in the Excel file.")
                self.status_var.set("No duplicate columns found.")
                return
                
            # Populate the results tree
            sheets_data = analysis["sheets_data"]
            for sheet_name, sheet_data in sheets_data.items():
                duplicate_columns = sheet_data["duplicate_columns"]
                
                for base_name, columns in duplicate_columns.items():
                    self.auto_results_tree.insert(
                        "", 
                        "end", 
                        values=(
                            sheet_name,
                            base_name, 
                            ", ".join(columns)
                        )
                    )
                    
            # Store analysis
            self.file_analysis = analysis
            
            # Update status
            total_duplicates = sum(len(sheet["duplicate_columns"]) for sheet in sheets_data.values())
            self.status_var.set(f"Found {total_duplicates} duplicate column groups. Choose a merge strategy and click 'Auto Merge'.")
            
            # Select the Auto Merge tab
            self.notebook.select(0)
            
            # Enable merge button
            self.auto_merge_button.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Analysis failed.")
    
    def auto_merge_columns(self):
        """Auto merge duplicate columns"""
        if not hasattr(self, 'file_analysis'):
            messagebox.showerror("Error", "Please analyze the file first.")
            return
            
        strategy = self.merge_strategy.get()
        self.status_var.set(f"Merging columns using '{strategy}' strategy...")
        
        try:
            # Perform the merge
            if self.merger.merge_columns(self.file_analysis, strategy):
                messagebox.showinfo(
                    "Success", 
                    "Duplicate columns merged successfully!\n\nYou can now use the other tabs to perform additional operations, or save the file."
                )
                self.status_var.set("Columns merged. You can now use other tabs or save the file.")
                
                # Update sheet selectors
                self.update_sheet_selectors()
                
                # Enable save button
                self.save_button.configure(state="normal")
            else:
                self.status_var.set("Failed to merge columns.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Failed to merge columns.")
    
    def open_manual_merge(self):
        """Open the manual merge window"""
        selected_sheet = self.sheet_var.get()
        if not selected_sheet or selected_sheet not in self.merger.current_sheets:
            messagebox.showerror("Error", "Please select a valid sheet.")
            return
            
        df = self.merger.current_sheets[selected_sheet]
        column_list = list(df.columns)
        
        if len(column_list) < 2:
            messagebox.showwarning("Warning", "Sheet has less than 2 columns. No merging possible.")
            return
            
        # Open the manual merge window
        ManualMergeWindow(self.root, self.merger, selected_sheet, column_list)
    
    def open_compare_columns(self):
        """Open the compare columns window"""
        selected_sheet = self.compare_sheet_var.get()
        if not selected_sheet or selected_sheet not in self.merger.current_sheets:
            messagebox.showerror("Error", "Please select a valid sheet.")
            return
            
        df = self.merger.current_sheets[selected_sheet]
        column_list = list(df.columns)
        
        if len(column_list) < 2:
            messagebox.showwarning("Warning", "Sheet has less than 2 columns. No comparison possible.")
            return
            
        # Open the compare columns window
        CompareColumnsWindow(self.root, self.merger, selected_sheet, column_list)
    
    def open_column_preview(self):
        """Open the column preview window"""
        selected_sheet = self.preview_sheet_selector.get()
        if not selected_sheet or selected_sheet not in self.merger.current_sheets:
            messagebox.showerror("Error", "Please select a valid sheet.")
            return
        
        selected_indices = self.preview_columns_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select a column to preview.")
            return
        
        # Get column name from listbox content rather than index
        selected_column = self.preview_columns_listbox.get(selected_indices[0])
        
        # Open the column preview window
        ColumnPreviewWindow(self.root, self.merger, selected_sheet, selected_column)
    
    def save_file(self):
        """Save the modified Excel file"""
        if not hasattr(self.merger, 'current_sheets') or not self.merger.current_sheets:
            messagebox.showerror("Error", "No data to save.")
            return
            
        self.status_var.set("Saving file...")
        
        try:
            # Save the file
            output_file = FileOperations.save_merged_file(self.merger)
            
            if output_file:
                messagebox.showinfo(
                    "Success", 
                    f"File saved successfully as:\n{output_file}"
                )
                
                # Ask if the user wants to open the file
                if messagebox.askyesno("Open File", "Do you want to open the merged file?"):
                    try:
                        FileOperations.open_file(output_file)
                    except Exception as e:
                        messagebox.showwarning("Warning", str(e))
                
                self.status_var.set(f"File saved as {os.path.basename(output_file)}")
            else:
                self.status_var.set("Save cancelled.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Save failed.")
    
    def refresh_after_update(self, event):
        """Refresh displays after a sheet update"""
        self.refresh_columns_display()
        self.refresh_compare_columns_display()
        self.refresh_preview_columns_display()
        self.save_button.configure(state="normal")
    
    