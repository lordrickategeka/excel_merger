import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class ColumnPreviewWindow:
    """
    Window for previewing column data and statistics, with option to delete the column
    """
    def __init__(self, parent, merger, sheet_name, column_name):
        self.parent = parent
        self.merger = merger
        self.sheet_name = sheet_name
        self.column_name = column_name
        
        # Create a new window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Column Preview: {column_name} - {sheet_name}")
        self.window.geometry("700x700")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        window_width = 700
        window_height = 700
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width / 2) - (window_width / 2)
        y = (screen_height / 2) - (window_height / 2)
        self.window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
        
        # Create frames
        header_frame = ttk.Frame(self.window, padding="10")
        header_frame.pack(fill="x")
        
        stats_frame = ttk.LabelFrame(self.window, text="Column Statistics", padding="10")
        stats_frame.pack(fill="x", padx=10, pady=5)
        
        sample_frame = ttk.LabelFrame(self.window, text="Sample Data", padding="10")
        sample_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        vis_frame = ttk.LabelFrame(self.window, text="Data Visualization", padding="10")
        vis_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        button_frame = ttk.Frame(self.window, padding="10")
        button_frame.pack(fill="x", side="bottom")
        
        # Header
        ttk.Label(
            header_frame, 
            text=f"Column: {column_name}", 
            font=("Helvetica", 12, "bold")
        ).pack(pady=5)
        
        # Analyze the column
        try:
            self.analysis = self.merger.analyze_column(sheet_name, column_name)
            
            if not self.analysis:
                messagebox.showerror("Error", f"Failed to analyze column '{column_name}'.")
                self.window.destroy()
                return
                
            # Stats display
            self.create_stats_display(stats_frame)
            
            # Sample data display
            self.create_sample_display(sample_frame)
            
            # Data visualization
            self.create_visualization(vis_frame)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error analyzing column: {str(e)}")
            self.window.destroy()
            return
        
        # Delete confirmation frame
        delete_frame = ttk.LabelFrame(self.window, text="Column Actions", padding="10")
        delete_frame.pack(fill="x", padx=10, pady=5)
        
        self.delete_warning = ttk.Label(
            delete_frame,
            text="Warning: Deleting a column cannot be undone. Make sure to save a backup of your file.",
            foreground="red",
            wraplength=650
        )
        self.delete_warning.pack(pady=5)
        
        # Delete button
        self.delete_button = ttk.Button(
            delete_frame,
            text="Delete This Column",
            command=self.confirm_delete_column,
            style="Delete.TButton"
        )
        self.delete_button.pack(pady=5)
        
        # Create a custom style for the delete button
        self.window.style = ttk.Style()
        self.window.style.configure("Delete.TButton", foreground="red")
        
        # Close button
        ttk.Button(button_frame, text="Close", command=self.window.destroy).pack(side="right", padx=5)
    
    def create_stats_display(self, parent_frame):
        """Create the statistics display frame"""
        stats_grid = ttk.Frame(parent_frame)
        stats_grid.pack(fill="x", pady=5)
        
        # Row 1
        ttk.Label(stats_grid, text="Total Rows:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(stats_grid, text=str(self.analysis["total_rows"])).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(stats_grid, text="Data Type:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, sticky="w", padx=5, pady=2)
        ttk.Label(stats_grid, text=self.analysis["data_type"]).grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        # Row 2
        ttk.Label(stats_grid, text="Non-Empty Values:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(stats_grid, text=f"{self.analysis['non_empty_count']} ({100 - self.analysis['empty_percentage']:.1f}%)").grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(stats_grid, text="Empty Values:", font=("Helvetica", 10, "bold")).grid(row=1, column=2, sticky="w", padx=5, pady=2)
        ttk.Label(stats_grid, text=f"{self.analysis['empty_count']} ({self.analysis['empty_percentage']:.1f}%)").grid(row=1, column=3, sticky="w", padx=5, pady=2)
        
        # Row 3
        ttk.Label(stats_grid, text="Unique Values:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(stats_grid, text=str(self.analysis["unique_count"])).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        # Data type specific stats
        if self.analysis["data_type"] == "Numeric" and "numeric_stats" in self.analysis:
            stats = self.analysis["numeric_stats"]
            
            ttk.Label(stats_grid, text="Min:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["min"])).grid(row=3, column=1, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Max:", font=("Helvetica", 10, "bold")).grid(row=3, column=2, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["max"])).grid(row=3, column=3, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Mean:", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{stats['mean']:.2f}" if stats["mean"] is not None else "N/A").grid(row=4, column=1, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Median:", font=("Helvetica", 10, "bold")).grid(row=4, column=2, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["median"])).grid(row=4, column=3, sticky="w", padx=5, pady=2)
            
        elif self.analysis["data_type"] == "Text" and "text_stats" in self.analysis:
            stats = self.analysis["text_stats"]
            
            ttk.Label(stats_grid, text="Min Length:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["min_length"])).grid(row=3, column=1, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Max Length:", font=("Helvetica", 10, "bold")).grid(row=3, column=2, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["max_length"])).grid(row=3, column=3, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Avg Length:", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{stats['avg_length']:.1f}" if stats["avg_length"] is not None else "N/A").grid(row=4, column=1, sticky="w", padx=5, pady=2)
            
        elif self.analysis["data_type"] == "Date/Time" and "date_stats" in self.analysis:
            stats = self.analysis["date_stats"]
            
            ttk.Label(stats_grid, text="Earliest:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["earliest"])).grid(row=3, column=1, sticky="w", padx=5, pady=2)
            
            ttk.Label(stats_grid, text="Latest:", font=("Helvetica", 10, "bold")).grid(row=3, column=2, sticky="w", padx=5, pady=2)
            ttk.Label(stats_grid, text=str(stats["latest"])).grid(row=3, column=3, sticky="w", padx=5, pady=2)
    
    def create_sample_display(self, parent_frame):
        """Create the sample data display"""
        # Create a Treeview to display sample data
        sample_columns = ["Row", "Value"]
        self.sample_tree = ttk.Treeview(parent_frame, columns=sample_columns, show="headings", height=5)
        
        self.sample_tree.heading("Row", text="Row #")
        self.sample_tree.heading("Value", text="Value")
        
        self.sample_tree.column("Row", width=50)
        self.sample_tree.column("Value", width=600)
        
        # Add scrollbar
        sample_scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=self.sample_tree.yview)
        self.sample_tree.configure(yscrollcommand=sample_scrollbar.set)
        
        self.sample_tree.pack(side="left", fill="both", expand=True)
        sample_scrollbar.pack(side="right", fill="y")
        
        # Controls for viewing more data
        controls_frame = ttk.Frame(parent_frame)
        controls_frame.pack(side="bottom", fill="x", pady=5)
        
        # Add buttons to view more data
        ttk.Button(controls_frame, text="View First 20 Rows", command=lambda: self.load_data_sample(0, 20)).pack(side="left", padx=5)
        ttk.Button(controls_frame, text="View More Rows", command=self.load_more_data).pack(side="left", padx=5)
        
        # Add search functionality
        search_frame = ttk.Frame(parent_frame)
        search_frame.pack(side="bottom", fill="x", pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", padx=5)
        ttk.Button(search_frame, text="Find", command=self.search_data).pack(side="left", padx=5)
        
        # Get the column data from the sheet and load sample
        self.load_data_sample(0, 20)
    
    def load_data_sample(self, start_row, num_rows):
        """Load a sample of data from the column"""
        # Clear existing data
        for item in self.sample_tree.get_children():
            self.sample_tree.delete(item)
        
        # Get the column data from the sheet
        df = self.merger.current_sheets[self.sheet_name]
        column_data = df[self.column_name]
        
        # Ensure we don't try to access rows that don't exist
        end_row = min(start_row + num_rows, len(column_data))
        
        # Add rows of data to the tree
        for i in range(start_row, end_row):
            value = column_data.iloc[i]
            # Format the value display
            if pd.isna(value):
                value_display = "(empty)"
            else:
                value_display = str(value)
                # Truncate very long values
                if len(value_display) > 100:
                    value_display = value_display[:100] + "..."
            
            self.sample_tree.insert("", "end", values=(i+1, value_display))
    
    def load_more_data(self):
        """Load more data rows, using a dialog to specify range"""
        # Create a dialog to get row range
        dialog = tk.Toplevel(self.window)
        dialog.title("Load Data Range")
        dialog.geometry("300x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # Center the dialog
        center_x = self.window.winfo_x() + (self.window.winfo_width() // 2) - 150
        center_y = self.window.winfo_y() + (self.window.winfo_height() // 2) - 75
        dialog.geometry(f"+{center_x}+{center_y}")
        
        # Add fields
        ttk.Label(dialog, text="Start Row:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        start_var = tk.StringVar(value="1")
        ttk.Entry(dialog, textvariable=start_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Number of Rows:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        count_var = tk.StringVar(value="20")
        ttk.Entry(dialog, textvariable=count_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        
        # Info about total rows
        total_rows = len(self.merger.current_sheets[self.sheet_name])
        ttk.Label(dialog, text=f"Total rows in sheet: {total_rows}").grid(row=2, column=0, columnspan=2, padx=5, pady=5)
        
        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        def load_range():
            try:
                start = int(start_var.get()) - 1  # Convert to 0-based index
                count = int(count_var.get())
                
                # Validate input
                if start < 0:
                    messagebox.showerror("Invalid Input", "Start row must be at least 1")
                    return
                
                if count <= 0:
                    messagebox.showerror("Invalid Input", "Number of rows must be positive")
                    return
                
                if start >= total_rows:
                    messagebox.showerror("Invalid Input", "Start row exceeds available data")
                    return
                
                # Load the data
                self.load_data_sample(start, count)
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid numbers")
        
        ttk.Button(button_frame, text="Load", command=load_range).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)
    
    def search_data(self):
        """Search for a value in the column data"""
        search_term = self.search_var.get().strip().lower()
        if not search_term:
            messagebox.showinfo("Search", "Please enter a search term")
            return
        
        # Get the column data
        df = self.merger.current_sheets[self.sheet_name]
        column_data = df[self.column_name]
        
        # Convert all data to strings for searching
        str_data = column_data.astype(str).str.lower()
        
        # Find matches
        matches = str_data.str.contains(search_term, na=False)
        match_indices = matches[matches].index.tolist()
        
        if not match_indices:
            messagebox.showinfo("Search Results", f"No matches found for '{search_term}'")
            return
        
        # Show results in a new dialog
        results_dialog = tk.Toplevel(self.window)
        results_dialog.title(f"Search Results for '{search_term}'")
        results_dialog.geometry("600x400")
        results_dialog.transient(self.window)
        
        # Center the dialog
        center_x = self.window.winfo_x() + (self.window.winfo_width() // 2) - 300
        center_y = self.window.winfo_y() + (self.window.winfo_height() // 2) - 200
        results_dialog.geometry(f"+{center_x}+{center_y}")
        
        # Create a frame for the results
        results_frame = ttk.Frame(results_dialog, padding="10")
        results_frame.pack(fill="both", expand=True)
        
        # Create a Treeview for the results
        results_columns = ["Row", "Value"]
        results_tree = ttk.Treeview(results_frame, columns=results_columns, show="headings", height=15)
        
        results_tree.heading("Row", text="Row #")
        results_tree.heading("Value", text="Value")
        
        results_tree.column("Row", width=50)
        results_tree.column("Value", width=500)
        
        # Add scrollbar
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
        results_tree.configure(yscrollcommand=results_scrollbar.set)
        
        results_tree.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")
        
        # Add results to the tree
        for idx in match_indices:
            row_num = idx + 1  # Convert to 1-based indexing for display
            value = column_data.iloc[idx]
            
            # Format the value display
            if pd.isna(value):
                value_display = "(empty)"
            else:
                value_display = str(value)
                # Truncate very long values
                if len(value_display) > 100:
                    value_display = value_display[:100] + "..."
            
            results_tree.insert("", "end", values=(row_num, value_display))
        
        # Add a close button
        ttk.Button(results_dialog, text="Close", command=results_dialog.destroy).pack(pady=10)
        
        # Show count
        ttk.Label(results_dialog, text=f"Found {len(match_indices)} matches.").pack(pady=5)
    
    def create_visualization(self, parent_frame):
        """Create data visualization based on column type"""
        # Get the column data
        df = self.merger.current_sheets[self.sheet_name]
        column_data = df[self.column_name]
        
        # Create a figure and axis
        figure = plt.Figure(figsize=(6, 4), dpi=100)
        ax = figure.add_subplot(111)
        
        # Different visualizations based on data type
        if self.analysis["data_type"] == "Numeric":
            # Histogram for numeric data
            try:
                ax.hist(column_data.dropna(), bins=20)
                ax.set_title(f"Distribution of {self.column_name}")
                ax.set_xlabel("Value")
                ax.set_ylabel("Frequency")
            except Exception as e:
                ax.text(0.5, 0.5, f"Could not create histogram: {str(e)}", 
                        horizontalalignment='center', verticalalignment='center')
                
        elif self.analysis["data_type"] in ["Text", "Mixed"]:
            # Bar chart of top values or empty/non-empty
            try:
                # For text data, show distribution of top 10 values
                if self.analysis["unique_count"] <= 10:
                    # If 10 or fewer unique values, show all of them
                    value_counts = column_data.value_counts().head(10)
                    value_counts.plot(kind='bar', ax=ax)
                    ax.set_title(f"Value Distribution: {self.column_name}")
                    ax.set_xlabel("Value")
                    ax.set_ylabel("Count")
                    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
                else:
                    # If more than 10 unique values, just show empty vs non-empty
                    empty_data = [self.analysis["empty_count"], self.analysis["non_empty_count"]]
                    ax.bar(["Empty", "Non-Empty"], empty_data)
                    ax.set_title(f"Empty vs. Non-Empty: {self.column_name}")
                    ax.set_ylabel("Count")
            except Exception as e:
                ax.text(0.5, 0.5, f"Could not create chart: {str(e)}", 
                        horizontalalignment='center', verticalalignment='center')
                
        elif self.analysis["data_type"] == "Date/Time":
            # Timeline or count by month/year for date data
            try:
                # Create timeline showing distribution over time
                column_data.dropna().groupby(column_data.dt.year).count().plot(kind='bar', ax=ax)
                ax.set_title(f"Date Distribution by Year: {self.column_name}")
                ax.set_xlabel("Year")
                ax.set_ylabel("Count")
            except Exception as e:
                ax.text(0.5, 0.5, f"Could not create timeline: {str(e)}", 
                        horizontalalignment='center', verticalalignment='center')
        else:
            # Generic empty vs non-empty chart
            empty_data = [self.analysis["empty_count"], self.analysis["non_empty_count"]]
            ax.bar(["Empty", "Non-Empty"], empty_data)
            ax.set_title(f"Empty vs. Non-Empty: {self.column_name}")
            ax.set_ylabel("Count")
        
        # Create canvas
        canvas = FigureCanvasTkAgg(figure, parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def confirm_delete_column(self):
        """Confirm and delete the column"""
        # Create confirmation dialog
        result = messagebox.askokcancel(
            "Confirm Delete",
            f"Are you sure you want to delete the column '{self.column_name}'?\n\nThis action cannot be undone.",
            icon="warning"
        )
        
        if result:
            try:
                # Get the dataframe
                df = self.merger.current_sheets[self.sheet_name]
                
                # Delete the column
                df = df.drop(columns=[self.column_name])
                
                # Update the dataframe in the merger
                self.merger.current_sheets[self.sheet_name] = df
                
                messagebox.showinfo("Success", f"Column '{self.column_name}' has been deleted.")
                
                # Signal to parent to refresh
                self.parent.event_generate("<<RefreshSheets>>")
                
                # Close the window
                self.window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete column: {str(e)}")
    
    