import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import threading
import queue

class TextToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Text to Excel Converter")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        self.files_queue = queue.Queue()
        self.selected_files = []
        self.delimiter = tk.StringVar(value=",")
        self.output_dir = tk.StringVar()
        self.source_folder = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tab control
        tab_control = ttk.Notebook(main_frame)
        
        # Individual files tab
        files_tab = ttk.Frame(tab_control)
        tab_control.add(files_tab, text="Select Files")
        
        # Folder scanning tab
        folder_tab = ttk.Frame(tab_control)
        tab_control.add(folder_tab, text="Scan Folder")
        
        tab_control.pack(expand=True, fill=tk.BOTH)
        
        # ------------------------ Files Tab UI -------------------------
        # Files selection section
        files_frame = ttk.LabelFrame(files_tab, text="Select Files", padding="10")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        btn_select = ttk.Button(files_frame, text="Select Text Files", command=self.select_files)
        btn_select.pack(side=tk.TOP, anchor=tk.W, pady=5)
        
        # Files listbox with scrollbar
        list_frame = ttk.Frame(files_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.files_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.files_listbox.yview)
        
        # Convert button for files tab
        btn_convert_files = ttk.Button(files_frame, text="Convert Selected Files", command=self.start_conversion)
        btn_convert_files.pack(side=tk.BOTTOM, anchor=tk.E, pady=10)
        
        # ------------------------ Folder Tab UI -------------------------
        folder_frame = ttk.LabelFrame(folder_tab, text="Select Source Folder", padding="10")
        folder_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        folder_select_frame = ttk.Frame(folder_frame)
        folder_select_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(folder_select_frame, text="Source Folder:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_select_frame, textvariable=self.source_folder, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(folder_select_frame, text="Browse", command=self.select_source_folder).pack(side=tk.LEFT)
        
        # Recursive checkbox
        recursive_check = ttk.Checkbutton(folder_frame, text="Include subfolders (recursive search)", variable=self.recursive_var)
        recursive_check.pack(anchor=tk.W, pady=5)
        
        # Scan button
        btn_scan = ttk.Button(folder_frame, text="Scan for Text Files", command=self.scan_folder)
        btn_scan.pack(anchor=tk.W, pady=5)
        
        # Scanned files display
        scan_list_frame = ttk.LabelFrame(folder_frame, text="Found Text Files")
        scan_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scan_scrollbar = ttk.Scrollbar(scan_list_frame)
        scan_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.scan_listbox = tk.Listbox(scan_list_frame)
        self.scan_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scan_listbox.config(yscrollcommand=scan_scrollbar.set)
        scan_scrollbar.config(command=self.scan_listbox.yview)
        
        # Convert button for folder tab
        btn_convert_folder = ttk.Button(folder_frame, text="Convert Found Files", command=self.start_conversion)
        btn_convert_folder.pack(side=tk.BOTTOM, anchor=tk.E, pady=10)
        
        # ------------------------ Common UI Elements -------------------------
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        # Delimiter option
        delimiter_frame = ttk.Frame(options_frame)
        delimiter_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(delimiter_frame, text="Delimiter:").pack(side=tk.LEFT)
        delimiter_options = [",", ";", "\\t", "|", " "]
        delimiter_combo = ttk.Combobox(delimiter_frame, textvariable=self.delimiter, values=delimiter_options, width=5)
        delimiter_combo.pack(side=tk.LEFT, padx=5)
        
        # Output directory option (only for individual files mode)
        output_frame = ttk.Frame(options_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output Directory (leave empty to use same location as source):").pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Browse", command=self.select_output_dir).pack(side=tk.LEFT)
        
        # Progress section
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, expand=True)
        
        # Main Convert button (larger and more prominent)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.btn_convert_main = ttk.Button(
            btn_frame, 
            text="CONVERT ALL FILES", 
            command=self.start_conversion,
            style="Accent.TButton"
        )
        self.btn_convert_main.pack(side=tk.RIGHT, padx=5, pady=5, ipadx=10, ipady=5)
        
        # Create a custom style for the main button
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", font=("Arial", 11, "bold"))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(fill=tk.X, pady=5)
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Text Files",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if files:
            self.selected_files = list(files)
            self.update_files_listbox()
    
    def select_source_folder(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)
    
    def scan_folder(self):
        source_folder = self.source_folder.get()
        if not source_folder or not os.path.isdir(source_folder):
            messagebox.showwarning("Invalid Folder", "Please select a valid source folder.")
            return
        
        # Clear previous scan results
        self.scan_listbox.delete(0, tk.END)
        self.selected_files = []
        
        # Start scanning in a separate thread to keep UI responsive
        self.status_var.set("Scanning for text files...")
        scan_thread = threading.Thread(target=self._scan_folder_thread)
        scan_thread.daemon = True
        scan_thread.start()
    
    def _scan_folder_thread(self):
        source_folder = self.source_folder.get()
        recursive = self.recursive_var.get()
        
        found_files = []
        
        try:
            if recursive:
                # Walk through all subdirectories
                for root, _, files in os.walk(source_folder):
                    for file in files:
                        if file.lower().endswith('.txt'):
                            file_path = os.path.join(root, file)
                            found_files.append(file_path)
            else:
                # Only look in the top directory
                for file in os.listdir(source_folder):
                    if file.lower().endswith('.txt'):
                        file_path = os.path.join(source_folder, file)
                        if os.path.isfile(file_path):
                            found_files.append(file_path)
            
            # Update UI with found files
            self.root.after(0, lambda: self._update_scan_results(found_files))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Scan Error", f"Error scanning folder: {str(e)}"))
            self.root.after(0, lambda: self.status_var.set("Scan failed."))
    
    def _update_scan_results(self, found_files):
        self.selected_files = found_files
        
        # Update scan listbox
        self.scan_listbox.delete(0, tk.END)
        for file_path in found_files:
            # Show relative path from source folder
            rel_path = os.path.relpath(file_path, self.source_folder.get())
            self.scan_listbox.insert(tk.END, rel_path)
        
        # Update status
        self.status_var.set(f"Found {len(found_files)} text files.")
    
    def update_files_listbox(self):
        self.files_listbox.delete(0, tk.END)
        for file in self.selected_files:
            self.files_listbox.insert(tk.END, os.path.basename(file))
    
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir.set(directory)
    
    def start_conversion(self):
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select or scan for text files first.")
            return
        
        # Put files in queue
        self.files_queue = queue.Queue()  # Reset queue
        for file in self.selected_files:
            self.files_queue.put(file)
        
        # Disable the convert buttons during conversion
        self.btn_convert_main.config(state=tk.DISABLED)
        
        # Start conversion thread
        conversion_thread = threading.Thread(target=self.process_queue)
        conversion_thread.daemon = True
        conversion_thread.start()
    
    def process_queue(self):
        total_files = len(self.selected_files)
        processed_files = 0
        failed_files = 0
        
        while not self.files_queue.empty():
            try:
                file = self.files_queue.get_nowait()
                self.status_var.set(f"Converting: {os.path.basename(file)}")
                
                success = self.convert_file(file)
                if not success:
                    failed_files += 1
                
                processed_files += 1
                progress_percentage = (processed_files / total_files) * 100
                self.progress_var.set(progress_percentage)
                
                self.files_queue.task_done()
                
            except queue.Empty:
                break
            except Exception as e:
                self.status_var.set(f"Error: {str(e)}")
                failed_files += 1
                processed_files += 1
        
        # Conversion completed
        self.btn_convert_main.config(state=tk.NORMAL)
        
        if failed_files > 0:
            self.status_var.set(f"Conversion completed with issues. {processed_files - failed_files}/{total_files} files converted successfully. {failed_files} files failed.")
        else:
            self.status_var.set(f"Conversion completed successfully. {processed_files}/{total_files} files converted.")
    
    def convert_file(self, input_file):
        try:
            # Determine output file path
            output_dir = self.output_dir.get()
            if output_dir:
                # If specific output directory is provided
                # Preserve folder structure relative to source folder
                if self.source_folder.get() and input_file.startswith(self.source_folder.get()):
                    rel_path = os.path.relpath(os.path.dirname(input_file), self.source_folder.get())
                    output_subdir = os.path.join(output_dir, rel_path)
                    
                    # Create subdirectory if it doesn't exist
                    os.makedirs(output_subdir, exist_ok=True)
                    
                    # Create output file path
                    name_without_ext = os.path.splitext(os.path.basename(input_file))[0]
                    output_file = os.path.join(output_subdir, f"{name_without_ext}.xlsx")
                else:
                    # If not from source folder scanning, just use the base name
                    name_without_ext = os.path.splitext(os.path.basename(input_file))[0]
                    output_file = os.path.join(output_dir, f"{name_without_ext}.xlsx")
            else:
                # Use same directory as input file
                output_file = os.path.splitext(input_file)[0] + '.xlsx'
            
            # Read the text file
            delimiter = self.delimiter.get()
            
            # Handle special case for tab delimiter
            if delimiter == "\\t":
                delimiter = "\t"
            
            # Read file with pandas
            df = pd.read_csv(input_file, delimiter=delimiter)
            
            # Write to Excel
            df.to_excel(output_file, index=False)
            
            return True
        except Exception as e:
            error_msg = f"Error converting {os.path.basename(input_file)}:\n{str(e)}"
            # Use after to avoid multiple dialog boxes freezing the UI
            self.root.after(0, lambda: messagebox.showerror("Conversion Error", error_msg))
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = TextToExcelConverter(root)
    root.mainloop()