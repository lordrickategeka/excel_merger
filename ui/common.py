import tkinter as tk
from tkinter import ttk

def center_window(window, width, height):
    """
    Center a window on the screen
    
    Args:
        window: Window to center
        width: Window width
        height: Window height
    """
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f"{width}x{height}+{int(x)}+{int(y)}")

def create_header(parent, title, description=None):
    """
    Create a header frame with title and optional description
    
    Args:
        parent: Parent widget
        title: Header title text
        description: Optional description text
        
    Returns:
        Header frame
    """
    header_frame = ttk.Frame(parent, padding="10")
    
    title_label = ttk.Label(
        header_frame, 
        text=title, 
        font=("Helvetica", 16)
    )
    title_label.pack(pady=10)
    
    if description:
        desc_label = ttk.Label(
            header_frame,
            text=description,
            wraplength=750,
            justify="center"
        )
        desc_label.pack(pady=5)
    
    return header_frame

def create_file_selector(parent, file_var, select_command):
    """
    Create a file selector frame
    
    Args:
        parent: Parent widget
        file_var: StringVar for file path
        select_command: Command to execute when Browse button is clicked
        
    Returns:
        File selector frame
    """
    file_frame = ttk.Frame(parent, padding="10")
    
    ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(file_frame, textvariable=file_var, width=70).grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(file_frame, text="Browse", command=select_command).grid(row=0, column=2, padx=5, pady=5)
    
    return file_frame

def create_sheet_selector(parent, sheet_var, values=None):
    """
    Create a sheet selector frame
    
    Args:
        parent: Parent widget
        sheet_var: StringVar for selected sheet
        values: Optional list of sheet names
        
    Returns:
        Tuple of (frame, combobox)
    """
    frame = ttk.Frame(parent, padding="5")
    
    ttk.Label(frame, text="Select Sheet:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    sheet_selector = ttk.Combobox(frame, textvariable=sheet_var, width=30, state="readonly")
    sheet_selector.grid(row=0, column=1, padx=5, pady=5)
    
    if values:
        sheet_selector['values'] = values
        if values:
            sheet_selector.current(0)
    
    return frame, sheet_selector

def create_status_bar(parent, status_var):
    """
    Create a status bar
    
    Args:
        parent: Parent widget
        status_var: StringVar for status text
        
    Returns:
        Status bar label
    """
    status_bar = ttk.Label(parent, textvariable=status_var, relief="sunken", anchor="w")
    status_bar.pack(side="bottom", fill="x")
    
    return status_bar

def create_columns_listbox(parent, height=10):
    """
    Create a columns listbox with scrollbar
    
    Args:
        parent: Parent widget
        height: Listbox height in lines
        
    Returns:
        Tuple of (frame, listbox)
    """
    frame = ttk.Frame(parent)
    
    listbox = tk.Listbox(frame, height=height, exportselection=0, selectmode=tk.MULTIPLE)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    listbox.configure(yscrollcommand=scrollbar.set)
    
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    return frame, listbox

def populate_columns_listbox(listbox, columns):
    """
    Populate a listbox with column names
    
    Args:
        listbox: Listbox widget
        columns: List of column names
    """
    # Clear existing items
    listbox.delete(0, tk.END)
    
    # Add columns
    for col in columns:
        listbox.insert(tk.END, col)

def create_button_frame(parent, padding="10"):
    """
    Create a button frame
    
    Args:
        parent: Parent widget
        padding: Frame padding
        
    Returns:
        Button frame
    """
    return ttk.Frame(parent, padding=padding)

def add_button(frame, text, command, side="left", padx=5):
    """
    Add a button to a frame
    
    Args:
        frame: Parent frame
        text: Button text
        command: Button command
        side: Pack side
        padx: Horizontal padding
        
    Returns:
        Button widget
    """
    button = ttk.Button(frame, text=text, command=command)
    button.pack(side=side, padx=padx)
    return button