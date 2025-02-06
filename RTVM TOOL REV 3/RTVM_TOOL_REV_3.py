import sys
import subprocess

def install_and_import(package_name, import_name=None):
    import_name = import_name or package_name
    try:
        __import__(import_name)
    except ImportError:
        print(f"Package '{import_name}' not found. Attempting to install '{package_name}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        try:
            __import__(import_name)
        except ImportError:
            print(f"Failed to install and import '{package_name}'. Please install it manually.")
            sys.exit(1)

# List of (package_name_on_pip, import_name_in_code)
required_packages = [
    ("pandas", "pandas"),
    ("numpy", "numpy"),
    ("matplotlib", "matplotlib"),
    ("openpyxl", "openpyxl"),
    ("pyspellchecker", "spellchecker"),  # Note the difference here
    # Tkinter is part of the standard library but may need installation on some systems
]

# Attempt to install and import each package
for package_name, import_name in required_packages:
    install_and_import(package_name, import_name)




import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import re
import os
import pandas as pd
import numpy as np  # Import numpy to check for NaN values
import matplotlib
matplotlib.use('TkAgg')
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import Counter  # Add this import
from openpyxl import load_workbook
from spellchecker import SpellChecker

class PatternDialog(tk.Toplevel):
    def __init__(self, master, app, obj_identifier, di_number, current_row):
        super().__init__(master)
        self.title("Pattern Generator")
        self.app = app  # Reference to the RTVMApp instance
        self.current_row = current_row  # Store the current row index

        # Initialize variables
        self.obj_identifier = obj_identifier
        self.di_number = di_number
        self.deletions = []

        # Object Identifier
        self.obj_identifier_label = tk.Label(self, text="Object Identifier")
        self.obj_identifier_label.grid(row=0, column=0, sticky="e")
        self.obj_identifier_entry = tk.Entry(self)
        self.obj_identifier_entry.grid(row=0, column=1, columnspan=2, sticky="w")
        self.obj_identifier_entry.insert(0, self.obj_identifier)

        # CDRL Name
        self.cdrl_name_label = tk.Label(self, text="(4) CDRL File Name")
        self.cdrl_name_label.grid(row=1, column=0, sticky="e")
        self.cdrl_name_entry = tk.Entry(self)
        self.cdrl_name_entry.grid(row=1, column=1, columnspan=2, sticky="w")

        # Detailed Location - Page/Sheet
        self.page_sheet_label = tk.Label(self, text="(4) Page/Sheet")
        self.page_sheet_label.grid(row=2, column=0, sticky="e")
        self.page_sheet_option_var = tk.StringVar(self)
        self.page_sheet_option_var.set("Page")  # default value
        self.page_sheet_option_menu = ttk.Combobox(
            self, textvariable=self.page_sheet_option_var,
            values=["Page", "Sheet"], width=8)
        self.page_sheet_option_menu.grid(row=2, column=1, sticky="w")
        self.page_sheet_entry = tk.Entry(self)
        self.page_sheet_entry.grid(row=2, column=2, sticky="w")

        # Detailed Location - Plan View/Section
        self.plan_view_label = tk.Label(self, text="(4) Plan View/Section")
        self.plan_view_label.grid(row=3, column=0, sticky="e")
        self.plan_view_option_var = tk.StringVar(self)
        self.plan_view_option_var.set("Plan View")  # default value
        self.plan_view_option_menu = ttk.Combobox(
            self, textvariable=self.plan_view_option_var,
            values=["Plan View", "Section"], width=8)
        self.plan_view_option_menu.grid(row=3, column=1, sticky="w")
        self.plan_view_entry = tk.Entry(self)
        self.plan_view_entry.grid(row=3, column=2, sticky="w")

        # Contractor Assessed Status
        self.status_label = tk.Label(self, text="(5) Contractor Assessed Status")
        self.status_label.grid(row=4, column=0, sticky="e")
        self.status_var = tk.StringVar(self)
        self.status_var.set("")  # Default to blank
        self.status_dropdown = ttk.Combobox(
            self, textvariable=self.status_var, values=["SAT", "UNSAT"])
        self.status_dropdown.grid(row=4, column=1, columnspan=2, sticky="w")

        # DI Number
        self.di_number_label = tk.Label(self, text="DI Number")
        self.di_number_label.grid(row=5, column=0, sticky="e")
        self.di_number_entry = tk.Entry(self)
        self.di_number_entry.grid(row=5, column=1, columnspan=2, sticky="w")
        self.di_number_entry.insert(0, self.di_number)

        # Buttons
        self.button_frame = tk.Frame(self)
        self.button_frame.grid(row=6, column=0, columnspan=3)

        # Generate Button
        self.generate_button = tk.Button(
            self.button_frame, text="(6) Generate Pattern", command=self.generate_pattern)
        self.generate_button.grid(row=0, column=0, padx=5, pady=5)

        # Copy Button
        self.copy_button = tk.Button(
            self.button_frame, text="(7) Copy to Clipboard", command=self.copy_to_clipboard)
        self.copy_button.grid(row=0, column=1, padx=5, pady=5)

        # Save to Excel Button
        self.save_button = tk.Button(
            self.button_frame, text="Save Generated Pattern to Excel", command=self.save_to_excel)
        self.save_button.grid(row=0, column=2, padx=5, pady=5)

        # Reset Button
        self.reset_button = tk.Button(
            self.button_frame, text="Reset", command=self.reset_fields)
        self.reset_button.grid(row=0, column=3, padx=5, pady=5)

        # New Button to create the 180-Vessel Version
        self.create_180_button = tk.Button(
            self.button_frame, text="Also create a 180-Vessel Version", command=self.create_180_version)
        self.create_180_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5)

        # Generated Pattern
        self.output_label = tk.Label(self, text="Generated Pattern: For column G")
        self.output_label.grid(row=7, column=0, columnspan=3, sticky="w")
        self.output_text = tk.Text(self, height=6, width=70)
        self.output_text.grid(row=8, column=0, columnspan=3)

        # Configure grid weights for proper resizing
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)


    def toggle_history_tables(self):
        if self.show_history_var.get() == 1:
            # Show the history tables and labels
            self.comment_history_label.grid()
            self.comment_history_table.grid()
            self.gov_comment_history_label.grid()
            self.gov_comment_history_table.grid()
            # Adjust column weights
            self.root.grid_columnconfigure(3, weight=1)
            self.root.grid_columnconfigure(4, weight=1)
        else:
            # Hide the history tables and labels
            self.comment_history_label.grid_remove()
            self.comment_history_table.grid_remove()
            self.gov_comment_history_label.grid_remove()
            self.gov_comment_history_table.grid_remove()
            # Adjust column weights
            self.root.grid_columnconfigure(3, weight=0)
            self.root.grid_columnconfigure(4, weight=0)

    def generate_pattern(self):
        # Clear output_text
        self.output_text.delete("1.0", tk.END)
        patterns = []

        # Check if we can generate the pattern
        can_generate_pattern = True
        error_messages = []

        if not self.status_var.get():
            error_messages.append("Contractor Assessed Status dropdown is blank.")
            can_generate_pattern = False

        if not self.page_sheet_entry.get().strip():
            error_messages.append("Page/Sheet input box is blank.")
            can_generate_patterns = False

        if not self.plan_view_entry.get().strip():
            error_messages.append("Plan View/Section input box is blank.")
            can_generate_patterns = False

        if can_generate_pattern:
            # Generate the first pattern
            obj_identifier = self.obj_identifier_entry.get().upper()
            cdrl_name = self.cdrl_name_entry.get().upper()
            page_sheet = self.page_sheet_entry.get()
            page_sheet_type = self.page_sheet_option_var.get()
            plan_view = self.plan_view_entry.get()
            plan_view_type = self.plan_view_option_var.get()
            status = self.status_var.get()

            detailed_location = f"{cdrl_name}, {page_sheet_type} {page_sheet}, {plan_view_type} {plan_view}"

            # First pattern line
            self.pattern1 = f"{obj_identifier};{detailed_location};{status}"

            patterns.append(self.pattern1)
        else:
            if error_messages:
                messagebox.showerror("Error", "\n".join(error_messages))
                return  # Exit the method to avoid adding empty patterns

        # Output patterns
        self.output_text.insert(tk.END, "\n".join(patterns))


    def create_180_version(self):
        # Check if the first pattern has been generated
        if not hasattr(self, 'pattern1'):
            messagebox.showerror("Error", "Please generate the initial pattern first.")
            return

        # Generate the second pattern
        obj_identifier = self.obj_identifier_entry.get().upper()
        cdrl_name = self.cdrl_name_entry.get().upper()
        page_sheet = self.page_sheet_entry.get()
        page_sheet_type = self.page_sheet_option_var.get()
        plan_view = self.plan_view_entry.get()
        plan_view_type = self.plan_view_option_var.get()
        status = self.status_var.get()
        di_number = self.di_number_entry.get()

        detailed_location = f"{cdrl_name}, {page_sheet_type} {page_sheet}, {plan_view_type} {plan_view}"

        # Replace 160-WLIC with 180-WLR
        detailed_location_wlr = detailed_location.replace("160-WLIC", "180-WLR")
        self.pattern2 = f"ADD;{di_number};{detailed_location_wlr};{status}"

        # Append the second pattern to the output
        self.output_text.insert(tk.END, "\n" + self.pattern2)

    def copy_to_clipboard(self):
        # Clear the clipboard
        self.clipboard_clear()

        # Get the content from the output text box
        patterns = self.output_text.get("1.0", tk.END).strip()

        # Copy the content to the clipboard
        self.clipboard_append(patterns)

        # Update the root window to ensure the clipboard retains the copied content
        self.update()

        # Print confirmation for debugging
        print("Copied to clipboard: \n" + patterns)

    def reset_fields(self):
        self.cdrl_name_entry.delete(0, tk.END)
        self.page_sheet_entry.delete(0, tk.END)
        self.plan_view_entry.delete(0, tk.END)
        self.status_var.set("")  # Set to blank on reset
        self.output_text.delete("1.0", tk.END)
        self.deletions.clear()

    def save_to_excel(self):
        # Get the generated pattern
        patterns = self.output_text.get("1.0", tk.END).strip()
        if not patterns:
            messagebox.showerror("Error", "No pattern generated to save.")
            return

        # Get existing content from cell G (column index 6) of the current row
        existing_content = self.app.df.iloc[self.current_row, 6]
        if pd.isna(existing_content):
            existing_content = ""
        elif not isinstance(existing_content, str):
            existing_content = str(existing_content)

        # Append the new patterns to the existing content
        if existing_content.strip():
            new_content = existing_content.strip() + "\n" + patterns
        else:
            new_content = patterns

        # Update the Excel file directly using openpyxl
        from openpyxl import load_workbook

        try:
            # Load the workbook
            wb = load_workbook(self.app.excel_file_path)
            ws = wb.active  # You may need to select the correct sheet if there are multiple

            # Calculate the Excel row number (considering headers)
            excel_row = self.current_row + 2  # Assuming header is on the first row

            # Update the cell in column G (which is column index 7 in openpyxl)
            ws.cell(row=excel_row, column=7, value=new_content)

            # Save the workbook
            wb.save(self.app.excel_file_path)

            # Update the DataFrame in memory
            self.app.df.iloc[self.current_row, 6] = new_content

            messagebox.showinfo("Success", "Pattern saved to Excel file successfully.")
            # Update the Contractor Proposed Change Request Input table
            self.app.update_proposed_changes_table()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save to Excel file: {e}")

import threading
class RTVMApp:
    def __init__(self, root):
        self.root = root
        root.title("RTVM Pattern Generator")

        # Initialize variables
        self.current_row = 0
        self.df = None
        self.deletions = []
        self.excel_file_path = ""  # Store the path to the Excel file
        self.pattern_dialog = None  # Reference to PatternDialog instance
        self.selected_base_path = None  # Will store user-selected base directory path
        # Add this line to initialize self.current_comments
        self.current_comments = {}  # To store comments for the current row

        # Initialize unique statuses
        self.unique_object_statuses = set()
        self.unique_contractor_statuses = set()
        self.unique_government_statuses = set()

        # Initialize filter variables
        self.filtered_row_indices = []
        self.current_filtered_index = 0

        # Adjust the window size if needed
        root.geometry("1600x600")  # Width x Height

        # Up and Down Buttons at the top along with Upload button
        self.button_frame_top = tk.Frame(root)
        self.button_frame_top.grid(row=0, column=0, columnspan=10, pady=10, sticky="ew")

        # Up and Down Buttons
        self.up_button = tk.Button(
            self.button_frame_top, text="Up", command=lambda: self.navigate_cells('up'))
        self.up_button.grid(row=0, column=0, padx=5, pady=5)

        self.down_button = tk.Button(
            self.button_frame_top, text="Down", command=lambda: self.navigate_cells('down'))
        self.down_button.grid(row=0, column=1, padx=5, pady=5)


        # Management Button
        self.management_button = tk.Button(
            self.button_frame_top, text="Management", command=self.open_management_window)
        self.management_button.grid(row=0, column=10, padx=5, pady=5)
        # Add the Tools Button
        self.tools_button = tk.Button(
             self.button_frame_top, text="Tools", command=self.open_tools_menu)
        self.tools_button.grid(row=0, column=11, padx=5, pady=5)

        # Jump to Cell Entry
        self.jump_to_label = tk.Label(self.button_frame_top, text="Jump to Row:")
        self.jump_to_label.grid(row=0, column=2, padx=5, pady=5)

        self.jump_to_var = tk.StringVar()
        self.jump_to_entry = tk.Entry(
            self.button_frame_top, textvariable=self.jump_to_var, width=5)
        self.jump_to_entry.grid(row=0, column=3, padx=5, pady=5)

        self.go_button = tk.Button(
            self.button_frame_top, text="Go", command=self.jump_to_cell)
        self.go_button.grid(row=0, column=4, padx=5, pady=5)

        # Upload Excel File Button on the same row
        self.upload_button = tk.Button(
            self.button_frame_top, text="Upload Excel File", command=self.upload_excel_file)
        self.upload_button.grid(row=0, column=5, padx=10, pady=5)

        # Apply Filter Button
        self.apply_filter_button = tk.Button(
            self.button_frame_top, text="Apply Filter", command=self.open_filter_dialog)
        self.apply_filter_button.grid(row=0, column=6, padx=5, pady=5)

        # Clear Filter Button
        self.clear_filter_button = tk.Button(
            self.button_frame_top, text="Clear Filter", command=self.clear_filters)
        self.clear_filter_button.grid(row=0, column=7, padx=5, pady=5)

        # Create Pie Charts Button
        self.create_pie_charts_button = tk.Button(
            self.button_frame_top, text="Create Pie Charts", command=self.open_pie_chart_window)
        self.create_pie_charts_button.grid(row=0, column=8, padx=5, pady=5)

        # Show/Hide History Tables Checkbox
        self.show_history_var = tk.IntVar(value=0)  # 0 means unchecked (hide tables by default)
        self.show_history_checkbox = tk.Checkbutton(
            self.button_frame_top,
            text="Show History Tables",
            variable=self.show_history_var,
            command=self.toggle_history_tables
        )
        self.show_history_checkbox.grid(row=0, column=9, padx=5, pady=5)

        # Show/Hide Progress Bar Checkbox
        self.show_progress_var = tk.IntVar(value=0)  # 0 means unchecked (hide progress bar by default)
        self.show_progress_checkbox = tk.Checkbutton(
            self.button_frame_top,
            text="Show Progress Bar",
            variable=self.show_progress_var,
            command=self.toggle_progress_bar
        )
        self.show_progress_checkbox.grid(row=1, column=9, padx=5, pady=5)

        # Row Indicator
        self.row_indicator_var = tk.StringVar(value="Row: 0")
        self.row_indicator_label = tk.Label(
            root, textvariable=self.row_indicator_var)
        self.row_indicator_label.grid(row=1, column=1, sticky="w")

        # Specification Text Label and Text Box
        self.spec_text_label = tk.Label(
            root, text="Specification Text")
        self.spec_text_label.grid(row=2, column=1, columnspan=4, sticky="w")

        self.spec_text_box = tk.Text(
            root, height=4, width=70, bg="lightblue")
        self.spec_text_box.grid(row=3, column=1, columnspan=9, sticky="nsew")
        self.spec_text_box.config(state=tk.DISABLED)

        # Version and POC Note
        self.version_note = tk.Label(root, text=(
            "Program Version 4\n"
            "POC: Eriks@birdon.us"
        ), justify="left", anchor="w")
        self.version_note.grid(row=1, column=6, rowspan=6, sticky="nw")

        # New Label Above Progress Bar
        self.progress_label = tk.Label(
            root, text="Progress Bar")
        self.progress_label.grid(row=4, column=0, sticky="w")

        # Progress Bar Table
        self.progress_table = ttk.Treeview(
            root, columns=("Row Number",), show="headings", height=20)
        self.progress_table.grid(row=5, column=0, sticky="nsew")
        self.progress_table.heading("Row Number", text="Row Number")
        self.progress_table.column("Row Number", width=50, anchor="center")

        # Bind click event
        self.progress_table.bind("<ButtonRelease-1>", self.on_progress_bar_click)


        # Initially hide the progress bar table
        self.progress_label.grid_remove()
        self.progress_table.grid_remove()

        # DI Number Breakdown Label
        self.table_label = tk.Label(
            root, text="DI Number Breakdown")
        self.table_label.grid(row=4, column=1, sticky="w")

        # DI Number Breakdown Table
        self.table = ttk.Treeview(root, columns=(
            "VeriDoc Number", "DI Number", "CDRL Subtitle", "Object Status",
            "Contractor Assessed Status", "Government Assessed Status"),
            show="headings")
        self.table.grid(row=5, column=1, sticky="nsew")

        # Configure headings
        self.table.heading("VeriDoc Number", text="VeriDoc Number")
        self.table.heading("DI Number", text="DI Number")
        self.table.heading("CDRL Subtitle", text="CDRL Subtitle")
        self.table.heading("Object Status", text="Object Status")
        self.table.heading("Contractor Assessed Status", text="Contractor Assessed Status")
        self.table.heading("Government Assessed Status", text="Government Assessed Status")

        # Comment Section Label
        self.comment_section_label = tk.Label(
            root, text="Comment from Birdon to USCG")
        self.comment_section_label.grid(row=4, column=2, sticky="w")

        # Comment Table
        self.comment_table = ttk.Treeview(
            root, columns=("Comments",), show="headings")
        self.comment_table.grid(row=5, column=2, sticky="nsew")
        self.comment_table.heading("Comments", text="Comments")

        # Bind the right-click event to the comment table
        self.comment_table.bind("<Button-3>", self.show_comment_table_context_menu)

        # Remove double-click editing if desired
        self.comment_table.unbind('<Double-1>')

        # Proposed Changes Label
        self.proposed_changes_label = tk.Label(
            root, text="Contractor Proposed Change Request Input")
        self.proposed_changes_label.grid(row=4, column=3, sticky="w")

        # Proposed Changes Table
        self.proposed_changes_table = ttk.Treeview(
            root, columns=("Pattern",), show="headings")
        self.proposed_changes_table.grid(
            row=5, column=3, sticky="nsew")
        self.proposed_changes_table.heading("Pattern", text="Pattern")

        # Comment History Label
        self.comment_history_label = tk.Label(
            root, text="Contractor Proposed Change Comment History")
        self.comment_history_label.grid(row=4, column=4, sticky="w")

        # Comment History Table
        self.comment_history_table = ttk.Treeview(
            root, columns=("History",), show="headings")
        self.comment_history_table.grid(row=5, column=4, sticky="nsew")
        self.comment_history_table.heading("History", text="History")

        # Government Comment History Label
        self.gov_comment_history_label = tk.Label(
            root, text="Government Adjudication Comment History")
        self.gov_comment_history_label.grid(row=4, column=5, sticky="w")

        # Government Comment History Table
        self.gov_comment_history_table = ttk.Treeview(
            root, columns=("History",), show="headings")
        self.gov_comment_history_table.grid(row=5, column=5, sticky="nsew")
        self.gov_comment_history_table.heading("History", text="History")

        # Configure column widths (Optional)
        self.table.column("VeriDoc Number", width=120)
        self.table.column("DI Number", width=100)
        self.table.column("CDRL Subtitle", width=150)
        self.table.column("Object Status", width=120)
        self.table.column("Contractor Assessed Status", width=180)
        self.table.column("Government Assessed Status", width=180)

        self.comment_table.column("Comments", width=200)
        self.proposed_changes_table.column("Pattern", width=400)
        self.comment_history_table.column("History", width=400)
        self.gov_comment_history_table.column("History", width=400)
        self.progress_table.column("Row Number", width=50)

        # Configure styles for highlighting
        self.style = ttk.Style()
        self.style.configure("Treeview", rowheight=25)

        # Create custom styles for the cells
        self.style.configure("GovAgree.Cell", background="green", foreground="white")
        self.style.configure("GovDisagree.Cell", background="red", foreground="white")
        self.style.configure("sat_row", background="green", foreground="white")
        self.style.configure("unsat_row", background="red", foreground="white")
        self.style.configure("tbd_row", background="yellow", foreground="black")
        self.style.configure("highlight", background="lightblue")

        # Keep track of highlighted items
        self.highlighted_items = []

        # Bind events
        self.table.bind("<<TreeviewSelect>>", self.on_table_row_select)
        self.table.bind("<Button-3>", self.show_table_context_menu)
        self.proposed_changes_table.bind("<Button-3>", self.show_proposed_changes_context_menu)

        # Configure grid weights for proper resizing
        # Initially set weight=0 for progress bar column
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=1)
        self.root.grid_columnconfigure(4, weight=1)
        self.root.grid_columnconfigure(5, weight=1)
        self.root.grid_rowconfigure(5, weight=1)

        # Initialize the visibility of history tables
        self.toggle_history_tables()

        # Initialize the visibility of progress bar
        self.toggle_progress_bar()

        #this prevents save functions from running multiple times
        self.save_lock = threading.Lock()  # Prevent simultaneous saves


    #this will allow the saving process to happen in its own thread in the background. 
    def save_comments_to_excel_background(self):
        def save_worker():
            try:
                with self.save_lock:
                    self.save_comments_to_excel()
            except Exception as e:
                # Optionally, if you need to show an error, schedule it on the main thread:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Background save error: {e}"))
        threading.Thread(target=save_worker, daemon=True).start()



    def open_tools_menu(self):
        # Create a new top-level window for tool selection
        self.tools_menu_window = tk.Toplevel(self.root)
        self.tools_menu_window.title("Select a Tool")
        self.tools_menu_window.attributes("-topmost", True)
        self.tools_menu_window.transient(self.root)
        self.tools_menu_window.lift()

        # Label
        tk.Label(self.tools_menu_window, text="Select a Tool:").pack(pady=10)

        # For future scalability, let's store tools in a dictionary
        # Key: Tool name (string), Value: function to open that tool
        self.available_tools = {
            "Compaire Tool": self.open_comair_tool_window,
            "Remove Previously Submitted Requests": self.open_remove_requests_tool_window,
            "RTVM Subset Management": self.open_rvtm_subset_management_window,
            "Disagreement Manager": self.open_disagreement_manager
        }
            # Later on, you can add more tools here:
            # "Another Tool": self.open_another_tool_window

        # Create a StringVar and Combobox for tool selection
        self.selected_tool = tk.StringVar(value="Compaire Tool")
        tool_list = list(self.available_tools.keys())
        self.tool_dropdown = ttk.Combobox(
            self.tools_menu_window, textvariable=self.selected_tool, values=tool_list, state="readonly")
        self.tool_dropdown.pack(pady=5)

        # Button to open the selected tool
        open_button = tk.Button(self.tools_menu_window, text="Open Selected Tool", command=self.open_selected_tool)
        open_button.pack(pady=10)

    def open_selected_tool(self):
        # Get the currently selected tool name
        tool_name = self.selected_tool.get()
        if tool_name in self.available_tools:
            # Call the corresponding function to open the tool
            self.tools_menu_window.destroy()  # Close the menu window
            self.available_tools[tool_name]()
        else:
            messagebox.showerror("Error", "Selected tool is not available.")





### START OF DIAGREEMENT TOOL 
    def open_disagreement_manager(self):
        if self.df is None:
            messagebox.showerror("Error", "No main Excel file is loaded. Please upload a main file first.")
            return

        # Filter rows
        # We need Government Assessed Status = "Disagree" and Object Status = "Accepted"
        # We'll use self.status_data which should have these fields
        # self.status_data is a list of dicts with keys like 'row_index', 'object_status', 'government_status', etc.
        filtered = [item for item in self.status_data 
                    if item['government_status'].strip().lower() == 'disagree' and 
                       item['object_status'].strip().lower() == 'accepted']

        if not filtered:
            messagebox.showinfo("No Disagreements", "No rows found with Government Assessed Status = Disagree and Object Status = Accepted.")
            return

        # Create the disagreement manager window
        self.disagreement_window = tk.Toplevel(self.root)
        self.disagreement_window.title("Disagreement Manager")

        # Create an instance of DisagreementManager class
        self.disagreement_manager = DisagreementManager(self.disagreement_window, self, filtered)


### END OF DIAGREEMENT TOOL             






     #### START OF REMOVAL TOOL
    def open_remove_requests_tool_window(self):
        # Create a new window
        self.remove_requests_window = tk.Toplevel(self.root)
        self.remove_requests_window.title("Remove Previously Submitted Contractor Proposed Change Request")

        # Explanation Label at the top
        explanation = ("This tool compares the currently loaded file with an older submission.\n"
                       "Any previously submitted contractor proposed change requests found in the old file\n"
                       "will be removed from the new file upon clicking 'Remove Previously Submitted Requests'.")
        tk.Label(self.remove_requests_window, text=explanation, justify="left").pack(pady=10, padx=10)

        # Frame for displaying current file and old file info
        files_frame = tk.Frame(self.remove_requests_window)
        files_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # Label for the currently loaded new file
        tk.Label(files_frame, text="Current (New) File:").grid(row=0, column=0, sticky="w")
        self.new_file_label_var = tk.StringVar(value=self.excel_file_path if self.excel_file_path else "No File Loaded")
        tk.Label(files_frame, textvariable=self.new_file_label_var).grid(row=0, column=1, sticky="w")

        # Label and button for uploading old file
        tk.Label(files_frame, text="Old File:").grid(row=1, column=0, sticky="w")
        self.old_file_label_var = tk.StringVar(value="No Old File Selected")
        tk.Label(files_frame, textvariable=self.old_file_label_var).grid(row=1, column=1, sticky="w")
        self.old_file_button = tk.Button(files_frame, text="Browse Old File", command=self.browse_old_file)
        self.old_file_button.grid(row=1, column=2, padx=5, sticky="w")

        # Button to remove previously submitted requests
        self.remove_requests_button = tk.Button(self.remove_requests_window, text="Remove Previously Submitted Requests",
                                                command=self.remove_previously_submitted_requests)
        self.remove_requests_button.pack(pady=20)

        # Create a frame for the progress bar
        self.progress_frame = tk.Frame(self.remove_requests_window)
        self.progress_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        self.progress_frame.grid_remove()  # Hide initially

        tk.Label(self.progress_frame, text="Processing...").pack(side=tk.LEFT, padx=5)
        self.removal_progress = ttk.Progressbar(self.progress_frame, orient='horizontal', length=300, mode='determinate')
        self.removal_progress.pack(side=tk.LEFT, padx=5)



    def browse_old_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            self.old_file_path = file_path
            self.old_file_label_var.set(file_path)

    def remove_previously_submitted_requests(self):
        # Check if main new file and old file are uploaded
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            messagebox.showerror("Error", "No main (new) file is loaded. Please upload a main file first.")
            return
        if not hasattr(self, 'old_file_path') or not self.old_file_path:
            messagebox.showerror("Error", "No old file is selected.")
            return

        try:
            old_df = pd.read_excel(self.old_file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read old Excel file: {e}")
            return

        # Identify columns in the current DF
        # Proposed Changes Column (G) is known to be at index 6
        proposed_changes_col = 6

        # Identify the "Contractor Proposed Change Comment Input" column by name
        comment_col_name = "Contractor Proposed Change Comment Input"
        if comment_col_name not in self.df.columns:
            messagebox.showerror("Error", f"Column '{comment_col_name}' not found in new file.")
            return
        comment_col = self.df.columns.get_loc(comment_col_name)

        # Validate column indexes in the old file
        if proposed_changes_col >= len(self.df.columns):
            messagebox.showerror("Error", "Proposed Changes column not found in new file.")
            return
        if proposed_changes_col >= len(old_df.columns):
            messagebox.showerror("Error", "Proposed Changes column not found in old file.")
            return

        if comment_col >= len(self.df.columns):
            messagebox.showerror("Error", f"'{comment_col_name}' column not found in new file.")
            return
        if comment_col >= len(old_df.columns):
            messagebox.showerror("Error", f"'{comment_col_name}' column not found in old file.")
            return

        # Convert object ID columns to string (Object ID is assumed to be column 0)
        object_id_col = 0
        self.df.iloc[:, object_id_col] = self.df.iloc[:, object_id_col].astype(str)
        old_df.iloc[:, object_id_col] = old_df.iloc[:, object_id_col].astype(str)

        # Create a dictionary from old_df for old patterns
        old_patterns_map = {}
        for idx, row in old_df.iterrows():
            obj_id = str(row.iloc[object_id_col])
            old_patterns_str = row.iloc[proposed_changes_col]
            if pd.isna(old_patterns_str):
                old_patterns_str = ""
            old_lines = [line.strip() for line in old_patterns_str.split('\n') if line.strip()]
            old_patterns_set = set(old_lines)
            old_patterns_map[obj_id] = {
                'patterns': old_patterns_set
            }

        # Also create a dictionary for old comments (from the old file)
        old_comments_map = {}
        for idx, row in old_df.iterrows():
            obj_id = str(row.iloc[object_id_col])
            old_comments_str = row.iloc[comment_col]
            if pd.isna(old_comments_str):
                old_comments_str = ""
            old_comment_lines = [line.strip() for line in old_comments_str.split('\n') if line.strip()]
            old_comments_set = set(old_comment_lines)
            if obj_id not in old_patterns_map:
                old_patterns_map[obj_id] = {'patterns': set()}
            # Add comments to the same map for convenience
            old_patterns_map[obj_id]['comments'] = old_comments_set

        changes_made = False

        # Remove previously submitted requests from Proposed Changes (column G)
        for idx, row in self.df.iterrows():
            obj_id = str(row.iloc[object_id_col])
            new_patterns_str = row.iloc[proposed_changes_col]
            if pd.isna(new_patterns_str):
                new_patterns_str = ""
            new_lines = [line.strip() for line in new_patterns_str.split('\n') if line.strip()]

            if obj_id in old_patterns_map:
                old_patterns_set = old_patterns_map[obj_id]['patterns']
                # Filter out any patterns that exist in old file
                filtered_lines = [line for line in new_lines if line not in old_patterns_set]
                if len(filtered_lines) != len(new_lines):
                    changes_made = True
                    updated_str = '\n'.join(filtered_lines)
                    self.df.iat[idx, proposed_changes_col] = updated_str

        # Remove previously submitted comments from "Contractor Proposed Change Comment Input" column
        for idx, row in self.df.iterrows():
            obj_id = str(row.iloc[object_id_col])
            new_comments_str = row.iloc[comment_col]
            if pd.isna(new_comments_str):
                new_comments_str = ""
            new_comment_lines = [line.strip() for line in new_comments_str.split('\n') if line.strip()]

            if obj_id in old_patterns_map and 'comments' in old_patterns_map[obj_id]:
                old_comments_set = old_patterns_map[obj_id]['comments']
                filtered_comment_lines = [line for line in new_comment_lines if line not in old_comments_set]
                if len(filtered_comment_lines) != len(new_comment_lines):
                    changes_made = True
                    updated_comments_str = '\n'.join(filtered_comment_lines)
                    self.df.iat[idx, comment_col] = updated_comments_str

        if not changes_made:
            messagebox.showinfo("No Changes", "No previously submitted requests or comments were found to remove.")
            return

        # If changes were made, update the Excel file preserving formatting
        from openpyxl import load_workbook

        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active  # Adjust if you need a specific sheet

            # Find the column indexes in the Excel sheet
            # Column G (Proposed Changes) = 7 in 1-based indexing
            # For 'Contractor Proposed Change Comment Input', find the correct column by header
            header_row = 1
            comment_col_letter = None
            for col in range(1, ws.max_column + 1):
                header_val = ws.cell(row=header_row, column=col).value
                if header_val == comment_col_name:
                    comment_col_letter = col
                    break

            # Update only the changed cells
            for idx, row in self.df.iterrows():
                new_content = row.iloc[proposed_changes_col]
                if pd.isna(new_content):
                    new_content = ""
                ws.cell(row=idx+2, column=7, value=new_content)  # Proposed Changes column (G)

                new_comment_content = row.iloc[comment_col]
                if pd.isna(new_comment_content):
                    new_comment_content = ""
                if comment_col_letter is not None:
                    ws.cell(row=idx+2, column=comment_col_letter, value=new_comment_content)

            wb.save(self.excel_file_path)

            # DataFrame in memory is already updated
            messagebox.showinfo("Success", "Previously submitted requests and comments have been removed and the file has been updated.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save updated file: {e}")


### END of removal Tool

### START OF COMPAIRE TOOL 
    def open_comair_tool_window(self):
        # This is your original `open_tools_window()` code
        # Renamed to `open_comair_tool_window()` for clarity.

        self.comair_window = tk.Toplevel(self.root)
        self.comair_window.title("Comair Tool")

        # Set window size (optional)
        self.comair_window.geometry("800x600")  # Adjust as needed

        # The rest of the original open_tools_window code goes here...
        # For example, your file comparison UI and logic:
        file_frame = tk.Frame(self.comair_window)
        file_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.file1_label = tk.Label(file_frame, text="Select First Excel File:")
        self.file1_label.grid(row=0, column=0, sticky="w")
        self.file1_entry = tk.Entry(file_frame, width=50)
        self.file1_entry.grid(row=0, column=1, padx=5)
        self.file1_button = tk.Button(file_frame, text="Browse", command=self.browse_file1)
        self.file1_button.grid(row=0, column=2, padx=5)

        self.file2_label = tk.Label(file_frame, text="Select Second Excel File:")
        self.file2_label.grid(row=1, column=0, sticky="w")
        self.file2_entry = tk.Entry(file_frame, width=50)
        self.file2_entry.grid(row=1, column=1, padx=5)
        self.file2_button = tk.Button(file_frame, text="Browse", command=self.browse_file2)
        self.file2_button.grid(row=1, column=2, padx=5)

        self.compare_button = tk.Button(self.comair_window, text="Compare Files", command=self.compare_files)
        self.compare_button.pack(pady=10)

        self.save_results_button = tk.Button(self.comair_window, text="Save Results", command=self.save_comparison_results)
        self.save_results_button.pack(pady=5)

        self.results_frame = tk.Frame(self.comair_window)
        self.results_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("Row", "Column", "Value in File 1", "Value in File 2")
        self.results_table = ttk.Treeview(self.results_frame, columns=columns, show="headings")
        self.results_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for col in columns:
            self.results_table.heading(col, text=col)
            self.results_table.column(col, anchor='center', width=150)

        scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_table.yview)
        self.results_table.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)



    def save_comparison_results(self):
        # Get the data from the results table
        rows = self.results_table.get_children()
        if not rows:
            messagebox.showinfo("No Data", "No differences to save.")
            return

        # Prepare data for DataFrame
        data = []
        for row_id in rows:
            row = self.results_table.item(row_id)['values']
            data.append(row)

        df = pd.DataFrame(data, columns=['Row', 'Column', 'Value in File 1', 'Value in File 2'])

        # Ask user for a file location
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                 title="Save Comparison Results")
        if file_path:
            try:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Results saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save results: {e}")

    def get_differences(self, df1, df2):
        try:
            # Align the DataFrames
            df1, df2 = df1.align(df2, join='outer', axis=1)
            df1.fillna('', inplace=True)
            df2.fillna('', inplace=True)

            # Use pandas built-in comparison
            diff = df1.compare(df2, keep_equal=False)
            diff.reset_index(inplace=True)

            # Prepare a list to collect differences
            differences = []

            for index, row in diff.iterrows():
                row_num = row['index'] + 2  # Adjust for human-readable indexing (assuming header row is row 1)
                for col in df1.columns:
                    if col in diff.columns.get_level_values(0):
                        val1 = diff.at[index, (col, 'self')]
                        val2 = diff.at[index, (col, 'other')]
                        differences.append({
                            'Row': row_num,
                            'Column': col,
                            'Value in File 1': val1,
                            'Value in File 2': val2
                        })

            # Create the DataFrame from the list of differences
            diff_df = pd.DataFrame(differences, columns=['Row', 'Column', 'Value in File 1', 'Value in File 2'])
            return diff_df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare Excel files: {e}")
            return pd.DataFrame()


    def browse_file1(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, file_path)

    def browse_file2(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, file_path)


    def compare_files(self):
        file1_path = self.file1_entry.get()
        file2_path = self.file2_entry.get()

        if not file1_path or not file2_path:
            messagebox.showerror("Error", "Please select both Excel files to compare.")
            return

        try:
            df1 = pd.read_excel(file1_path)
            df2 = pd.read_excel(file2_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel files: {e}")
            return

        # Clear previous results
        self.results_table.delete(*self.results_table.get_children())

        # Compare the dataframes
        try:
            diff_df = self.get_differences(df1, df2)
            if diff_df.empty:
                messagebox.showinfo("No Differences", "The two files are identical.")
            else:
                # Display differences in the results table
                for index, row in diff_df.iterrows():
                    self.results_table.insert("", "end", values=(row['Row'], row['Column'], row['Value in File 1'], row['Value in File 2']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare Excel files: {e}")




### End of compair tool ################################################################################################################################################################################################################################

### Start of RTVM Subsets tool pack ################################################################################################################################################################################################################################

    def open_rvtm_subset_management_window(self):
            self.rvtm_subset_window = tk.Toplevel(self.root)
            self.rvtm_subset_window.title("RTVM Subset Management")

            explanation = (
                "This RTVM Subset Management tool allows you to:\n"
                "- Create a summary report of verification data\n"
                "- Export photos of the generated report\n"
                "- Create subsets of the RTVM based on predefined SWBS groups\n"
                "- Recombine these subsets back into a single file\n"
                "- Merge a single subset into the main RTVM file\n\n"
                "The data is taken from the currently loaded main RTVM file.\n"
                "Please select a base location first."
            )
            tk.Label(self.rvtm_subset_window, text=explanation, justify="left").pack(pady=10, padx=10)

            top_frame = tk.Frame(self.rvtm_subset_window)
            top_frame.pack(fill='x', padx=10, pady=10)

            # Button to select the base location
            select_location_button = tk.Button(
                top_frame, text="Select Base Location", command=self.select_base_location
            )
            select_location_button.grid(row=0, column=0, padx=5, pady=5)

            # Display the currently loaded file name
            tk.Label(top_frame, text="Current (New) File:", anchor='w').grid(row=0, column=1, sticky='w')
            self.file_name_var = tk.StringVar(value=self.excel_file_path if self.excel_file_path else "No File Loaded")
            tk.Label(top_frame, textvariable=self.file_name_var, width=50, anchor='w').grid(row=0, column=2, padx=5, pady=5, sticky='ew')

            # Create Summary Report Button
            self.create_report_button = tk.Button(
                top_frame, text="Create Summary Report", command=self.create_summary_report
            )
            self.create_report_button.grid(row=0, column=3, padx=5, pady=5)

            # Export Photos Button
            self.export_photos_button = tk.Button(
                top_frame, text="Export Photos of Report", command=self.export_photos_of_report
            )
            self.export_photos_button.grid(row=0, column=4, padx=5, pady=5)

            # Create Subsets Button
            self.create_subsets_button = tk.Button(
                top_frame, text="Create Subsets", command=self.create_subsets
            )
            self.create_subsets_button.grid(row=0, column=5, padx=5, pady=5)

            # Recombine Subsets Button
            self.recombine_subsets_button = tk.Button(
                top_frame, text="Recombine Subsets", command=self.recombine_subsets
            )
            self.recombine_subsets_button.grid(row=0, column=6, padx=5, pady=5)

            # Merge Single Subset Button
            self.merge_single_subset_button = tk.Button(
                top_frame, text="Merge Single Subset", command=self.merge_single_subset
            )
            self.merge_single_subset_button.grid(row=0, column=7, padx=5, pady=5)

            self.set_assigned_verification_cells()


    def select_base_location(self):
            # Allow the user to select a folder where PMR-related files and subsets will be saved
            selected_directory = filedialog.askdirectory(title="Select Base Directory for PMR Files")
            if selected_directory:
                self.selected_base_path = selected_directory
                messagebox.showinfo("Base Location Selected", f"Base location set to: {self.selected_base_path}")
            else:
                messagebox.showwarning("No Selection", "No base location selected. Using current directory if needed.")

    def set_assigned_verification_cells(self):
        # Check if self.df is loaded
        if self.df is None:
            messagebox.showerror("Error", "No main file loaded. Please upload a main file before using this tool.")
            return

        assigned_column = "Assigned Verification Documents"
        if assigned_column in self.df.columns:
            self.assigned_verification_cells = self.df[assigned_column].tolist()
        else:
            messagebox.showerror("Error", f"Column '{assigned_column}' not found in the loaded DataFrame.")
            self.assigned_verification_cells = []


    # Below are the methods from the provided code snippet, adapted to use self.df, self.assigned_verification_cells,
    # self.excel_file_path, and the currently loaded main file. The upload functionality and separate class have been removed.

    def create_summary_report(self):
        # Collect data from all Assigned Verification Documents
        data_list = []
        for cell_content in self.assigned_verification_cells:
            # Convert cell_content to string to avoid float error
            if pd.isna(cell_content):
                continue
            cell_str = str(cell_content)
            if not cell_str or cell_str.strip().lower() == "nan":
                continue
            # Split entries separated by lines of underscores
            entries = cell_str.split('______________________')
            for entry in entries:
                entry = entry.strip()
                if not entry:
                    continue
                # Initialize variables
                data = {
                    'Object Identifier': "",
                    'DI Number': "",
                    'Government Assessed Status': "",
                    'Contractor Assessed Status': ""
                }
                # Split entry into lines
                lines = entry.split('\n')
                for line in lines:
                    line = line.strip()
                    if ':' in line:
                        key, value = line.split(':', 1)
                        key = key.strip()
                        value = value.strip()
                        if key in data:
                            data[key] = value
                data_list.append(data)

        # Now, data_list contains all the entries
        if not data_list:
            messagebox.showinfo("No Data", "No Assigned Verification Documents data found.")
            return

        # Convert to a DataFrame
        df_data = pd.DataFrame(data_list)

        # Now, group DI Numbers into SWBS groups as provided
        swbs_groups = {
            'SWBS 000': [
                '040-001', '042-001', '042-003', '042-005', '045-001',
                '068-001', '068-002', '068-003', '070-001', '073-001',
                '073-003', '073-006', '073-007', '073-008', '073-009',
                '076-002', '077-001', '077-002', '083-002', '085-004',
                '086-003', '088-001', '088-002', '088-005', '088-007',
                '092-001', '096-004'
            ],
            'SWBS 100': [
                '100-001', '100-002', '100-004', '100-006', '100-010',
                '100-011', '100-012', '100-013'
            ],
            'SWBS 200': [
                '200-001', '200-003', '233-001', '245-001', '245-002',
                '245-003', '249-001', '249-002', '249-003', '249-004',
                '259-001'
            ],
            'SWBS 202': [
                '202-012'
            ],
            'SWBS 300': [
                '300-001', '300-002', '300-003', '300-006', '300-007',
                '300-008', '300-009', '300-010', '300-011', '302-001',
                '310-001', '320-003', '303-001'
            ],
            'SWBS 400': [
                '400-001', '400-002', '400-003', '400-010', '400-011',
                '402-001', '402-002', '405-001', '407-001', '428-001',
                '432-001', '432-002', '435-001', '436-002', '440-001'
            ],
            'SWBS 500': [
                '508-001', '555-001', '580-001', '580-004', '583-001',
                '589-002', '593-002', '593-005','521-003'
            ],
            'SWBS 600': [
                '602-001', '604-001', '634-001', '640-002'
            ]
        }

        # Add DI Numbers that start with specific numbers
        for swbs in swbs_groups:
            if swbs == 'SWBS 000':
                starts_with = '0'
            elif swbs == 'SWBS 100':
                starts_with = '1'
            elif swbs == 'SWBS 200':
                starts_with = '2'
            elif swbs == 'SWBS 300':
                starts_with = '3'
            elif swbs == 'SWBS 400':
                starts_with = '4'
            elif swbs == 'SWBS 500':
                starts_with = '5'
            elif swbs == 'SWBS 600':
                starts_with = '6'
            else:
                starts_with = None

            if starts_with:
                # Get DI Numbers that start with the specified number
                di_numbers = df_data['DI Number'].unique()
                additional_di_numbers = [di for di in di_numbers if di.startswith(starts_with)]
                # Add them to the group if not already present
                for di in additional_di_numbers:
                    if di not in swbs_groups[swbs]:
                        swbs_groups[swbs].append(di)

        # Now, create a new window to display the pie charts
        self.report_window = tk.Toplevel(self.root)
        self.report_window.title("Summary Report")

        # Create a notebook to organize the charts
        notebook = ttk.Notebook(self.report_window)
        notebook.pack(fill='both', expand=True)

        # Create tabs for overall data
        overall_tabs = ttk.Notebook(self.report_window)
        notebook.add(overall_tabs, text='Overall Summary')
        self.figures_data = []
        # Overall Government Assessed Status
        gov_status_counts = df_data['Government Assessed Status'].value_counts()

        # Define colors
        status_colors = {
            'Disagree': 'red',
            'Agree': 'green',
            'Pending Review': 'orange',
            'Awaiting Input': 'blue'
        }

        gov_colors = [status_colors.get(status, 'grey') for status in gov_status_counts.index]

        gov_frame = ttk.Frame(overall_tabs)
        overall_tabs.add(gov_frame, text='Government Assessed Status')
        fig1, ax1 = plt.subplots(figsize=(6, 6))
        wedges1, texts1, autotexts1 = ax1.pie(
            gov_status_counts, labels=gov_status_counts.index,
            autopct='%1.1f%%', startangle=90, colors=gov_colors
        )
        ax1.axis('equal')
        ax1.set_title('Overall Government Assessed Status')

        # Store data for exporting
        self.figures_data = []
        self.figures_data.append({
            'swbs': 'Overall Status',
            'chart_type': 'Government Assessed Status',
            'counts': gov_status_counts.values,
            'labels': gov_status_counts.index.tolist(),
            'colors': gov_colors,
            'table_data': gov_status_counts.reset_index().values.tolist(),
            'table_columns': ['Status', 'Count']
        })

        canvas1 = FigureCanvasTkAgg(fig1, master=gov_frame)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill='both', expand=True)

        # Overall Contractor Assessed Status
        contractor_status_counts = df_data['Contractor Assessed Status'].value_counts()

        # Define colors for Contractor Assessed Status
        contractor_status_colors = {
            'Satisfactory': 'green',
            'Unsatisfactory': 'red',
            'TBD': 'grey'
        }

        contractor_colors = [contractor_status_colors.get(status, 'grey') for status in contractor_status_counts.index]

        contractor_frame = ttk.Frame(overall_tabs)
        overall_tabs.add(contractor_frame, text='Contractor Assessed Status')
        fig2, ax2 = plt.subplots(figsize=(6, 6))
        wedges2, texts2, autotexts2 = ax2.pie(
            contractor_status_counts, labels=contractor_status_counts.index,
            autopct='%1.1f%%', startangle=90, colors=contractor_colors
        )
        ax2.axis('equal')
        ax2.set_title('Overall Contractor Assessed Status')

        self.figures_data.append({
            'swbs': 'Overall Status',
            'chart_type': 'Contractor Assessed Status',
            'counts': contractor_status_counts.values,
            'labels': contractor_status_counts.index.tolist(),
            'colors': contractor_colors,
            'table_data': contractor_status_counts.reset_index().values.tolist(),
            'table_columns': ['Status', 'Count']
        })

        canvas2 = FigureCanvasTkAgg(fig2, master=contractor_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill='both', expand=True)

        # Overall DI Number Distribution
        di_number_counts = df_data['DI Number'].value_counts()
        di_frame = ttk.Frame(overall_tabs)
        overall_tabs.add(di_frame, text='DI Number Distribution')
        fig3, ax3 = plt.subplots(figsize=(6, 6))
        ax3.pie(
            di_number_counts, labels=di_number_counts.index,
            autopct='%1.1f%%', startangle=90
        )
        ax3.axis('equal')
        ax3.set_title('Overall DI Number Distribution')

        self.figures_data.append({
            'swbs': 'Overall Status',
            'chart_type': 'DI Number Distribution',
            'counts': di_number_counts.values,
            'labels': di_number_counts.index.tolist(),
            'colors': None,  # No specific colors
            'table_data': di_number_counts.reset_index().values.tolist(),
            'table_columns': ['DI Number', 'Count']
        })

        canvas3 = FigureCanvasTkAgg(fig3, master=di_frame)
        canvas3.draw()
        canvas3.get_tk_widget().pack(fill='both', expand=True)

        # For each SWBS group, create a tab
        for swbs, di_numbers in swbs_groups.items():
             # Strip spaces from swbs to avoid mismatch
            swbs = swbs.strip()

            # Filter the data for this SWBS group
            df_swbs = df_data[df_data['DI Number'].isin(di_numbers)]
            if df_swbs.empty:
                continue

            swbs_notebook = ttk.Notebook(self.report_window)
            notebook.add(swbs_notebook, text=swbs)

            # Government Assessed Status for SWBS
            gov_status_counts_swbs = df_swbs['Government Assessed Status'].value_counts()
            gov_colors_swbs = [status_colors.get(status, 'grey') for status in gov_status_counts_swbs.index]

            gov_frame_swbs = ttk.Frame(swbs_notebook)
            swbs_notebook.add(gov_frame_swbs, text='Government Assessed Status')
            fig_swbs_gov, ax_swbs_gov = plt.subplots(figsize=(6, 6))
            wedges_swbs_gov, texts_swbs_gov, autotexts_swbs_gov = ax_swbs_gov.pie(
                gov_status_counts_swbs, labels=gov_status_counts_swbs.index,
                autopct='%1.1f%%', startangle=90, colors=gov_colors_swbs
            )
            ax_swbs_gov.axis('equal')
            ax_swbs_gov.set_title(f'{swbs} - Government Assessed Status')

            self.figures_data.append({
                'swbs': swbs.replace('SWBS ', ''),
                'chart_type': 'Government Assessed Status',
                'counts': gov_status_counts_swbs.values,
                'labels': gov_status_counts_swbs.index.tolist(),
                'colors': gov_colors_swbs,
                'table_data': gov_status_counts_swbs.reset_index().values.tolist(),
                'table_columns': ['Status', 'Count']
            })

            canvas_swbs_gov = FigureCanvasTkAgg(fig_swbs_gov, master=gov_frame_swbs)
            canvas_swbs_gov.draw()
            canvas_swbs_gov.get_tk_widget().pack(fill='both', expand=True)

            # Contractor Assessed Status for SWBS
            contractor_status_counts_swbs = df_swbs['Contractor Assessed Status'].value_counts()
            contractor_colors_swbs = [contractor_status_colors.get(status, 'grey') for status in contractor_status_counts_swbs.index]

            contractor_frame_swbs = ttk.Frame(swbs_notebook)
            swbs_notebook.add(contractor_frame_swbs, text='Contractor Assessed Status')
            fig_swbs_contractor, ax_swbs_contractor = plt.subplots(figsize=(6, 6))
            wedges_swbs_contractor, texts_swbs_contractor, autotexts_swbs_contractor = ax_swbs_contractor.pie(
                contractor_status_counts_swbs, labels=contractor_status_counts_swbs.index,
                autopct='%1.1f%%', startangle=90, colors=contractor_colors_swbs
            )
            ax_swbs_contractor.axis('equal')
            ax_swbs_contractor.set_title(f'{swbs} - Contractor Assessed Status')

            self.figures_data.append({
                'swbs': swbs.replace('SWBS ', ''),
                'chart_type': 'Contractor Assessed Status',
                'counts': contractor_status_counts_swbs.values,
                'labels': contractor_status_counts_swbs.index.tolist(),
                'colors': contractor_colors_swbs,
                'table_data': contractor_status_counts_swbs.reset_index().values.tolist(),
                'table_columns': ['Status', 'Count']
            })

            canvas_swbs_contractor = FigureCanvasTkAgg(fig_swbs_contractor, master=contractor_frame_swbs)
            canvas_swbs_contractor.draw()
            canvas_swbs_contractor.get_tk_widget().pack(fill='both', expand=True)

            # DI Number Distribution for SWBS
            di_number_counts_swbs = df_swbs['DI Number'].value_counts()
            di_frame_swbs = ttk.Frame(swbs_notebook)
            swbs_notebook.add(di_frame_swbs, text='DI Number Distribution')
            fig_swbs_di, ax_swbs_di = plt.subplots(figsize=(6, 6))
            ax_swbs_di.pie(
                di_number_counts_swbs, labels=di_number_counts_swbs.index,
                autopct='%1.1f%%', startangle=90
            )
            ax_swbs_di.axis('equal')
            ax_swbs_di.set_title(f'{swbs} - DI Number Distribution')

            self.figures_data.append({
                'swbs': swbs.replace('SWBS ', ''),
                'chart_type': 'DI Number Distribution',
                'counts': di_number_counts_swbs.values,
                'labels': di_number_counts_swbs.index.tolist(),
                'colors': None,  # No specific colors
                'table_data': di_number_counts_swbs.reset_index().values.tolist(),
                'table_columns': ['DI Number', 'Count']
            })

            canvas_swbs_di = FigureCanvasTkAgg(fig_swbs_di, master=di_frame_swbs)
            canvas_swbs_di.draw()
            canvas_swbs_di.get_tk_widget().pack(fill='both', expand=True)

            
    def export_photos_of_report(self):
        if not hasattr(self, 'figures_data') or not self.figures_data:
            messagebox.showwarning("No Report Generated", "Please create the summary report first.")
            return

        if not self.selected_base_path:
            messagebox.showerror("No Base Location", "Please select a base location before exporting photos.")
            return

        # Prompt for PMR number
        pmr_number = simpledialog.askinteger("PMR Number", "Enter the PMR number:")
        if pmr_number is None:
            return  # User canceled

        pmr_folder = os.path.join(self.selected_base_path, f"PMR {pmr_number}")

        # Create the PMR folder if it doesn't exist
        if not os.path.exists(pmr_folder):
            os.makedirs(pmr_folder)

        for data in self.figures_data:
            swbs = data['swbs']
            chart_type = data['chart_type']
            counts = data['counts']
            labels = data['labels']
            colors = data['colors']
            table_data = data['table_data']
            table_columns = data['table_columns']

            # Strip spaces from swbs to ensure exact match
            swbs = swbs.strip()

            if swbs == 'Overall Status':
                swbs_folder = os.path.join(pmr_folder, "Overall Status")
            else:
                swbs_folder = os.path.join(pmr_folder, f"{swbs} SWBS", "Status Photos")

            if not os.path.exists(swbs_folder):
                os.makedirs(swbs_folder)

            # Recreate the figure
            fig, ax = plt.subplots(figsize=(6, 6))
            if colors:
                ax.pie(
                    counts, labels=labels,
                    autopct='%1.1f%%', startangle=90, colors=colors
                )
            else:
                ax.pie(
                    counts, labels=labels,
                    autopct='%1.1f%%', startangle=90
                )
            ax.axis('equal')
            ax.set_title(f"{swbs} - {chart_type}")

            # Save the figure
            filename = f"PMR {pmr_number} - SWBS {swbs} - {chart_type}.png"
            save_path = os.path.join(swbs_folder, filename)
        
            # Ensure directories exist before saving
            if not os.path.exists(swbs_folder):
                os.makedirs(swbs_folder)

            fig.savefig(save_path, bbox_inches='tight')
            plt.close(fig)

            # Prepare data for Excel export
            if swbs == 'Overall Status':
                excel_filename = f"PMR {pmr_number} - Overall Status - Raw Data Export.xlsx"
                excel_save_path = os.path.join(swbs_folder, excel_filename)
            else:
                excel_filename = f"PMR {pmr_number} - SWBS {swbs} - Raw Data Export.xlsx"
                excel_save_path = os.path.join(swbs_folder, excel_filename)

            # Write the data to Excel
            table_df = pd.DataFrame(table_data, columns=table_columns)
            sheet_name = chart_type
            if not os.path.exists(excel_save_path):
                with pd.ExcelWriter(excel_save_path, engine='openpyxl') as writer:
                    table_df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(excel_save_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    table_df.to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Export Completed", f"All charts have been exported to {pmr_folder}")



    def create_subsets(self):
        import shutil
        from openpyxl import load_workbook
        from copy import copy
        from tkinter import ttk, messagebox, simpledialog, filedialog
        import threading
        import queue
        import os

        if self.selected_base_path is None:
            messagebox.showerror("No Base Location", "Please select a base location before creating subsets.")
            return

        if not self.excel_file_path:
            messagebox.showwarning("No File", "Please upload an Excel file first.")
            return

        # Prompt the user for the PMR number
        pmr_number = simpledialog.askinteger("PMR Number", "Enter the PMR number:")
        if pmr_number is None:
            return  # User canceled

        pmr_folder = os.path.join(self.selected_base_path, f"PMR {pmr_number}")

        # Create the PMR folder if it doesn't exist
        if not os.path.exists(pmr_folder):
            os.makedirs(pmr_folder)

        # Check if template_file_path is defined
        if not hasattr(self, 'template_file_path') or not self.template_file_path:
            # Ask the user if they want to upload a template file
            if messagebox.askyesno("No Template File", "No template file is currently selected.\nWould you like to select one now?"):
                template_file_path = filedialog.askopenfilename(
                    title="Select Template Excel File",
                    filetypes=[("Excel files", "*.xls;*.xlsx")]
                )
                if not template_file_path:
                    messagebox.showwarning("No Template File", "No template Excel file was selected. Cannot proceed.")
                    return
                self.template_file_path = template_file_path
            else:
                # User chose not to select a template file
                return

        template_file_path = self.template_file_path

        # Define the SWBS groups and DI Numbers (unchanged)
        swbs_groups = {
            'SWBS 000': [
                '040-001', '042-001', '042-003', '042-005', '045-001',
                '068-001', '068-002', '068-003', '070-001', '073-001',
                '073-003', '073-006', '073-007', '073-008', '073-009',
                '076-002', '077-001', '077-002', '083-002', '085-004',
                '086-003', '088-001', '088-002', '088-005', '088-007',
                '092-001', '096-004'
            ],
            'SWBS 100': [
                '100-001', '100-002', '100-004', '100-006', '100-010',
                '100-011', '100-012', '100-013'
            ],
            'SWBS 200': [
                '200-001', '200-003', '233-001', '245-001', '245-002',
                '245-003', '249-001', '249-002', '249-003', '249-004',
                '259-001'
            ],
            'SWBS 202': [
                '202-012'
            ],
            'SWBS 300': [
                '300-001', '300-002', '300-003', '300-006', '300-007',
                '300-008', '300-009', '300-010', '300-011', '302-001',
                '310-001', '320-003', '303-001'
            ],
            'SWBS 400': [
                '400-001', '400-002', '400-003', '400-010', '400-011',
                '402-001', '402-002', '405-001', '407-001', '428-001',
                '432-001', '432-002', '435-001', '436-002', '440-001'
            ],
            'SWBS 500': [
                '508-001', '555-001', '580-001', '580-004', '583-001',
                '589-002', '593-002', '593-005','521-003'
            ],
            'SWBS 600': [
                '602-001', '604-001', '634-001', '640-002'
            ]
        }

        def process_data(progress_queue):
            print("Starting recombination process...")
            # Create a set of all DI Numbers for quick lookup
            all_di_numbers = set()
            di_number_to_swbs = {}
            for swbs, di_numbers in swbs_groups.items():
                for di_number in di_numbers:
                    all_di_numbers.add(di_number)
                    di_number_to_swbs[di_number] = swbs

            # Copy the template file to create subset files
            di_number_to_file_path = {}
            for swbs, di_numbers in swbs_groups.items():
                swbs_number = swbs.replace('SWBS ', '')
                swbs_folder = os.path.join(pmr_folder, f"{swbs_number} SWBS")
                if not os.path.exists(swbs_folder):
                    os.makedirs(swbs_folder)

                for di_number in di_numbers:
                    filename = f"{di_number} - Revision X - RTVM Subset.xlsx"
                    dest_file_path = os.path.join(swbs_folder, filename)

                    try:
                        shutil.copyfile(template_file_path, dest_file_path)
                        di_number_to_file_path[di_number] = dest_file_path
                    except Exception as e:
                        progress_queue.put(('error', f"An error occurred while copying the template file:\n{e}"))
                        return

            # Now, process the main Excel file
            try:
                print("Starting recombination process..2.")
                wb_main = load_workbook(filename=self.excel_file_path)
            except Exception as e:
                progress_queue.put(('error', f"Failed to open the main Excel file:\n{e}"))
                return

            if 'RTVM' not in wb_main.sheetnames:
                progress_queue.put(('error', "The sheet 'RTVM' was not found in the Excel file."))
                return
            ws_main = wb_main['RTVM']

            total_rows = ws_main.max_row - 1
            progress_queue.put(('total', total_rows))

            di_number_to_wb = {}
            di_number_to_ws = {}
            di_number_to_next_row = {}

            for idx, row in enumerate(ws_main.iter_rows(min_row=2, values_only=False), start=1):
                progress_queue.put(('progress', idx))

                cell_f = row[5]
                cell_value = cell_f.value

                if cell_value:
                    entries = cell_value.split('______________________')
                    di_numbers_in_cell = set()
                    for entry in entries:
                        entry = entry.strip()
                        if not entry:
                            continue
                        lines = entry.split('\n')
                        for line in lines:
                            line = line.strip()
                            if line.startswith('DI Number:'):
                                di_number = line[len('DI Number:'):].strip()
                                if di_number:
                                    di_numbers_in_cell.add(di_number)

                    for di_number in di_numbers_in_cell:
                        if di_number in all_di_numbers:
                            if di_number not in di_number_to_wb:
                                subset_file_path = di_number_to_file_path[di_number]
                                wb_subset = load_workbook(filename=subset_file_path)
                                if 'RTVM' in wb_subset.sheetnames:
                                    ws_subset = wb_subset['RTVM']
                                else:
                                    ws_subset = wb_subset.active
                                di_number_to_wb[di_number] = wb_subset
                                di_number_to_ws[di_number] = ws_subset
                                di_number_to_next_row[di_number] = 2
                            else:
                                wb_subset = di_number_to_wb[di_number]
                                ws_subset = di_number_to_ws[di_number]

                            next_row = di_number_to_next_row[di_number]

                            for cell in row:
                                new_cell = ws_subset.cell(row=next_row, column=cell.column, value=cell.value)
                                if cell.has_style:
                                    new_cell.font = copy(cell.font)
                                    new_cell.border = copy(cell.border)
                                    new_cell.fill = copy(cell.fill)
                                    new_cell.number_format = copy(cell.number_format)
                                    new_cell.protection = copy(cell.protection)
                                    new_cell.alignment = copy(cell.alignment)
                                if cell.hyperlink:
                                    new_cell.hyperlink = copy(cell.hyperlink)
                                if cell.comment:
                                    new_cell.comment = copy(cell.comment)

                            di_number_to_next_row[di_number] += 1

            for di_number, wb_subset in di_number_to_wb.items():
                subset_file_path = di_number_to_file_path[di_number]
                try:
                    wb_subset.save(subset_file_path)
                    wb_subset.close()
                except Exception as e:
                    progress_queue.put(('error', f"Failed to save subset file for DI Number {di_number}:\n{e}"))

            progress_queue.put(('done', None))

        progress_queue = queue.Queue()
        threading.Thread(target=process_data, args=(progress_queue,)).start()

        progress_window = tk.Toplevel()
        progress_window.title("Processing Rows")
        progress_label = tk.Label(progress_window, text="Processing rows...")
        progress_label.pack()
        progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=300, mode='determinate')
        progress_bar.pack()
        progress_info = tk.Label(progress_window, text="0%")
        progress_info.pack()

        total_rows = 0
        current_row = 0

        def update_progress():
            nonlocal total_rows, current_row
            try:
                while True:
                    message_type, data = progress_queue.get_nowait()
                    if message_type == 'total':
                        total_rows = data
                        progress_bar['maximum'] = total_rows
                    elif message_type == 'progress':
                        current_row = data
                        if total_rows > 0:
                            percent = int((current_row / total_rows) * 100)
                            progress_bar['value'] = current_row
                            progress_info.config(text=f"{percent}%")
                    elif message_type == 'error':
                        messagebox.showerror("Error", data)
                        progress_window.destroy()
                        return
                    elif message_type == 'done':
                        progress_bar['value'] = progress_bar['maximum']
                        progress_info.config(text="100%")
                        progress_window.destroy()
                        messagebox.showinfo("Subsets Created", f"The subsets have been created in {pmr_folder}")
                        return
            except queue.Empty:
                pass
            except Exception as e:
                print(f"Error in update_progress: {e}")
            progress_window.after(100, update_progress)

        update_progress()



    def recombine_subsets(self):
        import os
        import time
        from openpyxl import load_workbook
        from copy import copy
        from tkinter import ttk, messagebox, filedialog, simpledialog
        import tkinter as tk
        import threading
        import queue

        # Since you've already uploaded the main Excel file and stored it in self.excel_file_path,
        # we assume self.excel_file_path is valid and no need to check again.
        # If you still want to ensure it's loaded, you can handle it gracefully:
        if not self.excel_file_path:
            messagebox.showerror("Error", "No main Excel file is loaded. Please upload a main file first.")
            return

        # Prompt the user to select the template file if not already selected
        if not hasattr(self, 'template_file_path') or not self.template_file_path:
            template_file_path = filedialog.askopenfilename(
                title="Select the Template Excel File",
                filetypes=[("Excel files", "*.xls;*.xlsx")]
            )
            if not template_file_path:
                messagebox.showwarning("No Template File", "No template Excel file was selected.")
                return
            self.template_file_path = template_file_path
        else:
            template_file_path = self.template_file_path

        # Prompt the user to select the folder containing the subset files
        subset_folder = filedialog.askdirectory(title="Select the PMR Folder Containing Subset Files")
        if not subset_folder:
            messagebox.showwarning("No Folder Selected", "No folder was selected.")
            return

        # Prompt the user to select the output file path for the recombined Excel file
        output_file_path = filedialog.asksaveasfilename(
            title="Save Recombined Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_file_path:
            messagebox.showwarning("No Output File", "No output file was specified.")
            return

        def process_data(progress_queue):
            try:
                print("Starting recombination process...")

                # Load the main Excel file to get the structure and headers
                wb_main = load_workbook(filename=self.excel_file_path)
                if 'RTVM' in wb_main.sheetnames:
                    ws_main = wb_main['RTVM']
                else:
                    ws_main = wb_main.active
                print("Main workbook loaded.")

                # Load the template workbook to use as the recombined workbook
                wb_recombined = load_workbook(filename=template_file_path)
                if 'RTVM' in wb_recombined.sheetnames:
                    ws_recombined = wb_recombined['RTVM']
                else:
                    ws_recombined = wb_recombined.active
                print("Template workbook loaded as recombined workbook.")

                # Clear existing data in the recombined worksheet except the header
                ws_recombined.delete_rows(2, ws_recombined.max_row)
                print("Cleared existing data in the recombined worksheet.")

                # Read all subset files and collect data
                data_dict = {}  # Key: DOORS SPEC ID, Value: Row data (as a dict)
                columns_to_combine = [7, 8]  # Columns G and H (1-based indexing)

                print("Collecting subset files...")
                subset_files = []
                for root, dirs, files in os.walk(subset_folder):
                    if 'Status Photos' in dirs:
                        dirs.remove('Status Photos')
                    for file in files:
                        if file.endswith('.xlsx') or file.endswith('.xls'):
                            file_path = os.path.join(root, file)
                            subset_files.append(file_path)
                print(f"Total subset files found: {len(subset_files)}")

                total_files = len(subset_files)
                progress_queue.put(('total', total_files))

                current_file = 0

                for subset_file in subset_files:
                    current_file += 1
                    progress_queue.put(('progress', current_file))
                    print(f"Processing file {current_file}/{total_files}: {subset_file}")

                    wb_subset = load_workbook(filename=subset_file)
                    if 'RTVM' in wb_subset.sheetnames:
                        ws_subset = wb_subset['RTVM']
                    else:
                        ws_subset = wb_subset.active
                    print(f"Opened subset workbook: {subset_file}")

                    for row in ws_subset.iter_rows(min_row=2, values_only=False):
                        doors_spec_id_cell = row[0]
                        doors_spec_id = doors_spec_id_cell.value
                        if not doors_spec_id:
                            continue

                        row_values = {}
                        for cell in row:
                            col_idx = cell.column
                            row_values[col_idx] = cell.value

                        if doors_spec_id not in data_dict:
                            data_dict[doors_spec_id] = row_values
                        else:
                            # Combine columns G and H
                            for col_idx in columns_to_combine:
                                existing_value = data_dict[doors_spec_id].get(col_idx, '')
                                new_value = row_values.get(col_idx, '')
                                if existing_value and new_value and new_value not in existing_value:
                                    combined_value = f"{existing_value}\n{new_value}"
                                    data_dict[doors_spec_id][col_idx] = combined_value
                                elif not existing_value and new_value:
                                    data_dict[doors_spec_id][col_idx] = new_value

                    wb_subset.close()
                    print(f"Finished processing file: {subset_file}")

                print("All subset files processed.")

                print("Preparing to write data to recombined workbook...")
                progress_queue.put(('writing_start', None))

                doors_spec_id_order = []
                for row in ws_main.iter_rows(min_row=2, values_only=False):
                    doors_spec_id = row[0].value
                    if doors_spec_id:
                        doors_spec_id_order.append(doors_spec_id)

                total_rows = len(doors_spec_id_order)
                progress_queue.put(('writing_total', total_rows))
                print(f"Total rows to write: {total_rows}")

                next_row = 2

                for idx, doors_spec_id in enumerate(doors_spec_id_order, start=1):
                    progress_queue.put(('writing_progress', idx))
                    if idx % 100 == 0 or idx == total_rows:
                        print(f"Writing row {idx}/{total_rows}")
                    if doors_spec_id in data_dict:
                        row_values = data_dict[doors_spec_id]
                        for cell in ws_main[next_row]:
                            col_idx = cell.column
                            value = row_values.get(col_idx, cell.value)
                            new_cell = ws_recombined.cell(row=next_row, column=col_idx, value=value)
                            if cell.has_style:
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.number_format = copy(cell.number_format)
                                new_cell.protection = copy(cell.protection)
                                new_cell.alignment = copy(cell.alignment)
                            if cell.hyperlink:
                                new_cell.hyperlink = copy(cell.hyperlink)
                            if cell.comment:
                                new_cell.comment = copy(cell.comment)
                    else:
                        for cell in ws_main[next_row]:
                            new_cell = ws_recombined.cell(row=next_row, column=cell.column, value=cell.value)
                            if cell.has_style:
                                new_cell.font = copy(cell.font)
                                new_cell.border = copy(cell.border)
                                new_cell.fill = copy(cell.fill)
                                new_cell.number_format = copy(cell.number_format)
                                new_cell.protection = copy(cell.protection)
                                new_cell.alignment = copy(cell.alignment)
                            if cell.hyperlink:
                                new_cell.hyperlink = copy(cell.hyperlink)
                            if cell.comment:
                                new_cell.comment = copy(cell.comment)
                    next_row += 1

                print("Data written to recombined workbook.")

                wb_recombined.save(output_file_path)
                wb_recombined.close()
                print(f"Recombined workbook saved at: {output_file_path}")

                progress_queue.put(('done', None))
                print("Recombination process completed.")

            except Exception as e:
                progress_queue.put(('error', f"An error occurred during recombination:\n{e}"))
                print(f"An error occurred during recombination: {e}")

        progress_queue = queue.Queue()
        threading.Thread(target=process_data, args=(progress_queue,)).start()

        progress_window = tk.Toplevel()
        progress_window.title("Recombining Subsets")
        progress_label = tk.Label(progress_window, text="Processing subset files...")
        progress_label.pack()
        progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=300, mode='determinate')
        progress_bar.pack()
        progress_info = tk.Label(progress_window, text="0%")
        progress_info.pack()

        total_files = 0
        current_file = 0
        total_rows = 0
        current_row = 0
        write_start_time = None

        def update_progress():
            nonlocal total_files, current_file, total_rows, current_row, write_start_time
            try:
                while True:
                    message_type, data = progress_queue.get_nowait()
                    if message_type == 'total':
                        total_files = data
                        progress_bar['maximum'] = total_files
                        progress_label.config(text="Processing subset files...")
                    elif message_type == 'progress':
                        current_file = data
                        if total_files > 0:
                            percent = int((current_file / total_files) * 100)
                            progress_bar['value'] = current_file
                            progress_info.config(text=f"{percent}%")
                    elif message_type == 'writing_start':
                        current_row = 0
                        total_rows = 0
                        progress_bar['value'] = 0
                        progress_bar['maximum'] = 1
                        progress_label.config(text="Writing data to recombined workbook...")
                        progress_info.config(text="0%")
                        write_start_time = time.time()
                    elif message_type == 'writing_total':
                        total_rows = data
                        progress_bar['maximum'] = total_rows
                    elif message_type == 'writing_progress':
                        current_row = data
                        if total_rows > 0:
                            percent = int((current_row / total_rows) * 100)
                            progress_bar['value'] = current_row
                            elapsed_time = time.time() - write_start_time
                            rows_per_sec = current_row / elapsed_time if elapsed_time > 0 else 0
                            remaining_rows = total_rows - current_row
                            est_remaining_time = remaining_rows / rows_per_sec if rows_per_sec > 0 else 0
                            mins, secs = divmod(int(est_remaining_time), 60)
                            hours, mins = divmod(mins, 60)
                            est_time_str = f"Est. time remaining: {hours:d}h {mins:02d}m {secs:02d}s"
                            progress_info.config(text=f"{percent}% - {est_time_str}")
                    elif message_type == 'error':
                        messagebox.showerror("Error", data)
                        progress_window.destroy()
                        return
                    elif message_type == 'done':
                        progress_bar['value'] = progress_bar['maximum']
                        progress_info.config(text="100%")
                        progress_window.destroy()
                        messagebox.showinfo(
                            "Recombination Complete",
                            f"The subset files have been recombined into {output_file_path}"
                        )
                        return
            except queue.Empty:
                pass
            except Exception as e:
                print(f"Error in update_progress: {e}")
            progress_window.after(100, update_progress)

        update_progress()





    def merge_single_subset(self):
        from openpyxl import load_workbook
        from copy import copy
        import tkinter as tk
        from tkinter import filedialog, messagebox
        import threading
        import queue

        # Prompt the user to select the subset file to merge
        subset_file_path = filedialog.askopenfilename(
            title="Select the Subset Excel File to Merge",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not subset_file_path:
            messagebox.showwarning("No Subset File", "No subset Excel file was selected.")
            return

        # Prompt the user to select the main file to merge into
        main_file_path = filedialog.askopenfilename(
            title="Select the Main Excel File to Merge Into",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not main_file_path:
            messagebox.showwarning("No Main File", "No main Excel file was selected.")
            return

        # Prompt the user to select the output file path for the merged Excel file
        output_file_path = filedialog.asksaveasfilename(
            title="Save Merged Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_file_path:
            messagebox.showwarning("No Output File", "No output file was specified.")
            return

        # Function to perform the merge in a background thread
        def process_data(progress_queue):
            try:
                # Load the main Excel file
                wb_main = load_workbook(filename=main_file_path)
                if 'RTVM' in wb_main.sheetnames:
                    ws_main = wb_main['RTVM']
                else:
                    ws_main = wb_main.active  # Use the first sheet
                print("Main workbook loaded.")

                # Load the subset Excel file
                wb_subset = load_workbook(filename=subset_file_path)
                if 'RTVM' in wb_subset.sheetnames:
                    ws_subset = wb_subset['RTVM']
                else:
                    ws_subset = wb_subset.active  # Use the first sheet
                print("Subset workbook loaded.")

                # Build a mapping from DOORS SPEC ID to row in main workbook
                doors_spec_id_to_row_main = {}
                for row in ws_main.iter_rows(min_row=2, values_only=False):
                    doors_spec_id = row[0].value  # Column A
                    if doors_spec_id:
                        doors_spec_id_to_row_main[doors_spec_id] = row

                # For progress tracking
                total_rows = ws_subset.max_row - 1  # Exclude header
                progress_queue.put(('total', total_rows))
                current_row_num = 0

                # Iterate over the subset rows
                for row_subset in ws_subset.iter_rows(min_row=2, values_only=False):
                    current_row_num += 1
                    progress_queue.put(('progress', current_row_num))

                    doors_spec_id = row_subset[0].value  # Column A
                    if not doors_spec_id:
                        continue  # Skip rows without DOORS SPEC ID

                    if doors_spec_id in doors_spec_id_to_row_main:
                        # Merge data into the main workbook
                        row_main = doors_spec_id_to_row_main[doors_spec_id]

                        # For columns G and H (columns 7 and 8), combine the data
                        for col_idx in [7, 8]:  # Columns G and H
                            cell_main = row_main[col_idx - 1]  # 0-based index
                            cell_subset = row_subset[col_idx - 1]

                            value_main = cell_main.value or ''
                            value_subset = cell_subset.value or ''
                            if value_subset and value_subset not in value_main:
                                if value_main:
                                    combined_value = f"{value_main}\n{value_subset}"
                                else:
                                    combined_value = value_subset
                                cell_main.value = combined_value
                    else:
                        # DOORS SPEC ID not found in main workbook
                        # You may choose to add the new row or skip it
                        pass

                # Save the merged workbook
                wb_main.save(output_file_path)
                wb_main.close()
                wb_subset.close()
                print("Merged workbook saved.")

                # Signal completion
                progress_queue.put(('done', None))

            except Exception as e:
                progress_queue.put(('error', f"An error occurred during merging:\n{e}"))
                print(f"An error occurred during merging: {e}")

        # Create a queue to communicate with the background thread
        progress_queue = queue.Queue()

        # Start the background thread
        threading.Thread(target=process_data, args=(progress_queue,)).start()

        # Create a progress bar
        progress_window = tk.Toplevel()
        progress_window.title("Merging Subset")
        progress_label = tk.Label(progress_window, text="Merging subset file...")
        progress_label.pack()
        progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=300, mode='determinate')
        progress_bar.pack()
        progress_info = tk.Label(progress_window, text="0%")
        progress_info.pack()

        # Variables to keep track of progress
        total_rows = 0
        current_row = 0

        # Function to update the progress bar
        def update_progress():
            nonlocal total_rows, current_row
            try:
                while True:
                    message_type, data = progress_queue.get_nowait()
                    if message_type == 'total':
                        total_rows = data
                        progress_bar['maximum'] = total_rows
                    elif message_type == 'progress':
                        current_row = data
                        if total_rows > 0:
                            percent = int((current_row / total_rows) * 100)
                            progress_bar['value'] = current_row
                            progress_info.config(text=f"{percent}%")
                    elif message_type == 'error':
                        messagebox.showerror("Error", data)
                        progress_window.destroy()
                        return
                    elif message_type == 'done':
                        progress_bar['value'] = total_rows
                        progress_info.config(text="100%")
                        progress_window.destroy()
                        messagebox.showinfo(
                            "Merge Complete",
                            f"The subset file has been merged into {output_file_path}"
                        )
                        return
            except queue.Empty:
                pass
            except Exception as e:
                print(f"Error in update_progress: {e}")
            # Schedule the next check
            progress_window.after(100, update_progress)

        # Start updating the progress bar
        update_progress()


### End of RTVM Subset tool pack ################################################################################################################################################################################################################################           




    def open_management_window(self):
        # Create a new window
        self.management_window = tk.Toplevel(self.root)
        self.management_window.title("Management")

        # Set window size (optional)
        self.management_window.geometry("800x600")  # Adjust as needed

        # Create a frame for the buttons
        button_frame = tk.Frame(self.management_window)
        button_frame.pack(side=tk.TOP, fill=tk.X)

        # Export to Excel Button
        export_button = tk.Button(button_frame, text="Export to Excel", command=self.export_management_table)
        export_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Create the management table
        columns = ("CDRL NUMBER (DI Number)", "Accepted", "Depreciated", "Proposed Add", "Proposed Delete", "Awaiting Input")
        self.management_table = ttk.Treeview(self.management_window, columns=columns, show="headings")
        self.management_table.pack(fill=tk.BOTH, expand=True)

        # Configure headings with sorting
        for col in columns:
            self.management_table.heading(col, text=col, command=lambda c=col: self.sort_management_table(c))
            self.management_table.column(col, anchor='center', width=130)  # Adjust width as needed

        # Populate the management table
        self.populate_management_table()


    def populate_management_table(self):
        # Clear existing data in the management table
        for item in self.management_table.get_children():
            self.management_table.delete(item)

        # Initialize a dictionary to hold aggregated data
        aggregated_data = {}

        # Loop through the status data
        for item in self.status_data:
            di_number = item.get('di_number', 'Unknown')
            object_status = item.get('object_status', '').lower()
            contractor_status = item.get('contractor_status', '').lower()
            government_status = item.get('government_status', '').lower()

            # Initialize the DI Number entry if not present
            if di_number not in aggregated_data:
                aggregated_data[di_number] = {
                    'Accepted': 0,
                    'Depreciated': 0,
                    'Proposed Add': 0,
                    'Proposed Delete': 0,
                    'Awaiting Input': 0
                }

            # Update counts based on statuses
            if object_status == 'accepted':
                aggregated_data[di_number]['Accepted'] += 1
            elif object_status == 'depreciated':
                aggregated_data[di_number]['Depreciated'] += 1

            # Update counts for 'Proposed Add' and 'Proposed Delete' based on contractor_status
            if object_status == 'proposed add':
                aggregated_data[di_number]['Proposed Add'] += 1
            elif object_status == 'proposed remove' or object_status == 'proposed delete':
                aggregated_data[di_number]['Proposed Delete'] += 1

            # Update count for 'Awaiting Input' based on government_status
            if government_status == 'awaiting input':
                aggregated_data[di_number]['Awaiting Input'] += 1

        # Populate the management table with aggregated data
        for di_number, counts in aggregated_data.items():
            self.management_table.insert(
                "", "end",
                values=(
                    di_number,
                    counts['Accepted'],
                    counts['Depreciated'],
                    counts['Proposed Add'],
                    counts['Proposed Delete'],
                    counts['Awaiting Input']
                )
            )






    def export_management_table(self):
        # Prompt the user to select a file location
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                 title="Save Management Table as Excel File")
        if file_path:
            try:
                # Create a DataFrame from the management table
                columns = ["CDRL NUMBER (DI Number)", "Accepted", "Depreciated", "Proposed Add", "Proposed Delete", "Awaiting Input"]
                data = []
                for item in self.management_table.get_children():
                    values = self.management_table.item(item)['values']
                    data.append(values)
                df = pd.DataFrame(data, columns=columns)

                # Save the DataFrame to an Excel file
                df.to_excel(file_path, index=False)

                messagebox.showinfo("Success", f"Management table exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {e}")

    def sort_management_table(self, sort_column):
        # Determine the sort order
        if hasattr(self, 'sort_column') and self.sort_column == sort_column:
            # Toggle sort order
            self.sort_descending = not self.sort_descending
        else:
            # New sort column, default to ascending
            self.sort_descending = False
        self.sort_column = sort_column

        # Extract data from the table
        data = []
        for item in self.management_table.get_children():
            values = self.management_table.item(item)['values']
            data.append(values)

        # Get the index of the sort column
        columns = ["CDRL NUMBER (DI Number)", "Accepted", "Depreciated", "Proposed Add", "Proposed Delete", "Awaiting Input"]
        sort_col_index = columns.index(sort_column)

        # Convert the data to appropriate types for sorting
        for row in data:
            # Attempt to convert numerical values (except for the DI Number)
            for i in range(1, len(row)):
                try:
                    row[i] = int(row[i])
                except ValueError:
                    row[i] = 0  # Default to 0 if conversion fails

        # Sort the data
        data.sort(key=lambda x: x[sort_col_index], reverse=self.sort_descending)

        # Clear the table
        for item in self.management_table.get_children():
            self.management_table.delete(item)

        # Insert sorted data back into the table
        for row in data:
            self.management_table.insert("", "end", values=row)

    def on_progress_bar_click(self, event):
        # Identify the row under the cursor
        item_id = self.progress_table.identify_row(event.y)
        if item_id:
            # Get the row number from the item
            row_value = self.progress_table.item(item_id, 'values')[0]
            # Remove '>' if present
            row_value = row_value.lstrip('>')
            try:
                selected_row_number = int(row_value)
                # Adjust for zero-based index and header row
                new_row_index = selected_row_number - 2  # Assuming header is on Excel row 1
                max_row_index = len(self.df) - 1
                if 0 <= new_row_index <= max_row_index:
                    # Save any unsaved data before changing row
                    self.save_comments_to_excel()
                    # Update current row
                    self.current_row = new_row_index
                    # Update the UI
                    self.update_ui_after_navigation()
                    # Update row indicator
                    self.row_indicator_var.set(f"Row: {self.current_row + 2}")
                else:
                    messagebox.showerror(
                        "Invalid Row", "Selected row number is out of range.")
            except ValueError:
                pass  # If conversion to int fails, do nothing

    def toggle_progress_bar(self):
        if self.show_progress_var.get() == 1:
            # Show the progress bar table and label
            self.progress_label.grid()
            self.progress_table.grid()
            # Adjust column weights
            self.root.grid_columnconfigure(0, weight=1)
        else:
            # Hide the progress bar table and label
            self.progress_label.grid_remove()
            self.progress_table.grid_remove()
            # Adjust column weights
            self.root.grid_columnconfigure(0, weight=0)

    def populate_progress_table(self):
        if self.df is None:
            return
        try:
            # Clear existing items
            self.progress_table.delete(*self.progress_table.get_children())
            total_rows = len(self.df)
            for i in range(2, total_rows + 2):  # Rows are numbered starting from 2
                self.progress_table.insert("", "end", values=(str(i),))
        except Exception as e:
            print(f"Error populating progress table: {e}")

            
    def update_progress_bar_highlight(self):
        # First, remove existing highlights
        for item in self.progress_table.get_children():
            # Remove all custom tags
            self.progress_table.item(item, tags=())
            # Get the row number
            row_value = self.progress_table.item(item, 'values')[0]
            # Remove '>' if present
            if row_value.startswith('>'):
                row_value = row_value[1:]
                self.progress_table.item(item, values=(row_value,))
        # Now, highlight the current row and filtered rows
        current_row_number = self.current_row + 2  # Adjust for header
        for item in self.progress_table.get_children():
            row_value = self.progress_table.item(item, 'values')[0]
            df_row_index = int(row_value) - 2  # Adjust for DataFrame index
            tags = []
            # Check if this is the current row
            if int(row_value.lstrip('>')) == current_row_number:
                # Add '>' and 'current_row' tag
                self.progress_table.item(item, values=('>' + row_value.lstrip('>'),))
                tags.append('current_row')
            # Check if filters are applied and this row is in filtered rows
            elif hasattr(self, 'filtered_row_indices') and self.filtered_row_indices and df_row_index in self.filtered_row_indices:
                tags.append('filtered_row')
            # Set the tags
            self.progress_table.item(item, tags=tuple(tags))
        # Configure the highlight styles
        self.progress_table.tag_configure('current_row', background='lightblue')
        self.progress_table.tag_configure('filtered_row', background='yellow')


    def toggle_history_tables(self):
        if self.show_history_var.get() == 1:
            # Show the history tables and labels
            self.comment_history_label.grid()
            self.comment_history_table.grid()
            self.gov_comment_history_label.grid()
            self.gov_comment_history_table.grid()
            # Adjust column weights
            self.root.grid_columnconfigure(3, weight=1)
            self.root.grid_columnconfigure(4, weight=1)
            # Reset weights of columns 0-2
            self.root.grid_columnconfigure(0, weight=1)
            self.root.grid_columnconfigure(1, weight=1)
            self.root.grid_columnconfigure(2, weight=1)
        else:
            # Hide the history tables and labels
            self.comment_history_label.grid_remove()
            self.comment_history_table.grid_remove()
            self.gov_comment_history_label.grid_remove()
            self.gov_comment_history_table.grid_remove()
            # Adjust column weights
            self.root.grid_columnconfigure(3, weight=0)
            self.root.grid_columnconfigure(4, weight=0)
            # Increase weights of columns 0-2 to fill space
            self.root.grid_columnconfigure(0, weight=2)
            self.root.grid_columnconfigure(1, weight=2)
            self.root.grid_columnconfigure(2, weight=2)



    def upload_excel_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            try:
                # Load the Excel file
                self.df = pd.read_excel(file_path)
                self.excel_file_path = file_path  # Store the file path
                self.current_row = 0  # Reset to the first row

                # Debugging output: Print column names
                print(f"DataFrame columns: {self.df.columns.tolist()}")

                # Process the DataFrame to extract statuses
                self.process_statuses()

                # **Add this line to populate the progress bar table**
                self.populate_progress_table()

                # Update the UI
                self.update_ui_after_navigation()
                self.row_indicator_var.set(f"Row: {self.current_row + 2}")
                messagebox.showinfo(
                    "Success", "Excel file loaded successfully.")
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to load Excel file: {e}")


    def process_statuses(self):
        self.status_data = []  # List to store status info for each status entry
        self.unique_object_statuses = set()
        self.unique_contractor_statuses = set()
        self.unique_government_statuses = set()
        try:
            # Loop through the DataFrame
            for index, row in self.df.iterrows():
                data_cell = row[5]  # Adjust if the data is not in column F
                if pd.isna(data_cell):
                    continue
                elif not isinstance(data_cell, str):
                    data_cell = str(data_cell)
                # Split entries separated by lines of underscores
                entries = data_cell.split('______________________')
                for entry in entries:
                    entry = entry.strip()
                    if not entry:
                        continue
                    # Initialize variables
                    data = {
                        'Object Identifier': "",
                        'DI Number': "",
                        'Object Status': "",
                        'Contractor Assessed Status': "",
                        'Government Assessed Status': ""
                    }
                    # Split entry into lines
                    lines = entry.split('\n')
                    for line in lines:
                        line = line.strip()
                        if ':' in line:
                            key, value = line.split(':', 1)
                            key = key.strip()
                            value = value.strip()
                            if key in data:
                                data[key] = value
                    # Add to status_data
                    obj_id = data['Object Identifier']
                    di_number = data['DI Number']
                    object_status = data['Object Status']
                    contractor_status = data['Contractor Assessed Status']
                    government_status = data['Government Assessed Status']
                    self.unique_object_statuses.add(object_status)
                    self.unique_contractor_statuses.add(contractor_status)
                    self.unique_government_statuses.add(government_status)
                    self.status_data.append({
                        'row_index': index,
                        'veridoc_number': obj_id,
                        'di_number': di_number,
                        'object_status': object_status,
                        'contractor_status': contractor_status,
                        'government_status': government_status
                    })
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process statuses: {e}")



    def open_filter_dialog(self):
        if not self.unique_object_statuses and not self.unique_contractor_statuses and not self.unique_government_statuses:
            messagebox.showerror(
                "Error", "No status data available. Please upload an Excel file with valid data.")
            return
        # Create a new window
        self.filter_window = tk.Toplevel(self.root)
        self.filter_window.title("Apply Filter")

        # Object Status
        object_label = tk.Label(self.filter_window, text="Object Status")
        object_label.pack()
        self.object_status_var = tk.StringVar()
        self.object_status_dropdown = ttk.Combobox(
            self.filter_window,
            textvariable=self.object_status_var,
            values=["Any"] + sorted(self.unique_object_statuses),
            state="readonly"
        )
        self.object_status_dropdown.pack()
        self.object_status_var.set("Any")  # Set default value

        # Contractor Assessed Status
        contractor_label = tk.Label(self.filter_window, text="Contractor Assessed Status")
        contractor_label.pack()
        self.contractor_status_var = tk.StringVar()
        self.contractor_status_dropdown = ttk.Combobox(
            self.filter_window,
            textvariable=self.contractor_status_var,
            values=["Any"] + sorted(self.unique_contractor_statuses),
            state="readonly"
        )
        self.contractor_status_dropdown.pack()
        self.contractor_status_var.set("Any")  # Set default value

        # Government Assessed Status
        government_label = tk.Label(self.filter_window, text="Government Assessed Status")
        government_label.pack()
        self.government_status_var = tk.StringVar()
        self.government_status_dropdown = ttk.Combobox(
            self.filter_window,
            textvariable=self.government_status_var,
            values=["Any"] + sorted(self.unique_government_statuses),
            state="readonly"
        )
        self.government_status_dropdown.pack()
        self.government_status_var.set("Any")  # Set default value

        # Apply Button
        apply_btn = tk.Button(self.filter_window, text="Apply", command=self.apply_filters)
        apply_btn.pack(pady=10)
        
    def apply_filters(self):
        # Get selected statuses
        selected_object_status = self.object_status_var.get()
        selected_contractor_status = self.contractor_status_var.get()
        selected_government_status = self.government_status_var.get()
        # Save comments before changing current row
        self.save_comments_to_excel()
        # Treat "Any" as no filter
        if selected_object_status == "Any":
            selected_object_status = ''
        if selected_contractor_status == "Any":
            selected_contractor_status = ''
        if selected_government_status == "Any":
            selected_government_status = ''

        # Filter the status data
        self.filtered_row_indices = []
        filtered_veridoc_numbers = []
        for item in self.status_data:
            # Check if the statuses match
            object_match = True
            contractor_match = True
            government_match = True

            if selected_object_status:
                object_match = (item['object_status'] == selected_object_status)
            if selected_contractor_status:
                contractor_match = (item['contractor_status'] == selected_contractor_status)
            if selected_government_status:
                government_match = (item['government_status'] == selected_government_status)

            if object_match and contractor_match and government_match:
                self.filtered_row_indices.append(item['row_index'])
                # Collect VeriDoc numbers
                filtered_veridoc_numbers.append(item['veridoc_number'])

        # Count unique VeriDoc numbers
        unique_veridoc_numbers = set(filtered_veridoc_numbers)

        # Reset the current row index to 0 in the filtered list
        self.current_filtered_index = 0
        if self.filtered_row_indices:
            self.current_row = self.filtered_row_indices[self.current_filtered_index]
            self.update_ui_after_navigation()
            self.row_indicator_var.set(f"Row: {self.current_row + 2}")
            messagebox.showinfo(
                "Filter Applied",
                f"{len(self.filtered_row_indices)} rows match the selected filters.\n"
                f"{len(unique_veridoc_numbers)} unique VeriDoc(s) found."
            )
        else:
            messagebox.showinfo(
                "No Results", "No rows match the selected filters.")
        # Close the filter window
        self.filter_window.destroy()
        # Update the progress bar highlight
        self.update_progress_bar_highlight()
    def clear_filters(self):
        if hasattr(self, 'filtered_row_indices'):
            del self.filtered_row_indices
        self.current_row = 0
        self.update_ui_after_navigation()
        self.row_indicator_var.set(f"Row: {self.current_row + 2}")
        messagebox.showinfo("Filter Cleared", "Filters have been cleared.")
        # Update the progress bar highlight
        self.update_progress_bar_highlight()

    def navigate_cells(self, direction):
        if self.df is None:
            return

        # Instead of a blocking save, run the save operation in a background thread.
        self.save_comments_to_excel_background()

        if hasattr(self, 'filtered_row_indices') and self.filtered_row_indices:
            max_index = len(self.filtered_row_indices) - 1
            if direction == 'up':
                self.current_filtered_index = max(self.current_filtered_index - 1, 0)
            elif direction == 'down':
                self.current_filtered_index = min(self.current_filtered_index + 1, max_index)
            self.current_row = self.filtered_row_indices[self.current_filtered_index]
        else:
            max_row = len(self.df) - 1
            if direction == 'up':
                self.current_row = max(self.current_row - 1, 0)
            elif direction == 'down':
                self.current_row = min(self.current_row + 1, max_row)
        self.update_ui_after_navigation()
        self.row_indicator_var.set(f"Row: {self.current_row + 2}")


    def jump_to_cell(self):
        try:
            # Save comments before navigating away
            self.save_comments_to_excel()
            row_number = int(self.jump_to_var.get()) - 2  # Adjust for header
            if 0 <= row_number < len(self.df):
                self.current_row = row_number
                self.update_ui_after_navigation()
                self.row_indicator_var.set(f"Row: {self.current_row + 2}")
            else:
                messagebox.showerror(
                    "Invalid Row", "Row number out of range.")
        except ValueError:
            messagebox.showerror(
                "Invalid Input", "Please enter a valid row number.")


    def extract_info(self, data):
        # Clear highlights before clearing tables
        self.clear_highlights()

        # Clear the tables for new extraction
        self.table.delete(*self.table.get_children())
        self.comment_table.delete(*self.comment_table.get_children())

        # Check if data is empty
        if not data.strip():
            messagebox.showwarning("No Data", "No data available to extract.")
            return

        # Get comments from column H (index 7)
        comments_content = self.df.iloc[self.current_row, 7]
        if pd.isna(comments_content):
            comments_content = ""
        elif not isinstance(comments_content, str):
            comments_content = str(comments_content)

        # Split into lines and create a dictionary
        comment_lines = comments_content.strip().split('\n')
        comments_dict = {}
        for line in comment_lines:
            if ' - ' in line:
                obj_id, comment_text = line.split(' - ', 1)
                comments_dict[obj_id.strip()] = comment_text.strip()

        # Store comments for the current row
        self.current_comments = comments_dict

        # Regex pattern to match full entries including Object Identifier, DI Number, etc.
        pattern = (r'Object Identifier:\s*(?P<obj_id>WCC-VERI-DOC-(?P<veridoc_num>\d+)).*?'
                   r'DI Number:\s*(?P<di_num>\d+-\d+).*?'
                   r'CDRL Subtitle:\s*(?P<cdrl_subtitle>.*?)\n.*?'
                   r'Object Status:\s*(?P<object_status>.*?)\n.*?'
                   r'Contractor Assessed Status:\s*(?P<contractor_status>.*?)\n.*?'
                   r'Government Assessed Status:\s*(?P<government_status>.*?)\n')

        matches = re.finditer(pattern, data, re.DOTALL)

        count = 0  # Counter for the number of rows
        for match in matches:
            obj_id = match.group('obj_id')
            veridoc_num = match.group('veridoc_num')
            di_num = match.group('di_num')
            cdrl_subtitle = match.group('cdrl_subtitle').strip()
            object_status = match.group('object_status').strip()
            contractor_status = match.group('contractor_status').strip()
            government_status = match.group('government_status').strip()

            # Only insert into the table if we have a valid DI Number and CDRL Subtitle
            # This ensures that lines that are purely comments (without a proper DI Number or subtitle)
            # will not be inserted.
            if not di_num or not cdrl_subtitle:
                # Skip this match if it doesn't contain the required info
                continue

            # Insert the item into the data display table
            item_id = self.table.insert(
                "", "end",
                values=(obj_id, di_num, cdrl_subtitle, object_status, contractor_status, government_status)
            )

            # Get existing comment if any
            if obj_id in comments_dict:
                comment_text = f"{obj_id} - {comments_dict[obj_id]}"
            else:
                # Prepopulate the comment with the Object Identifier followed by " - "
                comment_text = f"{obj_id} - "

            self.comment_table.insert("", "end", values=(comment_text,))

            # Apply conditional formatting
            self.apply_conditional_formatting(item_id, contractor_status, government_status)

            # Check if Object Status is "DEPRECIATED"
            if object_status.upper() == "DEPRECIATED":
                messagebox.showwarning(
                    "Warning",
                    "The Object Status is marked as DEPRECIATED.\nA CDRL that the government concurs should be removed "
                    "is pending the full removal of the object from DOORs. No DLOC edits will be accepted for a "
                    "VERI-DOC in this status."
                )

            count += 1  # Increment the counter

        # Adjust the table height based on the number of rows
        self.adjust_table_height(count)

    def adjust_table_height(self, row_count):
        # Set a maximum height to prevent the table from becoming too large
        max_height = 20  # You can adjust this value as needed
        height = min(row_count, max_height)
        self.table.config(height=height)
        self.comment_table.config(height=height)

    def apply_conditional_formatting(self, item_id, contractor_status, government_status):
        # Apply conditional formatting based on Contractor Assessed Status
        if contractor_status == "SAT":
            self.table.item(item_id, tags=("sat_row",))
        elif contractor_status == "UNSAT":
            self.table.item(item_id, tags=("unsat_row",))
        elif contractor_status == "TBD":
            self.table.item(item_id, tags=("tbd_row",))

        # Apply conditional formatting based on Government Assessed Status
        if government_status == "Agree":
            self.table.set(item_id, "#6", government_status)
            self.table.tag_configure("GovAgree.Cell", background="green", foreground="white")
            self.table.set(item_id, "Government Assessed Status", government_status)
            current_tags = self.table.item(item_id, "tags")
            if isinstance(current_tags, str):
                current_tags = (current_tags,)
            elif current_tags is None:
                current_tags = ()
            self.table.item(item_id, tags=current_tags + ("GovAgree.Cell",))
        elif government_status == "Disagree":
            self.table.set(item_id, "#6", government_status)
            self.table.tag_configure("GovDisagree.Cell", background="red", foreground="white")
            self.table.set(item_id, "Government Assessed Status", government_status)
            current_tags = self.table.item(item_id, "tags")
            if isinstance(current_tags, str):
                current_tags = (current_tags,)
            elif current_tags is None:
                current_tags = ()
            self.table.item(item_id, tags=current_tags + ("GovDisagree.Cell",))

        # Configure row tags
        self.table.tag_configure("sat_row", background="green", foreground="white")
        self.table.tag_configure("unsat_row", background="red", foreground="white")
        self.table.tag_configure("tbd_row", background="yellow", foreground="black")

    def update_proposed_changes_table(self):
        # Clear the table
        self.proposed_changes_table.delete(
            *self.proposed_changes_table.get_children())

        # Get the content from cell G (column index 6)
        content = self.df.iloc[self.current_row, 6]
        if pd.isna(content):
            content = ""
        elif not isinstance(content, str):
            content = str(content)

        # Split the content into lines and populate the table
        lines = content.strip().split('\n')
        for line in lines:
            if line.strip():
                self.proposed_changes_table.insert(
                    "", "end", values=(line.strip(),))

    def show_table_context_menu(self, event):
        # Identify the row under the cursor
        row_id = self.table.identify_row(event.y)
        if row_id:
            # Get values from the selected row
            item = self.table.item(row_id)
            values = item['values']
            if values:
                object_status = values[3]  # Object Status is at index 3
                if object_status.lower() == "depreciated":
                    # Show message and do not display context menu
                    messagebox.showinfo("Information", "This item is locked as it is depreciated. You do not need to do anything with this")
                else:
                    # Create a context menu
                    menu = tk.Menu(self.root, tearoff=0)
                    menu.add_command(
                        label="Set DI Number to match CDRL", command=lambda: self.set_di_number(row_id))
                    menu.add_command(
                        label="Delete DI Number", command=lambda: self.delete_di_number(row_id))
                    menu.post(event.x_root, event.y_root)
            else:
                # If no values are present, you might want to handle this case
                messagebox.showerror("Error", "No data available for the selected row.")


    def set_di_number(self, row_id):
        # Get values from the selected row
        item = self.table.item(row_id)
        values = item['values']
        if values:
            obj_id = values[0]  # VeriDoc Number (Object Identifier)
            di_num = values[1]
            # Use obj_id directly since it's already extracted

            # Open the PatternDialog pop-up window
            pattern_dialog = PatternDialog(
                self.root, self, obj_id, di_num, self.current_row)
            self.pattern_dialog = pattern_dialog  # Keep reference to the dialog
            # Pass deletions if needed
            pattern_dialog.deletions = self.deletions

            # Optionally, check for Government Assessed Status
            government_status = values[5]
            if government_status == "Pending Review":
                messagebox.showinfo("Information",
                                    "USCG is currently reviewing this line item, no action is required.")
            # Additional actions can be added here

        else:
            messagebox.showerror(
                "Error", "No data available for the selected row.")

    def delete_di_number(self, row_id):
        # Get values from the selected row
        item = self.table.item(row_id)
        values = item['values']
        if values:
            obj_id = values[0]  # VeriDoc Number (Object Identifier)
            # Create deletion pattern
            del_pattern = f"DEL; {obj_id}"

            # Save the deletion pattern to Excel
            self.save_deletion_to_excel(del_pattern)

            # Update the Proposed Changes Table
            self.update_proposed_changes_table()

            messagebox.showinfo("Success", "DI Number deleted and changes saved to Excel file successfully.")
        else:
            messagebox.showerror("Error", "No data available for the selected row.")


    def save_deletion_to_excel(self, del_pattern):
        # Get existing content from cell G (column index 6) of the current row
        existing_content = self.df.iloc[self.current_row, 6]
        if pd.isna(existing_content):
            existing_content = ""
        elif not isinstance(existing_content, str):
            existing_content = str(existing_content)

        # Append the new deletion pattern to the existing content
        if existing_content.strip():
            new_content = existing_content.strip() + "\n" + del_pattern
        else:
            new_content = del_pattern

        # Update the Excel file directly using openpyxl
        from openpyxl import load_workbook

        try:
            # Load the workbook
            wb = load_workbook(self.excel_file_path)
            ws = wb.active  # You may need to select the correct sheet if there are multiple

            # Calculate the Excel row number (considering headers)
            excel_row = self.current_row + 2  # Assuming header is on the first row

            # Update the cell in column G (which is column index 7 in openpyxl)
            ws.cell(row=excel_row, column=7, value=new_content)

            # Save the workbook
            wb.save(self.excel_file_path)

            # Update the DataFrame in memory
            self.df.iloc[self.current_row, 6] = new_content

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save deletion to Excel file: {e}")

    def on_table_row_select(self, event):
        # Clear previous highlights
        self.clear_highlights()

        # Get selected item
        selected_item = self.table.selection()
        if selected_item:
            selected_item = selected_item[0]
            # Get VeriDoc Number from the selected row
            values = self.table.item(selected_item, 'values')
            veridoc_number = values[0]

            # Highlight rows in DI Number Breakdown table with the same VeriDoc Number
            for item in self.table.get_children():
                item_values = self.table.item(item, 'values')
                if item_values[0] == veridoc_number:
                    # Get current tags and ensure they are a tuple
                    current_tags = self.table.item(item, "tags")
                    if isinstance(current_tags, str):
                        current_tags = (current_tags,)
                    elif current_tags is None:
                        current_tags = ()
                    # Add the "highlight" tag
                    self.table.item(item, tags=current_tags + ("highlight",))
                    self.highlighted_items.append(('table', item))

            # Highlight corresponding rows in Comments table
            for item in self.comment_table.get_children():
                comment = self.comment_table.item(item, 'values')[0]
                # Extract the VeriDoc Number from the comment
                comment_veridoc_number = comment.split(' - ')[0]
                if comment_veridoc_number == veridoc_number:
                    # Get current tags and ensure they are a tuple
                    current_tags = self.comment_table.item(item, "tags")
                    if isinstance(current_tags, str):
                        current_tags = (current_tags,)
                    elif current_tags is None:
                        current_tags = ()
                    # Add the "highlight" tag
                    self.comment_table.item(item, tags=current_tags + ("highlight",))
                    self.highlighted_items.append(('comment_table', item))

            # Highlight corresponding rows in Contractor Proposed Change Request Input table
            for item in self.proposed_changes_table.get_children():
                pattern = self.proposed_changes_table.item(item, 'values')[0]
                # Extract the VeriDoc Number from the pattern
                pattern_veridoc_number = self.extract_veridoc_number_from_pattern(
                    pattern)
                if pattern_veridoc_number == veridoc_number:
                    # Get current tags and ensure they are a tuple
                    current_tags = self.proposed_changes_table.item(item, "tags")
                    if isinstance(current_tags, str):
                        current_tags = (current_tags,)
                    elif current_tags is None:
                        current_tags = ()
                    # Add the "highlight" tag
                    self.proposed_changes_table.item(item, tags=current_tags + ("highlight",))
                    self.highlighted_items.append(('proposed_changes_table', item))

            # Highlight corresponding rows in Contractor Proposed Change Comment History table
            for item in self.comment_history_table.get_children():
                history_entry = self.comment_history_table.item(item, 'values')[0]
                # Extract the VeriDoc Number from the history entry if possible
                history_veridoc_number = self.extract_veridoc_number_from_pattern(history_entry)
                if history_veridoc_number == veridoc_number:
                    # Get current tags and ensure they are a tuple
                    current_tags = self.comment_history_table.item(item, "tags")
                    if isinstance(current_tags, str):
                        current_tags = (current_tags,)
                    elif current_tags is None:
                        current_tags = ()
                    # Add the "highlight" tag
                    self.comment_history_table.item(item, tags=current_tags + ("highlight",))
                    self.highlighted_items.append(('comment_history_table', item))

            # Configure the highlight style
            self.table.tag_configure("highlight", background="lightblue")
            self.comment_table.tag_configure("highlight", background="lightblue")
            self.proposed_changes_table.tag_configure("highlight", background="lightblue")
            self.comment_history_table.tag_configure("highlight", background="lightblue")


    def extract_veridoc_number_from_pattern(self, pattern):
        # Try to extract the VeriDoc Number using regex
        match = re.search(r'(WCC-VERI-DOC-\d+)', pattern)
        if match:
            return match.group(1)
        else:
            return None

    def clear_highlights(self):
        # Remove highlight from previously highlighted items
        for table_name, item in self.highlighted_items:
            try:
                if table_name == 'table':
                    # Remove 'highlight' tag while preserving other tags
                    tags = list(self.table.item(item, 'tags'))
                    if 'highlight' in tags:
                        tags.remove('highlight')
                    self.table.item(item, tags=tuple(tags))
                elif table_name == 'comment_table':
                    self.comment_table.item(item, tags=())
                elif table_name == 'proposed_changes_table':
                    self.proposed_changes_table.item(item, tags=())
            except tk.TclError:
                # Item no longer exists, ignore the error
                pass
        self.highlighted_items = []

    def on_comment_double_click(self, event):
        # Identify the row and column under the cursor
        region = self.comment_table.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.comment_table.identify_row(event.y)
            column = self.comment_table.identify_column(event.x)
            x, y, width, height = self.comment_table.bbox(row_id, column)

            # Get the current value
            item = self.comment_table.item(row_id)
            current_text = item['values'][0] if item['values'] else ''

            # Create an entry widget over the cell
            self.entry_popup = tk.Entry(self.comment_table)
            self.entry_popup.place(x=x, y=y, width=width, height=height)
            self.entry_popup.insert(0, current_text)
            self.entry_popup.focus_set()
            self.entry_popup.bind(
                "<Return>", lambda event: self.on_entry_confirm(row_id))
            self.entry_popup.bind(
                "<FocusOut>", lambda event: self.on_entry_confirm(row_id))

    def on_entry_confirm(self, row_id):
        new_text = self.entry_popup.get()
        self.comment_table.item(row_id, values=(new_text,))
        self.entry_popup.destroy()

        # Update self.current_comments
        # Get the index of the item in comment_table
        item_index = self.comment_table.get_children().index(row_id)
        # Get the corresponding item in self.table
        table_items = self.table.get_children()
        if item_index < len(table_items):
            table_row_id = table_items[item_index]
            table_values = self.table.item(table_row_id, 'values')
            obj_id = table_values[0]  # VeriDoc Number
            # Extract the comment text from new_text
            if ' - ' in new_text:
                _, comment_text = new_text.split(' - ', 1)
                self.current_comments[obj_id] = comment_text.strip()
            else:
                # No ' - ', treat the entire text as comment
                self.current_comments[obj_id] = new_text.strip()


    def show_proposed_changes_context_menu(self, event):
        # Identify the row under the cursor
        row_id = self.proposed_changes_table.identify_row(event.y)
        if row_id:
            # Create a context menu
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(
                label="Delete Row and Save to Excel", command=lambda: self.delete_proposed_change_row(row_id))
            menu.post(event.x_root, event.y_root)

    def delete_proposed_change_row(self, row_id):
        # Get the pattern from the selected row
        item = self.proposed_changes_table.item(row_id)
        pattern_to_delete = item['values'][0]
        if pattern_to_delete:
            # Remove the pattern from the DataFrame
            content = self.df.iloc[self.current_row, 6]
            if pd.isna(content):
                content = ""
            elif not isinstance(content, str):
                content = str(content)

            # Split the content into lines
            lines = content.strip().split('\n')
            if pattern_to_delete in lines:
                lines.remove(pattern_to_delete)
                # Join the remaining lines
                new_content = '\n'.join(lines)

                # Update the Excel file directly using openpyxl
                from openpyxl import load_workbook

                try:
                    # Load the workbook
                    wb = load_workbook(self.excel_file_path)
                    ws = wb.active  # You may need to select the correct sheet if there are multiple

                    # Calculate the Excel row number (considering headers)
                    excel_row = self.current_row + 2  # Assuming header is on the first row

                    # Update the cell in column G (which is column index 7 in openpyxl)
                    ws.cell(row=excel_row, column=7, value=new_content)

                    # Save the workbook
                    wb.save(self.excel_file_path)

                    # Update the DataFrame in memory
                    self.df.iloc[self.current_row, 6] = new_content

                    messagebox.showinfo(
                        "Success", "Row deleted and changes saved to Excel file successfully.")
                    # Update the Proposed Changes Table
                    self.update_proposed_changes_table()
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Failed to save to Excel file: {e}")
            else:
                messagebox.showerror(
                    "Error", "Pattern not found in the Excel data.")
        else:
            messagebox.showerror(
                "Error", "No pattern found in the selected row.")


    def update_ui_after_navigation(self):
            if self.df is not None and self.current_row < len(self.df):
                # Get data from column F (index 5)
                data = self.df.iloc[self.current_row, 5]
                # Update Specification Text Box with content from column B (index 1)
                spec_text = self.df.iloc[self.current_row, 1]
                if pd.isna(spec_text):
                    spec_text = ""
                elif not isinstance(spec_text, str):
                    spec_text = str(spec_text)
                self.spec_text_box.config(state=tk.NORMAL)
                self.spec_text_box.delete("1.0", tk.END)
                self.spec_text_box.insert(tk.END, spec_text)
                self.spec_text_box.config(state=tk.DISABLED)
                # Check if data is NaN or not a string
                if pd.isna(data):
                    data = ""
                elif not isinstance(data, str):
                    data = str(data)
                # Run extract_info automatically
                self.extract_info(data)
                # Update the Proposed Changes Table
                self.update_proposed_changes_table()
                # **Update the Comment History Table**
                self.update_comment_history_table()
                # Update the progress bar highlight
                self.update_progress_bar_highlight()
            else:
                messagebox.showerror(
                    "Error", "No data available at this row.")

    def update_comment_history_table(self):
        # Clear the table
        self.comment_history_table.delete(
            *self.comment_history_table.get_children())

        # Get the content from column I (index 8)
        content = self.df.iloc[self.current_row, 8]
        if pd.isna(content):
            content = ""
        elif not isinstance(content, str):
            content = str(content)

        # Split the content into lines and populate the table
        lines = content.strip().split('\n')
        for line in lines:
            if line.strip():
                self.comment_history_table.insert(
                    "", "end", values=(line.strip(),))


    def open_pie_chart_window(self):
        if self.df is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        self.pie_chart_window = tk.Toplevel(self.root)
        self.pie_chart_window.title("Pie Charts")

        # Create filter frame (this stays the same)
        filter_frame = tk.Frame(self.pie_chart_window)
        filter_frame.pack(side=tk.TOP, fill=tk.X)

        # Object Status Filter
        tk.Label(filter_frame, text="Object Status").pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_object_status_var = tk.StringVar()
        self.pie_object_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_object_status_var,
            values=["Any"] + sorted(self.unique_object_statuses),
            state="readonly",
            width=15
        )
        self.pie_object_status_dropdown.pack(side=tk.LEFT)
        self.pie_object_status_var.set("Any")

        # Contractor Assessed Status Filter
        tk.Label(filter_frame, text="Contractor Assessed Status").pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_contractor_status_var = tk.StringVar()
        self.pie_contractor_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_contractor_status_var,
            values=["Any"] + sorted(self.unique_contractor_statuses),
            state="readonly",
            width=15
        )
        self.pie_contractor_status_dropdown.pack(side=tk.LEFT)
        self.pie_contractor_status_var.set("Any")

        # Government Assessed Status Filter
        tk.Label(filter_frame, text="Government Assessed Status").pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_government_status_var = tk.StringVar()
        self.pie_government_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_government_status_var,
            values=["Any"] + sorted(self.unique_government_statuses),
            state="readonly",
            width=15
        )
        self.pie_government_status_dropdown.pack(side=tk.LEFT)
        self.pie_government_status_var.set("Any")

        # Update Charts Button
        tk.Button(filter_frame, text="Update Charts", command=self.update_pie_charts).pack(side=tk.LEFT, padx=10)

        # Now create a frame to hold four charts in a 2×2 grid
        charts_frame = tk.Frame(self.pie_chart_window)
        charts_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Create the four figures and embed them using grid.
        self.object_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.object_status_canvas = FigureCanvasTkAgg(self.object_status_fig, charts_frame)
        self.object_status_canvas.get_tk_widget().grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.contractor_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.contractor_status_canvas = FigureCanvasTkAgg(self.contractor_status_fig, charts_frame)
        self.contractor_status_canvas.get_tk_widget().grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.government_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.government_status_canvas = FigureCanvasTkAgg(self.government_status_fig, charts_frame)
        self.government_status_canvas.get_tk_widget().grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        self.gov_req_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.gov_req_status_canvas = FigureCanvasTkAgg(self.gov_req_status_fig, charts_frame)
        self.gov_req_status_canvas.get_tk_widget().grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

        # Optionally configure the grid so cells expand equally.
        charts_frame.columnconfigure(0, weight=1)
        charts_frame.columnconfigure(1, weight=1)
        charts_frame.rowconfigure(0, weight=1)
        charts_frame.rowconfigure(1, weight=1)

        # Create a frame for the counts table (same as before)
        counts_frame = tk.Frame(self.pie_chart_window)
        counts_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.counts_table = ttk.Treeview(counts_frame, columns=("Status", "Count", "Total Count"), show="headings")
        self.counts_table.heading("Status", text="Status")
        self.counts_table.heading("Count", text="Count")
        self.counts_table.heading("Total Count", text="Total Count")
        self.counts_table.pack(fill=tk.BOTH, expand=True)

        self.compute_total_counts()
        self.update_pie_charts()

    def update_pie_charts(self):
        # Get the selected filters
        selected_object_status = self.pie_object_status_var.get()
        selected_contractor_status = self.pie_contractor_status_var.get()
        selected_government_status = self.pie_government_status_var.get()

        if selected_object_status == "Any":
            selected_object_status = ''
        if selected_contractor_status == "Any":
            selected_contractor_status = ''
        if selected_government_status == "Any":
            selected_government_status = ''

        # Filter the status data based on the selected filters (for the first three charts)
        filtered_items = []
        for item in self.status_data:
            object_match = True
            contractor_match = True
            government_match = True
            if selected_object_status:
                object_match = (item['object_status'] == selected_object_status)
            if selected_contractor_status:
                contractor_match = (item['contractor_status'] == selected_contractor_status)
            if selected_government_status:
                government_match = (item['government_status'] == selected_government_status)
            if object_match and contractor_match and government_match:
                filtered_items.append(item)

        filtered_statuses = {
            'object_status': [item['object_status'] for item in filtered_items],
            'contractor_status': [item['contractor_status'] for item in filtered_items],
            'government_status': [item['government_status'] for item in filtered_items]
        }

        # Plot the first three charts using the filtered data
        self.plot_pie_chart(self.object_status_fig, filtered_statuses['object_status'], 'Object Status')
        self.plot_pie_chart(self.contractor_status_fig, filtered_statuses['contractor_status'], 'Contractor Assessed Status')
        self.plot_pie_chart(self.government_status_fig, filtered_statuses['government_status'], 'Government Assessed Status')

        # --- Now compute the per-requirement overall government status ---
        # We will group by row index (each row is one requirement).
        from collections import defaultdict, Counter
        req_status = defaultdict(list)
        for item in self.status_data:
            row_idx = item['row_index']
            status = item['government_status']
            if isinstance(status, str):
                req_status[row_idx].append(status.strip().lower())

        # Define the order of precedence (lower number = higher precedence)
        precedence = {"agree": 1, "disagree": 2, "pending review": 3, "awaiting input": 4}

        overall_status_list = []
        for row_idx, statuses in req_status.items():
            overall = None
            min_rank = float('inf')
            for s in statuses:
                rank = precedence.get(s, 100)  # default to a high rank if unknown
                if rank < min_rank:
                    min_rank = rank
                    overall = s
            if overall:
                # Capitalize the first letter (so "agree" becomes "Agree", etc.)
                overall_status_list.append(overall.capitalize())

        # Plot the new pie chart using the overall status per requirement.
        self.plot_pie_chart(self.gov_req_status_fig, overall_status_list, 'Government Assessed Status Per Requirement')

        # Redraw all canvases
        self.object_status_canvas.draw()
        self.contractor_status_canvas.draw()
        self.government_status_canvas.draw()
        self.gov_req_status_canvas.draw()

        # Update counts table (using filtered_statuses as before)
        self.update_counts_table(filtered_statuses)


    def plot_pie_chart(self, fig, data_list, title):
        # Clear the figure
        fig.clear()
        from collections import Counter
        counter = Counter(data_list)
        labels = list(counter.keys())
        sizes = list(counter.values())

        # Define your desired color mapping:
        status_color_map = {
            'Agree': 'green',
            'Disagree': 'red',
            'Awaiting input': 'yellow',
            'Pending eview': 'blue'
        }

        # We standardize each label by capitalizing it before checking the mapping.
        colors = []
        for label in labels:
            label_cap = label.capitalize()
            colors.append(status_color_map.get(label_cap, 'grey'))



        # Handle the case of no data.
        if not sizes:
            labels = ['No Data']
            sizes = [1]
            colors = ['lightgrey']

        # Create the pie chart with the specified colors.
        ax = fig.add_subplot(111)
        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=colors)
        ax.axis('equal')
        ax.set_title(title)

    def update_counts_table(self, filtered_statuses):
        # Clear the table
        self.counts_table.delete(*self.counts_table.get_children())

        # Prepare data for the table
        status_types = {
            'Object Status': 'object_status',
            'Contractor Assessed Status': 'contractor_status',
            'Government Assessed Status': 'government_status'
        }

        for status_type, key in status_types.items():
            # Insert a heading row for the status type
            self.counts_table.insert("", "end", values=(status_type, "", ""), tags=('heading',))
            # Count the occurrences in filtered data
            filtered_counter = Counter(filtered_statuses[key])
            # Total counts from the entire dataset
            total_counter = self.total_counts[key]

            # Get all unique statuses to ensure alignment
            all_statuses = set(filtered_counter.keys()).union(set(total_counter.keys()))

            for status in all_statuses:
                filtered_count = filtered_counter.get(status, 0)
                total_count = total_counter.get(status, 0)
                self.counts_table.insert("", "end", values=(status, filtered_count, total_count))
            # Insert an empty row for spacing
            self.counts_table.insert("", "end", values=("", "", ""))
        # Configure the heading row style
        self.counts_table.tag_configure('heading', font=('Arial', 10, 'bold'))


    def compute_total_counts(self):
        # Initialize dictionaries to store total counts
        self.total_counts = {
            'object_status': {},
            'contractor_status': {},
            'government_status': {}
        }

        # Collect all statuses from the entire dataset
        total_statuses = {
            'object_status': [],
            'contractor_status': [],
            'government_status': []
        }
        for item in self.status_data:
            total_statuses['object_status'].append(item['object_status'])
            total_statuses['contractor_status'].append(item['contractor_status'])
            total_statuses['government_status'].append(item['government_status'])

        # Count occurrences for each status type
        for key in self.total_counts.keys():
            counter = Counter(total_statuses[key])
            self.total_counts[key] = counter

    def save_comments_to_excel(self):
        # Prepare the content for the 'Contractor Proposed Change Comment Input' column
        comment_lines = []
        for obj_id, comment_text in self.current_comments.items():
            comment_line = f"{obj_id} - {comment_text}"
            comment_lines.append(comment_line)
        comments_content = '\n'.join(comment_lines)

        # Use openpyxl to write to column H (index 8)
        from openpyxl import load_workbook

        try:
            # Load the workbook
            wb = load_workbook(self.excel_file_path)
            ws = wb.active  # Or specify the sheet if needed

            # Calculate the Excel row number (considering headers)
            excel_row = self.current_row + 2  # Adjust if your header is on a different row

            # Find the column index for 'Contractor Proposed Change Comment Input'
            column_letter = None
            for col in ws.iter_cols(1, ws.max_column):
                if col[0].value == 'Contractor Proposed Change Comment Input':
                    column_letter = col[0].column_letter
                    break

            if column_letter is None:
                messagebox.showerror("Error", "Column 'Contractor Proposed Change Comment Input' not found in Excel file.")
                return

            # Update the cell in the correct column
            ws[f"{column_letter}{excel_row}"] = comments_content

            # Save the workbook
            wb.save(self.excel_file_path)

            # Update the DataFrame in memory
            self.df.at[self.current_row, 'Contractor Proposed Change Comment Input'] = comments_content

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save comments to Excel file: {e}")


    def update_gov_comment_history_table(self):
        print("Updating Government Adjudication Comment History Table")
        # Clear the table
        self.gov_comment_history_table.delete(*self.gov_comment_history_table.get_children())

        # Get the content from column J (index 9)
        try:
            content = self.df.iloc[self.current_row, 9]
        except IndexError:
            messagebox.showerror("Error", "Column J not found in Excel file.")
            return

        if pd.isna(content):
            content = ""
        elif not isinstance(content, str):
            content = str(content)

        # Debugging output
        print(f"Content from column J at row {self.current_row}: {content}")

        # Split the content into lines and populate the table
        lines = content.strip().split('\n')
        for line in lines:
            if line.strip():
                self.gov_comment_history_table.insert("", "end", values=(line.strip(),))



    def on_comment_double_click(self, event):
        # Disable double-click editing
        pass  # Do nothing on double-click

    def show_comment_table_context_menu(self, event):
        # Identify the row under the cursor
        row_id = self.comment_table.identify_row(event.y)
        if row_id:
            # Create a context menu
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(
                label="Edit Comment", command=lambda: self.open_comment_editor(row_id))
            menu.post(event.x_root, event.y_root)

    def open_comment_editor(self, row_id):
        # Get the current comment text
        item = self.comment_table.item(row_id)
        current_text = item['values'][0] if item['values'] else ''

        # Create a new window
        comment_window = tk.Toplevel(self.root)
        comment_window.title("Edit Comment")

        # Initialize the spell checker
        spell = SpellChecker()

        # Large text box
        text_box = tk.Text(comment_window, width=60, height=20)
        text_box.pack(padx=10, pady=10)
        text_box.insert(tk.END, current_text)

        # Bind the key release event for automatic spell checking
        text_box.bind("<KeyRelease>", lambda event: self.check_spelling(text_box, spell))

        # Configure the tag for misspelled words
        text_box.tag_config("misspelled", underline=True, foreground="red")

        # Save Comment button
        save_button = tk.Button(comment_window, text="Save Comment",
                                command=lambda: self.save_comment(row_id, text_box.get("1.0", tk.END), comment_window))
        save_button.pack(pady=10)

    def check_spelling(self, text_widget, spell):
        # Get the content of the text widget
        content = text_widget.get("1.0", tk.END)

        # Split the content into words with positions
        words = content.split()
        index = "1.0"

        # Clear previous tags
        text_widget.tag_remove("misspelled", "1.0", tk.END)

        for word in words:
            # Clean the word by removing punctuation
            clean_word = ''.join(char for char in word if char.isalpha())

            # Get the start and end index of the word
            start_index = text_widget.search(word, index, stopindex=tk.END)
            if not start_index:
                continue
            end_index = f"{start_index}+{len(word)}c"

            # Update the index for the next search
            index = end_index

            if clean_word.lower() not in spell:
                # Highlight the misspelled word
                text_widget.tag_add("misspelled", start_index, end_index)

    def on_right_click(self, event, text_widget, spell):
        # Get the index of the mouse click
        index = text_widget.index(f"@{event.x},{event.y}")

        # Get the word at that index
        word_start = text_widget.search(r'\m', index, backwards=True, regexp=True)
        word_end = text_widget.search(r'\M', index, forwards=True, regexp=True)
        word = text_widget.get(word_start, word_end)

        # Clean the word
        clean_word = ''.join(char for char in word if char.isalpha())

        # Check if the word is misspelled
        if clean_word.lower() not in spell:
            # Get suggestions
            suggestions = spell.candidates(clean_word.lower())

            # Create a context menu
            menu = tk.Menu(self.root, tearoff=0)
            for suggestion in suggestions:
                menu.add_command(label=suggestion, command=lambda s=suggestion: self.replace_word(text_widget, word_start, word_end, s))
            menu.post(event.x_root, event.y_root)

    def save_comment(self, row_id, new_text, window):
        # Update the comment in the comment_table
        self.comment_table.item(row_id, values=(new_text.strip(),))

        # Close the comment window
        window.destroy()

        # Update self.current_comments
        # Get the index of the item in comment_table
        item_index = self.comment_table.get_children().index(row_id)
        # Get the corresponding item in self.table
        table_items = self.table.get_children()
        if item_index < len(table_items):
            table_row_id = table_items[item_index]
            table_values = self.table.item(table_row_id, 'values')
            obj_id = table_values[0]  # VeriDoc Number
            # Extract the comment text from new_text
            if ' - ' in new_text:
                _, comment_text = new_text.split(' - ', 1)
                self.current_comments[obj_id] = comment_text.strip()
            else:
                # No ' - ', treat the entire text as comment
                self.current_comments[obj_id] = new_text.strip()

        # Save comments to Excel file immediately
        self.save_comments_to_excel()
        

import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from textwrap import wrap
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os
from datetime import datetime
from textwrap import wrap


class DisagreementManager:
    def __init__(self, master, app, disagreement_items):
        """
        master: the parent window
        app: the instance of RTVMApp
        disagreement_items: list of items from self.status_data that match the criteria
        """
        self.master = master
        self.app = app
        self.disagreement_items = disagreement_items
        self.current_index = 0
        self.output_folder = None  # Folder where PDFs will be saved

        # Make the window stay on top
        self.master.attributes("-topmost", True)

        # GUI layout
        self.button_frame = tk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X)

        self.up_button = tk.Button(self.button_frame, text="Up", command=self.prev_item)
        self.up_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.down_button = tk.Button(self.button_frame, text="Down", command=self.next_item)
        self.down_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.generate_pdf_button = tk.Button(self.button_frame, text="Generate PDF", command=self.generate_pdf_for_current)
        self.generate_pdf_button.pack(side=tk.LEFT, padx=5, pady=5)

        # New button to create all reports
        self.create_all_button = tk.Button(self.button_frame, text="Create All Reports", command=self.create_all_reports)
        self.create_all_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.select_folder_button = tk.Button(self.button_frame, text="Select Output Folder", command=self.select_output_folder)
        self.select_folder_button.pack(side=tk.LEFT, padx=5, pady=5)

        #Button to create the report in the format of an excel documetn. 
        self.generate_xlsx_button = tk.Button(self.button_frame, text="Generate Excel", command=self.generate_excel_for_current)
        self.generate_xlsx_button.pack(side=tk.LEFT, padx=5, pady=5)


        # Frame for preview
        self.preview_frame = tk.Frame(master)
        self.preview_frame.pack(fill=tk.BOTH, expand=True)

        # A text box to show preview
        self.preview_text = tk.Text(self.preview_frame, wrap="word")
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        self.show_item()

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder for PDFs")
        if folder:
            self.output_folder = folder
            messagebox.showinfo("Folder Selected", f"PDFs will be saved to: {folder}")

    def prev_item(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.show_item()

    def next_item(self):
        if self.current_index < len(self.disagreement_items) - 1:
            self.current_index += 1
            self.show_item()

    def show_item(self):
        # Clear the preview
        self.preview_text.delete("1.0", tk.END)

        # Get current row_index
        item = self.disagreement_items[self.current_index]
        row_index = item['row_index']

        # Set current_row in the app and update UI so that extract_info() is called
        self.app.current_row = row_index
        self.app.update_ui_after_navigation()

        # Now self.app.table contains all lines (agreements and disagreements) for this row
        # Also self.app.spec_text_box, self.app.comment_table, etc. are updated

        # For preview, we can just show the spec text and the lines we have in self.app.table.
        spec_text = self.app.spec_text_box.get("1.0", tk.END).strip()
        self.preview_text.insert(tk.END, "Specification Text:\n" + spec_text + "\n\n---\n\n")

        # Insert DI Number Breakdown lines
        self.preview_text.insert(tk.END, "DI Number Breakdown:\n")
        for line_id in self.app.table.get_children():
            values = self.app.table.item(line_id, 'values')
            # Each line is a tuple like (VeriDoc, DI Number, CDRL Subtitle, Object Status, Contractor Status, Government Status)
            self.preview_text.insert(tk.END, ", ".join(str(v) for v in values) + "\n")

        self.preview_text.insert(tk.END, "\n---\n\nComments:\n")
        # Insert comments
        for c_id in self.app.comment_table.get_children():
            c_values = self.app.comment_table.item(c_id, 'values')
            if c_values:
                self.preview_text.insert(tk.END, c_values[0] + "\n")



### This creates the diagreement reports as an excel file. 

    def generate_excel_for_current(self, show_popup=True):
        """
        Creates an Excel file (.xlsx) containing the same data that would normally go into the PDF.
        """
        # 1. Gather the same data as in generate_pdf_for_current
        item = self.disagreement_items[self.current_index]
        row_index = item['row_index']

        # Update UI and load current row data
        self.app.current_row = row_index
        self.app.update_ui_after_navigation()

        # DOORS SPEC ID
        doors_spec_id = self.app.df.iloc[row_index, 0]
        if pd.isna(doors_spec_id):
            doors_spec_id = ""
        elif not isinstance(doors_spec_id, str):
            doors_spec_id = str(doors_spec_id)

        # Spec text
        spec_text = self.app.spec_text_box.get("1.0", "end").strip()

        # Contractor Proposed Change Comment History (col 8)
        contractor_history_content = ""
        if len(self.app.df.columns) > 8:
            val = self.app.df.iloc[row_index, 8]
            if pd.isna(val):
                val = ""
            contractor_history_content = str(val)

        # Government Adjudication Comment History (col 9)
        gov_history_content = ""
        if len(self.app.df.columns) > 9:
            val = self.app.df.iloc[row_index, 9]
            if pd.isna(val):
                val = ""
            gov_history_content = str(val)

        # Table data from self.app.table
        items = self.app.table.get_children()
        breakdown_data = [
            ["VeriDoc Number", "DI Number", "CDRL Subtitle", "Gov. Assessed Status"]
        ]
        for line_id in items:
            values = self.app.table.item(line_id, 'values')
            # (VeriDoc, DI Number, CDRL Subtitle, Object Status, Contractor Status, Government Status)
            # We'll store relevant columns
            breakdown_data.append([values[0], values[1], values[2], values[5]])

        # Count how many are "agree" vs "disagree"
        agree_count = 0
        disagree_count = 0
        for i in range(1, len(breakdown_data)):
            gov_status = str(breakdown_data[i][3]).lower()
            if gov_status == "agree":
                agree_count += 1
            elif gov_status == "disagree":
                disagree_count += 1

        # Extract specifically which lines are "disagree"
        disagreement_rows = []
        for i in range(1, len(breakdown_data)):
            gov_status = str(breakdown_data[i][3]).lower()
            if gov_status == "disagree":
                disagreement_rows.append(breakdown_data[i])

        # 2. Prepare an Excel workbook with openpyxl
        wb = Workbook()
        ws = wb.active
        ws.title = "Disagreement Report"

        # Basic styling references
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))

        # 3. Fill some header info
        current_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws["A1"] = "Contract: 70Z02323D93270001"
        ws["A2"] = f"Date/Time: {current_dt}"
        ws["A3"] = ("Distribution Statement D: Distribution authorized to DHS/CG/DOD "
                    "and their contractors only due to administrative or operational use.\n"
                    "Other requests shall be referred to COMMANDANT (CG-9327).")
        ws["A4"] = ("DESTRUCTION NOTICE: Destroy this document by any method that "
                    "will prevent disclosure of contents or reconstruction of the document.")
        for row in range(1, 5):
            ws.row_dimensions[row].height = 30  # Make them a bit taller

        # 4. Summary info table: DOORS SPEC, row, total agreements, disagreements
        ws["A6"] = "DOORS SPEC ID"
        ws["B6"] = "Excel Row"
        ws["C6"] = "Total Agreements"
        ws["D6"] = "Total Disagreements"

        ws["A7"] = doors_spec_id
        ws["B7"] = str(row_index + 2)  # +2 for Excel-based row counting
        ws["C7"] = str(agree_count)
        ws["D7"] = str(disagree_count)

        for col in range(1, 5):
            cell = ws.cell(row=6, column=col)
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = thin_border

            cell2 = ws.cell(row=7, column=col)
            cell2.border = thin_border

        # 5. Write out the Specification Text
        spec_start = 9
        ws.cell(spec_start, 1, "Specification Text:")
        ws.cell(spec_start, 1).font = bold_font
        # Put the spec text in the next row
        ws.cell(spec_start+1, 1, spec_text)
        # Optionally wrap text
        ws.cell(spec_start+1, 1).alignment = Alignment(wrap_text=True)

        # 6. Comments
        comment_start = spec_start + 3
        ws.cell(comment_start, 1, "Contractor Proposed Change Comment History:")
        ws.cell(comment_start, 1).font = bold_font

        # We'll split by new lines and write each line in a new row
        c_lines = contractor_history_content.strip().split('\n')
        row_c = comment_start + 1
        for cline in c_lines:
            if not cline.strip() or '_____' in cline:
                continue
            ws.cell(row_c, 1, cline)
            row_c += 1

        # Then government
        row_c += 2
        ws.cell(row_c, 1, "Government Adjudication Comment History:")
        ws.cell(row_c, 1).font = bold_font
        row_c += 1
        g_lines = gov_history_content.strip().split('\n')
        for gline in g_lines:
            if not gline.strip() or '_____' in gline:
                continue
            ws.cell(row_c, 1, gline)
            row_c += 1

        # 7. Breakdown table
        # Let's put it under everything else
        breakdown_start = row_c + 2
        ws.cell(breakdown_start, 1, "Breakdown Data:")
        ws.cell(breakdown_start, 1).font = bold_font
        breakdown_start += 1

        # Insert headers
        for col_idx, header in enumerate(breakdown_data[0], start=1):
            cell = ws.cell(breakdown_start, col_idx, header)
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = thin_border

        # Insert data rows
        for i in range(1, len(breakdown_data)):
            row_data = breakdown_data[i]
            row_num = breakdown_start + i
            for col_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row_num, col_idx, val)
                cell.alignment = Alignment(wrap_text=True)
                cell.border = thin_border
            
                # Simple color-coding if Government Status is "disagree" or "agree"
                if col_idx == 4:  # Gov. Status
                    if str(val).lower() == "disagree":
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        cell.font = Font(color="FFFFFF")  # White text
                    elif str(val).lower() == "agree":
                        cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                        cell.font = Font(color="FFFFFF")  # White text
                    elif str(val).lower() == "pending review":
                        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow
                        cell.font = Font(color="000000")  # Black text

        # 8. Disagreement Comments
        disagree_section = breakdown_start + len(breakdown_data) + 2
        ws.cell(disagree_section, 1, "Disagreement Comments:")
        ws.cell(disagree_section, 1).font = bold_font
        disagree_section += 1
        if not disagreement_rows:
            ws.cell(disagree_section, 1, "No items are marked as 'Disagree'.")
            disagree_section += 1
        else:
            for drow in disagreement_rows:
                # drow is something like [veridoc, di_num, cdrl_subtitle, gov_status]
                veridoc = drow[0]
                di_num = drow[1]

                ws.cell(disagree_section, 1, f"VeriDoc: {veridoc}, DI Number: {di_num}")
                disagree_section += 1

                # Optionally see if you can find relevant lines in gov_lines
                related_gov_lines = [gl for gl in g_lines if di_num in gl]
                if related_gov_lines:
                    for gl in related_gov_lines:
                        ws.cell(disagree_section, 1, f"Gov comment: {gl}")
                        disagree_section += 1
                else:
                    ws.cell(disagree_section, 1, "No specific government comment found for this line.")
                    disagree_section += 1

                disagree_section += 1  # Spacing

        # 9. Optional "form fields" replaced with placeholders
        # Excel doesn't have fillable forms in the same sense, but you can label cells
        form_start = disagree_section + 2
        ws.cell(form_start, 1, "PWG assigned to resolve disagreement:")
        ws.cell(form_start, 2, "")  # user can fill in
        form_start += 1

        ws.cell(form_start, 1, "USCG POC assigned to resolve disagreement:")
        ws.cell(form_start, 2, "")
        form_start += 1

        ws.cell(form_start, 1, "BIRDON POC assigned to resolve disagreement:")
        ws.cell(form_start, 2, "")
        form_start += 2

        ws.cell(form_start, 1, "General Comments:")
        ws.cell(form_start+1, 1, "Enter general comments in the cell below:")
        form_start += 3

        # 10. Finally, save the workbook
        # Decide filename
        filename = f"Disagreement Report - WCC-SPEC-{doors_spec_id}.xlsx"
        # If user selected a folder
        if self.output_folder:
            xlsx_path = os.path.join(self.output_folder, filename)
        else:
            xlsx_path = filename

        try:
            wb.save(xlsx_path)
            if show_popup:
                from tkinter import messagebox
                messagebox.showinfo("Excel Generated", f"Excel saved as {xlsx_path}")
        except Exception as e:
            if show_popup:
                messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")




###This Creates the Diagreement report PDF's 

    def generate_pdf_for_current(self, show_popup=True):
        from datetime import datetime
        import os
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import Paragraph, Table, TableStyle
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.pdfbase import acroform
        from textwrap import wrap
        from tkinter import messagebox

        # 1. Gather the same data as before
        item = self.disagreement_items[self.current_index]
        row_index = item['row_index']

        # Update UI and load current row data
        self.app.current_row = row_index
        self.app.update_ui_after_navigation()

        # Get DOORS SPEC ID
        doors_spec_id = self.app.df.iloc[row_index, 0]
        if pd.isna(doors_spec_id):
            doors_spec_id = ""
        elif not isinstance(doors_spec_id, str):
            doors_spec_id = str(doors_spec_id)

        # Get specification text
        spec_text = self.app.spec_text_box.get("1.0", "end").strip()

        # Contractor Proposed Change Comment History
        contractor_history_content = ""
        if len(self.app.df.columns) > 8:
            val = self.app.df.iloc[row_index, 8]
            if pd.isna(val):
                val = ""
            contractor_history_content = str(val)

        # Government Adjudication Comment History
        gov_history_content = ""
        if len(self.app.df.columns) > 9:
            val = self.app.df.iloc[row_index, 9]
            if pd.isna(val):
                val = ""
            gov_history_content = str(val)

        # Page setup
        width, height = letter
        left_margin = 72
        right_margin = 72
        top_margin = 50
        bottom_margin = 72
        usable_width = width - (left_margin + right_margin)

        # Collect table data from self.app.table
        items = self.app.table.get_children()
        breakdown_data = [
            ["VeriDoc Number", "DI Number", "CDRL Subtitle", "Government Assessed Status"]
        ]
        for line_id in items:
            values = self.app.table.item(line_id, 'values')
            # (VeriDoc, DI Number, CDRL Subtitle, Object Status, Contractor Status, Government Status)
            breakdown_data.append([values[0], values[1], values[2], values[5]])

        # Count agreements and disagreements
        agree_count = 0
        disagree_count = 0
        for i in range(1, len(breakdown_data)):
            gov_status = str(breakdown_data[i][3]).strip().lower()
            if gov_status == "agree":
                agree_count += 1
            elif gov_status == "disagree":
                disagree_count += 1

        # Extract disagreeing items
        disagreement_rows = []
        for i in range(1, len(breakdown_data)):
            gov_status = str(breakdown_data[i][3]).strip().lower()
            if gov_status == "disagree":
                disagreement_rows.append(breakdown_data[i])

        # Prepare the PDF
        filename = f"Disagreement Report - WCC-SPEC-{doors_spec_id}.pdf"
        pdf_path = (
            os.path.join(self.output_folder, filename) if self.output_folder else filename
        )
        c = canvas.Canvas(pdf_path, pagesize=letter)
        form = acroform.AcroForm(c)

        # Some helper styling
        styles = getSampleStyleSheet()
        styleN = styles["Normal"]

        def wrap_text_to_pdf(c, text, x, y, max_width):
            """Helper to wrap lines at a certain width."""
            chars_per_line = int(max_width / 6)  # Approx. for 12-pt text
            wrapped_lines = wrap(text, width=chars_per_line)
            for wline in wrapped_lines:
                if y < bottom_margin:
                    c.showPage()
                    c.setFont("Helvetica", 12)
                    y = height - top_margin
                c.drawString(x, y, wline)
                y -= 14
            return y

        # Start writing content
        c.setFont("Helvetica", 8)
        y = height - top_margin
        current_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.drawString(left_margin, y, f"Date/Time: {current_dt}")
        y -= 10
        c.drawString(left_margin, y, "Contract: 70Z02323D93270001")
        y -= 10

        distribution_text = (
            "DISTRIBUTION STATEMENT D: DISTRIBUTION AUTHORIZED TO DHS/CG/DOD AND THEIR "
            "CONTRACTORS ONLY DUE TO ADMINISTRATIVE OR OPERATIONAL USE (5 OCT 2022). "
            "OTHER REQUESTS SHALL BE REFERRED TO COMMANDANT (CG-9327)."
        )

        destruction_text = (
            "DESTRUCTION NOTICE: DESTROY THIS DOCUMENT BY ANY METHOD THAT WILL "
            "PREVENT DISCLOSURE OF CONTENTS OR RECONSTRUCTION OF THE DOCUMENT."
        )

        # Wrap distribution text
        y = wrap_text_to_pdf(c, distribution_text, left_margin, y, usable_width)
        y -= 10
        # Wrap destruction text
        y = wrap_text_to_pdf(c, destruction_text, left_margin, y, usable_width)

        # DOORS SPEC ID summary table
        from reportlab.platypus import Table, TableStyle
        from reportlab.lib import colors

        id_table_data = [
            ["DOORS SPEC ID", "Excel Row", "Total Agreements", "Total Disagreements"],
            [doors_spec_id, str(row_index + 2), str(agree_count), str(disagree_count)],
        ]

        id_table = Table(id_table_data, colWidths=[130, 60, 100, 120])
        id_style = TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
        id_table.setStyle(id_style)
        w_id, h_id = id_table.wrap(usable_width, 50)

        c.setFont("Helvetica", 12)
        if y - h_id < bottom_margin:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = height - top_margin

        id_table.drawOn(c, left_margin, y - h_id)
        y = y - h_id - 30

        # Specification Text
        c.drawString(left_margin, y, "Specification Text:")
        y -= 20
        y = wrap_text_to_pdf(c, spec_text, left_margin, y, usable_width)

        # ------------------------------
        # Create a two-column "Comments" table for contractor vs. government
        # ------------------------------
        y -= 30
        c.drawString(left_margin, y, "Comments:")
        y -= 20

        # Split lines and remove blank/underscore lines
        contractor_lines = [
            line.strip()
            for line in contractor_history_content.split("\n")
            if line.strip() and "_____" not in line
        ]
        gov_lines = [
            line.strip()
            for line in gov_history_content.split("\n")
            if line.strip() and "_____" not in line
        ]

        comments_data = [
            [
                "Contractor Proposed Change Comment History",
                "Government Adjudication Comment History",
            ]
        ]

        max_len = max(len(contractor_lines), len(gov_lines))
        for i in range(max_len):
            c_text = contractor_lines[i] if i < len(contractor_lines) else ""
            g_text = gov_lines[i] if i < len(gov_lines) else ""
            # Wrap them as Paragraphs for safer text wrapping in the table
            comments_data.append([Paragraph(c_text, styleN), Paragraph(g_text, styleN)])

        # Build the table
        comments_table = Table(comments_data, colWidths=[usable_width / 2, usable_width / 2])
        comments_table_style = TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ]
        )
        comments_table.setStyle(comments_table_style)

        # Compute space needed, do page break if needed
        w_comments, h_comments = comments_table.wrap(usable_width, y)
        if y - h_comments < bottom_margin:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = height - top_margin
        comments_table.drawOn(c, left_margin, y - h_comments)
        y -= h_comments + 20

        # Breakdown table
        approx_char_width = 6
        max_lengths = [0, 0, 0, 0]
        for row in breakdown_data:
            for j, val in enumerate(row):
                length = len(str(val))
                if length > max_lengths[j]:
                    max_lengths[j] = length
        column_widths = [length * approx_char_width for length in max_lengths]

        # Cap the CDRL column at 200 px
        if column_widths[2] > 200:
            column_widths[2] = 200

        # Convert the CDRL column to a Paragraph for wrapping
        for i in range(1, len(breakdown_data)):
            cdrl_text = breakdown_data[i][2]
            breakdown_data[i][2] = Paragraph(cdrl_text, styleN)

        t = Table(breakdown_data, colWidths=column_widths)
        style = TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )

        # Color-coding for Government Status
        for i in range(1, len(breakdown_data)):
            gov_status = str(breakdown_data[i][3]).strip().lower()
            if gov_status == "disagree":
                style.add("BACKGROUND", (3, i), (3, i), colors.red)
                style.add("TEXTCOLOR", (3, i), (3, i), colors.white)
            elif gov_status == "agree":
                style.add("BACKGROUND", (3, i), (3, i), colors.green)
                style.add("TEXTCOLOR", (3, i), (3, i), colors.white)
            elif gov_status == "pending review":
                style.add("BACKGROUND", (3, i), (3, i), colors.yellow)
                style.add("TEXTCOLOR", (3, i), (3, i), colors.black)

        t.setStyle(style)

        w, h = t.wrap(usable_width, 400)
        if y - h < bottom_margin:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = height - top_margin
        t.drawOn(c, left_margin, y - h)
        y = y - h - 60

        # Disagreement Comments
        if disagreement_rows:
            c.drawString(left_margin, y, "Disagreement Comments:")
            y -= 20

            for d_row in disagreement_rows:
                veridoc = str(d_row[0]).strip()
                di_num = str(d_row[1]).strip()

                if y < bottom_margin:
                    c.showPage()
                    c.setFont("Helvetica", 12)
                    y = height - top_margin

                c.line(left_margin, y, width - right_margin, y)  # horizontal line
                y -= 10
                c.drawString(left_margin, y, f"VeriDoc: {veridoc}")
                y -= 14
                c.drawString(left_margin, y, f"DI Number: {di_num}")
                y -= (14 * 2)

                if y < bottom_margin:
                    c.showPage()
                    c.setFont("Helvetica", 12)
                    y = height - top_margin

                c.drawString(left_margin, y, "Government Comments:")
                y -= 14

                # Filter relevant lines that mention this di_num
                related_gov_lines = [gl for gl in gov_lines if di_num in gl]
                if not related_gov_lines:
                    c.setFillColor(colors.red)
                    c.drawString(
                        left_margin,
                        y,
                        "No specific government comments related to this item.",
                    )
                    c.setFillColor(colors.black)
                    y -= 14
                else:
                    for gl in related_gov_lines:
                        y = wrap_text_to_pdf(c, gl, left_margin, y, usable_width)
        else:
            # If no disagreements
            c.drawString(left_margin, y, "No items are marked 'Disagree' in this row.")
            y -= 20

        # -----------------------------------------------------------------------
        # General Comments text box
        # -----------------------------------------------------------------------
        single_line_height = 20
        if y < 200:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = height - top_margin

        c.drawString(left_margin, y, "General Comments:")
        general_box_width = usable_width
        general_box_height = 80
        form.textfield(
            name="generalComments",
            tooltip="General Comments",
            x=left_margin,
            y=y - general_box_height,
            width=general_box_width,
            height=general_box_height,
            borderStyle="inset",
            borderWidth=1,
            fillColor=colors.white,
        )
        y -= (general_box_height + 40)

        # -----------------------------------------------------------------------
        # Options for Birdon (title) + Three Checkboxes
        # -----------------------------------------------------------------------
        c.drawString(left_margin, y, "Options for Birdon")
        y -= 20

        # --- 1) Disagreement Not Clear...
        c.drawString(left_margin, y, "Disagreement Not Clear Send to USCG for Clarification")
        form.checkbox(
            name="disagreementNotClear",
            tooltip="Check if the disagreement is not clear and needs USCG clarification",
            x=left_margin + 420,
            y=y - 8,
            size=12,
            borderWidth=1,
            checked=False,
            buttonStyle="check",
        )
        y -= 20

        # --- 2) Disagreement can be resolved...
        c.drawString(
            left_margin,
            y,
            "Disagreement can be resolved with updated locations flag for RTVM",
        )
        form.checkbox(
            name="disagreementResolvedLocations",
            tooltip="Check if the disagreement can be resolved with updated locations in RTVM",
            x=left_margin + 420,
            y=y - 8,
            size=12,
            borderWidth=1,
            checked=False,
            buttonStyle="check",
        )
        y -= 20

        # --- 3) Disagreement can not be resolved...
        c.drawString(left_margin, y, "Disagreement can not be resolved at this time")
        form.checkbox(
            name="disagreementNotResolved",
            tooltip="Check if the disagreement cannot be resolved at this time",
            x=left_margin + 420,
            y=y - 8,
            size=12,
            borderWidth=1,
            checked=False,
            buttonStyle="check",
        )
        y -= 40

        # -----------------------------------------------------------------------
        # USCG Responce (title) + a text box
        # -----------------------------------------------------------------------
        c.drawString(left_margin, y, "USCG Responce:")
        uscg_box_height = 60
        form.textfield(
            name="uscgResponceBox",
            tooltip="Enter USCG Responce",
            x=left_margin,
            y=y - uscg_box_height,
            width=usable_width,
            height=uscg_box_height,
            borderStyle="inset",
            borderWidth=1,
            fillColor=colors.white,
        )
        y -= (uscg_box_height + 40)

        # -----------------------------------------------------------------------
        # USCG Signature (approved to disregard disagreement) + Date of Resolution
        # -----------------------------------------------------------------------
        c.drawString(
            left_margin, y, "USCG Signature (approved to disregard disagreement):"
        )
        form.textfield(
            name="uscgSignature",
            tooltip="USCG Signature",
            x=left_margin + 364,
            y=y - 12,
            width=150,
            height=single_line_height,
            borderStyle="inset",
            borderWidth=1,
            fillColor=colors.white,
        )
        y -= 40

        c.drawString(left_margin, y, "Date of Resolution:")
        form.textfield(
            name="resolutionDate",
            tooltip="Date of Resolution",
            x=left_margin + 228,
            y=y - 12,
            width=200,
            height=single_line_height,
            borderStyle="inset",
            borderWidth=1,
            fillColor=colors.white,
        )
        y -= 40

        # End the page
        c.showPage()
        c.save()

        if show_popup:
            messagebox.showinfo("PDF Generated", f"PDF saved as {pdf_path}")





    def wrap_text_to_pdf(self, c, text, x, y, max_width):
        from textwrap import wrap
        lines = wrap(text, width=80)
        for line in lines:
            c.drawString(x, y, line)
            y -= 14
        return y

    def create_all_reports(self):
        if not self.disagreement_items:
            messagebox.showinfo("No Disagreements", "No disagreement items available.")
            return

        # When generating all reports, do not show the PDF generated popup each time.
        for i in range(len(self.disagreement_items)):
            self.current_index = i
            self.generate_pdf_for_current(show_popup=False)

        # After all are done, you can show a single message
        messagebox.showinfo("All Reports Created", "All disagreement reports have been generated.")

    #################################################################################################END END END ##########################################################

if __name__ == "__main__":
    root = tk.Tk()
    app = RTVMApp(root)
    root.mainloop()
    
