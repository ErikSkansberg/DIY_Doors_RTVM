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
from tkinter import ttk, messagebox, filedialog
import re
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
                data = row[5]  # Column F
                if pd.isna(data):
                    data = ""
                elif not isinstance(data, str):
                    data = str(data)
                # Use regex to extract statuses and VeriDoc Number
                matches = re.finditer(
                    r'Object Identifier:\s*(?P<obj_id>WCC-VERI-DOC-\d+).*?Object Status:\s*(?P<object_status>.*?)\n.*?Contractor Assessed Status:\s*(?P<contractor_status>.*?)\n.*?Government Assessed Status:\s*(?P<government_status>.*?)\n',
                    data, re.DOTALL
                )
                found_match = False
                for match in matches:
                    found_match = True
                    obj_id = match.group('obj_id').strip()
                    object_status = match.group('object_status').strip()
                    contractor_status = match.group('contractor_status').strip()
                    government_status = match.group('government_status').strip()
                    self.unique_object_statuses.add(object_status)
                    self.unique_contractor_statuses.add(contractor_status)
                    self.unique_government_statuses.add(government_status)
                    self.status_data.append({
                        'row_index': index,
                        'veridoc_number': obj_id,
                        'object_status': object_status,
                        'contractor_status': contractor_status,
                        'government_status': government_status
                    })
                # If no matches found, add an entry with empty statuses
                if not found_match:
                    self.status_data.append({
                        'row_index': index,
                        'veridoc_number': '',
                        'object_status': '',
                        'contractor_status': '',
                        'government_status': ''
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
        # Save comments before navigating away
        self.save_comments_to_excel()
        # Continue with navigation
        if hasattr(self, 'filtered_row_indices') and self.filtered_row_indices:
            max_index = len(self.filtered_row_indices) - 1
            if direction == 'up':
                self.current_filtered_index = max(
                    self.current_filtered_index - 1, 0)
            elif direction == 'down':
                self.current_filtered_index = min(
                    self.current_filtered_index + 1, max_index)
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
            messagebox.showwarning(
                "No Data", "No data available to extract.")
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

        # Extract and populate the table with all DI Numbers
        matches = re.finditer(
            r'Object Identifier:\s*(?P<obj_id>WCC-VERI-DOC-(?P<veridoc_num>\d+)).*?DI Number:\s*(?P<di_num>\d+-\d+).*?CDRL Subtitle:\s*(?P<cdrl_subtitle>.*?)\n.*?Object Status:\s*(?P<object_status>.*?)\n.*?Contractor Assessed Status:\s*(?P<contractor_status>.*?)\n.*?Government Assessed Status:\s*(?P<government_status>.*?)\n',
            data, re.DOTALL
        )

        count = 0  # Counter for the number of rows
        for match in matches:
            obj_id = match.group('obj_id')
            veridoc_num = match.group('veridoc_num')
            di_num = match.group('di_num')
            cdrl_subtitle = match.group('cdrl_subtitle').strip()
            object_status = match.group('object_status').strip()
            contractor_status = match.group('contractor_status').strip()
            government_status = match.group('government_status').strip()

            # Insert the item into the data display table
            item_id = self.table.insert("", "end", values=(obj_id, di_num, cdrl_subtitle, object_status,
                                                           contractor_status, government_status))

            # Get existing comment if any
            if obj_id in comments_dict:
                comment_text = f"{obj_id} - {comments_dict[obj_id]}"
            else:
                # Prepopulate the comment with the Object Identifier followed by " - "
                comment_text = f"{obj_id} - "

            self.comment_table.insert("", "end", values=(comment_text,))

            # Apply conditional formatting
            self.apply_conditional_formatting(
                item_id, contractor_status, government_status)

            # Check if Object Status is "DEPRECIATED"
            if object_status.upper() == "DEPRECIATED":
                messagebox.showwarning("Warning",
                                       "The Object Status is marked as DEPRECIATED \n A CDRL that the government concurs should be removed from the SPEC line item and is pending the full removal of the object from DOORs. No DLOC edits will be accepted for a VERI-DOC in this status.")

            count += 1  # Increment the counter

        # Adjust the table height based on the number of rows
        self.adjust_table_height(count)


        count = 0  # Counter for the number of rows
        for match in matches:
            obj_id = match.group('obj_id')
            veridoc_num = match.group('veridoc_num')
            di_num = match.group('di_num')
            cdrl_subtitle = match.group('cdrl_subtitle').strip()
            object_status = match.group('object_status').strip()
            contractor_status = match.group('contractor_status').strip()
            government_status = match.group('government_status').strip()

            # Insert the item into the data display table
            item_id = self.table.insert("", "end", values=(obj_id, di_num, cdrl_subtitle, object_status,
                                                           contractor_status, government_status))
            # Prepopulate the comment with the Object Identifier followed by " - "
            comment_initial_text = f"{obj_id} - "
            self.comment_table.insert("", "end", values=(comment_initial_text,))

            # Apply conditional formatting
            self.apply_conditional_formatting(
                item_id, contractor_status, government_status)

            # Check if Object Status is "DEPRECIATED"
            if object_status.upper() == "DEPRECIATED":
                messagebox.showwarning("Warning",
                                       "The Object Status is marked as DEPRECIATED \n A CDRL that the government concurs should be removed from the SPEC line item and is pending the full removal of the object from DOORs. No DLOC edits will be accepted for a VERI-DOC in this status.")

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

        # Create a new window
        self.pie_chart_window = tk.Toplevel(self.root)
        self.pie_chart_window.title("Pie Charts")

        # Create filter frame
        filter_frame = tk.Frame(self.pie_chart_window)
        filter_frame.pack(side=tk.TOP, fill=tk.X)

        # Object Status Filter
        object_label = tk.Label(filter_frame, text="Object Status")
        object_label.pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_object_status_var = tk.StringVar()
        self.pie_object_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_object_status_var,
            values=["Any"] + sorted(self.unique_object_statuses),
            state="readonly",
            width=15
        )
        self.pie_object_status_dropdown.pack(side=tk.LEFT)
        self.pie_object_status_var.set("Any")  # Default value

        # Contractor Assessed Status Filter
        contractor_label = tk.Label(filter_frame, text="Contractor Assessed Status")
        contractor_label.pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_contractor_status_var = tk.StringVar()
        self.pie_contractor_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_contractor_status_var,
            values=["Any"] + sorted(self.unique_contractor_statuses),
            state="readonly",
            width=15
        )
        self.pie_contractor_status_dropdown.pack(side=tk.LEFT)
        self.pie_contractor_status_var.set("Any")  # Default value

        # Government Assessed Status Filter
        government_label = tk.Label(filter_frame, text="Government Assessed Status")
        government_label.pack(side=tk.LEFT, padx=5, pady=5)
        self.pie_government_status_var = tk.StringVar()
        self.pie_government_status_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.pie_government_status_var,
            values=["Any"] + sorted(self.unique_government_statuses),
            state="readonly",
            width=15
        )
        self.pie_government_status_dropdown.pack(side=tk.LEFT)
        self.pie_government_status_var.set("Any")  # Default value

        # Update Charts Button
        update_button = tk.Button(filter_frame, text="Update Charts", command=self.update_pie_charts)
        update_button.pack(side=tk.LEFT, padx=10)

        # Create frames for the pie charts
        charts_frame = tk.Frame(self.pie_chart_window)
        charts_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Object Status Pie Chart
        self.object_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.object_status_canvas = FigureCanvasTkAgg(self.object_status_fig, charts_frame)
        self.object_status_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Contractor Assessed Status Pie Chart
        self.contractor_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.contractor_status_canvas = FigureCanvasTkAgg(self.contractor_status_fig, charts_frame)
        self.contractor_status_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Government Assessed Status Pie Chart
        self.government_status_fig = plt.Figure(figsize=(4, 4), dpi=100)
        self.government_status_canvas = FigureCanvasTkAgg(self.government_status_fig, charts_frame)
        self.government_status_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a frame for the counts table
        counts_frame = tk.Frame(self.pie_chart_window)
        counts_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Create the counts table with an additional column for Total Count
        self.counts_table = ttk.Treeview(counts_frame, columns=("Status", "Count", "Total Count"), show="headings")
        self.counts_table.heading("Status", text="Status")
        self.counts_table.heading("Count", text="Count")
        self.counts_table.heading("Total Count", text="Total Count")
        self.counts_table.pack(fill=tk.BOTH, expand=True)

        # Compute total counts before updating pie charts
        self.compute_total_counts()

        # Initially update the pie charts and counts table
        self.update_pie_charts()


    def update_pie_charts(self):
        # Get selected statuses
        selected_object_status = self.pie_object_status_var.get()
        selected_contractor_status = self.pie_contractor_status_var.get()
        selected_government_status = self.pie_government_status_var.get()

        # Treat "Any" as no filter
        if selected_object_status == "Any":
            selected_object_status = ''
        if selected_contractor_status == "Any":
            selected_contractor_status = ''
        if selected_government_status == "Any":
            selected_government_status = ''

        # Filter the status data
        filtered_items = []
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
                filtered_items.append(item)

        # Collect statuses from filtered items
        filtered_statuses = {
            'object_status': [item['object_status'] for item in filtered_items],
            'contractor_status': [item['contractor_status'] for item in filtered_items],
            'government_status': [item['government_status'] for item in filtered_items]
        }

        # Generate pie charts
        self.plot_pie_chart(self.object_status_fig, filtered_statuses['object_status'], 'Object Status')
        self.plot_pie_chart(self.contractor_status_fig, filtered_statuses['contractor_status'], 'Contractor Assessed Status')
        self.plot_pie_chart(self.government_status_fig, filtered_statuses['government_status'], 'Government Assessed Status')

        # Draw the canvases
        self.object_status_canvas.draw()
        self.contractor_status_canvas.draw()
        self.government_status_canvas.draw()

        # Update the counts table
        self.update_counts_table(filtered_statuses)



    def plot_pie_chart(self, fig, data_list, title):
        # Clear the figure
        fig.clear()

        # Count the occurrences
        from collections import Counter
        counter = Counter(data_list)
        labels = list(counter.keys())
        sizes = list(counter.values())

        # Handle empty data
        if not sizes:
            labels = ['No Data']
            sizes = [1]

        # Create pie chart
        ax = fig.add_subplot(111)
        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
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
        





if __name__ == "__main__":
    root = tk.Tk()
    app = RTVMApp(root)
    root.mainloop()