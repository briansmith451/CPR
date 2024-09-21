import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import pyodbc
import pandas as pd
import re
import random
import threading
import logging

# ------------------------- Setup Logging -------------------------
logging.basicConfig(
    filename='app.log',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)

# ------------------------- Helper Functions -------------------------
def is_valid_email(email):
    """Validate the email address using regex."""
    regex = r'^[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,63}$'
    return re.match(regex, email, re.IGNORECASE) is not None

def is_valid_phone(phone):
    """
    Validate the phone number.
    - If 'DSN' appears anywhere in the phone number, accept it.
    - US numbers: Exactly 10 digits after removing non-digit characters.
    - International numbers: Accept numbers starting with specified country codes.
    """
    if not phone:
        return False  # Empty phone number is considered invalid

    phone = phone.strip()

    # If 'DSN' appears anywhere in the phone number, accept it
    if 'DSN' in phone.upper():
        return True

    # Remove parentheses
    phone = phone.replace('(', '').replace(')', '')

    # Prepare to check if phone number starts with allowed country codes
    cleaned_phone_for_code_check = re.sub(r'[^\d+]', '', phone)

    allowed_country_codes = ['+33', '33', '+44', '44', '+49', '49', '+61', '61', '+64', '64']

    # Check if the phone number starts with any of the allowed country codes
    is_international = any(cleaned_phone_for_code_check.startswith(code) for code in allowed_country_codes)

    if is_international:
        # For international numbers starting with specified country codes, accept as valid
        return True
    else:
        # For other numbers, remove all non-digit characters
        digits_only = re.sub(r'\D', '', phone)

        # Check if length is exactly 10 digits (US number)
        return len(digits_only) == 10

def clean_phone_number(phone_number):
    """
    Clean the phone number according to the specified rules.

    - If 'DSN' appears anywhere, keep the phone number as is after trimming whitespace.
    - Remove parentheses around area code.
    - For numbers starting with specified country codes:
        - Remove parentheses.
        - Do not remove spaces or dashes.
    - For US numbers:
        - Remove spaces.
        - Remove all non-digit characters except '-'.
        - Ensure format is XXX-XXX-XXXX.
    """
    allowed_country_codes = ['+33', '33', '+44', '44', '+49', '49', '+61', '61', '+64', '64']

    phone_number = str(phone_number).strip()

    # If 'DSN' appears, return the phone number as is (after trimming whitespace)
    if 'DSN' in phone_number.upper():
        return phone_number

    # Remove parentheses
    phone_number = phone_number.replace('(', '').replace(')', '')

    # Prepare to check if phone number starts with allowed country codes
    cleaned_phone_for_code_check = re.sub(r'[^\d+]', '', phone_number)

    # Check if the phone number starts with any of the allowed country codes
    is_international = any(cleaned_phone_for_code_check.startswith(code) for code in allowed_country_codes)

    if is_international:
        # For international numbers, remove parentheses but keep spaces and dashes
        pass  # No further action needed
    else:
        # For US numbers
        # Remove spaces
        phone_number = phone_number.replace(' ', '')

        # Remove all non-digit characters except '-'
        phone_number = re.sub(r'[^0-9-]', '', phone_number)

        # Remove existing dashes to reformat
        digits_only = re.sub(r'\D', '', phone_number)

        if len(digits_only) == 10:
            # Format as XXX-XXX-XXXX
            phone_number = f"{digits_only[:3]}-{digits_only[3:6]}-{digits_only[6:]}"
        elif len(digits_only) == 7:
            # Format as XXX-XXXX
            phone_number = f"{digits_only[:3]}-{digits_only[3:]}"
        else:
            # Cannot format properly, keep digits only
            phone_number = digits_only

    return phone_number

def clean_data(value, data_type):
    """Clean data based on the expected data type."""
    if pd.isna(value):
        if data_type == str:
            return ""
        elif data_type in [float, int]:
            return 0
    try:
        return data_type(value)
    except ValueError:
        return data_type() if data_type in [float, int] else ""
    

class CombinedApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("CDAO Personnel Data Manager")
        self.base_width = 720  # Set new width to 720 pixels
        self.base_height = 920
        self.geometry(f"{self.base_width}x{self.base_height}")

        self.iconbitmap(r"C:\Users\Brian Smith\Desktop\favicon.ico")

        # Get the screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate position to center the window
        x = (screen_width // 2) - (self.base_width // 2)
        y = (screen_height // 2) - (self.base_height // 2) - 30  # Adjust 30 pixels higher

        # Set the window size and position
        self.geometry(f"{self.base_width}x{self.base_height}+{x}+{y}")

        self.base_font_size = 10
        self.zoom_step = 1
        self.current_font_size = self.base_font_size

        # Create the notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both')

        self.data_finder_tab = ttk.Frame(self.notebook)
        self.data_writer_tab = ttk.Frame(self.notebook)
        self.delete_data_tab = ttk.Frame(self.notebook)

        # Add tabs in the desired order: Data Finder -> Data Writer -> Edit Data
        self.notebook.add(self.data_finder_tab, text="Data Finder")
        self.notebook.add(self.data_writer_tab, text="Data Writer")
        self.notebook.add(self.delete_data_tab, text="Edit Data")

        # Initialize the DatabaseWriter and other tabs, passing self for app reference
        self.data_writer = DatabaseWriter(self.data_writer_tab, self)
        self.data_finder = PersonnelDataFinder(self.data_finder_tab, self)
        self.delete_data = EditDataTab(self.delete_data_tab, self)

        self.bind("<Control-MouseWheel>", self.zoom)

    def zoom(self, event):
        if event.delta > 0:
            self.current_font_size += self.zoom_step
        else:
            self.current_font_size -= self.zoom_step

        self.current_font_size = max(8, min(self.current_font_size, 30))

        self.data_writer.update_font_size(self.current_font_size)
        self.data_finder.update_font_size(self.current_font_size)
        self.delete_data.update_font_size(self.current_font_size)

        self.update_window_size()

    def update_window_size(self):
        scale_factor = self.current_font_size / self.base_font_size
        new_width = int(self.base_width * scale_factor)
        new_height = int(self.base_height * scale_factor)

        # Recalculate the position for the centered window
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (new_width // 2)
        y = (screen_height // 2) - (new_height // 2)

        self.geometry(f"{new_width}x{new_height}+{x}+{y}")


class DatabaseWriter:

    def __init__(self, frame, app):
        self.frame = frame
        self.app = app
        self.current_row = 0  # Initialize row counter

        # Create a canvas and scrollbar
        self.canvas = tk.Canvas(self.frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a scrollable frame inside the canvas
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Bind scrolling events
        self.frame.bind_all("<MouseWheel>", self.on_mouse_wheel)
        self.frame.bind_all("<Button-4>", self.on_mouse_wheel)  # For Linux
        self.frame.bind_all("<Button-5>", self.on_mouse_wheel)  # For Linux

        # Configure the scrollable frame to adjust the scroll region
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)

        # Initialize the UI
        self.create_writer_ui()

        # Connect to the database
        try:
            self.conn = pyodbc.connect(
                "Driver={ODBC Driver 18 for SQL Server};"
                "Server=tcp:idsfuturetech-sql-server.database.windows.net,1433;"
                "Database=ids.personnel;"
                "Uid=sqladmin;"
                "Pwd=YourSecurePassword123!;"
                "Encrypt=yes;"
                "TrustServerCertificate=no;"
                "Connection Timeout=30;"
            )
            self.cursor = self.conn.cursor()
            logging.info("Database connection established.")
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to connect to the database: {e}")
            logging.error(f"Database connection failed: {e}")

    def on_frame_configure(self, event):
        """Reset the scroll region to encompass the inner frame."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_mouse_wheel(self, event):
        """Handle mouse wheel scrolling."""
        if event.delta:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        elif event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def create_writer_ui(self):
        """Create the UI components for data entry."""
        self.font = ("Segoe UI", self.app.current_font_size)

        # Load and display the logo at the top
        try:
            logo_image = Image.open(r"C:\Users\Brian Smith\Desktop\logo.png")
            logo_image = logo_image.resize((700, int(700 * logo_image.height / logo_image.width)), Image.LANCZOS)
            logo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(self.scrollable_frame, image=logo)
            logo_label.image = logo  # Keep a reference to avoid garbage collection
            logo_label.pack(pady=10)
        except Exception as e:
            logging.error(f"Error loading logo: {e}")
            print(f"Error loading logo: {e}")

        # Create the form frame
        self.form_frame = ttk.Frame(self.scrollable_frame)
        self.form_frame.pack(pady=10, padx=110)

        # Create form entries using row counter
        self.entry_primary_govt_org = self.create_row(self.form_frame, "Primary Government Organization:", self.font)
        self.entry_directorate = self.create_row(self.form_frame, "Directorate:", self.font)
        self.entry_dept_div_branch = self.create_row(self.form_frame, "Dept/Div/Branch:", self.font)
        self.entry_secondary_govt_org = self.create_row(self.form_frame, "Secondary Gov't Org:", self.font)
        self.entry_civilian_company = self.create_row(self.form_frame, "Civilian Company:", self.font)
        self.entry_first_name = self.create_row(self.form_frame, "First Name:", self.font)
        self.entry_last_name = self.create_row(self.form_frame, "Last Name:", self.font)
        self.entry_callsign_nickname = self.create_row(self.form_frame, "Callsign/Nickname:", self.font)
        self.entry_rank = self.create_row(self.form_frame, "Rank:", self.font)
        self.entry_duty_position = self.create_row(self.form_frame, "Duty Position:", self.font)
        self.entry_commercial_number = self.create_row(self.form_frame, "Commercial #:", self.font)
        self.entry_cell_number = self.create_row(self.form_frame, "Cell #:", self.font)
        self.entry_svoip = self.create_row(self.form_frame, "SVOIP:", self.font)
        self.entry_company_email = self.create_row(self.form_frame, "Company Email:", self.font)
        self.entry_nipr_email = self.create_row(self.form_frame, "NIPR Email:", self.font)
        self.entry_sipr_email = self.create_row(self.form_frame, "SIPR Email:", self.font)

        # Create country dropdown
        self.create_country_dropdown(self.form_frame, self.font)

        # Button frame for buttons to be aligned evenly
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(pady=20)

        # Add Record button
        self.add_button = ttk.Button(button_frame, text="Add Record", command=self.add_record, bootstyle="success")
        self.add_button.grid(row=0, column=0, padx=10)

        # Clear Record button
        self.clear_button = ttk.Button(button_frame, text="Clear Record", command=self.clear_fields, bootstyle="warning")
        self.clear_button.grid(row=0, column=1, padx=10)

        # Load Excel button placed to the right of Clear Record button
        self.load_excel_button = ttk.Button(button_frame, text="Load Excel Data", command=self.load_excel_file, bootstyle="primary")
        self.load_excel_button.grid(row=0, column=2, padx=10)

    def create_row(self, parent, label_text, font):
        """Create a label and entry widget in the form."""
        label = ttk.Label(parent, text=label_text, font=font)
        label.grid(row=self.current_row, column=0, sticky='e', padx=5, pady=5)
        entry = ttk.Entry(parent, font=font)
        entry.grid(row=self.current_row, column=1, padx=5, pady=5, sticky='w')
        self.current_row += 1  # Increment row counter
        return entry

    def create_country_dropdown(self, parent, font):
        """Create a dropdown for selecting the country."""
        preferred_countries = [
            "United States", 
            "Australia", 
            "France", 
            "New Zealand", 
            "United Kingdom"
        ]

        label = ttk.Label(parent, text="Country:", font=font)
        label.grid(row=self.current_row, column=0, sticky='e', padx=5, pady=5)

        self.country_var = tk.StringVar()
        self.country_combobox = ttk.Combobox(parent, textvariable=self.country_var, values=preferred_countries, font=font)
        self.country_combobox.grid(row=self.current_row, column=1, padx=5, pady=5, sticky='w')
        self.current_row += 1  # Increment row counter

    def add_record(self):
        """Fetch data from form, validate, and insert into the database."""
        # Fetching and cleaning values from the input fields
        primary_govt_org = clean_data(self.entry_primary_govt_org.get(), str)
        directorate = clean_data(self.entry_directorate.get(), str)
        dept_div_branch = clean_data(self.entry_dept_div_branch.get(), str)
        secondary_govt_org = clean_data(self.entry_secondary_govt_org.get(), str)
        civilian_company = clean_data(self.entry_civilian_company.get(), str)
        first_name = clean_data(self.entry_first_name.get(), str)
        last_name = clean_data(self.entry_last_name.get(), str)
        callsign_nickname = clean_data(self.entry_callsign_nickname.get(), str)
        rank = clean_data(self.entry_rank.get(), str)
        duty_position = clean_data(self.entry_duty_position.get(), str)
        commercial_number = clean_phone_number(clean_data(self.entry_commercial_number.get(), str))
        cell_number = clean_phone_number(clean_data(self.entry_cell_number.get(), str))
        svoip = clean_data(self.entry_svoip.get(), str)
        company_email = clean_data(self.entry_company_email.get(), str).strip()
        nipr_email = clean_data(self.entry_nipr_email.get(), str).strip()
        sipr_email = clean_data(self.entry_sipr_email.get(), str).strip()
        country = clean_data(self.country_combobox.get(), str)

        # Debugging: Log fetched values
        logging.debug(f"Primary Gov't Org: '{primary_govt_org}'")
        logging.debug(f"Civilian Company: '{civilian_company}'")

        # Generate a new 7-digit random ID
        record_id = random.randint(1000000, 9999999)

        # Check if essential fields are filled (first_name and last_name)
        if first_name and last_name:
            # Initialize a list to keep track of invalid fields
            invalid_fields = []

            # Validate Company Email
            if company_email and not is_valid_email(company_email):
                invalid_fields.append("Company Email")

            # Validate NIPR Email
            if nipr_email and not is_valid_email(nipr_email):
                invalid_fields.append("NIPR Email")

            # Validate SIPR Email
            if sipr_email and not is_valid_email(sipr_email):
                invalid_fields.append("SIPR Email")

            # Validate Commercial Number
            if commercial_number and not is_valid_phone(commercial_number):
                invalid_fields.append("Commercial Number")

            # Validate Cell Number
            if cell_number and not is_valid_phone(cell_number):
                invalid_fields.append("Cell Number")

            if invalid_fields:
                # Create an error message listing all invalid fields
                error_message = "The following fields have invalid formats:\n" + "\n".join(invalid_fields)
                messagebox.showerror("Invalid Input Format", error_message)
                
                # Highlight the invalid fields by changing their background color
                for field in invalid_fields:
                    if field == "Company Email":
                        self.entry_company_email.config(background="pink")
                    elif field == "NIPR Email":
                        self.entry_nipr_email.config(background="pink")
                    elif field == "SIPR Email":
                        self.entry_sipr_email.config(background="pink")
                    elif field == "Commercial Number":
                        self.entry_commercial_number.config(background="pink")
                    elif field == "Cell Number":
                        self.entry_cell_number.config(background="pink")
                return  # Exit the method to prevent insertion

            # Reset background colors in case they were previously highlighted
            self.entry_company_email.config(background="white")
            self.entry_nipr_email.config(background="white")
            self.entry_sipr_email.config(background="white")
            self.entry_commercial_number.config(background="white")
            self.entry_cell_number.config(background="white")

            try:
                # Insert the record into the database
                self.cursor.execute("""
                    INSERT INTO employees_v2 
                    ([primary_govt_org], [directorate], [dept_div_branch], [secondary_govt_org], [civilian_company], 
                    [first_name], [last_name], [callsign_nickname], [rank], [duty_position], [commercial_number], 
                    [cell_number], [svoip], [company_email], [nipr_email], [sipr_email], [country], [id])
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    primary_govt_org, directorate, dept_div_branch, secondary_govt_org, civilian_company, 
                    first_name, last_name, callsign_nickname, rank, duty_position, commercial_number, 
                    cell_number, svoip, company_email, nipr_email, sipr_email, country, record_id
                ))
                self.conn.commit()
                logging.info(f"Record added successfully with ID {record_id}")
                messagebox.showinfo("Success", "Record added successfully")
                
                # Optional: Clear fields after successful insertion
                self.clear_fields()

            except pyodbc.IntegrityError as ie:
                logging.error(f"IntegrityError: {ie}")
                messagebox.showerror("Integrity Error", f"Failed to add record: {ie}")
            except pyodbc.ProgrammingError as e:
                logging.error(f"ProgrammingError: {e}")
                messagebox.showerror("Error", f"Failed to add record: {e}")
            except pyodbc.Error as e:
                logging.error(f"Database Error: {e}")
                messagebox.showerror("Database Error", f"Database error: {e}")
            except Exception as ex:
                logging.exception("Unexpected error occurred while adding a record.")
                messagebox.showerror("Error", f"An unexpected error occurred: {ex}")
        else:
            messagebox.showwarning("Input Error", "Please enter both first name and last name.")

    def clear_fields(self):
        """Clear all input fields in the form."""
        self.entry_primary_govt_org.delete(0, 'end')
        self.entry_directorate.delete(0, 'end')
        self.entry_dept_div_branch.delete(0, 'end')
        self.entry_secondary_govt_org.delete(0, 'end')
        self.entry_civilian_company.delete(0, 'end')
        self.entry_first_name.delete(0, 'end')
        self.entry_last_name.delete(0, 'end')
        self.entry_callsign_nickname.delete(0, 'end')
        self.entry_rank.delete(0, 'end')
        self.entry_duty_position.delete(0, 'end')
        self.entry_commercial_number.delete(0, 'end')
        self.entry_cell_number.delete(0, 'end')
        self.entry_svoip.delete(0, 'end')
        self.entry_company_email.delete(0, 'end')
        self.entry_nipr_email.delete(0, 'end')
        self.entry_sipr_email.delete(0, 'end')
        self.country_combobox.set('')

        # Reset background colors of email fields
        self.entry_company_email.config(background="white")
        self.entry_nipr_email.config(background="white")
        self.entry_sipr_email.config(background="white")
        
        # Reset background colors of phone fields
        self.entry_commercial_number.config(background="white")
        self.entry_cell_number.config(background="white")

    def update_font_size(self, font_size):
        """Update the font size of all form widgets."""
        font = ("Segoe UI", font_size)
        for widget in self.form_frame.winfo_children():
            if isinstance(widget, ttk.Label) or isinstance(widget, ttk.Entry):
                widget.configure(font=font)
        # Update other components if necessary

    def load_excel_file(self):
        """Handle the loading of an Excel file."""
        # Open file dialog to select an Excel file
        excel_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Select an Excel File"
        )

        if excel_path:
            # Create a thread to load Excel file without freezing the UI
            threading.Thread(target=self.load_and_insert_excel_data, args=(excel_path,), daemon=True).start()

    def show_progress_bar_window(self):
        """Create and display a progress bar window."""
        # Create a new top-level window for the progress bar
        self.progress_window = tk.Toplevel(self.frame)
        self.progress_window.title("Loading Progress")

        # Set the size of the window
        window_width = 300
        window_height = 150

        # Get the screen's width and height
        screen_width = self.progress_window.winfo_screenwidth()
        screen_height = self.progress_window.winfo_screenheight()

        # Calculate the x and y coordinates to center the window
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        # Set the size and position of the window
        self.progress_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Add a label that instructs the user to wait
        self.progress_label = ttk.Label(self.progress_window, text="Please wait...", font=("Segoe UI", 12))
        self.progress_label.pack(pady=10)

        # Add the progress bar widget to the new window
        self.progress_bar = ttk.Progressbar(self.progress_window, orient="horizontal", length=250, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Add a label to display the percentage completed
        self.percentage_label = ttk.Label(self.progress_window, text="0%", font=("Segoe UI", 10))
        self.percentage_label.pack(pady=5)

    def close_progress_bar_window(self):
        """Close the progress bar window."""
        if hasattr(self, 'progress_window') and self.progress_window:
            self.progress_window.destroy()

    def load_and_insert_excel_data(self, excel_path):
        """Load data from an Excel file and insert it into the database."""
        # Show the progress bar in a new window
        self.show_progress_bar_window()

        # Read the Excel file into a pandas DataFrame
        try:
            df = pd.read_excel(excel_path)
            logging.info(f"Excel file '{excel_path}' loaded successfully.")
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to read the Excel file: {e}")
            logging.error(f"Failed to read Excel file '{excel_path}': {e}")
            self.close_progress_bar_window()
            return

        # Get the total number of rows to use for progress updates
        total_rows = len(df)

        # Lists to track duplicate and invalid records
        duplicate_records = []
        fuzzy_match_records = []

        # Mapping headers from Excel to database columns (remove 'ID')
        header_mapping = {
            "Primary Gov't Org": "primary_govt_org",
            "Directorate": "directorate",
            "Dept/Div/Branch": "dept_div_branch",
            "Secondary Gov't Org": "secondary_govt_org",
            "Civ. Company": "civilian_company",
            "First Name": "first_name",
            "Last Name": "last_name",
            "Callsign/Nickname": "callsign_nickname",
            "Rank": "rank",
            "Duty Position": "duty_position",
            "Commercial #": "commercial_number",
            "Cell #": "cell_number",
            "SVOIP": "svoip",
            "Company Email": "company_email",
            "NIPR Email": "nipr_email",
            "SIPR Email": "sipr_email",
            "Country": "country"
        }

        # Check if all required headers are present
        missing_headers = [header for header in header_mapping.keys() if header not in df.columns]
        if missing_headers:
            messagebox.showerror("Header Error", f"The following required headers are missing in the Excel file: {', '.join(missing_headers)}")
            logging.error(f"Missing headers in Excel file: {', '.join(missing_headers)}")
            self.close_progress_bar_window()
            return

        # Rename columns based on mapping (no ID included)
        df.rename(columns=header_mapping, inplace=True)

        # Loop through each row and insert data into the database
        for index, row in df.iterrows():
            try:
                # Clean and sanitize each field
                primary_govt_org = clean_data(row.get('primary_govt_org', ''), str)
                directorate = clean_data(row.get('directorate', ''), str)
                dept_div_branch = clean_data(row.get('dept_div_branch', ''), str)
                secondary_govt_org = clean_data(row.get('secondary_govt_org', ''), str)
                civilian_company = clean_data(row.get('civilian_company', ''), str)
                first_name = clean_data(row.get('first_name', ''), str)
                last_name = clean_data(row.get('last_name', ''), str)
                callsign_nickname = clean_data(row.get('callsign_nickname', ''), str)
                rank = clean_data(row.get('rank', ''), str)
                duty_position = clean_data(row.get('duty_position', ''), str)

                # Clean phone numbers by removing non-numeric characters (preserving '+' if present)
                commercial_number = clean_phone_number(clean_data(row.get('commercial_number', ''), str))
                cell_number = clean_phone_number(clean_data(row.get('cell_number', ''), str))

                # Strip and clean email fields, converting any non-string values to strings first
                company_email = clean_data(row.get('company_email', ''), str).strip()
                nipr_email = clean_data(row.get('nipr_email', ''), str).strip()
                sipr_email = clean_data(row.get('sipr_email', ''), str).strip()
                svoip = clean_data(row.get('svoip', ''), str).strip()  # Add svoip handling here

                country = clean_data(row.get('country', ''), str)

                # Initialize a list to keep track of invalid fields for this record
                invalid_fields = []

                # Validate Company Email
                if company_email and not is_valid_email(company_email):
                    invalid_fields.append("Company Email")

                # Validate NIPR Email
                if nipr_email and not is_valid_email(nipr_email):
                    invalid_fields.append("NIPR Email")

                # Validate SIPR Email
                if sipr_email and not is_valid_email(sipr_email):
                    invalid_fields.append("SIPR Email")

                # Validate Commercial Number
                if commercial_number and not is_valid_phone(commercial_number):
                    invalid_fields.append("Commercial Number")

                # Validate Cell Number
                if cell_number and not is_valid_phone(cell_number):
                    invalid_fields.append("Cell Number")

                if invalid_fields:
                    duplicate_records.append(f"Row {index + 2}: Invalid {', '.join(invalid_fields)}")
                    logging.error(f"Row {index + 2}: Invalid {', '.join(invalid_fields)}")
                    continue  # Skip inserting this record

                # Check if an exact record with the same first_name, last_name, and primary_govt_org exists
                self.cursor.execute("""
                    SELECT first_name, last_name, primary_govt_org 
                    FROM employees_v2 
                    WHERE first_name = ? AND last_name = ? AND primary_govt_org = ?
                """, (first_name, last_name, primary_govt_org))

                exact_match = self.cursor.fetchone()

                if exact_match:
                    duplicate_records.append(f"Row {index + 2}: {first_name} {last_name}, {primary_govt_org}")
                    logging.info(f"Duplicate found for Row {index + 2}: {first_name} {last_name}, {primary_govt_org}")
                else:
                    # Handle other logic like fuzzy matching, etc.

                    # Generate a new 7-digit random ID
                    record_id = random.randint(1000000, 9999999)

                    # Insert the cleaned row into the employees_v2 table
                    self.cursor.execute("""
                        INSERT INTO employees_v2 
                        ([primary_govt_org], [directorate], [dept_div_branch], [secondary_govt_org], [civilian_company], 
                        [first_name], [last_name], [callsign_nickname], [rank], [duty_position], [commercial_number], 
                        [cell_number], [svoip], [company_email], [nipr_email], [sipr_email], [country], [id])
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        primary_govt_org, directorate, dept_div_branch, secondary_govt_org, civilian_company,
                        first_name, last_name, callsign_nickname, rank, duty_position, commercial_number,
                        cell_number, svoip, company_email, nipr_email, sipr_email, country, record_id
                    ))
                    logging.info(f"Inserted Row {index + 2} with ID {record_id}")

                # Update the progress bar and percentage label
                progress = (index + 1) / total_rows * 100
                self.progress_bar['value'] = progress  # Update the progress bar
                self.percentage_label.config(text=f"{int(progress)}%")  # Update the percentage label

                # Ensure the UI is updated
                self.progress_window.update_idletasks()

            except Exception as e:
                logging.error(f"Skipping record at Row {index + 2} due to error: {e}")
                continue

        try:
            self.conn.commit()
            logging.info("All eligible records from Excel have been inserted successfully.")
        except pyodbc.Error as e:
            messagebox.showerror("Database Commit Error", f"Failed to commit changes: {e}")
            logging.error(f"Failed to commit changes: {e}")

        # Close the progress bar window
        self.close_progress_bar_window()

        # Inform the user about duplicates or fuzzy matches
        if duplicate_records:
            duplicate_message = (
                "The following records were not added due to invalid phone or email formats:\n\n"
                + "\n".join(duplicate_records)
                + "\n\n"
            )
            if fuzzy_match_records:
                duplicate_message += (
                    "The following records were flagged for near-duplicate matches and required confirmation:\n\n"
                    + "\n".join(fuzzy_match_records)
                    + "\n\n"
                )
            messagebox.showinfo("Import Complete", f"Data imported successfully!\n\n{duplicate_message}")
        else:
            messagebox.showinfo("Success", "Data imported successfully!")

# ------------------------- End of DatabaseWriter Class -------------------------




class PersonnelDataFinder:
    def __init__(self, frame, app):
        self.frame = frame
        self.app = app
        self.conn = None
        self.create_finder_ui()

    def create_finder_ui(self):
        """Initialize and set up the UI for the Data Finder tab."""
        self.font = ("Segoe UI", self.app.current_font_size)

        # Create the main canvas that holds everything
        self.canvas = tk.Canvas(self.frame, bd=0, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Add a scrollbar on the right side of the canvas
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a scrollable frame inside the canvas to hold the content
        self.scrollable_frame = tk.Frame(self.canvas)

        # Bind the configure event to set the proper size and position of the scrollable frame
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)

        # Initially create the window at the top-left corner (we'll reposition it after calculating widths)
        self.scrollable_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Load and display the logo at the top
        self.create_logo()

        # Create the frame for search fields
        self.search_frame = tk.Frame(self.scrollable_frame)
        self.search_frame.pack(pady=10, padx=10)

        # Add the Wildcard Search field first
        self.entry_wildcard_search = self.create_search_row("Wildcard Search:", self.font)

        # Add search fields for various search criteria
        self.entry_first_name = self.create_search_row("First Name:", self.font)
        self.entry_last_name = self.create_search_row("Last Name:", self.font)
        self.entry_primary_govt_org = self.create_search_row("Primary Gov't Org:", self.font)
        # Removed the following two lines to eliminate "Directorate" and "Dept/Div/Branch" fields
        # self.entry_directorate = self.create_search_row("Directorate:", self.font)
        # self.entry_dept_div_branch = self.create_search_row("Dept/Div/Branch:", self.font)
        self.entry_civilian_company = self.create_search_row("Civilian Company:", self.font)
        self.entry_callsign_nickname = self.create_search_row("Callsign/Nickname:", self.font)
        self.entry_rank = self.create_search_row("Rank:", self.font)
        self.entry_duty_position = self.create_search_row("Duty Position:", self.font)
        self.entry_country = self.create_search_row("Country:", self.font)

        # Create the results frame to display the search results
        self.results_frame = tk.Frame(self.scrollable_frame)
        self.results_frame.pack(pady=10, padx=0, expand=True, fill="both")  # Full width and height

        # Add a text widget to display the search results
        self.results_text = tk.Text(
            self.results_frame,
            height=20,
            font=self.font,
            wrap="none",
            bd=0,
            highlightthickness=0
        )
        self.results_text.pack(expand=True, fill="both", padx=0, pady=0)  # Full expansion horizontally and vertically

        # Set up the database connection and search bindings
        self.setup_database_connection()
        self.setup_search_bindings()

        # Ensure proper alignment after the layout is set
        self.canvas.update_idletasks()

        # Recalculate the position to align it to the right
        canvas_width = self.canvas.winfo_width()
        frame_width = self.scrollable_frame.winfo_width()

        if frame_width < canvas_width:
            self.canvas.coords(self.scrollable_window, (canvas_width - frame_width, 0))

    def create_search_row(self, label_text, font):
        """Create a label and entry widget in the search frame."""
        label = ttk.Label(self.search_frame, text=label_text, font=font)
        label.grid(row=len(self.search_frame.grid_slaves()) // 2, column=0, sticky='e', padx=5, pady=5)
        entry = ttk.Entry(self.search_frame, font=font)
        entry.grid(row=len(self.search_frame.grid_slaves()) // 2, column=1, padx=5, pady=5, sticky='w')
        return entry

    def create_logo(self):
        """Load and display the logo image."""
        try:
            logo_image = Image.open(r"C:\Users\Brian Smith\Desktop\logo.png")  # Update the path if necessary
            logo_image = logo_image.resize(
                (700, int(700 * logo_image.height / logo_image.width)),
                Image.LANCZOS
            )
            logo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(self.scrollable_frame, image=logo)
            logo_label.image = logo  # Keep a reference to avoid garbage collection
            logo_label.pack(pady=10)
        except Exception as e:
            logging.error(f"Error loading logo: {e}")
            print(f"Error loading logo: {e}")

    def setup_database_connection(self):
        """Establish a connection to the Azure SQL database."""
        try:
            self.conn = pyodbc.connect(
                "Driver={ODBC Driver 18 for SQL Server};"
                "Server=tcp:idsfuturetech-sql-server.database.windows.net,1433;"
                "Database=ids.personnel;"
                "Uid=sqladmin;"
                "Pwd=YourSecurePassword123!;"  # Ensure this password is correct and secure
                "Encrypt=yes;"
                "TrustServerCertificate=no;"
                "Connection Timeout=30;"
            )
            logging.info("Database connection established successfully.")
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to connect to the database: {e}")
            logging.error(f"Database connection failed: {e}")

    def setup_search_bindings(self):
        """Bind key release events to update the search results dynamically."""
        # Bind the Wildcard Search field
        self.entry_wildcard_search.bind("<KeyRelease>", self.update_search_results)

        # Bind the remaining search fields
        self.entry_first_name.bind("<KeyRelease>", self.update_search_results)
        self.entry_last_name.bind("<KeyRelease>", self.update_search_results)
        self.entry_primary_govt_org.bind("<KeyRelease>", self.update_search_results)
        self.entry_civilian_company.bind("<KeyRelease>", self.update_search_results)
        self.entry_callsign_nickname.bind("<KeyRelease>", self.update_search_results)
        self.entry_rank.bind("<KeyRelease>", self.update_search_results)
        self.entry_duty_position.bind("<KeyRelease>", self.update_search_results)
        self.entry_country.bind("<KeyRelease>", self.update_search_results)

    def update_search_results(self, event=None):
        """Update the search results based on the current input in search fields."""
        search_filters = {
            "first_name": self.entry_first_name.get(),
            "last_name": self.entry_last_name.get(),
            "primary_govt_org": self.entry_primary_govt_org.get(),
            # "directorate": self.entry_directorate.get(),  # Removed
            # "dept_div_branch": self.entry_dept_div_branch.get(),  # Removed
            "civilian_company": self.entry_civilian_company.get(),
            "callsign_nickname": self.entry_callsign_nickname.get(),
            "rank": self.entry_rank.get(),
            "duty_position": self.entry_duty_position.get(),
            "country": self.entry_country.get(),
        }

        wildcard_search = self.entry_wildcard_search.get()

        # Base query to select all relevant fields
        query = """
            SELECT first_name, last_name, primary_govt_org, civilian_company, 
                   callsign_nickname, rank, duty_position, 
                   commercial_number, cell_number, svoip, company_email, 
                   nipr_email, sipr_email, country, id
            FROM employees_v2
            WHERE 1=1
        """

        params = []

        # Add specific search filters
        for field, value in search_filters.items():
            if value:
                query += f" AND {field} LIKE ?"
                params.append(f"%{value}%")

        # Add wildcard search across all fields (excluding removed fields)
        if wildcard_search:
            wildcard_conditions = []
            fields_to_search = [
                'first_name', 'last_name', 'primary_govt_org',
                'civilian_company', 'callsign_nickname', 'rank', 'duty_position',
                'commercial_number', 'cell_number', 'svoip', 'company_email',
                'nipr_email', 'sipr_email', 'country', 'id'
            ]
            for field in fields_to_search:
                wildcard_conditions.append(f"{field} LIKE ?")
                params.append(f"%{wildcard_search}%")
            query += " AND (" + " OR ".join(wildcard_conditions) + ")"

        # Fetch and display all fields
        if self.conn:
            try:
                cursor = self.conn.cursor()
                cursor.execute(query, params)
                results = cursor.fetchall()

                self.results_text.delete(1.0, tk.END)
                for row in results:
                    formatted_result = (
                        f"First Name: {row[0] if row[0] is not None else ''}\n"
                        f"Last Name: {row[1] if row[1] is not None else ''}\n"
                        f"Primary Gov't Org: {row[2] if row[2] is not None else ''}\n"
                        # f"Directorate: {row[3] if row[3] is not None else ''}\n"  # Removed
                        # f"Dept/Div/Branch: {row[4] if row[4] is not None else ''}\n"  # Removed
                        f"Civilian Company: {row[3] if row[3] is not None else ''}\n"
                        f"Callsign/Nickname: {row[4] if row[4] is not None else ''}\n"
                        f"Rank: {row[5] if row[5] is not None else ''}\n"
                        f"Duty Position: {row[6] if row[6] is not None else ''}\n"
                        f"Commercial #: {row[7] if row[7] is not None else ''}\n"
                        f"Cell #: {row[8] if row[8] is not None else ''}\n"
                        f"SVOIP: {row[9] if row[9] is not None else ''}\n"
                        f"Company Email: {row[10] if row[10] is not None else ''}\n"
                        f"NIPR Email: {row[11] if row[11] is not None else ''}\n"
                        f"SIPR Email: {row[12] if row[12] is not None else ''}\n"
                        f"Country: {row[13] if row[13] is not None else ''}\n"
                        f"ID: {row[14] if row[14] is not None else ''}\n"
                    )

                    self.results_text.insert(tk.END, formatted_result + "\n\n")
            except Exception as e:
                logging.error(f"Error fetching search results: {e}")
                messagebox.showerror("Search Error", f"An error occurred while fetching results: {e}")

    def on_frame_configure(self, event):
        """Reset the scroll region to encompass the inner frame."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))






class EditDataTab:
    def __init__(self, frame, app):
        self.frame = frame
        self.app = app
        self.conn = None
        self.create_delete_ui()  # This will also be renamed appropriately for the edit UI


    def create_delete_ui(self):
        self.font = ("Segoe UI", self.app.current_font_size)

        # Create the main canvas that holds everything
        self.canvas = tk.Canvas(self.frame, bd=0, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Add a scrollbar on the right side of the canvas
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a scrollable frame inside the canvas to hold the content
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Adjust the scroll region when the frame's size changes
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)

        # Load and display the logo at the top
        self.create_logo()

        # Create the frame for search fields
        self.search_frame = ttk.Frame(self.scrollable_frame)
        self.search_frame.pack(pady=10, padx=10)

        # Add search fields for various search criteria (First Name, Last Name, etc.)
        self.entry_first_name = self.create_search_row("First Name:", self.font)
        self.entry_last_name = self.create_search_row("Last Name:", self.font)

        # Create the button frame to hold the Edit and Delete buttons
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(pady=20, padx=10)

        # Add the "Edit Record" button
        edit_button = ttk.Button(button_frame, text="Edit Record", bootstyle="primary", command=self.ask_edit_id)
        edit_button.grid(row=0, column=0, padx=10)  # Grid geometry used for even spacing

        # Add the "Delete Record" button
        delete_button = ttk.Button(button_frame, text="Delete Record", bootstyle="danger", command=self.ask_delete_id)
        delete_button.grid(row=0, column=1, padx=10)  # Place it next to Edit Record with grid

        # Create the results frame to display the search results (return data window)
        self.results_frame = ttk.Frame(self.scrollable_frame)
        self.results_frame.pack(pady=10, padx=0, expand=True, fill="both")  # Full width and height

        # Add a text widget to display the search results
        self.results_text = tk.Text(self.results_frame, height=20, font=self.font, wrap="none", bd=0, highlightthickness=0)
        self.results_text.pack(expand=True, fill="both", padx=0, pady=0)  # Full expansion horizontally and vertically

        # Set up the database connection and search bindings
        self.setup_database_connection()
        self.setup_search_bindings()


    def create_logo(self):
        try:
            logo_image = Image.open(r"C:\Users\Brian Smith\Desktop\logo.png")
            logo_image = logo_image.resize((700, int(700 * logo_image.height / logo_image.width)), Image.LANCZOS)
            logo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(self.scrollable_frame, image=logo)
            logo_label.image = logo
            logo_label.pack(pady=10)

        except Exception as e:
            print(f"Error loading logo: {e}")

    def create_search_row(self, label_text, font):
        label = ttk.Label(self.search_frame, text=label_text, font=font)
        label.grid(row=len(self.search_frame.grid_slaves()) // 2, column=0, sticky='e', padx=5, pady=5)
        entry = ttk.Entry(self.search_frame, font=font)
        entry.grid(row=len(self.search_frame.grid_slaves()) // 2, column=1, padx=5, pady=5, sticky='w')
        return entry

    def setup_database_connection(self):
        try:
            self.conn = pyodbc.connect(
                "Driver={ODBC Driver 18 for SQL Server};"
                "Server=tcp:idsfuturetech-sql-server.database.windows.net,1433;"
                "Database=ids.personnel;"
                "Uid=sqladmin;"
                "Pwd=YourSecurePassword123!;"
                "Encrypt=yes;"
                "TrustServerCertificate=no;"
                "Connection Timeout=30;"
            )
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to connect to the database: {e}")

    def setup_search_bindings(self):
        self.entry_first_name.bind("<KeyRelease>", self.update_search_results)
        self.entry_last_name.bind("<KeyRelease>", self.update_search_results)

    def update_search_results(self, event=None):
        first_name = self.entry_first_name.get()
        last_name = self.entry_last_name.get()

        query = "SELECT first_name, last_name, primary_govt_org, directorate, dept_div_branch, civilian_company, callsign_nickname, rank, duty_position, country, id FROM employees_v2 WHERE 1=1"
        params = []

        if first_name:
            query += " AND first_name LIKE ?"
            params.append(f"%{first_name}%")
        if last_name:
            query += " AND last_name LIKE ?"
            params.append(f"%{last_name}%")

        if self.conn:
            cursor = self.conn.cursor()
            cursor.execute(query, params)
            results = cursor.fetchall()

            self.results_text.delete(1.0, tk.END)
            for row in results:
                formatted_result = (
                    f"First Name: {row[0] if row[0] is not None else ''}\n"
                    f"Last Name: {row[1] if row[1] is not None else ''}\n"
                    f"Primary Gov't Org: {row[2] if row[2] is not None else ''}\n"
                    f"Directorate: {row[3] if row[3] is not None else ''}\n"
                    f"Dept/Div/Branch: {row[4] if row[4] is not None else ''}\n"
                    f"Civilian Company: {row[5] if row[5] is not None else ''}\n"
                    f"Callsign/Nickname: {row[6] if row[6] is not None else ''}\n"
                    f"Rank: {row[7] if row[7] is not None else ''}\n"
                    f"Duty Position: {row[8] if row[8] is not None else ''}\n"
                    f"Country: {row[9] if row[9] is not None else ''}\n"
                    f"ID # TO EDIT/DELETE RECORD: {row[10] if row[10] is not None else ''}\n"
                )
                self.results_text.insert(tk.END, formatted_result + "\n\n")

    # Method to handle the Edit Record button click
    def ask_edit_id(self):
        edit_window = tk.Toplevel(self.frame)
        edit_window.title("Edit Record")

        fixed_frame = ttk.Frame(edit_window, width=300, height=300)
        fixed_frame.pack_propagate(False)
        fixed_frame.pack()

        label = ttk.Label(fixed_frame, text="Enter the 6-digit ID of the record to be edited:")
        label.pack(pady=20)

        id_entry = ttk.Entry(fixed_frame)
        id_entry.pack(pady=10)

        confirm_button = ttk.Button(fixed_frame, text="Confirm", command=lambda: self.on_confirm_edit(edit_window, id_entry))
        confirm_button.pack(pady=20)

        id_entry.focus_set()  # Automatically set focus to the ID entry field

        # Bind the 'Enter' key to trigger the confirm button
        edit_window.bind('<Return>', lambda event: self.on_confirm_edit(edit_window, id_entry))


    def on_confirm_edit(self, edit_window, id_entry):
        id_to_edit = id_entry.get()  # Get the ID entered by the user
        edit_window.destroy()  # Close the window after getting the ID
        self.perform_edit(id_to_edit)  # Call the edit function with the entered ID


    # Function to perform the edit (to be further implemented)
    def perform_edit(self, id_to_edit):
        if id_to_edit:
            # SQL query to retrieve the correct fields in the correct order
            query = """
            SELECT first_name, last_name, primary_govt_org, directorate, dept_div_branch, 
                civilian_company, callsign_nickname, rank, duty_position, 
                commercial_number, cell_number, svoip, company_email, 
                nipr_email, sipr_email, country, id
            FROM employees_v2
            WHERE id = ?
            """
            cursor = self.conn.cursor()
            cursor.execute(query, id_to_edit)
            result = cursor.fetchone()



            if result:
                # Open a new window for editing
                edit_window = tk.Toplevel(self.frame)
                edit_window.title("Edit Record")
                edit_window.geometry("400x750")

                # Create a frame to hold the form entries
                form_frame = ttk.Frame(edit_window)
                form_frame.pack(pady=10, padx=10)

                # Correct mapping of each database result to the respective field
                entries = {}

                # Mapping the database fields to the form in the correct order
                labels_entries = [
                    ("First Name", result[0]),  # First name from result[0]
                    ("Last Name", result[1]),   # Last name from result[1]
                    ("Primary Gov't Org", result[2]),  # Primary Govt Org from result[2]
                    ("Directorate", result[3]),
                    ("Dept/Div/Branch", result[4]),
                    ("Civilian Company", result[5]),
                    ("Callsign/Nickname", result[6]),
                    ("Rank", result[7]),
                    ("Duty Position", result[8]),
                    ("Commercial #", result[9]),
                    ("Cell #", result[10]),
                    ("SVOIP", result[11]),
                    ("Company Email", result[12]),
                    ("NIPR Email", result[13]),
                    ("SIPR Email", result[14]),
                    ("Country", result[15])
                ]

                # Create the labels and entry fields, inserting the correct data into each
                for i, (label_text, value) in enumerate(labels_entries):
                    label = ttk.Label(form_frame, text=f"{label_text}:")
                    label.grid(row=i, column=0, sticky='e', padx=5, pady=5)

                    entry = ttk.Entry(form_frame, width=40)
                    entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
                    entry.insert(0, value)  # Insert the correct value for each field
                    entries[label_text] = entry

                # Focus on the First Name field by default
                entries["First Name"].focus_set()

                # Create an 'Update Record' button to save changes
                update_button = ttk.Button(edit_window, text="Update Record", bootstyle="success", command=lambda: self.save_updated_record(id_to_edit, entries))
                update_button.pack(pady=20)

                # Bind the 'Enter' key to trigger the update button
                edit_window.bind('<Return>', lambda event: self.save_updated_record(id_to_edit, entries))



    def save_updated_record(self, id_to_edit, entries):
        # Gather updated data from entry fields
        updated_data = {field: entry.get() for field, entry in entries.items()}

        # Initialize a list to keep track of invalid fields
        invalid_fields = []

        # Validate Company Email
        company_email = updated_data.get("Company Email", "")
        if company_email and not is_valid_email(company_email):
            invalid_fields.append("Company Email")

        # Validate NIPR Email
        nipr_email = updated_data.get("NIPR Email", "")
        if nipr_email and not is_valid_email(nipr_email):
            invalid_fields.append("NIPR Email")

        # Validate SIPR Email
        sipr_email = updated_data.get("SIPR Email", "")
        if sipr_email and not is_valid_email(sipr_email):
            invalid_fields.append("SIPR Email")

        # Validate Commercial Number
        commercial_number = updated_data.get("Commercial #", "")
        if commercial_number and not is_valid_phone(commercial_number):
            invalid_fields.append("Commercial Number")

        # Validate Cell Number
        cell_number = updated_data.get("Cell #", "")
        if cell_number and not is_valid_phone(cell_number):
            invalid_fields.append("Cell Number")

        if invalid_fields:
            # Create an error message listing all invalid fields
            error_message = "The following fields have invalid formats:\n" + "\n".join(invalid_fields)
            messagebox.showerror("Invalid Input Format", error_message)
            
            # Highlight the invalid fields by changing their background color
            for field in invalid_fields:
                if field == "Company Email":
                    entries["Company Email"].config(background="pink")
                elif field == "NIPR Email":
                    entries["NIPR Email"].config(background="pink")
                elif field == "SIPR Email":
                    entries["SIPR Email"].config(background="pink")
                elif field == "Commercial Number":
                    entries["Commercial #"].config(background="pink")
                elif field == "Cell Number":
                    entries["Cell #"].config(background="pink")
            return  # Exit the method to prevent updating

        # Reset background colors in case they were previously highlighted
        entries["Company Email"].config(background="white")
        entries["NIPR Email"].config(background="white")
        entries["SIPR Email"].config(background="white")
        entries["Commercial #"].config(background="white")
        entries["Cell #"].config(background="white")

        try:
            # Clean phone numbers
            cleaned_commercial_number = clean_phone_number(commercial_number)
            cleaned_cell_number = clean_phone_number(cell_number)

            # SQL query to update the record in the database
            update_query = """
            UPDATE employees_v2
            SET first_name = ?, last_name = ?, primary_govt_org = ?, directorate = ?, dept_div_branch = ?, 
                civilian_company = ?, callsign_nickname = ?, rank = ?, duty_position = ?, 
                commercial_number = ?, cell_number = ?, svoip = ?, company_email = ?, 
                nipr_email = ?, sipr_email = ?, country = ?
            WHERE id = ?
            """
            params = [
                updated_data["First Name"],
                updated_data["Last Name"],
                updated_data["Primary Gov't Org"],
                updated_data["Directorate"],
                updated_data["Dept/Div/Branch"],
                updated_data["Civilian Company"],
                updated_data["Callsign/Nickname"],
                updated_data["Rank"],
                updated_data["Duty Position"],
                cleaned_commercial_number,
                cleaned_cell_number,
                updated_data["SVOIP"],
                company_email,
                nipr_email,
                sipr_email,
                updated_data["Country"],
                id_to_edit
            ]
            cursor = self.conn.cursor()
            cursor.execute(update_query, params)
            self.conn.commit()

            messagebox.showinfo("Success", "Record updated successfully!")

        except pyodbc.ProgrammingError as e:
            messagebox.showerror("Error", f"Failed to update record: {e}")
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Database error: {e}")

    def ask_delete_id(self):
        delete_window = tk.Toplevel(self.frame)
        delete_window.title("Delete Record")

        fixed_frame = ttk.Frame(delete_window, width=300, height=300)
        fixed_frame.pack_propagate(False)
        fixed_frame.pack()

        label = ttk.Label(fixed_frame, text="Enter the 6-digit ID of the record to be deleted:")
        label.pack(pady=20)

        id_entry = ttk.Entry(fixed_frame)
        id_entry.pack(pady=10)

        id_entry.focus_set()  # Automatically set focus to the ID entry field

        confirm_button = ttk.Button(fixed_frame, text="Confirm", command=lambda: self.on_confirm_delete(delete_window, id_entry))
        confirm_button.pack(pady=20)

        # Bind the 'Enter' key to trigger the confirm button
        delete_window.bind('<Return>', lambda event: self.on_confirm_delete(delete_window, id_entry))

    def on_confirm_delete(self, delete_window, id_entry):
        id_to_delete = id_entry.get()  # Get the ID entered by the user
        delete_window.destroy()  # Close the window after getting the ID
        self.perform_deletion(id_to_delete)  # Call the deletion function with the entered ID

    def perform_deletion(self, id_to_delete):
        if id_to_delete:
            # SQL query to retrieve the correct fields in the correct order for deletion confirmation
            query = """
            SELECT first_name, last_name, primary_govt_org, directorate, dept_div_branch, 
                civilian_company, callsign_nickname, rank, duty_position, 
                commercial_number, cell_number, svoip, company_email, 
                nipr_email, sipr_email, country, id
            FROM employees_v2
            WHERE id = ?
            """
            cursor = self.conn.cursor()
            cursor.execute(query, id_to_delete)
            result = cursor.fetchone()

            if result:
                # Open a new window to confirm deletion
                confirm_window = tk.Toplevel(self.frame)
                confirm_window.title("Confirm Deletion")
                confirm_window.geometry("400x750")

                form_frame = ttk.Frame(confirm_window)
                form_frame.pack(pady=10, padx=10)

                # Map each database field to the correct form field for confirmation
                fields = [
                    ("First Name", result[0]),  # First name
                    ("Last Name", result[1]),   # Last name
                    ("Primary Gov't Org", result[2]),  # Primary Government Org
                    ("Directorate", result[3]),
                    ("Dept/Div/Branch", result[4]),
                    ("Civilian Company", result[5]),
                    ("Callsign/Nickname", result[6]),
                    ("Rank", result[7]),
                    ("Duty Position", result[8]),
                    ("Commercial #", result[9]),
                    ("Cell #", result[10]),
                    ("SVOIP", result[11]),
                    ("Company Email", result[12]),
                    ("NIPR Email", result[13]),
                    ("SIPR Email", result[14]),
                    ("Country", result[15])
                ]

                # Create form fields to display the data for confirmation
                for i, (label_text, value) in enumerate(fields):
                    label = ttk.Label(form_frame, text=f"{label_text}:")
                    label.grid(row=i, column=0, sticky='e', padx=5, pady=5)

                    entry = ttk.Entry(form_frame, width=40)
                    entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
                    entry.insert(0, value)  # Insert the correct value for each field
                    entry.config(state="readonly")  # Make the fields readonly

                # Ask the user to confirm deletion
                confirm_button = ttk.Button(confirm_window, text="Delete Record", bootstyle="danger", command=lambda: self.delete_record(id_to_delete, confirm_window))
                confirm_button.pack(pady=20)


    def delete_record(self, id_to_delete, confirm_window):
        query = "DELETE FROM employees_v2 WHERE id = ?"
        cursor = self.conn.cursor()
        cursor.execute(query, id_to_delete)
        self.conn.commit()
        
        # Close the confirmation window and show a success message
        confirm_window.destroy()
        messagebox.showinfo("Success", "The record has been deleted successfully.")

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_font_size(self, font_size):
        font = ("Segoe UI", font_size)
        self.font = font
        self.results_text.config(font=font)
        self.update_search_results()


if __name__ == "__main__":
    app = CombinedApp()
    app.mainloop()