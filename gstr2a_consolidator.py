import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from customtkinter import *
import customtkinter as ctk
from datetime import datetime
import re
from tkinter import font

def show_popup(message, title,width=300, height=150):
    """Create a custom popup window to display a message with specified size."""
    popup = ctk.CTkToplevel()  # Use CTkToplevel for a custom popup
    popup.title(title)

    # Set the size of the popup window
    popup.geometry(f"{width}x{height}")

    # Create a label to display the message with text wrapping
    label = ctk.CTkLabel(popup, text=message, padx=20, pady=20, wraplength=width - 40)
    label.pack(expand=True)

    button_frame = ctk.CTkFrame(popup)
    button_frame.pack(padx=10, pady=10)

    # Create a button to close the popup
    ok_button = ctk.CTkButton(button_frame, text="OK", command=popup.destroy, corner_radius=10)
    ok_button.pack()
    # Create a button to close the popup
    # button = tk.Button(popup, text="OK", command=popup.destroy, padx=10, pady=5)
    # button.pack(pady=10)

# Define the pattern for the filename
filename_pattern = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]{3}_[0-1][0-9][0-9]{4}_R2A\.xlsx$')

def is_valid_filename(filename):
    """Check if the filename matches the required format."""
    basename = os.path.basename(filename)
    return bool(filename_pattern.match(basename))


# Function to convert month number to month name
def month_number_to_name(month_num: str) -> str:
    return datetime.strptime(month_num, "%m").strftime("%B")




# Store the actual file paths separately
selected_files = {}

# Function to select multiple Excel files
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    file_list.delete("1.0", tk.END)  # Clear the text box before displaying new files

    if files:
        files = sort_files(files)
        file_names = [os.path.basename(file) for file in files]
        year = None
        months = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
        month_files = {month: None for month in months}

        for file in files:
            if not is_valid_filename(file):
                file_name_show = os.path.basename(file)
                show_popup(f"The file '{file_name_show}' does not match the required format. Please reselect. For help, click on Info.", "Invalid File")
                return  # Stop processing and return to reselect files

            match = re.search(r'_(\d{2})(\d{4})_', file)
            if match:
                month_num = match.group(1)
                year = match.group(2)
                month_name = month_number_to_name(month_num)
                # Handle year change for March
                if month_name == "March":
                    year = str(int(year) + 1)
                month_files[month_name] = file

        start_year = int(year) - 1 if year else None
        end_year = int(year) if year else None

        if start_year and end_year:
            for month in months:
                month_file = month_files[month]
                if month_file:
                    file_name = get_month_file_name(month_file)
                    # Store the actual file path in a dictionary
                    selected_files[month] = month_file
                    file_list.insert(tk.END, file_name + '\n', 'success')  # Apply success tag for correct files
                else:
                    month_year_str = f"{month}, {start_year if month in months[:9] else end_year}"
                    file_name = f"{month_year_str:<25} - File missing for this month."
                    file_list.insert(tk.END, file_name + '\n', 'error')  # Apply error tag for missing files
        else:
            file_list.insert(tk.END, "No files found for the selected financial year.", 'error')

# Function to get month and file name
def get_month_file_name(file: str, total_width: int = 25) -> str:
    file_name = os.path.basename(file)
    match = re.search(r'_(\d{2})(\d{4})_', file_name)
    if match:
        month_num = match.group(1)  # Extract month (MM)
        year = match.group(2)       # Extract year (YYYY)
        month_name = month_number_to_name(month_num)
        month_year_str = f"{month_name}, {year}"
        padding_spaces = total_width - len(month_year_str)
        formatted_string = f"{month_year_str:<{total_width}} - {file_name}"
    else:
        formatted_string = "Incorrect file format."
    return formatted_string


# Function to process selected files using the actual file paths
def process_selected_files():
    for month in selected_files:
        actual_file_path = selected_files[month]

# Sorting files based on month and year
def sort_files(files):
    return sorted(files, key=lambda f: re.search(r'_(\d{2})(\d{4})_', os.path.basename(f)).groups())

HELP_TEXT  = "No files selected. Use 'Browse Files' to add files."
# Define headers for each type
HEADERS = {
    "B2B": ["GSTIN of supplier", "Trade/Legal name of the Supplier", "Invoice number", "Invoice type", "Invoice Date", "Invoice Value (₹)", "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period", "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which Amended", "Effective date of cancellation", "Source", "IRN", "IRN date"],
    "B2BA": ["Original Invoice No.", "Original Invoice Date", "GSTIN of supplier", "Trade/Legal name of the Supplier", "Invoice type", "Invoice number", "Invoice Date", "Invoice Value (₹)", "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period", "GSTR-3B Filing Status", "Effective date of cancellation", "Amendment made, if any", "Original tax period in which reported"],
    "CDNR": ["GSTIN of supplier", "Trade/Legal name of the Supplier", "Note Type", "Note number", "Note Supply type", "Note Date", "Note Value (₹)", "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period", "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which Amended", "Effective date of cancellation", "Source", "IRN", "IRN date"],
    "CDNRA": ["Note Type", "Note Number", "Note Date", "GSTIN of supplier", "Trade/Legal name of the Supplier", "Note Supply type", "Note Value (₹)", "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period", "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which reported earlier", "Effective date of cancellation"],
    "ECO": ["GSTIN of ECO", "Trade/Legal name of the ECO", "Document number", "Document type", "Document Date", "Document Value (₹)", "Place of supply Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A Filing Status", "GSTR-1/IFF/GSTR-1A Filing Date", "GSTR-1/IFF/GSTR-1A Filing Period", "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which Amended", "Effective date of cancellation", "Source", "IRN", "IRN date"],
    "ECOA": ["Document number", "Document Date", "GSTIN of ECO", "Trade/Legal name of the ECO", "Document type", "Document Value (₹)", "Place of supply", "Rate (%)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)", "GSTR-1/IFF/GSTR-1A Filing Status", "GSTR-1/IFF/GSTR-1A Filing Date", "GSTR-1/IFF/GSTR-1A Filing Period", "GSTR-3B Filing Status", "Effective date of cancellation", "Amendment made if any", "Original tax period in which reported"],
    "ISD": ["Eligibility of ITC", "GSTIN of ISD", "Trade/Legal name of the ISD", "ISD Document type", "ISD Invoice number", "ISD Invoice date", "ISD credit note number", "ISD credit note date", "Original Invoice Number", "Original invoice date", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)", "ISD GSTR-6 Filing status", "Amendment made if any", "Tax Period in which Amended"],
    "ISDA": ["Original ISD Document type", "Original Document Number", "Original Document date", "Eligibility of ITC", "GSTIN of ISD", "Trade/Legal name of the ISD", "ISD Document type", "ISD Invoice number", "ISD Invoice date", "ISD credit note number", "ISD credit note date", "Original Invoice Number", "Original invoice date", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)", "ISD GSTR-6 Filing status", "Amendment made if any", "Original tax period in which reported"],
    "TDS": ["GSTIN of Deductor", "Deductor's Name", "Tax period of GSTR 7", "Taxable value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)"],
    "TDSA": ["GSTIN of Deductor", "Deductor's Name", "Tax period of original GSTR 7", "Tax period of amended GSTR 7", "Revised taxable value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)"],
    "TCS": ["GSTIN of E-com. Operator", "E-com. Operator's name", "Tax period of GSTR 8", "Gross Value of supplies (₹)", "Value of supplies returned (₹)", "Net amount liable for TCS (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)"],
    "IMPG": ["Reference date (ICEGATE)", "Port code", "BE Number", "Date", "Taxable value (₹)", "Integrated tax (₹)", "Cess (₹)", "Amended (Yes)"],
    "IMPG SEZ": ["GSTIN of supplier", "Trade/Legal name", "Reference date (ICEGATE)", "Port code", "BE Number", "Date", "Taxable value (₹)", "Integrated tax (₹)", "Cess (₹)", "Amended (Yes)"]
}

SKIP_ROWS = {
    "B2B": 6, "B2BA": 7, "CDNR": 6, "CDNRA": 7, "ECO": 6, "ECOA": 7, "ISD": 6, "ISDA": 7,
    "TDS": 6, "TDSA": 6, "TCS": 6, "IMPG": 6, "IMPG SEZ": 6
}

ALL_SHEETS = list(HEADERS.keys())


# Function to show the 'About' window
def show_about():
    about_window = ctk.CTkToplevel()  # Create a new top-level window (pop-up)
    about_window.title("About")
    about_window.geometry("300x200")  # Set window size

    # Add labels for information in the 'About' window
    ctk.CTkLabel(about_window, text="GSTR2A Consolidator", font=("Arial", 16)).pack(pady=10)
    ctk.CTkLabel(about_window, text="Version 1.0", font=("Arial", 12)).pack(pady=5)
    ctk.CTkLabel(about_window, text="Developed by: Udit Vashisht", font=("Arial", 12)).pack(pady=5)
    ctk.CTkLabel(about_window, text="This app consolidates the monthly GSTR2A downloaded from GST BO into one excel file.", font=("Arial", 10)).pack(pady=5)

    # Add a Close button to close the About window
    ctk.CTkButton(about_window, text="Close", command=about_window.destroy).pack(pady=20)

def exit_app():
    """Exit the application."""
    root.quit()


def clear_files():
    """Clear the Listbox input."""
    file_list.delete("1.0", tk.END)
    file_list.insert(tk.END, HELP_TEXT)

# Function to select output file location (file will be saved automatically)
def get_output_file_name(files):
    """Generate the output file name based on GSTIN and Financial Year, and save it to the same directory as the input files."""
    # Get the directory of the first file
    input_directory = os.path.dirname(files[0])

    # Extract the GSTIN and Financial Year from the first file name
    first_file = os.path.basename(files[0])
    gstin, fy = extract_gstin_and_fy(first_file)

    # Generate the output file name
    output_filename = f"{gstin}_2A_FY_{fy}.xlsx"

    # Return the full path of the output file in the same directory
    return os.path.join(input_directory, output_filename)

def extract_gstin_and_fy(filename):
    """Extract GSTIN and Financial Year from the filename."""
    gstin = filename[:15]
    month_year = filename.split('_')[1]
    month = int(month_year[:2])
    year = int(month_year[2:6])
    fy = f"{year-1}_{year}" if month <= 3 else f"{year}_{year+1}"
    return gstin, fy

def sort_files(files):
    """
    Sort files by year and month, with 042018 first and 032019 last.
    """
    sorted_files = sorted(files, key=lambda x: (x.split('_')[1][2:], x.split('_')[1][:2]))
    return sorted_files

# Function to process Excel files and concatenate data
def process_files(files, sheet_name, rows_to_skip):
    try:
        dataframes = []
        for file in files:
            if file.endswith('.xlsx'):
                df = create_df(sheet_name, file, rows_to_skip)
                dataframes.append(df)
            else:
                raise ValueError(f"{file} is not an Excel file.")

        final_dataframe = pd.concat(dataframes, ignore_index=True)
        if final_dataframe.empty:
            final_dataframe = pd.DataFrame(columns = HEADERS[sheet_name])
        final_dataframe.columns = HEADERS[sheet_name]
        return final_dataframe
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to create a dataframe from an Excel file
def create_df(sheet_name, file, rows_to_skip=6):
    df = pd.read_excel(file, sheet_name=sheet_name, skiprows=rows_to_skip, header=None)
    rows_to_keep = []
    for i in range(len(df)):
        if i < len(df) - 1 and df.iloc[i + 1].isnull().all():
            continue
        else:
            rows_to_keep.append(df.iloc[i])
    new_df = pd.DataFrame(rows_to_keep)
    new_df = new_df.dropna(how='all')
    if sheet_name in ["B2B", "B2BA", "CDNR", "CDNRA"]:
        new_df = new_df[:-1]
    if new_df.empty :
        new_df = pd.DataFrame()
    return new_df

# Function to trigger the file processing
def start_processing():
    ALL_SHEETS_DICT = {}
    displayed_files = file_list.get("1.0", tk.END)  # These are the customized file names
    # Check if no files are selected
    if not displayed_files or displayed_files.startswith("No files selected"):
        show_popup("No files selected, please select files to process.", "No File")
        return

    # Collect the actual file paths from the selected_files dictionary
    actual_files = [selected_files[month] for month in selected_files if selected_files[month]]  # Use actual file paths
    actual_files = sort_files(actual_files)  # Sort the actual files

    output_file = get_output_file_name(actual_files)  # Use actual file paths for output file name

    for sheet in ALL_SHEETS:
        dataframe = process_files(actual_files, sheet, SKIP_ROWS[sheet])  # Pass actual files for processing
        ALL_SHEETS_DICT[sheet] = dataframe

    # Use pandas ExcelWriter to create an Excel file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Iterate over the dictionary and write each DataFrame to a separate sheet
        for sheet_name, df in ALL_SHEETS_DICT.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # Access the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # Set text wrapping format
            wrap_format = workbook.add_format({'text_wrap': True})

            # Adjust column width and apply the wrap format to all columns in the DataFrame
            for col_num, column in enumerate(df.columns):
                # Get the maximum length of data in each column, including the header
                max_len = max(df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
                worksheet.set_column(col_num, col_num, max_len, wrap_format)

    success_message = f"Success, the consolidated file is saved as {os.path.basename(output_file)} in the working folder."
    show_popup(success_message, "Success")
    clear_files()

def open_info_window():
    info_window = ctk.CTkToplevel()  # Create a new top-level window
    info_window.title("Information")
    info_window.geometry("500x300")

    # Create Tabview for "About" and "How to Use" sections
    tabview = ctk.CTkTabview(info_window)
    tabview.pack(expand=True, fill="both", padx=20, pady=20)

    # Add tabs
    tabview.add("About")
    tabview.add("How to Use")

    # --- About Section ---
    about_text = ("GST2A Consolidator\n"
                  "Version 1.0\n\n"
                  "Developed by: Udit Vashisht\n"
                  "Email: udit.vashisht@gov.in\n"
                  "This app helps to consolidate the monthly GSTR2A files "
                  "downloaded from GST-BO into one annual file containing "
                  "data of all the months.")

    # Create a Textbox for the About section to enable text wrapping
    about_textbox = ctk.CTkTextbox(tabview.tab("About"), wrap=tk.WORD, width=480, height=200)
    about_textbox.pack(padx=10, pady=10)
    about_textbox.insert(tk.END, about_text)  # Insert the About text
    about_textbox.configure(state='disabled')  # Make it read-only

    # --- How to Use Section ---
    how_to_use_text = ("1. Click the 'Browse Files' button to select monthly GSTR2A files.\n"
                       "2. The files must be downloaded from GST-BO.\n"
                       "3. Although you can consolidate more than 12 files, but the app is designed to work better with 12 files of a single financial year.\n"
                       "4. You must not change the file name of the excel file downloaded from GST-BO. It must be in the format *****GSTIN*****_MMYYYY_2A.xlsx.\n"
                       "4. Click 'Process Files' to process the selected files.\n"
                       "5. Click 'Clear Input' to reset the input file list.\n"
                       "6. The output file will be saved in the same directory as the input files.")

    # Create a Textbox for the How to Use section to enable text wrapping
    how_to_use_textbox = ctk.CTkTextbox(tabview.tab("How to Use"), wrap=tk.WORD, width=480, height=200)
    how_to_use_textbox.pack(padx=10, pady=10)
    how_to_use_textbox.insert(tk.END, how_to_use_text)  # Insert the How to Use text
    how_to_use_textbox.configure(state='disabled')  # Make it read-only

# Main Application Window
root = ctk.CTk()
root.geometry("600x450")
root.title("GSTR2A Consolidator")

# Set the appearance mode
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Use CTkFont instead of tkinter font
fixed_width_font = ctk.CTkFont(family="Courier", size=14)

# Create the instructions label
instructions_label = ctk.CTkLabel(root, text="Please select the Monthly GSTR2A (GST-BO) Excel files to process. Use the buttons below to manage your files.", font=("Arial", 16), wraplength=400, justify=tk.CENTER)
instructions_label.pack(padx=10, pady=10, fill=tk.X)

# Creating a frame for file selection and listing
file_list = ctk.CTkTextbox(root, width=500, height=200, font=fixed_width_font)
file_list.pack(padx=10, pady=10, fill=tk.X)
file_list.insert(tk.END, "No files selected. Use 'Browse Files' to add files.\n")

# Configure tags for different text colors
file_list.tag_config('success', foreground='green')  # Green for correct files
file_list.tag_config('error', foreground='red')      # Red for errors or missing files

# Create buttons
button_frame_1 = ctk.CTkFrame(root)
button_frame_1.pack(padx = 10,pady=10)
button_frame_2= ctk.CTkFrame(root)
button_frame_2.pack(padx = 10,pady=10)

# First row buttons
browse_button = ctk.CTkButton(button_frame_1, text="Browse Files", command=select_files, corner_radius=10)
browse_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

process_button = ctk.CTkButton(button_frame_1, text="Process Files", command=start_processing, corner_radius=10)
process_button.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

clear_button = ctk.CTkButton(button_frame_1, text="Clear Input", command=clear_files, corner_radius=10)
clear_button.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")

# Second row buttons
info_button = ctk.CTkButton(button_frame_2, text="Info", command=open_info_window, corner_radius=10)
info_button.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

exit_button = ctk.CTkButton(button_frame_2, text="Exit", command=exit_app, corner_radius=10, fg_color="red")
exit_button.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

root.mainloop()
