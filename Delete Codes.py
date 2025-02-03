import win32com.client as win32
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def choose_file():
    """
    Opens a file dialog for the user to select a file.

    Returns:
        str: The full path to the selected file.
    """
    # Hide the main Tkinter window
    root = Tk()
    root.withdraw()

    # Open file dialog
    print("Please select the Excel file...")
    file_path = askopenfilename(
        filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")],  # Restrict to Excel files
        title="Select an Excel file"
    )
    if not file_path:
        raise FileNotFoundError("No file was selected!")
    return file_path

def clean_data(file_path, data_sheet_name="Data", col_index=4, criteria=None):
    """
    Cleans the data in the specified Excel file by deleting rows based on criteria.

    Args:
        file_path (str): Path to the Excel file.
        data_sheet_name (str): Name of the sheet containing the data.
        col_index (int): Index of the column to filter (1-based).
        criteria (list): List of values to delete rows for.
    """
    if criteria is None:
        criteria = []

    # Start Excel application
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Set to False to run in the background

    try:
        # Open the workbook
        print("Opening workbook for cleaning...")
        workbook = excel.Workbooks.Open(file_path)

        # Get the data sheet
        print(f"Accessing sheet: {data_sheet_name}")
        ws_data = workbook.Worksheets(data_sheet_name)

        # Apply AutoFilter to the column
        print(f"Applying AutoFilter to column index {col_index} with criteria: {criteria}")
        ws_data.Columns(col_index).AutoFilter(Field=1, Criteria1=criteria, Operator=7)  # 7 = xlFilterValues

        # Delete all visible rows (excluding the header)
        print("Deleting filtered rows...")
        visible_rows = ws_data.Rows("2:" + str(ws_data.Rows.Count)).SpecialCells(12)  # 12 = xlCellTypeVisible
        visible_rows.Delete()

        # Turn off AutoFilter
        ws_data.AutoFilterMode = False

        # Save and close the workbook
        print("Saving cleaned data...")
        workbook.Save()
        print("Data cleaned successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")


# Main Program
if __name__ == "__main__":
    try:
        # Choose a file dynamically
        file_path = choose_file()

        # Clean the data using the optimized approach
        clean_data(file_path, col_index=4, criteria=["D", "DAM", "BSM", "TR", "bsm", "PM","WAR"])
    except Exception as e:
        print(f"An error occurred: {e}")

