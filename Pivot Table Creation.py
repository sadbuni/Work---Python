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

def create_pivot_table(file_path, data_sheet_name="Data", pivot_sheet_name="Pivot Tables"):
    """
    Creates pivot tables in the specified Excel file.

    Args:
        file_path (str): Path to the Excel file.
        data_sheet_name (str): Name of the sheet containing the data.
        pivot_sheet_name (str): Name of the sheet where pivot tables will be created.
    """
    # Start Excel application
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Set to False to run in the background

    try:
        # Open the workbook
        print("Opening workbook...")
        workbook = excel.Workbooks.Open(file_path)

        # Get the data sheet
        print(f"Accessing sheet: {data_sheet_name}")
        ws_data = workbook.Worksheets(data_sheet_name)

        # Clear or create the pivot table sheet
        try:
            ws_pivot = workbook.Worksheets(pivot_sheet_name)
            print(f"Clearing existing sheet: {pivot_sheet_name}")
            ws_pivot.Cells.Clear()  # Clear any existing data
        except Exception:
            print(f"Creating new sheet: {pivot_sheet_name}")
            ws_pivot = workbook.Sheets.Add()
            ws_pivot.Name = pivot_sheet_name

        # Determine the range of data
        print("Determining data range...")
        last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp
        last_col = ws_data.Cells(1, ws_data.Columns.Count).End(-4159).Column  # -4159 is xlToLeft
        print(f"Data range: Rows 1-{last_row}, Columns 1-{last_col}")

        # Build the full range object
        data_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(last_row, last_col))

        # Create the first Pivot Table
        pivot_cache1 = workbook.PivotCaches().Create(SourceType=1, SourceData=data_range)
        pivot_table1 = pivot_cache1.CreatePivotTable(TableDestination=ws_pivot.Cells(1, 1), TableName="PivotTable1")
        pivot_table1.PivotFields("Code (Sell Branch)").Orientation = 1  # Row field
        pivot_table1.PivotFields("Code (Adjustment Type)").Orientation = 2  # Column field
        pivot_table1.AddDataField(pivot_table1.PivotFields("Cost"), "Sum of Cost", -4157)
        pivot_table1.NullString = "$0.00"
        pivot_table1.PivotFields("Sum of Cost").NumberFormat = "$#,##0.00_);($#,##0.00)"
        pivot_table1.TableStyle2 = "PivotStyleLight16"

        # Rename cells for first Pivot Table
        ws_pivot.Range("A1").Value = "ADJ$"
        ws_pivot.Range("B1").Value = "ADJ Code"
        ws_pivot.Range("A2").Value = "BR#"

        # Calculate the row below the first pivot table
        first_pivot_end_row = ws_pivot.Cells(ws_pivot.Rows.Count, 1).End(-4162).Row
        second_pivot_start_row = first_pivot_end_row + 2

        # Create the second Pivot Table
        pivot_cache2 = workbook.PivotCaches().Create(SourceType=1, SourceData=data_range)
        pivot_table2 = pivot_cache2.CreatePivotTable(TableDestination=ws_pivot.Cells(second_pivot_start_row, 1),
                                                     TableName="PivotTable2")
        pivot_table2.PivotFields("Code (Sell Branch)").Orientation = 1  # Row field
        pivot_table2.PivotFields("Division (Product)").Orientation = 2  # Column field
        pivot_table2.AddDataField(pivot_table2.PivotFields("Cost"), "Sum of Cost", -4157)
        pivot_table2.NullString = "$0.00"
        pivot_table2.PivotFields("Sum of Cost").NumberFormat = "$#,##0.00_);($#,##0.00)"
        pivot_table2.TableStyle2 = "PivotStyleLight16"

        # Rename cells for second Pivot Table
        ws_pivot.Cells(second_pivot_start_row, 1).Value = "ADJ$"
        ws_pivot.Cells(second_pivot_start_row, 2).Value = "Division"
        ws_pivot.Cells(second_pivot_start_row + 1, 1).Value = "BR#"

        # Calculate the row below the second pivot table
        second_pivot_end_row = ws_pivot.Cells(ws_pivot.Rows.Count, 1).End(-4162).Row
        third_pivot_start_row = second_pivot_end_row + 2

        # Create the third Pivot Table
        print("Creating Third Pivot Table...")
        pivot_cache3 = workbook.PivotCaches().Create(SourceType=1, SourceData=data_range)
        pivot_table3 = pivot_cache3.CreatePivotTable(TableDestination=ws_pivot.Cells(third_pivot_start_row, 1),
                                                     TableName="PivotTable3")

        # Configure the third Pivot Table
        print("Configuring Third Pivot Table...")
        pivot_table3.PivotFields("Code (Sell Branch)").Orientation = 1  # Top-level row field
        pivot_table3.PivotFields("Code (Written By)").Orientation = 1  # Nested row field
        pivot_table3.AddDataField(pivot_table3.PivotFields("Cost"), "ADJ $", -4157)  # Sum field
        pivot_table3.AddDataField(pivot_table3.PivotFields("Cost"), "# of ADJ", -4112)  # Count field

        # Rename "Row Labels" to "ADJ By" while keeping the layout consistent
        ws_pivot.Cells(third_pivot_start_row + 1, 1).Value = "ADJ By"

        # Set blank values to display "$0.00"
        pivot_table3.NullString = "$0.00"

        # Format the Third Pivot Table
        pivot_table3.PivotFields("ADJ $").NumberFormat = "$#,##0.00_);($#,##0.00)"
        pivot_table3.TableStyle2 = "PivotStyleLight16"

        # Align all data to the left
        print("Aligning data to the left...")
        ws_pivot.Cells.HorizontalAlignment = -4131  # -4131 is xlLeft

        # Save the workbook
        print("Saving workbook...")
        workbook.Save()
        print(f"Three Pivot Tables created successfully in the sheet: {pivot_sheet_name}")

    except Exception as e:
        print(f"An error occurred: {e}")

# Main Program
if __name__ == "__main__":
    try:
        # Choose a file dynamically
        file_path = choose_file()

        # Create pivot tables
        create_pivot_table(file_path)
    except Exception as e:
        print(f"An error occurred: {e}")
