import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def choose_file():
    """
    Opens a file dialog for the user to select a file.

    Returns:
        str: The full path to the selected file.
    """
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(
        filetypes=[("Excel Files", "*.xlsx;*.xlsm")],
        title="Select an Excel file"
    )
    if not file_path:
        raise FileNotFoundError("No file was selected!")
    return file_path


def split_by_sell_branch(file_path, branch_groups):
    """
    Splits the Excel file into separate workbooks based on multiple lists of Sell Branch codes.

    Args:
        file_path (str): Path to the Excel file.
        branch_groups (dict): Dictionary where keys are group names and values are lists of Sell Branch codes.
    """
    df = pd.read_excel(file_path, engine='openpyxl')
    output_dir = os.path.dirname(file_path)

    for group_name, branches in branch_groups.items():
        df_group = df[df["Code (Sell Branch)"].isin(branches)]
        if not df_group.empty:
            output_path = os.path.join(output_dir, f"SellBranch_{group_name}.xlsx")
            df_group.to_excel(output_path, index=False)
            print(f"Created file: {output_path}")


if __name__ == "__main__":
    try:
        file_path = choose_file()
        branch_groups = {
            "Group_A": [1,2,3],
            "Group_B": [4,5,6],
            "Group_C": [7,8,9]
        }  # Example groups of branch codes
        split_by_sell_branch(file_path, branch_groups)
        print("Splitting complete.")
    except Exception as e:
        print(f"An error occurred: {e}")
