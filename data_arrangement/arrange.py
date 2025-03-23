import shutil
from pathlib import Path
import pandas as pd

def fetch_and_rename_excel_files():
    # Define source and destination directories
    source_base = Path("S:/Desktop/your source path")
    dest_dir = Path("S:/Desktop/combined sheet")
    
    # Create destination directory if it doesn't exist
    dest_dir.mkdir(parents=True, exist_ok=True)
    
    # Process the initial directory without a number suffix
    process_directory(source_base / "path" , dest_dir)
    
    # Process directories with number suffixes starting from 1
    i = 1
    while True:
        dir_path = source_base / f"path{i}/Outputs_of_path{i}"
        if not dir_path.exists():
            print(f"Directory {dir_path} does not exist. Stopping.")
            break  # Stop if the directory doesn't exist
        process_directory(dir_path, dest_dir)
        i += 1

    print("the path is",dir_path)

def process_directory(dir_path, dest_dir):
    print(f"Processing directory: {dir_path}")  # Added log
    excel_file = dir_path / "combined result.xlsx"
    if not excel_file.exists():
        print(f"Warning: The file {excel_file} does not exist in {dir_path}.")
        return
    
    # Generate the new filename based on the directory name
    original_dir_name = dir_path.name
    parts = original_dir_name.split('_')
    parts[0] = 'Output'  # Replace 'Outputs' with 'Output'
    new_part = ' '.join(parts)
    new_filename = f"Combined result of {new_part}.xlsx"
    dest_path = dest_dir / new_filename
    
    # Copy the file to the destination with the new name
    print(f"Copying file from: {excel_file} to {dest_path}")
    shutil.copy(excel_file, dest_path)
    print(f"Copied: {excel_file} -> {dest_path}")


def create_master_sheet():
    dest_dir = Path("S:\\Desktop\\combined sheet")
    master_path = dest_dir / "master_sheet.xlsx"
    
    # Get all fetched Excel files matching the pattern
    fetched_files = list(dest_dir.glob("Combined result of Output of path*.xlsx"))
    
    if not fetched_files:
        print("No Excel files found to create master sheet.")
        return
    
    # Create a master Excel workbook
    with pd.ExcelWriter(master_path, engine='openpyxl') as writer:
        for file in fetched_files:
            # Extract the sheet name (e.g., "path", "path1")
            sheet_name = file.stem.split("Combined result of Output of ")[-1].strip()
            # Read the Excel file
            df = pd.read_excel(file)
            # Write to the master workbook as a new worksheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\nMaster workbook created at: {master_path}")

if __name__ == "__main__":
    fetch_and_rename_excel_files()
    create_master_sheet()
