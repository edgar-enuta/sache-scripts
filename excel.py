import os
from dotenv import load_dotenv
from openpyxl import Workbook
from datetime import datetime

load_dotenv()

EXCEL_PATH = os.getenv("EXCEL_PATH")

def generate_timestamped_filename(base_path=EXCEL_PATH):
    """
    Generate a timestamped Excel filename based on the current date and time.
    Format: YYYY-MM-DD_HH-MM-SS_original_filename.xlsx
    """
    if not base_path:
        raise ValueError("EXCEL_PATH not set in .env file.")
    
    # Get current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    # Extract directory and filename components
    dir_name = os.path.dirname(base_path)
    base_name = os.path.basename(base_path)
    name, ext = os.path.splitext(base_name)
    
    # Create timestamped filename
    timestamped_filename = f"{timestamp}_{name}{ext}"
    
    # Create full path
    timestamped_path = os.path.join(dir_name, timestamped_filename)
    
    # Ensure directory exists
    if dir_name and not os.path.exists(dir_name):
        os.makedirs(dir_name)
    
    return timestamped_path

def export_to_excel(data, excel_path=None):
    """
    Creates a new Excel file with a timestamped filename for each run.
    Each dict in data is written as a row to the Excel file.
    Each dict's keys are column names; values are written to the corresponding columns.
    Creates the file and columns if they do not exist.
    """
    if not data:
        print("No data to export. Skipping Excel file creation.")
        return None
    
    # Generate timestamped filename for this run
    if excel_path is None:
        excel_path = generate_timestamped_filename()
    
    print(f"Creating Excel file: {excel_path}")

    # Determine all columns from all dicts in data
    columns = set()
    for row in data:
        columns.update(row.keys())
    columns = list(columns)

    # Always create a new workbook with timestamped filename
    wb = Workbook()
    ws = wb.active
    ws.append(columns)

    for row in data:
        # Ensure order matches columns
        ws.append([row.get(col, "") for col in columns])

    wb.save(excel_path)
    print(f"Excel file created successfully with {len(data)} rows")
    return excel_path
