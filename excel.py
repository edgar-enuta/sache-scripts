import os
from dotenv import load_dotenv
from openpyxl import Workbook
from datetime import datetime
import yaml

load_dotenv()

EXCEL_PATH = os.getenv("EXCEL_PATH")

def get_column_order(config_path="field_config.yaml"):
    """
    Get the preferred column order from the config file.
    System columns (order_number, email_date) come first, followed by pattern fields.
    """
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    
    # Define system columns that should appear first
    system_fields = ['order_number', 'email_date']
    ordered_columns = []
    
    # Add system columns first
    for field in system_fields:
        if field in config:
            excel_column = config[field].get('excel_column')
            if excel_column:
                ordered_columns.append(excel_column)
    
    # Add pattern-matching fields
    for field, props in config.items():
        if field not in system_fields and props.get('pattern') and props.get('excel_column'):
            ordered_columns.append(props['excel_column'])
    
    return ordered_columns

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

    # Get preferred column order from configuration
    preferred_columns = get_column_order()
    
    # Determine all columns present in the data
    all_columns = set()
    for row in data:
        all_columns.update(row.keys())
    
    # Create final column list: preferred columns first, then any extra columns
    columns = []
    # Add preferred columns that exist in data
    for col in preferred_columns:
        if col in all_columns:
            columns.append(col)
    
    # Add any remaining columns that weren't in the preferred list
    for col in sorted(all_columns):  # Sort for consistent ordering
        if col not in columns:
            columns.append(col)

    # Always create a new workbook with timestamped filename
    wb = Workbook()
    ws = wb.active
    ws.append(columns)

    for row in data:
        # Ensure order matches columns
        ws.append([row.get(col, "") for col in columns])

    wb.save(excel_path)
    print(f"Excel file created successfully with {len(data)} rows and columns: {', '.join(columns)}")
    return excel_path
