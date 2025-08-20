import yaml
import re
from datetime import datetime

def extract_fields_from_emails(emails, config_path="field_config.yaml"):
  """
  For each email, search for each field's pattern (if it has excel_column) in the email text.
  Also adds system-generated columns like order number and email date.
  Save any found value in a dict under the key from excel_column.
  Returns a list of dicts (one per email).
  """
  with open(config_path, "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

  results = []
  order_counter = 1  # Start order numbering from 1
  
  for email in emails:
    row = {}
    text = f"{email.get('subject', '')}\n{email.get('body', '')}"
    
    # Process pattern-matching fields
    for field, props in config.items():
      pattern = props.get('pattern')
      excel_column = props.get('excel_column')
      if pattern and excel_column:
        match = re.search(pattern, text)
        if match:
          row[excel_column] = match.group(1) if match.groups() else match.group(0)
    
    # Add system-generated columns
    # Order Number
    order_number_column = config.get('order_number', {}).get('excel_column', 'OrderNo')
    row[order_number_column] = order_counter
    order_counter += 1
    
    # Email Date
    email_date_column = config.get('email_date', {}).get('excel_column', 'EmailDate')
    email_date = email.get('date', '')
    if email_date:
      # Try to parse and format the email date
      try:
        # Email date format can vary, try to parse it
        # If parsing fails, use the raw date string
        parsed_date = datetime.strptime(email_date.split(' (')[0], '%a, %d %b %Y %H:%M:%S %z')
        row[email_date_column] = parsed_date.strftime('%Y-%m-%d %H:%M:%S')
      except:
        # If parsing fails, use the raw date or current time
        try:
          # Try alternative parsing
          from email.utils import parsedate_tz, mktime_tz
          date_tuple = parsedate_tz(email_date)
          if date_tuple:
            timestamp = mktime_tz(date_tuple)
            parsed_date = datetime.fromtimestamp(timestamp)
            row[email_date_column] = parsed_date.strftime('%Y-%m-%d %H:%M:%S')
          else:
            row[email_date_column] = email_date  # Use raw date as fallback
        except:
          row[email_date_column] = email_date  # Use raw date as final fallback
    else:
      # If no date available, use current timestamp as fallback
      row[email_date_column] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    print(f"Extracted fields: {row}")
    results.append(row)
  return results
