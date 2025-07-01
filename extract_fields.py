import yaml
import re

def extract_fields_from_emails(emails, config_path="field_config.yaml"):
  """
  For each email, search for each field's pattern (if it has excel_column) in the email text.
  Save any found value in a dict under the key from excel_column.
  Returns a list of dicts (one per email).
  """
  with open(config_path, "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

  results = []
  for email in emails:
    row = {}
    text = f"{email.get('subject', '')}\n{email.get('body', '')}"
    for field, props in config.items():
      pattern = props.get('pattern')
      excel_column = props.get('excel_column')
      if pattern and excel_column:
        match = re.search(pattern, text)
        if match:
          row[excel_column] = match.group(1) if match.groups() else match.group(0)
    print(f"Extracted fields: {row}")
    results.append(row)
  return results
