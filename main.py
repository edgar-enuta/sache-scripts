import imap
from extract_fields import extract_fields_from_emails
from excel_export import export_to_excel

if __name__ == "__main__":
  imap.init()
  imap.authenticate()
  # TODO: rollback to previous state in case of errors
  emails = imap.get_unread_emails()
  extracted = extract_fields_from_emails(emails)
  export_to_excel(extracted)
  # TODO: flag emails once confirmed processed
  imap.logout()