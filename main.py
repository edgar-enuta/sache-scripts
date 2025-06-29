import imap
from extract_fields import extract_fields_from_emails


if __name__ == "__main__":
  imap.init()
  imap.authenticate()
  emails = imap.get_unread_emails()
  extracted = extract_fields_from_emails(emails)
  print(extracted)
  imap.logout()