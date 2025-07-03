import imap
from extract_fields import extract_fields_from_emails
import excel
import os
from dotenv import load_dotenv

load_dotenv()
EXCEL_PATH = os.getenv("EXCEL_PATH")

def flag_emails_as_processed(emails):
    # Example: mark as flagged (\Flagged) or as seen (\Seen)
    for email in emails:
        # You need to store the email UID or ID in the email dict for this to work
        # This is a placeholder; actual implementation depends on your imap.py
        pass


if __name__ == "__main__":
    imap.init()
    imap.authenticate()
    emails = []
    extracted = []
    backup_path = None
    try:
        emails = imap.get_unread_emails()
        extracted = extract_fields_from_emails(emails)
        backup_path = excel.backup_excel(EXCEL_PATH)
        excel.export_to_excel(extracted)
        flag_emails_as_processed(emails)
    except Exception as e:
        print(f"Error occurred: {e}. Rolling back changes.")
        excel.restore_excel(backup_path, EXCEL_PATH)
    finally:
        imap.logout()