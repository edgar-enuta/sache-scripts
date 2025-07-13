import imap
from extract_fields import extract_fields_from_emails
import excel
import os
from dotenv import load_dotenv

load_dotenv()
EXCEL_PATH = os.getenv("EXCEL_PATH")

def flag_emails_as_processed(emails):
    # Mark each email as flagged (\Flagged) using imap
    for email in emails:
        uid = email.get('uid')
        if uid:
            try:
                imap.mail.store(uid, '+FLAGS', '\\Flagged')
            except Exception as e:
                print(f"Failed to flag email {uid}: {e}")
        else:
            subj = email.get('subject', '[No Subject]')
            sender = email.get('from', '[No From]')
            print(f"No email id found for flagging. Subject: {subj}, From: {sender}")

def unflag_emails(emails):
    # Remove the \Flagged flag from each email using imap
    for email in emails:
        uid = email.get('uid')
        if uid:
            try:
                imap.mail.store(uid, '-FLAGS', '\\Flagged')
            except Exception as e:
                print(f"Failed to unflag email {uid}: {e}")
        else:
            subj = email.get('subject', '[No Subject]')
            sender = email.get('from', '[No From]')
            print(f"No email id found for unflagging. Subject: {subj}, From: {sender}")


if __name__ == "__main__":
    imap.init()
    if not imap.authenticate():
        print("Login failed. Halting program.")
        exit(1)
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
        unflag_emails(emails)
    finally:
        imap.logout()