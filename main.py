import argparse
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


def process_emails():
    """
    Main processing function that handles the email-to-Excel workflow.
    Returns a tuple: (success: bool, created_file_path: str or None, processed_emails_count: int)
    """
    emails = []
    extracted_data = []
    created_excel_path = None
    
    try:
        # Step 1: Fetch unread emails
        print("Fetching unread emails...")
        emails = imap.get_unread_emails()
        
        if not emails:
            print("No unread emails found matching the filter criteria.")
            return True, None, 0
        
        print(f"Found {len(emails)} emails to process")
        
        # Step 2: Extract data from emails
        print("Extracting data from emails...")
        extracted_data = extract_fields_from_emails(emails)
        
        if not extracted_data or all(not row for row in extracted_data):
            print("No valid data could be extracted from the emails.")
            # Still flag the emails as processed to avoid reprocessing
            flag_emails_as_processed(emails)
            return True, None, len(emails)
        
        # Step 3: Create Excel file
        print("Creating Excel file...")
        created_excel_path = excel.export_to_excel(extracted_data)
        
        if not created_excel_path:
            raise Exception("Failed to create Excel file")
        
        # Step 4: Mark emails as processed only after successful Excel creation
        print("Marking emails as processed...")
        flag_emails_as_processed(emails)
        
        return True, created_excel_path, len(emails)
        
    except Exception as e:
        print(f"Error during processing: {e}")
        
        # Cleanup: Remove the Excel file if it was created
        if created_excel_path and os.path.exists(created_excel_path):
            try:
                os.remove(created_excel_path)
                print(f"Cleaned up incomplete Excel file: {created_excel_path}")
            except Exception as cleanup_error:
                print(f"Warning: Could not clean up Excel file: {cleanup_error}")
        
        # Cleanup: Unflag any emails that might have been flagged
        if emails:
            print("Unflagging emails due to processing failure...")
            unflag_emails(emails)
        
        return False, None, len(emails)


def run_email():
    """Run the email module. Returns True on success."""
    print("=== Email Processing Started ===")
    imap.init()
    if not imap.authenticate():
        print("Login failed. Halting program.")
        return False

    try:
        success, created_file, processed_count = process_emails()

        if success:
            if created_file:
                print(f"\nProcessed {processed_count} emails and created Excel file:")
                print(f"   {created_file}")
            else:
                print(f"\nProcessed {processed_count} emails (no data to export)")
        else:
            print(f"\nFailed: Processing failed for {processed_count} emails")
        return success

    except Exception as e:
        print(f"\nCritical error in email module: {e}")
        return False
    finally:
        try:
            imap.logout()
        except Exception as e:
            print(f"Warning: Error during logout: {e}")
        print("=== Email Processing Completed ===")


def run_scrape():
    """Run the scrape module. Returns True on success."""
    from scrape import process_scrape

    print("=== Scrape Processing Started ===")
    try:
        success, created_file, count = process_scrape()

        if success:
            if created_file:
                print(f"\nScraped {count} items and created Excel file:")
                print(f"   {created_file}")
            else:
                print(f"\nScrape completed ({count} items, no data to export)")
        else:
            print(f"\nFailed: Scrape processing failed")
        return success

    except Exception as e:
        print(f"\nCritical error in scrape module: {e}")
        return False
    finally:
        print("=== Scrape Processing Completed ===")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Email-to-Excel and web scraping automation tool.")
    parser.add_argument(
        "command",
        nargs="?",
        default="all",
        choices=["all", "email", "scrape"],
        help="Which module to run (default: all)",
    )
    args = parser.parse_args()

    any_failed = False

    if args.command in ("all", "email"):
        if not run_email():
            any_failed = True

    if args.command in ("all", "scrape"):
        if not run_scrape():
            any_failed = True

    if any_failed:
        exit(1)