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


if __name__ == "__main__":
    print("=== Email to Excel Processing Started ===")
    
    # Initialize IMAP connection
    imap.init()
    if not imap.authenticate():
        print("Login failed. Halting program.")
        exit(1)
    
    try:
        success, created_file, processed_count = process_emails()
        
        if success:
            if created_file:
                print(f"\n✅ SUCCESS: Processed {processed_count} emails and created Excel file:")
                print(f"   📄 {created_file}")
            else:
                print(f"\n✅ SUCCESS: Processed {processed_count} emails (no data to export)")
        else:
            print(f"\n❌ FAILED: Processing failed for {processed_count} emails")
            exit(1)
            
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        exit(1)
    finally:
        try:
            imap.logout()
            print("\n=== Email to Excel Processing Completed ===")
        except Exception as e:
            print(f"Warning: Error during logout: {e}")