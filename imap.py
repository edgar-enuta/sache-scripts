import imaplib
import getpass
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
import yaml
import re

IMAP_SERVER = None
IMAP_PORT = None
mail = None

def init():
  global IMAP_SERVER, IMAP_PORT
  load_dotenv()
  IMAP_SERVER = os.getenv("IMAP_SERVER")
  IMAP_PORT = int(os.getenv("IMAP_PORT", 993))

def authenticate():
  global mail
  print("Login to your email account at mail.mondoparts.ro")
  user = os.getenv("EMAIL") or input("Email: ")
  password = os.getenv("EMAIL_PASSWORD") or getpass.getpass("Password: ")

  try:
    # Connect to the server
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(user, password)
    print("Login successful!\n")
  except imaplib.IMAP4.error as e:
      print(f"IMAP error: {e}")
  except Exception as e:
      print(f"An error occurred: {e}")

def fetch_email_by_id(email_id):
  """
  Fetch the raw email data for a given email ID.
  Returns the raw bytes of the email, or None if fetch fails.
  Does not mark the email as read.
  """
  global mail
  status, msg_data = mail.fetch(email_id, "(BODY.PEEK[])")
  if status != "OK" or not msg_data or not msg_data[0]:
    print(f"Failed to fetch email {email_id.decode()}")
    return None
  return msg_data[0][1]

def parse_email_from_bytes(raw_bytes):
  """
  Parse an email.message.Message object from raw bytes and extract sender, subject, and body.
  Returns a dict: {'from': ..., 'subject': ..., 'body': ...}
  """
  msg = email.message_from_bytes(raw_bytes)
  # Decode subject
  subject, encoding = decode_header(msg["Subject"])[0]
  if isinstance(subject, bytes):
    subject = subject.decode(encoding or "utf-8", errors="ignore")
  from_ = msg.get("From")
  date = msg.get("Date")
  # Extract body (prefer plain text)
  body = ""
  parts = msg.walk() if msg.is_multipart() else [msg]
  for part in parts:
    content_type = part.get_content_type()
    content_disposition = str(part.get("Content-Disposition"))
    if content_type == "text/plain" and "attachment" not in content_disposition:
      charset = part.get_content_charset() or "utf-8"
      payload = part.get_payload(decode=True)
      try:
        body = payload.decode(charset, errors="ignore")
      except Exception:
        body = payload.decode("utf-8", errors="ignore")
      break
  return {"from": from_, "subject": subject, "body": body, "date": date}

def parse_title_from_bytes(raw_bytes):
    """
    Extract and decode the subject/title from raw email bytes.
    Returns the decoded subject as a string.
    """
    msg = email.message_from_bytes(raw_bytes)
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8", errors="ignore")
    return subject

def print_email(email_obj):
  """
  Print the details of an email object (dict with keys: from, subject, body, date).
  """
  print(f"From: {email_obj['from']}")
  print(f"Subject: {email_obj['subject']}")
  print(f"Date: {email_obj['date']}")
  print(f"Body: {email_obj['body'][:200]}")  # Print first 200 chars of body
  print("-" * 40)

def filter_by_title(title, config_path="field_config.yaml"):
    """
    Returns True if the title matches the email_title pattern in field_config.yaml, else False.
    """
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    pattern = config.get("email_title", {}).get("pattern")
    if not pattern:
        return False
    return re.search(pattern, title) is not None

def get_unread_emails(limit=None):
  """
  Retrieve unread and unflagged emails from the inbox.
  If limit is given, only fetch up to that many emails.
  Returns a list of email objects.
  """
  global mail
  if mail is None:
    print("Not authenticated. Please login first.")
    authenticate()
    return []

  emails = []
  try:
    mail.select("inbox")
    # Only fetch emails that are UNSEEN and UNFLAGGED
    status, messages = mail.search(None, "UNSEEN", "UNFLAGGED")
    if status != "OK":
      print("Failed to search for unread emails.")
      return []

    email_ids = messages[0].split()
    print(f"Unread and unflagged emails: {len(email_ids)}\n")

    if limit is not None:
      email_ids = email_ids[:limit]

    for num in email_ids:
      raw_bytes = fetch_email_by_id(num)
      if not raw_bytes:
        continue
      title = parse_title_from_bytes(raw_bytes)
      print(f"Filtering email: {title}")
      if filter_by_title(title):
        print(f"Email matches filter")
        email_obj = parse_email_from_bytes(raw_bytes)
        # Add the email id (UID) to the email object for later flagging
        email_obj['uid'] = num
        emails.append(email_obj)
    return emails
  except imaplib.IMAP4.error as e:
    print(f"IMAP error: {e}")
    return []
  except Exception as e:
    print(f"An error occurred: {e}")
    return []

def logout():
  print("Logging out...")
  mail.logout()