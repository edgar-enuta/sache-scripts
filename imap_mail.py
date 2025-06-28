import imaplib
import getpass
import email
from email.header import decode_header
import os
from dotenv import load_dotenv

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
  """
  global mail
  status, msg_data = mail.fetch(email_id, "(RFC822)")
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

def get_unread_emails(limit=None):
  """
  Retrieve and print unread emails from the inbox.
  If limit is given, only fetch up to that many emails.
  """
  global mail
  if mail is None:
    print("Not authenticated. Please login first.")
    authenticate()
    return

  try:
    mail.select("inbox")
    status, messages = mail.search(None, "UNSEEN")
    if status != "OK":
      print("Failed to search for unread emails.")
      return

    email_ids = messages[0].split()
    print(f"Unread emails: {len(email_ids)}\n")

    if limit is not None:
      email_ids = email_ids[:limit]

    for num in email_ids:
      raw_bytes = fetch_email_by_id(num)
      if not raw_bytes:
        continue
      email_obj = parse_email_from_bytes(raw_bytes)
      print(f"From: {email_obj['from']}")
      print(f"Subject: {email_obj['subject']}")
      print(f"Body: {email_obj['body'][:200]}")  # Print first 200 chars of body
      print("-" * 40)
  except imaplib.IMAP4.error as e:
    print(f"IMAP error: {e}")
  except Exception as e:
    print(f"An error occurred: {e}")

def logout():
  print("Logging out...")
  mail.logout()