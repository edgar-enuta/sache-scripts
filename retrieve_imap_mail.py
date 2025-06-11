import imaplib
import getpass
import email
from email.header import decode_header

# IMAP server details
IMAP_SERVER = "mail.mondoparts.ro"
IMAP_PORT = 993

def authenticate():
  print("Login to your email account at mail.mondoparts.ro")
  user = input("Email: ")
  password = getpass.getpass("Password: ")

  try:
    # Connect to the server
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(user, password)
    print("Login successful!\n")
    print(mail.list())

    #     # Select the mailbox you want to use
    #     mail.select("inbox")

    #     # Search for all emails
    #     status, messages = mail.search(None, "ALL")
    #     if status != "OK":
    #         print("No messages found!")
    #         return

    #     email_ids = messages[0].split()
    #     print(f"Total emails: {len(email_ids)}\n")

    #     # Fetch the latest 5 emails
    #     for num in email_ids[-5:]:
    #         status, msg_data = mail.fetch(num, "(RFC822)")
    #         if status != "OK":
    #             print(f"Failed to fetch email {num.decode()}")
    #             continue
    #         msg = email.message_from_bytes(msg_data[0][1])
    #         subject, encoding = decode_header(msg["Subject"])[0]
    #         if isinstance(subject, bytes):
    #             subject = subject.decode(encoding or "utf-8", errors="ignore")
    #         from_ = msg.get("From")
    #         print(f"From: {from_}")
    #         print(f"Subject: {subject}")
    #         print("-" * 40)

    mail.logout()
  except imaplib.IMAP4.error as e:
      print(f"IMAP error: {e}")
  except Exception as e:
      print(f"An error occurred: {e}")
