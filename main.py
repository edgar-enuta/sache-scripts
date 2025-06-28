import imap


if __name__ == "__main__":
  imap.init()
  imap.authenticate()
  emails = imap.get_unread_emails(limit=5)
  for email in emails:
    imap.print_email(email)
  imap.logout()