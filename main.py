import imap_mail as imap_mail


if __name__ == "__main__":
  imap_mail.init()
  imap_mail.authenticate()
  imap_mail.get_unread_emails(5)
  imap_mail.logout()