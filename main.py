from fetch_unread_emails import fetch_unread_emails
from user_sign_in import get_access_token


if __name__ == "__main__":
  # user sign in
  try:
    access_token = get_access_token()
    # fetch_unread_emails(access_token)
  except Exception as e:
    print(f"Error: {e}")