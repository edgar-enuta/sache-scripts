import requests
import msal
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Access credentials from environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

# Microsoft Graph API endpoints
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.ReadWrite"]
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0/me/messages"

def get_access_token():
    """Authenticate and get an access token using delegated authentication flow."""
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
    )

    # Get the authorization URL
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow. Check your credentials.")

    print("Please go to the following URL and enter the code to authenticate:")
    print(flow["verification_uri"])
    print("Code:", flow["user_code"])

    # Wait for the user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Could not acquire access token. Check your credentials.")

def fetch_unread_emails(access_token):
    """Fetch unread emails and mark them as read."""
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"$filter": "isRead eq false"}
    response = requests.get(GRAPH_API_ENDPOINT, headers=headers, params=params)

    if response.status_code == 200:
        emails = response.json().get("value", [])
        for email in emails:
            print(f"Subject: {email['subject']}")
            print(f"From: {email['from']['emailAddress']['address']}")
            print(f"Received: {email['receivedDateTime']}")
            print("-" * 50)

            # Mark the email as read
            mark_as_read_url = f"{GRAPH_API_ENDPOINT}/{email['id']}"
            mark_as_read_response = requests.patch(
                mark_as_read_url,
                headers=headers,
                json={"isRead": True},
            )
            if mark_as_read_response.status_code == 200:
                print(f"Email with subject '{email['subject']}' marked as read.")
            else:
                print(f"Failed to mark email as read: {mark_as_read_response.text}")
    else:
        print(f"Failed to fetch emails: {response.text}")