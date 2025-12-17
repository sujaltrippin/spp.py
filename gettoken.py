from __future__ import print_function
import os.path
import base64
import email
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying scopes, delete the token.json file.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    """Download all attachments from your Gmail inbox."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('oauth.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    results = service.users().messages().list(userId='me', q="has:attachment newer_than:7d").execute()
    messages = results.get('messages', [])

    if not messages:
        print("No messages found.")
        return

    for msg in messages:
        msg_id = msg['id']
        msg = service.users().messages().get(userId='me', id=msg_id).execute()
        for part in msg['payload'].get('parts', []):
            if part['filename']:
                attachment_id = part['body']['attachmentId']
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=msg_id, id=attachment_id
                ).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'])
                path = part['filename']
                with open(path, 'wb') as f:
                    f.write(file_data)
                print(f"Downloaded: {path}")

if __name__ == '__main__':
    main()
