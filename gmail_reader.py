import os
import base64
import shutil
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GmailDownloader:
    def __init__(
        self,
        credentials_path,
        token_path,
        download_folder,
        query='has:attachment',
        scopes=['https://www.googleapis.com/auth/gmail.readonly']
    ):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.download_folder = download_folder
        self.query = query
        self.scopes = scopes
        self.service = None

    # =========================
    #  AUTHENTICATION
    # =========================
    def authenticate(self):
        """Authenticates and returns Gmail API service."""
        creds = None

        # Load existing token if available
        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)

        # If no token or invalid → refresh or run OAuth
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print("🔐 Running OAuth consent flow...")
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path, self.scopes
                )
                creds = flow.run_local_server(port=8080, access_type='offline', prompt='consent')

                with open(self.token_path, 'w') as token:
                    token.write(creds.to_json())

                print("✅ Token saved:", self.token_path)

        self.service = build('gmail', 'v1', credentials=creds)
        return self.service

    # =========================
    #  UTILITY
    # =========================
    def clean_folder(self):
        """Clears old attachments."""
        if os.path.exists(self.download_folder):
            shutil.rmtree(self.download_folder)
        os.makedirs(self.download_folder, exist_ok=True)

    # =========================
    #  FETCH LATEST EMAIL
    # =========================
    def get_latest_email(self):
        """Returns latest email that matches search query."""
        results = self.service.users().messages().list(
            userId='me',
            q=self.query,
            maxResults=1,
            labelIds=['INBOX']
        ).execute()

        messages = results.get('messages', [])
        return messages[0] if messages else None

    # =========================
    #  DOWNLOAD ATTACHMENTS
    # =========================
    def download_attachments(self, msg):
        """Downloads all attachments from given message."""
        msg_data = self.service.users().messages().get(
            userId='me',
            id=msg['id']
        ).execute()

        payload = msg_data.get('payload', {})
        parts = payload.get('parts', [])

        count = 0

        for part in parts:
            filename = part.get('filename')
            body = part.get('body', {})

            if filename and 'attachmentId' in body:
                att_id = body['attachmentId']

                attachment = self.service.users().messages().attachments().get(
                    userId='me',
                    messageId=msg['id'],
                    id=att_id
                ).execute()

                data = base64.urlsafe_b64decode(attachment['data'])
                filepath = os.path.join(self.download_folder, filename)

                with open(filepath, 'wb') as f:
                    f.write(data)

                print(f"📎 Saved: {filepath}")
                count += 1

        if count == 0:
            print("⚠️ No attachments found.")

    # =========================
    #  MAIN FUNCTION CALL
    # =========================
    def download_latest(self):
        """Main function to authenticate, find latest email, and download attachments."""
        print("🔐 Authenticating Gmail...")
        self.authenticate()

        print("🔍 Searching for latest email...")
        msg = self.get_latest_email()

        if not msg:
            print("⚠️ No matching emails found.")
            return False

        print("🧹 Cleaning folder...")
        self.clean_folder()

        print("⬇️ Downloading attachments...")
        self.download_attachments(msg)

        print("🎯 Done!")
        return True
