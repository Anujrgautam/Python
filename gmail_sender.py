import os
import base64
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GmailSender:
    def __init__(self, credentials_path, token_path):
        self.credentials_path = credentials_path
        self.token_path = token_path

        # REQUIRED SCOPES
        self.scopes = [
            "https://www.googleapis.com/auth/gmail.send",
        ]

        self.service = None

    def authenticate(self):
        creds = None

        # Load existing token
        if os.path.exists(self.token_path):
            try:
                creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
            except Exception:
                creds = None  # Invalid token file → regenerate

        # Generate or refresh token
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                print("🔄 Refreshing Gmail token...")
                creds.refresh(Request())
            else:
                print("🔐 Running OAuth consent flow for Gmail Sender...")
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path,
                    self.scopes
                )
                creds = flow.run_local_server(
                    port=8082,
                    access_type='offline',
                    prompt='consent'
                )

                with open(self.token_path, "w") as token:
                    token.write(creds.to_json())
                print("✅ New sender token saved:", self.token_path)

        # Build Gmail API service
        self.service = build("gmail", "v1", credentials=creds)
        return self.service

    def send_email(self, to, subject, message_text):
        service = self.authenticate()

        message = MIMEText(message_text)
        message["to"] = to
        message["subject"] = subject

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

        try:
            result = (
                service.users()
                .messages()
                .send(userId="me", body={"raw": raw})
                .execute()
            )
            print(f"📨 Email sent to {to}! Message ID: {result.get('id')}")
            return True

        except Exception as e:
            print(f"❌ Failed to send email to {to}: {e}")
            return False
