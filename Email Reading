from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os
import base64
import shutil

# === CONFIGURATION ===
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_PATH = r"D:\Upload Bank Statement\credentials.json"
TOKEN_PATH = r"D:\Upload Bank Statement\token.json"
DOWNLOAD_FOLDER = r"D:\Upload Bank Statement\Attachments"

# Gmail search filter — edit as needed
query = 'subject:"Statements" has:attachment'


def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("🔐 Running OAuth consent flow...")
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_PATH, SCOPES
            )
            creds = flow.run_local_server(port=8080, access_type='offline', prompt='consent')
            with open(TOKEN_PATH, 'w') as token:
                token.write(creds.to_json())
            print("✅ Token saved at:", TOKEN_PATH)

    return build('gmail', 'v1', credentials=creds)


def clean_download_folder(folder):
    """Deletes all files in the download folder before saving new ones."""
    if os.path.exists(folder):
        print("🧹 Clearing old attachments...")
        shutil.rmtree(folder)
    os.makedirs(folder)
    print("✅ Folder cleaned:", folder)


def get_latest_email(service, query):
    """Fetch the most recent email matching the query."""
    results = service.users().messages().list(
        userId='me', q=query, maxResults=1, labelIds=['INBOX']
    ).execute()
    messages = results.get('messages', [])
    return messages[0] if messages else None


def download_attachments(service, msg, folder):
    """Download all attachments from the latest email."""
    msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
    payload = msg_data.get('payload', {})
    headers = payload.get('headers', [])
    subject = next((h['value'] for h in headers if h['name'] == 'Subject'), "(No Subject)")

    print(f"\n📧 Latest Email: {subject}")

    parts = payload.get('parts', [])
    for part in parts:
        filename = part.get('filename')
        body = part.get('body', {})
        if filename and 'attachmentId' in body:
            att_id = body['attachmentId']
            att = service.users().messages().attachments().get(
                userId='me', messageId=msg['id'], id=att_id
            ).execute()
            data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
            filepath = os.path.join(folder, filename)
            with open(filepath, 'wb') as f:
                f.write(data)
            print(f"📎 Saved: {filepath}")

    print("\n🎯 Done! All attachments from the latest email downloaded.")


if __name__ == "__main__":
    service = get_gmail_service()

    print("🔍 Searching for the most recent matching email...")
    msg = get_latest_email(service, query)

    if not msg:
        print("⚠️ No matching recent email found.")
    else:
        clean_download_folder(DOWNLOAD_FOLDER)
        download_attachments(service, msg, DOWNLOAD_FOLDER)
