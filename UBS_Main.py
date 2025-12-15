from gmail_reader import GmailDownloader
from gmail_sender import GmailSender
from sap_automation import run_sap_upload
from sap_killer import close_sap

USER_EMAIL = "anuj.gautam.ext@holcim.com"
DEV_EMAIL = "anuj.gautam.ext@holcim.com"

sender = GmailSender(
    credentials_path=r"D:\Upload Bank Statement\credentials_gmail.json",
    token_path=r"D:\Upload Bank Statement\token_sender.json"
)

def send_success():
    body = (
        "Dear Team,\n\n"
        "The Upload Bank Statement bot has completed successfully.\n"
        "Please refer to the bot log file for the status of each transaction.\n"
        "All uploaded files are available in the processed folder.\n\n"
        "Regards,\n"
        "UBS Bot"
    )

    sender.send_email(USER_EMAIL, "Upload Bank Statement Bot – SUCCESS", body)
    sender.send_email(DEV_EMAIL, "Upload Bank Statement Bot – SUCCESS", body)


def send_failure(reason):
    body = (
        "Dear Team,\n\n"
        "The Upload Bank Statement bot has FAILED.\n\n"
        "Failure Reason:\n"
        f"{reason}\n\n"
        "Recommended Actions:\n"
        "- Verify SAP layout and RFEBKA00 availability\n"
        "- Check variant configuration\n"
        "- Review the bot log file for details\n\n"
        "Regards,\n"
        "UBS Bot"
    )

    sender.send_email(USER_EMAIL, "Upload Bank Statement Bot – FAILED", body)
    sender.send_email(DEV_EMAIL, "Upload Bank Statement Bot – FAILED", body)


try:
    print("📥 Step 1: Downloading invoices...")

    reader = GmailDownloader(
        credentials_path=r"D:\Upload Bank Statement\credentials_gmail.json",
        token_path=r"D:\Upload Bank Statement\token_reader.json",
        download_folder=r"D:\Upload Bank Statement\Attachments",
        query='subject:"Daily MT940 Statements" has:attachment'
    )

    print("📤 Step 2: Uploading to SAP...")

    sap_result = run_sap_upload(r"D:\Upload Bank Statement\Creds.txt")

    if sap_result:
        send_success()
    else:
        send_failure(
            "SAP upload failed. RFEBKA00 was not found, "
            "variant did not load, or no valid files were processed."
        )

except Exception as e:
    print("❌ ERROR:", str(e))
    send_failure(str(e))

finally:
    close_sap()
