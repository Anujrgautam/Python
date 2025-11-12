import win32com.client
import subprocess
import sys
import time
from datetime import datetime
import os
import pyautogui
import pygetwindow as gw
from time import sleep
from PIL import ImageGrab

# === CONFIGURATION ===
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_CONNECTION_NAME = "QR-5 Quality ECC6 Core BRT"
CLIENT = "200"
USER = "Poojnai"
PASSWORD = "Pratiba@90"
LANGUAGE = "EN"
VARIANT_NAME = "IRISKOVA"
INVOICE_DIR = r"D:\Upload Bank Statement\Invoices"
MAX_WAIT = 60
LOG_FILE = r"D:\Upload Bank Statement\Upload_Log.txt"
UPLOAD_DIR = r"D:\Upload Bank Statement\Invoices"
SCREENSHOT_DIR = r"D:\Upload Bank Statement\Screenshots"

# ======================


def launch_sap_logon():
    print("🚀 Launching SAP Logon...")
    try:
        subprocess.Popen(SAP_LOGON_PATH)
        time.sleep(5)
    except Exception as e:
        print(f"❌ Could not start SAP Logon: {e}")
        sys.exit(1)


def connect_to_sap():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        return application
    except Exception as e:
        print(f"❌ Unable to attach to SAP GUI scripting engine: {e}")
        sys.exit(1)


def open_connection(application, connection_name):
    try:
        print(f"🔗 Opening SAP connection: {connection_name}")
        conn = None
        for i in range(application.Connections.Count):
            if connection_name.lower() in application.Children(i).Name.lower():
                conn = application.Children(i)
                break
        if not conn:
            conn = application.OpenConnection(connection_name, True)
        return conn
    except Exception as e:
        print(f"❌ Failed to open SAP connection: {e}")
        sys.exit(1)


def login_to_sap(conn):
    try:
        session = conn.Children(0)
        print("🔐 Logging in...")
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = CLIENT
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = PASSWORD
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = LANGUAGE
        session.findById("wnd[0]").sendVKey(0)

        sleep(3)
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            pass

        print("⏳ Waiting for SAP Easy Access screen...")
        for _ in range(MAX_WAIT):
            try:
                if "Easy Access" in session.findById("wnd[0]").Text:
                    print(f"✅ SAP Easy Access detected.")
                    break
            except:
                pass
            sleep(1)
        return session
    except Exception as e:
        print(f"❌ Login failed: {e}")
        sys.exit(1)


def wait_until_ready(session, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        try:
            if not session.Busy and not session.Info.IsLowSpeedConnection:
                return True
        except:
            pass
        sleep(0.5)
    return False


def wait_for_popup(session, timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        try:
            if session.Children.Count > 1:
                popup = session.findById("wnd[1]")
                if popup:
                    return popup
        except:
            pass
        sleep(1)
    return None


# === MAIN EXECUTION ===
launch_sap_logon()
sleep(5)
application = connect_to_sap()
connection = open_connection(application, SAP_CONNECTION_NAME)
session = login_to_sap(connection)

# === Run FF.5 ===
print("➡️ Opening transaction FF.5...")
session.findById("wnd[0]/tbar[0]/okcd").text = "FF.5"
session.findById("wnd[0]").sendVKey(0)
wait_until_ready(session, 10)

print("📋 Selecting RFEBKA00 program...")
for row in range(3, 8):
    try:
        label = session.findById(f"wnd[0]/usr/lbl[62,{row}]")
        if "RFEBKA00" in label.Text:
            label.setFocus()
            session.findById("wnd[0]").sendVKey(2)
            print("✅ Selected RFEBKA00 successfully")
            break
    except:
        continue

wait_until_ready(session, 8)
session.findById("wnd[0]").maximize()
sleep(2)

# === Variant Selection ===
print("➡️ Opening 'Get Variant' popup...")
session.findById("wnd[0]/tbar[1]/btn[17]").press()
popup = wait_for_popup(session, 10)
if not popup:
    print("❌ Could not open variant popup.")
    sys.exit(1)

try:
    print("✏️ Filling 'Created by' field...")
    popup.findById("usr/txtENAME-LOW").text = VARIANT_NAME.upper()
    popup.findById("tbar[0]/btn[8]").press()
    print(f"✅ Entered Created by: {VARIANT_NAME.upper()}")
except Exception as e:
    print(f"❌ Error filling variant: {e}")
    sys.exit(1)

def log_message(file_name, msg_type, message):
    """Append messages to log file with timestamps."""
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} | File: {file_name} | Type: {msg_type} | Message: {message}\n")

def handle_all_popups(session, current_file, log_file_path):

    messages = []
    popup_index = 0

    def log_message(file, msg_type, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file_path, "a", encoding="utf-8") as log:
            log.write(f"{timestamp} | File: {os.path.basename(file)} | Type: {msg_type} | Message: {message}\n")

    while True:
        time.sleep(1)
        try:
            popup = session.findById("wnd[1]")
            full_text = []

            # Loop through all elements inside the popup window
            try:
                for child in popup.findById("usr").Children:
                    try:
                        # Some are label/text elements
                        if hasattr(child, "Text") and child.Text.strip():
                            full_text.append(child.Text.strip())
                    except:
                        continue
            except:
                pass

            # Fallback to popup title or text if nothing found
            if not full_text:
                try:
                    full_text = [popup.Text.strip()]
                except:
                    full_text = ["<No text captured>"]

            message_text = " | ".join(full_text)
            popup_index += 1
            messages.append(message_text)

            print(f"💬 Popup {popup_index}: {message_text}")
            log_message(current_file, "Information", message_text)

            # Optional: save a screenshot
            try:
                os.makedirs(SCREENSHOT_DIR, exist_ok=True)
                base = os.path.splitext(os.path.basename(current_file))[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                popup_path = os.path.join(SCREENSHOT_DIR, f"{base}_popup_{popup_index}_{timestamp}.png")
                img = ImageGrab.grab()
                img.save(popup_path)
                print(f"📸 Screenshot saved: {popup_path}")
            except Exception as e:
                print(f"⚠️ Could not capture popup screenshot: {e}")

            # Close popup (Enter or green check)
            try:
                popup.sendVKey(0)
            except:
                try:
                    popup.findById("tbar[0]/btn[0]").press()
                except:
                    pass

        except:
            break  # No more popups

    if messages:
        print(f"✅ Handled {len(messages)} popup(s).")
    else:
        print("ℹ️ No popups detected.")

    return messages



def capture_screenshot_and_exit(session, current_file):
    """
    Takes screenshot of the current SAP screen and exits twice to return to Easy Access.
    """
    try:
        # Ensure folder exists
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)

        # Build filename
        base = os.path.splitext(os.path.basename(current_file))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(SCREENSHOT_DIR, f"{base}_{timestamp}.png")

        # Capture screenshot
        img = ImageGrab.grab()
        img.save(screenshot_path)
        print(f"📸 Screenshot saved: {screenshot_path}")

        # Try to close/exit the window
        for i in range(2):
            try:
                session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Exit button
                print(f"↩️ Clicked Exit ({i+1}/2)")
                time.sleep(2)
            except:
                pass

    except Exception as e:
        print(f"⚠️ Failed to capture screenshot or exit: {e}")

def upload_statements(session):
    """Main function to upload all valid .txt files >1KB from the upload folder."""
    print(f"\n📂 Uploading from: {UPLOAD_DIR}")

    files_to_upload = [
        os.path.join(UPLOAD_DIR, f)
        for f in os.listdir(UPLOAD_DIR)
        if f.lower().endswith(".txt") and os.path.getsize(os.path.join(UPLOAD_DIR, f)) > 1024
    ]

    if not files_to_upload:
        print("⚠️ No valid .txt files > 1KB found.")
        return

    print(f"📄 Found {len(files_to_upload)} file(s) to upload.\n")

    for file_path in files_to_upload:
        try:
            print(f"📤 Uploading: {file_path}")
            field = session.findById("wnd[0]/usr/ctxtAUSZFILE")
            field.setFocus()
            field.text = file_path
            print("✅ File path entered.")

            # Press F8 to execute
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(3)

            # Handle possible info/error popups
            if handle_all_popups(session, file_path, LOG_FILE):
                print("⚠️ Popups handled, moving to next file.\n")
                continue

            # If no popup, consider success
            log_message(os.path.basename(file_path), "Success", "File uploaded successfully")
            print("✅ Upload successful.\n")

        except Exception as e:
            error_msg = f"Error processing {file_path}: {e}"
            print(f"❌ {error_msg}")
            log_message(os.path.basename(file_path), "Error", error_msg)
            continue


# === Run Upload ===
upload_statements(session)
