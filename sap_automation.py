import win32com.client
import subprocess
import sys
import time
from datetime import datetime
import os
from time import sleep
from PIL import ImageGrab
import re
import shutil


# ---------------------------------------
# Helper to load credentials from TXT
# ---------------------------------------
def load_creds(creds_path):
    creds = type("Creds", (), {})()
    if not os.path.exists(creds_path):
        print(f"❌ Creds file not found: {creds_path}")
        sys.exit(1)

    with open(creds_path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            # allow entries with or without spaces around '='
            if "=" in line:
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()

                # Try to eval python literal (so r"..." or "..." would work). If fails, keep as raw string.
                try:
                    # protect against unquoted windows paths being interpreted as names
                    # only eval if it looks like a Python literal
                    if (value.startswith('"') or value.startswith("'") or value.startswith("r'") or value.startswith(
                            'r"')) \
                            or value.isdigit() or (value.upper() in ["TRUE", "FALSE"]):
                        evaluated = eval(value)
                        setattr(creds, key, evaluated)
                    else:
                        # keep raw string (most likely a path without quotes)
                        setattr(creds, key, value)
                except Exception:
                    setattr(creds, key, value)
    return creds


# ---------------------------------------
# Last-serials helpers
# ---------------------------------------
def read_last_serials(path):
    """Read transaction numbers from last_transaction.txt file"""
    serials = {}
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if not line or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k = k.strip().upper()
                    try:
                        serials[k] = int(v.strip())
                    except:
                        serials[k] = 0
        else:
            print(f"⚠️ Last transaction file not found: {path}")
    except Exception as e:
        print(f"⚠️ Could not read last-serials file '{path}': {e}")
    return serials


def write_last_serials(path, serials):
    """Write updated transaction numbers back to last_transaction.txt file"""
    try:
        dirn = os.path.dirname(path)
        if dirn:
            os.makedirs(dirn, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            for k in sorted(serials.keys()):
                f.write(f"{k}={serials[k]}\n")
    except Exception as e:
        print(f"❌ Failed to write last-serials file '{path}': {e}")


# ---------------------------------------
# SAP connection helpers
# ---------------------------------------
def launch_sap_logon(path):
    print("🚀 Launching SAP Logon...")
    try:
        subprocess.Popen(path)
        time.sleep(5)
    except Exception as e:
        print(f"❌ Could not start SAP Logon: {e}")
        sys.exit(1)


def connect_to_sap():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        return sap_gui_auto.GetScriptingEngine
    except Exception as e:
        print(f"❌ Unable to attach to SAP GUI scripting engine: {e}")
        sys.exit(1)


def open_connection(app, connection_name):
    try:
        print(f"🔗 Opening SAP connection: {connection_name}")
        conn = None
        # iterate connections safely (some COM collections are 0..Count-1)
        try:
            count = app.Connections.Count
        except Exception:
            count = 0
        for i in range(count):
            try:
                name = app.Children(i).Name
            except Exception:
                name = ""
            if connection_name.lower() in name.lower():
                conn = app.Children(i)
                break
        if not conn:
            conn = app.OpenConnection(connection_name, True)
        return conn
    except Exception as e:
        print(f"❌ Failed to open SAP connection: {e}")
        sys.exit(1)


def login_to_sap(conn, creds):
    try:
        session = conn.Children(0)
        print("🔐 Logging in...")
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = creds.CLIENT
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = creds.USER
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = creds.PASSWORD
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = creds.LANGUAGE
        session.findById("wnd[0]").sendVKey(0)
        sleep(3)
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            pass

        print("⏳ Waiting for SAP Easy Access...")
        max_wait = getattr(creds, "MAX_WAIT", 30)
        for _ in range(int(max_wait)):
            try:
                if "Easy Access" in session.findById("wnd[0]").Text:
                    print("✅ SAP Easy Access detected.")
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
            if not session.Busy:
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
                return popup
        except:
            pass
        sleep(1)
    return None


# ---------------------------------------
# Logging & Popups
# ---------------------------------------
def log_message(log_file, file_name, msg_type, message):
    try:
        os.makedirs(os.path.dirname(log_file) or ".", exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} | {file_name} | {msg_type} | {message}\n")
    except Exception as e:
        print(f"⚠️ Could not write to log file '{log_file}': {e}")


def handle_all_popups(session, current_file, log_file, screenshot_dir):
    messages = []
    popup_index = 0
    while True:
        sleep(1)
        try:
            popup = session.findById("wnd[1]")
            full_text = []
            try:
                for child in popup.findById("usr").Children:
                    if hasattr(child, "Text") and child.Text.strip():
                        full_text.append(child.Text.strip())
            except:
                pass
            if not full_text:
                try:
                    full_text = [popup.Text.strip()]
                except:
                    full_text = ["<No text captured>"]
            message_text = " | ".join(full_text)
            popup_index += 1
            messages.append(message_text)
            print(f"💬 Popup {popup_index}: {message_text}")
            log_message(log_file, current_file, "Information", message_text)
            try:
                os.makedirs(screenshot_dir, exist_ok=True)
                base = os.path.splitext(os.path.basename(current_file))[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                popup_path = os.path.join(screenshot_dir, f"{base}_popup_{popup_index}_{timestamp}.png")
                ImageGrab.grab().save(popup_path)
                print(f"📸 Screenshot saved: {popup_path}")
            except Exception:
                pass
            try:
                popup.sendVKey(0)
            except:
                try:
                    popup.findById("tbar[0]/btn[0]").press()
                except:
                    pass
        except:
            break
    return messages


# ---------------------------------------
# Utility to find currency in filename
# ---------------------------------------
def find_currency_in_filename(filename):
    """Detect currency code (AZN, EUR, USD, CHF, GBP, RUB) in filename"""
    name_upper = filename.upper()
    currencies = ["AZN", "EUR", "USD", "CHF", "GBP", "RUB"]
    for cur in currencies:
        if cur in name_upper:
            return cur
    return None


# ---------------------------------------
# Modified rewrite_28C_line function
# ---------------------------------------
def rewrite_28C_line(file_path, new_transaction_no):
    """
    Rewrite the :28C: line on the 4th line with new transaction number.
    Preserves the suffix (e.g., /4) that was already there.

    Example:
        Original: :28C:1077/4
        Modified: :28C:12352/4

    Args:
        file_path: Path to the file
        new_transaction_no: New transaction number to use
    """
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
    except Exception as e:
        raise RuntimeError(f"Could not read file {file_path}: {e}")

    if len(lines) < 4:
        raise RuntimeError(f"File has less than 4 lines: {file_path}")

    # The 4th line (index 3)
    line_28c = lines[3].strip()

    if not line_28c.startswith(":28C:"):
        raise RuntimeError(f"4th line does not start with :28C: in file {file_path}")

    # Extract the suffix (normally /4 or similar)
    match = re.match(r":28C:\d+(.*)", line_28c)
    if not match:
        raise RuntimeError(f"Could not parse :28C: line format in file {file_path}")

    suffix = match.group(1)  # This will be something like "/4"

    # Rewrite the 4th line with new transaction number
    lines[3] = f":28C:{new_transaction_no}{suffix}\n"

    try:
        with open(file_path, "w", encoding="utf-8", errors="ignore") as f:
            f.writelines(lines)
    except Exception as e:
        raise RuntimeError(f"Could not write updated file {file_path}: {e}")


# ---------------------------------------
# File move helper
# ---------------------------------------
def move_to_processed(src_file, processed_dir):
    try:
        os.makedirs(processed_dir, exist_ok=True)
        dest = os.path.join(processed_dir, os.path.basename(src_file))
        shutil.move(src_file, dest)
        print(f"📁 Moved to processed folder → {dest}")
    except Exception as e:
        print(f"⚠️ Could not move file to processed folder: {e}")


# ---------------------------------------
# Main SAP Upload function
# ---------------------------------------
def run_sap_upload(creds_path):
    """
    Returns:
        True  -> SAP upload completed successfully for at least one file
        False -> Any failure occurred
    """

    try:
        creds = load_creds(creds_path)

        # Resolve paths
        last_serials_path = getattr(creds, "LAST_SERIALS_PATH", None)
        if not last_serials_path:
            print("❌ LAST_SERIALS_PATH not found in credentials file")
            return False

        processed_dir_from_creds = getattr(creds, "PROCESSED_DIR", None)
        if not processed_dir_from_creds:
            upload_dir_guess = getattr(creds, "UPLOAD_DIR", None) or getattr(creds, "INVOICE_DIR", None)
            processed_dir_from_creds = os.path.join(upload_dir_guess, "processed") if upload_dir_guess else "processed"

        # Load last transaction numbers
        last_serials = read_last_serials(last_serials_path)
        print(f"📊 Loaded transaction numbers: {last_serials}")

        # Start SAP
        launch_sap_logon(creds.SAP_LOGON_PATH)
        sleep(4)
        app = connect_to_sap()
        conn = open_connection(app, creds.SAP_CONNECTION_NAME)
        session = login_to_sap(conn, creds)

        # Open FF.5
        print("➡️ Opening transaction FF.5...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "FF.5"
        session.findById("wnd[0]").sendVKey(0)
        wait_until_ready(session, 10)

        # Select RFEBKA00
        print("📋 Selecting RFEBKA00 program...")
        found = False
        for row in range(3, 20):
            try:
                label = session.findById(f"wnd[0]/usr/lbl[62,{row}]")
                if "RFEBKA00" in label.Text:
                    label.setFocus()
                    session.findById("wnd[0]").sendVKey(2)
                    found = True
                    break
            except:
                pass

        if not found:
            print("❌ RFEBKA00 not found — fix SAP layout.")
            return False   # RETURN FAILURE

        wait_until_ready(session)

        # Variant popup
        print("➡️ Opening 'Get Variant' popup...")
        try:
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
        except:
            print("❌ Could not press variant button")
            return False

        popup = wait_for_popup(session, 10)
        if not popup:
            print("❌ Variant popup did not appear.")
            return False

        try:
            popup.findById("usr/txtENAME-LOW").text = creds.VARIANT_NAME.upper()
            popup.findById("tbar[0]/btn[8]").press()
        except:
            print("❌ Could not apply variant.")
            return False

        sleep(1)

        # Gather files
        upload_dir = getattr(creds, "UPLOAD_DIR", getattr(creds, "INVOICE_DIR", None))
        if not upload_dir or not os.path.isdir(upload_dir):
            print(f"❌ Upload directory not found: {upload_dir}")
            return False

        print(f"\n📂 Scanning upload folder: {upload_dir}")

        files_to_upload = []
        for f in os.listdir(upload_dir):
            full_path = os.path.join(upload_dir, f)
            if f.lower().endswith(".txt") and os.path.getsize(full_path) > 1024:
                files_to_upload.append(full_path)

        if not files_to_upload:
            print("⚠️ No valid files found.")
            return False

        # Track if ANY file was successfully uploaded
        success_count = 0

        # Upload loop
        for file_path in files_to_upload:
            filename = os.path.basename(file_path)
            print(f"\n📤 Preparing to upload: {filename}")

            currency = find_currency_in_filename(filename)
            if not currency:
                print(f"⚠️ Currency not found in {filename}")
                continue

            current_serial = last_serials.get(currency, 0)
            next_serial = current_serial + 1

            try:
                rewrite_28C_line(file_path, next_serial)
            except Exception as e:
                print(f"❌ Failed to rewrite :28C: {e}")
                continue

            # Upload file
            try:
                field = session.findById("wnd[0]/usr/ctxtAUSZFILE")
                field.setFocus()
                field.text = file_path
                session.findById("wnd[0]").sendVKey(8)
                sleep(2)

                # Handle popups
                handle_all_popups(session, filename, getattr(creds, "LOG_FILE", "upload.log"),
                                  getattr(creds, "SCREENSHOT_DIR", "."))

                # Return to main page
                try:
                    for _ in range(3):
                        session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        sleep(1)
                except:
                    pass

                # Mark success
                last_serials[currency] = next_serial
                write_last_serials(last_serials_path, last_serials)

                move_to_processed(file_path, processed_dir_from_creds)

                print("✅ Upload successful.")
                success_count += 1

            except Exception as e:
                print(f"❌ Upload failed for {filename}: {e}")
                continue

        # FINAL OUTCOME
        if success_count > 0:
            print(f"🎉 {success_count} file(s) uploaded successfully.")
            return True
        else:
            print("❌ No files were successfully uploaded.")
            return False

    except Exception as main_e:
        print(f"❌ Fatal error in run_sap_upload: {main_e}")
        return False
