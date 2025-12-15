import win32com.client
import subprocess
import sys
import time
from time import sleep

# ---------------------------------------
# LOAD CREDS FROM TXT FILE
# ---------------------------------------
def load_sap_creds(creds_path):
    creds = type("Creds", (), {})()   # create empty object

    try:
        with open(creds_path, "r") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue

                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()

                    # evaluate Python literal (handles quotes & r"")
                    try:
                        value = eval(value)
                    except:
                        pass

                    setattr(creds, key, value)

        return creds

    except Exception as e:
        print(f"❌ Error loading creds file: {e}")
        sys.exit(1)


# ---------------------------------------
# Launch SAP Logon
# ---------------------------------------
def launch_sap_logon(path):
    print("🚀 Launching SAP Logon...")
    try:
        subprocess.Popen(path)
        time.sleep(5)
    except Exception as e:
        print(f"❌ Could not start SAP Logon: {e}")
        sys.exit(1)


# ---------------------------------------
# Connect to SAP GUI Scripting Engine
# ---------------------------------------
def connect_to_sap():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        return sap_gui_auto.GetScriptingEngine
    except Exception as e:
        print(f"❌ Unable to attach to SAP GUI scripting engine: {e}")
        sys.exit(1)


# ---------------------------------------
# Open SAP connection
# ---------------------------------------
def open_connection(app, connection_name):
    try:
        print(f"🔗 Opening SAP connection: {connection_name}")

        # Check already open connections
        for i in range(app.Connections.Count):
            if connection_name.lower() in app.Children(i).Name.lower():
                return app.Children(i)

        # If not open → open new one
        return app.OpenConnection(connection_name, True)

    except Exception as e:
        print(f"❌ Failed to open SAP connection: {e}")
        sys.exit(1)


# ---------------------------------------
# Login
# ---------------------------------------
def login_to_sap(conn, creds):
    try:
        session = conn.Children(0)
        print("🔐 Logging into SAP...")

        # Enter login details
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = creds.CLIENT
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = creds.USER
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = creds.PASSWORD
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = creds.LANGUAGE

        session.findById("wnd[0]").sendVKey(0)
        sleep(2)

        # Accept "last login" popup if exists
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            pass

        print("⏳ Waiting for SAP Easy Access...")

        # --- WAIT FOR SAP EASY ACCESS SCREEN ---
        for _ in range(creds.MAX_WAIT):
            try:
                win_title = session.findById("wnd[0]").Text
                if "Easy Access" in win_title:
                    break
            except:
                pass
            sleep(1)

        # --- EXTRA STABILITY FIX ---
        wait_until_ready(session)
        sleep(1)

        # Close any hidden popups
        try:
            if session.Children.Count > 1:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        # Make sure OKCode field is focused and EMPTY
        try:
            ok = session.findById("wnd[0]/tbar[0]/okcd")
            ok.text = ""
            ok.setFocus()
        except:
            pass

        sleep(0.5)
        wait_until_ready(session)

        print("✅ SAP Login Successful & Stable!")
        return session, creds

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

# ---------------------------------------
# PUBLIC FUNCTION → Call This
# ---------------------------------------
def get_sap_session(creds_path):
    creds = load_sap_creds(creds_path)

    launch_sap_logon(creds.SAP_LOGON_PATH)
    app = connect_to_sap()
    conn = open_connection(app, creds.SAP_CONNECTION_NAME)
    session = login_to_sap(conn, creds)

    return session, creds

