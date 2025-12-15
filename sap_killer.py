import os

def close_sap():
    print("🛑 Closing SAP…")
    os.system("taskkill /IM saplogon.exe /F")
    # os.system("taskkill /IM sapshcut.exe /F")
    # os.system("taskkill /IM sapgui.exe /F")
    print("✅ SAP closed.")
