import os
import datetime
import win32com.client
import shutil
import zipfile
import subprocess

# === CONFIGURATION ===
TARGET_FOLDER_NAME = "PnL"
SAVE_DIR = r"C:\Your\Target\Directory"  # <-- CHANGE THIS TO YOUR DESIRED DIRECTORY
FILENAMES = [
    "RACE_Clean_P&L_CEP and BHW including FX Sensitivities",
    "RACE_Clean_PnL_CGME with EUR",
    "RACE_Clean_PnL_CGML_CGME",
    "RACE_Clean_PnL_CGML_CGME_TB"
]

# Ensure save dir exists and is clean
os.makedirs(SAVE_DIR, exist_ok=True)

# === GET TARGET DATE ===
today = datetime.datetime.now()
if today.weekday() == 0:  # Monday
    target_date = today - datetime.timedelta(days=2)
else:
    target_date = today

# === CONNECT TO OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
pnl_folder = None
for folder in inbox.Folders:
    if folder.Name == TARGET_FOLDER_NAME:
        pnl_folder = folder
        break
if not pnl_folder:
    raise Exception(f"Folder '{TARGET_FOLDER_NAME}' not found in Inbox.")

# === HELPER: Save and Unzip ===
def save_and_unpack(attachment, base_name):
    zip_path = os.path.join(SAVE_DIR, base_name + ".zip")
    xlsx_path = os.path.join(SAVE_DIR, base_name + ".xlsx")

    # Remove old files
    for path in [zip_path, xlsx_path]:
        if os.path.exists(path):
            os.remove(path)

    # Save ZIP
    attachment.SaveAsFile(zip_path)

    # Unpack using SecureZIP CLI (assumed available in PATH)
    try:
        subprocess.run(["securezip", "-extract", zip_path, "-directory", SAVE_DIR, "-overwrite"], check=True)
    except Exception as e:
        print(f"Failed to unpack {zip_path}: {e}")

# === SEARCH EMAILS AND PROCESS ===
for base_name in FILENAMES:
    matching_mails = [
        mail for mail in pnl_folder.Items
        if mail.Subject.strip() == base_name and mail.ReceivedTime.date() == target_date.date()
    ]
    if not matching_mails:
        print(f"No matching mail found for '{base_name}' on {target_date.date()}")
        continue

    # Use latest if multiple
    mail = sorted(matching_mails, key=lambda m: m.ReceivedTime, reverse=True)[0]

    for attachment in mail.Attachments:
        if attachment.FileName.lower().endswith(".zip") and base_name in attachment.FileName:
            print(f"Processing attachment: {attachment.FileName}")
            save_and_unpack(attachment, base_name)

