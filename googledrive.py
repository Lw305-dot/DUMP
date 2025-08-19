import os
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ================= CONFIG =================
SERVICE_ACCOUNT_FILE = '/workspaces/DUMP/service_account.json'
SCOPES = ['https://www.googleapis.com/auth/drive.file']

MAIN_FOLDER_ID = "1TfaHhqs2qiCKBtsm8Z20mT-XcY7hehg_"
QR_FOLDER_ID   = "1r0awgcwbUbvSIGsQ_bP_sHIUJyBwV"

# Example file to upload
file_path = '/workspaces/DUMP/Employee_Reports13/sample.xlsx'
# =========================================

def authenticate_drive():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('drive', 'v3', credentials=creds)
    return service

def check_permissions(folder_id):
    """Check if service account has Editor access to the folder."""
    service = authenticate_drive()
    try:
        permissions = service.permissions().list(fileId=folder_id).execute()
    except Exception as e:
        print(f"❌ Cannot access folder: {e}")
        return False
    
    has_access = False
    print("Folder Permissions:")
    for p in permissions.get('permissions', []):
        email = p.get('emailAddress')
        role = p.get('role')
        print(f" - {email} ({role})")
        if email and role.lower() in ['owner', 'writer', 'editor']:
            has_access = True
    return has_access

def list_files(folder_id):
    """List files visible to the service account in the folder."""
    service = authenticate_drive()
    results = service.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if not files:
        print("No files found in this folder.")
    else:
        print("Files in folder:")
        for f in files:
            print(f" - {f['name']} (ID: {f['id']})")

def upload_to_drive(file_path, folder_id):
    """Upload file only if access is valid."""
    if not check_permissions(folder_id):
        print(f"❌ Service account does not have editor access to folder {folder_id}.")
        return None

    service = authenticate_drive()
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"✅ Uploaded {file_path} → Drive ID: {file.get('id')}")
    return file.get('id')

# ====== RUN TEST ======
print("Checking MAIN_FOLDER_ID permissions...")
list_files(MAIN_FOLDER_ID)

print("\nAttempting to upload file...")
upload_to_drive(file_path, MAIN_FOLDER_ID)

###################3
# from google.oauth2 import service_account
# from googleapiclient.discovery import build
# from pathlib import Path

# SERVICE_ACCOUNT_FILE = '/workspaces/DUMP/service_account.json'
# SCOPES = ['https://www.googleapis.com/auth/drive.file']
# MAIN_FOLDER_ID = "1TfaHhqs2qiCKBtsm8Z20mT-XcY7hehg_"

# def authenticate_drive():
#     creds = service_account.Credentials.from_service_account_file(
#         SERVICE_ACCOUNT_FILE, scopes=SCOPES
#     )
#     service = build('drive', 'v3', credentials=creds)
#     return service

# def list_folder_permissions(folder_id):
#     service = authenticate_drive()
#     permissions = service.permissions().list(fileId=folder_id).execute()
#     print("Folder Permissions:")
#     for p in permissions.get('permissions', []):
#         print(f" - {p.get('emailAddress')} ({p.get('role')})")

# def list_files_in_folder(folder_id):
#     service = authenticate_drive()
#     results = service.files().list(
#         q=f"'{folder_id}' in parents",
#         fields="files(id, name)"
#     ).execute()
#     files = results.get('files', [])
#     if not files:
#         print("No files found in this folder.")
#     else:
#         print("Files in folder:")
#         for f in files:
#             print(f" - {f['name']} (ID: {f['id']})")

# # Run checks
# list_folder_permissions(MAIN_FOLDER_ID)
# list_files_in_folder(MAIN_FOLDER_ID)
