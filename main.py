
import subprocess
import sys
import os
import re
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Google Drive configuration
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credential.json'
TOKEN_FILE = 'token.json'
DRIVE_FOLDER_ID = 'Folder ID'

def check_adb_availability():
    try:
        subprocess.run(["adb", "--version"], check=True, capture_output=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False



# def check_device_connected():
#     result = subprocess.run(["adb", "devices"], capture_output=True, text=True)
#     return "device" in result.stdout and not "unauthorized" in result.stdout



def check_device_connected():
    """Check device connection and display name if connected"""
    try:
        # Check connection status
        devices = subprocess.run(["adb", "devices"], capture_output=True, text=True, timeout=5)
        if "device" not in devices.stdout or "unauthorized" in devices.stdout:
            print("No device connected")
            return False
            
        # Get device name if connected
        name = subprocess.check_output(["adb", "shell", "getprop", "ro.product.model"], 
                                      text=True, timeout=5).strip()
        print(f"Device connected: {name}")
        return True
        
    except Exception:
        print("No device connected")
        return False





# def get_device_name():
#     try:
#         name = subprocess.check_output(["adb", "shell", "getprop", "ro.product.model"], text=True).strip()
#         return re.sub(r'[^a-zA-Z0-9_-]', '_', name)
#     except subprocess.CalledProcessError:
#         return "unknown_device"

def google_drive_auth():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(service, file_path):
    file_name = os.path.basename(file_path)
    media = MediaFileUpload(file_path, mimetype='application/xml')
    
    try:
        response = service.files().list(
            q=f"name='{file_name}' and '{DRIVE_FOLDER_ID}' in parents",
            spaces='drive',
            fields='files(id, name)'
        ).execute()
        
        if len(response.get('files', [])) > 0:
            file_id = response['files'][0]['id']
            service.files().update(
                fileId=file_id,
                media_body=media
            ).execute()
        else:
            file_metadata = {
                'name': file_name,
                'parents': [DRIVE_FOLDER_ID]
            }
            service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
        return True
    except Exception as e:
        print(f" Google Drive Error: {str(e)}")
        return False



def find_and_pull_xml():
    base_path = "/sdcard/Android/data/your directory/"
    downloads = Path.home() / "Downloads/path"
    
    try:
        # Corrected find command to locate all XML files
        find_cmd = f'find "{base_path}" -name "*.xml"'
        result = subprocess.run(["adb", "shell", find_cmd], check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError:
        print(" No XML files found or search error")
        return

    files = result.stdout.strip().split('\n')
    if not files or files[0] == '':
        print(" No XML files found")
        return

    # Initialize Google Drive service once
    try:
        drive_service = google_drive_auth()
    except Exception as e:
        print(f"Google Drive authentication failed: {str(e)}")
        drive_service = None

    for remote_path in files:
        try:
            original_filename = os.path.basename(remote_path)
            base_name, extension = os.path.splitext(original_filename)
            extension = extension.lstrip('.')  # Remove leading dot

            # Handle duplicate filenames
            index = 0
            while True:
                if index == 0:
                    filename = f"{base_name}.{extension}"
                else:
                    filename = f"{base_name}({index}).{extension}"
                
                dest_path = downloads / filename
                if not dest_path.exists():
                    break
                index += 1

            # Pull file from device
            subprocess.run(["adb", "pull", remote_path, str(dest_path)], check=True)
            print(f" Saved locally: {filename}")

            # Upload to Google Drive
            if drive_service:
                if upload_to_drive(drive_service, str(dest_path)):
                    print(f" Uploaded to Drive: {filename}")
                else:
                    print(f" Failed to upload: {filename}")

        except subprocess.CalledProcessError as e:
            print(f" Failed to copy {remote_path}: {e.stderr}")
        except Exception as e:
            print(f"Error processing {remote_path}: {str(e)}")




# def find_and_pull_xml():
#     base_path = "/sdcard/Android/data/Your directory/"
#     downloads = Path.home() / "Downloads/path"
#     device_name = get_device_name()
    
#     try:
#         find_cmd = f'find "{base_path}" -name last-saved.xml'
#         find_cmd = f'find "{base_path}" -name"*.xml"'
#         result = subprocess.run(["adb", "shell", find_cmd], check=True, capture_output=True, text=True)
#     except subprocess.CalledProcessError:
#         print(" No last_saved.xml files found or search error")
#         return

#     files = result.stdout.strip().split('\n')
#     if not files or files[0] == '':
#         print(" No last_saved.xml files found")
#         return

#     # Initialize Google Drive service once
#     try:
#         drive_service = google_drive_auth()
#     except Exception as e:
#         print(f"Google Drive authentication failed: {str(e)}")
#         drive_service = None

#     for remote_path in files:
#         try:
#             dir_path = os.path.dirname(remote_path)
#             project_folder = os.path.basename(dir_path)
            
#             clean_project = re.sub(r'[^a-zA-Z0-9_-]', '_', project_folder)
#             base_filename = f"{device_name}_{clean_project}_last-saved"
#             extension = "xml"

#             index = 0
#             while True:
#                 if index == 0:
#                     filename = f"{base_filename}.{extension}"
#                 else:
#                     filename = f"{base_filename}({index}).{extension}"
                
#                 dest_path = downloads / filename
#                 if not dest_path.exists():
#                     break
#                 index += 1

#             # Pull file from device
#             subprocess.run(["adb", "pull", remote_path, str(dest_path)], check=True)
#             print(f" Saved locally: {filename}")

            
#             if drive_service:
#                 if upload_to_drive(drive_service, str(dest_path)):
#                     print(f" Uploaded to Drive: {filename}")
#                 else:
#                     print(f"  Failed to upload: {filename}")

#         except subprocess.CalledProcessError as e:
#             print(f" Failed to copy {remote_path}: {e.stderr}")
#         except Exception as e:
#             print(f"Error processing {remote_path}: {str(e)}")

def main():
    if not check_adb_availability():
        print(" ADB not found. Install Android SDK Platform-Tools and add to PATH")
        sys.exit(1)

    if not check_device_connected():
        print(" No authorized device connected")
        sys.exit(1)

    print(" Searching for .xml files...")
    find_and_pull_xml()
    print("\nOperation completed. Check your Downloads folder and Google Drive.")

if __name__ == "__main__":
    main()



