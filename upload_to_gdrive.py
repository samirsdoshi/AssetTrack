#!/usr/bin/env python3
"""
Upload backup files to Google Drive
Uploads asset*.sql files and Asset.xlsx to specified Google Drive folder
"""

import os
import glob
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pickle

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Google Drive folder ID
FOLDER_ID = '1143-kZ1KCLy8yQsL8Dkms1mowIndLfRu'


def get_credentials():
    """Get valid user credentials from storage or run auth flow"""
    creds = None
    
    # The file token.pickle stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("ERROR: credentials.json not found!")
                print("\nTo use Google Drive API, you need to:")
                print("1. Go to https://console.cloud.google.com/")
                print("2. Create a new project or select existing one")
                print("3. Enable Google Drive API")
                print("4. Create OAuth 2.0 credentials (Desktop app)")
                print("5. Download credentials.json to this directory")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def find_existing_file(service, file_name, folder_id):
    """Check if a file with the same name exists in the folder"""
    try:
        query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
        results = service.files().list(
            q=query,
            fields='files(id, name)',
            pageSize=1
        ).execute()
        
        files = results.get('files', [])
        if files:
            return files[0]['id']
        return None
    except Exception as e:
        print(f"  Warning: Could not check for existing file: {e}")
        return None


def upload_file(service, file_path, folder_id):
    """Upload a file to Google Drive, overwriting if it exists"""
    file_name = os.path.basename(file_path)
    
    print(f"Uploading {file_name}...")
    
    # Check if file already exists
    existing_file_id = find_existing_file(service, file_name, folder_id)
    
    media = MediaFileUpload(file_path, resumable=True)
    
    try:
        if existing_file_id:
            # Update existing file
            print(f"  File exists - updating (ID: {existing_file_id})...")
            file = service.files().update(
                fileId=existing_file_id,
                media_body=media,
                fields='id, name, webViewLink'
            ).execute()
            print(f"✓ Updated: {file.get('name')}")
        else:
            # Create new file
            print(f"  Creating new file...")
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink'
            ).execute()
            print(f"✓ Uploaded: {file.get('name')}")
        
        print(f"  File ID: {file.get('id')}")
        print(f"  Link: {file.get('webViewLink')}")
        return file
    except Exception as e:
        print(f"✗ Error uploading {file_name}: {e}")
        return None


def main():
    """Main function to upload backup files"""
    print("=" * 60)
    print("Google Drive Backup Upload")
    print("=" * 60)
    
    # Get credentials
    creds = get_credentials()
    if not creds:
        return
    
    # Build the Drive service
    service = build('drive', 'v3', credentials=creds)
    
    # Find files to upload
    sql_files = glob.glob('backup/asset*.sql')
    excel_file = 'Asset.xlsx'
    
    files_to_upload = []
    
    # Add SQL backup files
    if sql_files:
        files_to_upload.extend(sql_files)
        print(f"\nFound {len(sql_files)} SQL backup file(s)")
    else:
        print("\nNo asset*.sql files found")
    
    # Add Excel file
    if os.path.exists(excel_file):
        files_to_upload.append(excel_file)
        print(f"Found {excel_file}")
    else:
        print(f"{excel_file} not found")
    
    if not files_to_upload:
        print("\nNo files to upload!")
        return
    
    print(f"\nUploading {len(files_to_upload)} file(s) to Google Drive folder...")
    print(f"Folder ID: {FOLDER_ID}\n")
    
    # Upload each file
    uploaded_count = 0
    for file_path in files_to_upload:
        result = upload_file(service, file_path, FOLDER_ID)
        if result:
            uploaded_count += 1
        print()
    
    print("=" * 60)
    print(f"Upload complete: {uploaded_count}/{len(files_to_upload)} files uploaded successfully")
    print("=" * 60)


if __name__ == '__main__':
    main()
