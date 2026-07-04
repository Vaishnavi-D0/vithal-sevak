"""
drive_helper.py
Uploads captured/selected member photos to a shared Google Drive folder,
and lets the app browse/download photos already sitting in that folder.
Uses the same service account credentials as the Sheets integration.
"""

import io
import os

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

CREDENTIALS_FILE = "credentials.json"
DRIVE_FOLDER_ID = "1X6OtZxVXiUVk7LgGtzVPFvvoTi1dtTRk"

SCOPES = ["https://www.googleapis.com/auth/drive"]

_service = None


def _get_service():
    global _service
    if _service is None:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        _service = build("drive", "v3", credentials=creds)
    return _service


def upload_photo(local_path, folder_id=DRIVE_FOLDER_ID):
    """Uploads a local photo file to the Drive folder.
    Returns (file_id, web_view_link)."""
    service = _get_service()
    filename = os.path.basename(local_path)
    file_metadata = {"name": filename, "parents": [folder_id]}
    media = MediaFileUpload(local_path, resumable=False)
    try:
        file = service.files().create(
            body=file_metadata, media_body=media, fields="id, webViewLink"
        ).execute()
    except Exception as e:
        raise RuntimeError(f"Drive upload failed: {e}")
    return file.get("id"), file.get("webViewLink")


def list_photos(folder_id=DRIVE_FOLDER_ID):
    """Returns a list of {'id', 'name'} dicts for image files in the folder."""
    service = _get_service()
    try:
        query = (
            f"'{folder_id}' in parents and trashed = false "
            "and mimeType contains 'image/'"
        )
        results = service.files().list(
            q=query, fields="files(id, name)", pageSize=1000,
            orderBy="name"
        ).execute()
    except Exception as e:
        raise RuntimeError(f"Could not list Drive photos: {e}")
    return results.get("files", [])


def download_photo(file_id, dest_path):
    """Downloads a Drive file to dest_path. Returns dest_path."""
    service = _get_service()
    try:
        request = service.files().get_media(fileId=file_id)
        with io.FileIO(dest_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
    except Exception as e:
        raise RuntimeError(f"Drive download failed: {e}")
    return dest_path
